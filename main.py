from __future__ import annotations

import asyncio
import html
import io
import ipaddress
import json
import os
import re
import secrets
import socket
from datetime import datetime, timezone
from typing import Iterable
from urllib.parse import urlparse

import httpx
from fastapi import BackgroundTasks, Depends, FastAPI, Form, HTTPException, Request
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from openpyxl import Workbook
from sqlalchemy import DateTime, Float, ForeignKey, Integer, String, Text, create_engine, select
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column, relationship, sessionmaker
from starlette.middleware.sessions import SessionMiddleware

APP_TITLE = "Парсер компаний — Lite"
TWOGIS_URL = "https://catalog.api.2gis.com/3.0/items"


def env(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()


def env_bool(name: str, default: bool = False) -> bool:
    value = env(name, "true" if default else "false").lower()
    return value in {"1", "true", "yes", "on", "да"}


def normalize_database_url(url: str) -> str:
    if not url:
        return "sqlite:///./leadparser.sqlite3"
    if url.startswith("postgres://"):
        return "postgresql+psycopg://" + url[len("postgres://") :]
    if url.startswith("postgresql://") and "+psycopg" not in url:
        return "postgresql+psycopg://" + url[len("postgresql://") :]
    return url


DATABASE_URL = normalize_database_url(env("DATABASE_URL"))
ADMIN_PASSWORD = env("ADMIN_PASSWORD", "change-me")
SESSION_SECRET = env("SESSION_SECRET") or secrets.token_urlsafe(32)
TWOGIS_API_KEY = env("TWOGIS_API_KEY")
DATA_LICENSE_ACKNOWLEDGED = env_bool("DATA_LICENSE_ACKNOWLEDGED", False)
TELEGRAM_BOT_TOKEN = env("TELEGRAM_BOT_TOKEN")
TELEGRAM_ADMIN_CHAT_ID = env("TELEGRAM_ADMIN_CHAT_ID")


class Base(DeclarativeBase):
    pass


class Job(Base):
    __tablename__ = "jobs"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    city: Mapped[str] = mapped_column(String(200))
    categories: Mapped[str] = mapped_column(Text)
    mode: Mapped[str] = mapped_column(String(30), default="no_site")
    max_results: Mapped[int] = mapped_column(Integer, default=50)
    status: Mapped[str] = mapped_column(String(30), default="pending")
    found_count: Mapped[int] = mapped_column(Integer, default=0)
    saved_count: Mapped[int] = mapped_column(Integer, default=0)
    error: Mapped[str | None] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime(timezone=True), default=lambda: datetime.now(timezone.utc))
    finished_at: Mapped[datetime | None] = mapped_column(DateTime(timezone=True), nullable=True)
    leads: Mapped[list["Lead"]] = relationship(back_populates="job", cascade="all, delete-orphan")


class Lead(Base):
    __tablename__ = "leads"

    id: Mapped[int] = mapped_column(Integer, primary_key=True)
    job_id: Mapped[int] = mapped_column(ForeignKey("jobs.id", ondelete="CASCADE"), index=True)
    source_id: Mapped[str] = mapped_column(String(200), default="")
    name: Mapped[str] = mapped_column(String(500))
    category: Mapped[str] = mapped_column(Text, default="")
    address: Mapped[str] = mapped_column(Text, default="")
    phones: Mapped[str] = mapped_column(Text, default="")
    emails: Mapped[str] = mapped_column(Text, default="")
    messengers: Mapped[str] = mapped_column(Text, default="")
    website: Mapped[str | None] = mapped_column(Text, nullable=True)
    rating: Mapped[float | None] = mapped_column(Float, nullable=True)
    audit_score: Mapped[int | None] = mapped_column(Integer, nullable=True)
    audit_notes: Mapped[str | None] = mapped_column(Text, nullable=True)
    draft_text: Mapped[str | None] = mapped_column(Text, nullable=True)
    job: Mapped[Job] = relationship(back_populates="leads")


connect_args = {"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}
engine = create_engine(DATABASE_URL, pool_pre_ping=True, connect_args=connect_args)
SessionLocal = sessionmaker(bind=engine, autoflush=False, expire_on_commit=False)

app = FastAPI(title=APP_TITLE)
app.add_middleware(
    SessionMiddleware,
    secret_key=SESSION_SECRET,
    same_site="lax",
    https_only=False,
    max_age=60 * 60 * 12,
)


@app.on_event("startup")
def startup() -> None:
    Base.metadata.create_all(bind=engine)


def get_db() -> Iterable[Session]:
    with SessionLocal() as db:
        yield db


def logged_in(request: Request) -> bool:
    return request.session.get("auth") is True


def require_login(request: Request) -> None:
    if not logged_in(request):
        raise HTTPException(status_code=303, headers={"Location": "/login"})


def esc(value: object) -> str:
    return html.escape("" if value is None else str(value))


def layout(title: str, body: str, request: Request | None = None, refresh: int | None = None) -> str:
    refresh_tag = f'<meta http-equiv="refresh" content="{refresh}">' if refresh else ""
    logout = (
        '<form method="post" action="/logout"><button class="secondary">Выйти</button></form>'
        if request and logged_in(request)
        else ""
    )
    return f"""<!doctype html>
<html lang="ru"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
{refresh_tag}<title>{esc(title)}</title>
<style>
:root {{ color-scheme: dark; --bg:#0e1117; --card:#171b22; --line:#2b313b; --text:#f4f6f8; --muted:#aab2bf; --accent:#7c5cff; --ok:#37c978; --bad:#ff5d6c; }}
* {{ box-sizing:border-box }} body {{ margin:0; background:var(--bg); color:var(--text); font:16px/1.45 system-ui,-apple-system,Segoe UI,sans-serif }}
a {{ color:#a997ff }} .wrap {{ max-width:1180px; margin:0 auto; padding:24px }}
header {{ display:flex; align-items:center; justify-content:space-between; gap:16px; margin-bottom:24px }}
.card {{ background:var(--card); border:1px solid var(--line); border-radius:16px; padding:20px; margin-bottom:18px }}
.grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(230px,1fr)); gap:14px }}
label {{ display:block; color:var(--muted); margin-bottom:6px }} input,textarea,select {{ width:100%; background:#0c0f14; color:var(--text); border:1px solid var(--line); border-radius:10px; padding:11px }}
textarea {{ min-height:110px; resize:vertical }} button,.button {{ display:inline-block; border:0; border-radius:10px; background:var(--accent); color:white; padding:11px 16px; cursor:pointer; text-decoration:none; font-weight:650 }}
.secondary {{ background:#29303a }} .danger {{ background:#7a2833 }} table {{ width:100%; border-collapse:collapse; font-size:14px }} th,td {{ text-align:left; vertical-align:top; padding:10px; border-bottom:1px solid var(--line) }} th {{ color:var(--muted) }}
.badge {{ display:inline-block; padding:4px 8px; border-radius:999px; background:#29303a; font-size:12px }} .ok {{ color:var(--ok) }} .bad {{ color:var(--bad) }} .muted {{ color:var(--muted) }}
pre {{ white-space:pre-wrap; word-break:break-word; background:#0c0f14; padding:12px; border-radius:10px }}
@media(max-width:700px) {{ .wrap {{ padding:14px }} .table-wrap {{ overflow:auto }} table {{ min-width:850px }} }}
</style></head><body><div class="wrap"><header><div><h1 style="margin:0">{esc(APP_TITLE)}</h1><div class="muted">2ГИС → лиды → Excel</div></div>{logout}</header>{body}</div></body></html>"""


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request, error: str = "") -> str:
    if logged_in(request):
        return RedirectResponse("/", status_code=303)
    error_html = f'<p class="bad">{esc(error)}</p>' if error else ""
    body = f"""<div class="card" style="max-width:460px;margin:60px auto">
<h2>Вход</h2>{error_html}<form method="post" action="/login"><label>Пароль администратора</label>
<input name="password" type="password" required autofocus><div style="height:14px"></div><button>Войти</button></form></div>"""
    return layout("Вход", body)


@app.post("/login", response_class=HTMLResponse)
def login(request: Request, password: str = Form(...)):
    if not secrets.compare_digest(password, ADMIN_PASSWORD):
        return HTMLResponse(login_page(request, "Неверный пароль"), status_code=401)
    request.session.clear()
    request.session["auth"] = True
    return RedirectResponse("/", status_code=303)


@app.post("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login", status_code=303)


def status_badge(status: str) -> str:
    labels = {"pending": "ожидает", "running": "работает", "done": "готово", "failed": "ошибка"}
    cls = "ok" if status == "done" else "bad" if status == "failed" else ""
    return f'<span class="badge {cls}">{esc(labels.get(status, status))}</span>'


@app.get("/", response_class=HTMLResponse)
def dashboard(request: Request, db: Session = Depends(get_db)):
    require_login(request)
    jobs = list(db.scalars(select(Job).order_by(Job.id.desc()).limit(100)).all())
    rows = "".join(
        f"<tr><td><a href='/jobs/{j.id}'>#{j.id}</a></td><td>{esc(j.city)}</td><td>{esc(j.categories)}</td><td>{'Без сайта' if j.mode == 'no_site' else 'Аудит сайтов'}</td><td>{status_badge(j.status)}</td><td>{j.saved_count}</td></tr>"
        for j in jobs
    ) or '<tr><td colspan="6" class="muted">Задач пока нет</td></tr>'
    settings = f"""<div class="card"><h2>Состояние подключения</h2><div class="grid">
<div>Ключ 2ГИС: <b class="{'ok' if TWOGIS_API_KEY else 'bad'}">{'есть' if TWOGIS_API_KEY else 'не задан'}</b></div>
<div>Сохранение данных разрешено: <b class="{'ok' if DATA_LICENSE_ACKNOWLEDGED else 'bad'}">{'да' if DATA_LICENSE_ACKNOWLEDGED else 'нет'}</b></div>
<div>Telegram: <b class="{'ok' if TELEGRAM_BOT_TOKEN and TELEGRAM_ADMIN_CHAT_ID else 'muted'}">{'настроен' if TELEGRAM_BOT_TOKEN and TELEGRAM_ADMIN_CHAT_ID else 'не настроен'}</b></div>
</div></div>"""
    form = """<div class="card"><h2>Новый поиск</h2><form method="post" action="/jobs">
<div class="grid"><div><label>Город</label><input name="city" placeholder="Москва" required></div>
<div><label>Режим</label><select name="mode"><option value="no_site">Без сайта, но с контактами</option><option value="audit_sites">Есть сайт — сделать технический аудит</option></select></div>
<div><label>Максимум на категорию</label><input name="max_results" type="number" min="1" max="100" value="30"></div></div>
<div style="height:14px"></div><label>Категории — через запятую или каждую с новой строки</label>
<textarea name="categories" placeholder="стоматология&#10;автосервис&#10;салон красоты" required></textarea><div style="height:14px"></div><button>Запустить поиск</button></form></div>"""
    history = f"""<div class="card"><h2>История</h2><div class="table-wrap"><table><thead><tr><th>ID</th><th>Город</th><th>Категории</th><th>Режим</th><th>Статус</th><th>Лидов</th></tr></thead><tbody>{rows}</tbody></table></div></div>"""
    return layout("Панель", settings + form + history, request)


@app.post("/jobs")
def create_job(
    request: Request,
    background_tasks: BackgroundTasks,
    city: str = Form(...),
    categories: str = Form(...),
    mode: str = Form("no_site"),
    max_results: int = Form(30),
    db: Session = Depends(get_db),
):
    require_login(request)
    category_list = [x.strip() for x in re.split(r"[,\n]+", categories) if x.strip()]
    if not category_list:
        raise HTTPException(400, "Укажи хотя бы одну категорию")
    job = Job(
        city=city.strip(),
        categories="\n".join(dict.fromkeys(category_list)),
        mode="audit_sites" if mode == "audit_sites" else "no_site",
        max_results=max(1, min(100, max_results)),
    )
    db.add(job)
    db.commit()
    db.refresh(job)
    background_tasks.add_task(run_job, job.id)
    return RedirectResponse(f"/jobs/{job.id}", status_code=303)


@app.get("/jobs/{job_id}", response_class=HTMLResponse)
def job_page(job_id: int, request: Request, db: Session = Depends(get_db)):
    require_login(request)
    job = db.get(Job, job_id)
    if not job:
        raise HTTPException(404, "Задача не найдена")
    leads = list(db.scalars(select(Lead).where(Lead.job_id == job_id).order_by(Lead.id)).all())
    rows = ""
    for lead in leads:
        score = "—" if lead.audit_score is None else str(lead.audit_score)
        site = f'<a href="{esc(lead.website)}" target="_blank" rel="noopener">открыть</a>' if lead.website else "—"
        rows += f"""<tr><td>{esc(lead.name)}</td><td>{esc(lead.category)}</td><td>{esc(lead.address)}</td>
<td>{esc(lead.phones)}</td><td>{esc(lead.emails)}</td><td>{site}</td><td>{esc(lead.rating)}</td><td>{score}</td><td>{esc(lead.audit_notes or '')}</td></tr>"""
    if not rows:
        rows = '<tr><td colspan="9" class="muted">Пока нет результатов</td></tr>'
    error = f'<div class="card"><h3 class="bad">Ошибка</h3><pre>{esc(job.error)}</pre></div>' if job.error else ""
    refresh = 4 if job.status in {"pending", "running"} else None
    body = f"""<div class="card"><a href="/">← Назад</a><h2>Задача #{job.id}</h2>
<div class="grid"><div>Город: <b>{esc(job.city)}</b></div><div>Статус: {status_badge(job.status)}</div><div>Найдено: <b>{job.found_count}</b></div><div>Сохранено: <b>{job.saved_count}</b></div></div>
<p class="muted">{esc(job.categories).replace(chr(10), ', ')}</p>
<a class="button" href="/jobs/{job.id}/export.xlsx">Скачать Excel</a></div>{error}
<div class="card"><div class="table-wrap"><table><thead><tr><th>Компания</th><th>Категория</th><th>Адрес</th><th>Телефоны</th><th>Email</th><th>Сайт</th><th>Рейтинг</th><th>Оценка</th><th>Что не так</th></tr></thead><tbody>{rows}</tbody></table></div></div>"""
    return layout(f"Задача #{job.id}", body, request, refresh)


@app.get("/jobs/{job_id}/export.xlsx")
def export_job(job_id: int, request: Request, db: Session = Depends(get_db)):
    require_login(request)
    job = db.get(Job, job_id)
    if not job:
        raise HTTPException(404, "Задача не найдена")
    leads = list(db.scalars(select(Lead).where(Lead.job_id == job_id).order_by(Lead.id)).all())
    wb = Workbook()
    ws = wb.active
    ws.title = "Лиды"
    headers = ["Компания", "Категория", "Адрес", "Телефоны", "Email", "Мессенджеры", "Сайт", "Рейтинг", "Оценка сайта", "Проблемы", "Черновик обращения"]
    ws.append(headers)
    for lead in leads:
        ws.append([lead.name, lead.category, lead.address, lead.phones, lead.emails, lead.messengers, lead.website or "", lead.rating, lead.audit_score, lead.audit_notes or "", lead.draft_text or ""])
    for col in ws.columns:
        width = min(55, max(12, max(len(str(cell.value or "")) for cell in col) + 2))
        ws.column_dimensions[col[0].column_letter].width = width
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f'attachment; filename="leads-{job_id}.xlsx"'})


def clean_list(values: Iterable[str]) -> list[str]:
    result: list[str] = []
    for value in values:
        value = value.strip()
        if value and value not in result:
            result.append(value)
    return result


def normalize_phone(value: str) -> str:
    digits = re.sub(r"\D", "", value)
    if len(digits) == 11 and digits.startswith("8"):
        digits = "7" + digits[1:]
    return "+" + digits if digits else ""


def normalize_url(value: str) -> str:
    value = value.strip()
    if not value:
        return ""
    if not re.match(r"^https?://", value, re.I):
        value = "https://" + value
    return value


def parse_company(item: dict, city: str, fallback_category: str) -> dict:
    phones: list[str] = []
    emails: list[str] = []
    messengers: list[str] = []
    websites: list[str] = []
    for group in item.get("contact_groups") or []:
        for contact in group.get("contacts") or []:
            ctype = str(contact.get("type") or "").lower()
            value = str(contact.get("value") or contact.get("text") or contact.get("url") or "").strip()
            if ctype in {"phone", "fax"}:
                phones.append(normalize_phone(value))
            elif ctype in {"email", "mail"}:
                emails.append(value.lower())
            elif ctype in {"website", "site", "url"}:
                websites.append(normalize_url(value))
            elif ctype in {"whatsapp", "telegram", "viber", "vk", "vkontakte", "instagram"}:
                messengers.append(value)
    rubrics = item.get("rubrics") or []
    category = ", ".join(str(r.get("name") or r.get("alias") or "") for r in rubrics).strip(", ") or fallback_category
    rating_raw = (item.get("reviews") or {}).get("general_rating")
    try:
        rating = float(rating_raw) if rating_raw is not None else None
    except (TypeError, ValueError):
        rating = None
    return {
        "source_id": str(item.get("id") or ""),
        "name": str(item.get("name") or "Без названия"),
        "category": category,
        "address": str(item.get("full_address_name") or item.get("address_name") or city),
        "phones": clean_list(phones),
        "emails": clean_list(emails),
        "messengers": clean_list(messengers),
        "website": next(iter(clean_list(websites)), None),
        "rating": rating,
    }


async def search_twogis(city: str, category: str, has_site: bool, limit: int) -> list[dict]:
    page_size = min(50, limit)
    results: list[dict] = []
    async with httpx.AsyncClient(timeout=25) as client:
        page = 1
        while len(results) < limit:
            params = {
                "q": f"{category} {city}",
                "type": "branch",
                "has_site": str(has_site).lower(),
                "page": page,
                "page_size": page_size,
                "fields": "items.contact_groups,items.rubrics,items.reviews,items.full_address_name",
                "key": TWOGIS_API_KEY,
            }
            response = await client.get(TWOGIS_URL, params=params)
            if response.status_code >= 400:
                detail = response.text[:500]
                raise RuntimeError(f"2ГИС вернул HTTP {response.status_code}: {detail}")
            items = response.json().get("result", {}).get("items", [])
            if not items:
                break
            results.extend(parse_company(item, city, category) for item in items)
            if len(items) < page_size:
                break
            page += 1
    return results[:limit]


def hostname_is_public(hostname: str) -> bool:
    try:
        addresses = socket.getaddrinfo(hostname, None)
    except OSError:
        return False
    for address in addresses:
        ip = ipaddress.ip_address(address[4][0])
        if ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved or ip.is_multicast:
            return False
    return True


async def audit_site(url: str) -> tuple[int, str]:
    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname or not hostname_is_public(parsed.hostname):
        return 0, "Небезопасный или недоступный адрес"
    score = 100
    issues: list[str] = []
    if parsed.scheme != "https":
        score -= 20
        issues.append("нет HTTPS")
    try:
        async with httpx.AsyncClient(timeout=12, follow_redirects=True, headers={"User-Agent": "Mozilla/5.0 SiteAudit/1.0"}) as client:
            response = await client.get(url)
        text = response.text[:1_500_000]
        lower = text.lower()
        if response.status_code >= 400:
            score -= 40
            issues.append(f"HTTP {response.status_code}")
        if "name=\"viewport\"" not in lower and "name='viewport'" not in lower:
            score -= 20
            issues.append("не найдена мобильная адаптация viewport")
        if not re.search(r"<title[^>]*>\s*[^<]{3,}", text, re.I):
            score -= 10
            issues.append("нет нормального title")
        if not re.search(r"<h1[^>]*>.*?</h1>", text, re.I | re.S):
            score -= 10
            issues.append("нет заголовка H1")
        if not re.search(r"<(form|button)[^>]*>", text, re.I):
            score -= 10
            issues.append("не найдена форма или кнопка действия")
        if len(text.encode("utf-8", errors="ignore")) > 1_000_000:
            score -= 5
            issues.append("тяжёлая главная страница")
        years = [int(y) for y in re.findall(r"(?:©|copyright)[^\n<]{0,30}(20\d{2})", lower, re.I)]
        current_year = datetime.now().year
        if years and max(years) < current_year - 2:
            score -= 10
            issues.append(f"старый год в подвале: {max(years)}")
    except Exception as exc:
        return 10, f"сайт не удалось проверить: {type(exc).__name__}"
    return max(0, score), ", ".join(issues) if issues else "явных технических проблем не найдено"


def draft_for(company: dict, score: int | None, notes: str | None) -> str:
    if company.get("website"):
        problem = notes or "есть возможности для улучшения сайта"
        return f"Здравствуйте! Посмотрел сайт компании «{company['name']}». Заметил: {problem}. Могу предложить короткий план обновления без обязательств. Кому можно его отправить?"
    return f"Здравствуйте! Увидел компанию «{company['name']}» в 2ГИС. Заметил, что отдельный сайт не указан. Могу бесплатно набросать структуру простой страницы с услугами и контактами. Кому можно отправить пример?"


async def telegram_notify(text: str) -> None:
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_ADMIN_CHAT_ID:
        return
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    try:
        async with httpx.AsyncClient(timeout=15) as client:
            await client.post(url, json={"chat_id": TELEGRAM_ADMIN_CHAT_ID, "text": text[:3900]})
    except Exception:
        pass


async def run_job(job_id: int) -> None:
    with SessionLocal() as db:
        job = db.get(Job, job_id)
        if not job:
            return
        job.status = "running"
        db.commit()
    try:
        if not TWOGIS_API_KEY:
            raise RuntimeError("Не задан TWOGIS_API_KEY в настройках Render")
        if not DATA_LICENSE_ACKNOWLEDGED:
            raise RuntimeError("DATA_LICENSE_ACKNOWLEDGED=false. Поставь true только если твой доступ 2ГИС разрешает сохранять данные")
        with SessionLocal() as db:
            job = db.get(Job, job_id)
            if not job:
                return
            city = job.city
            categories = [x.strip() for x in job.categories.splitlines() if x.strip()]
            mode = job.mode
            limit = job.max_results
        collected: list[dict] = []
        for category in categories:
            companies = await search_twogis(city, category, has_site=(mode == "audit_sites"), limit=limit)
            collected.extend(companies)
        seen: set[str] = set()
        unique: list[dict] = []
        for company in collected:
            key = company["source_id"] or (company["name"].lower() + "|" + (company["phones"][0] if company["phones"] else company["address"].lower()))
            if key in seen:
                continue
            seen.add(key)
            has_contact = bool(company["phones"] or company["emails"] or company["messengers"])
            if mode == "no_site" and (company["website"] or not has_contact):
                continue
            if mode == "audit_sites" and not company["website"]:
                continue
            unique.append(company)
        with SessionLocal() as db:
            job = db.get(Job, job_id)
            if not job:
                return
            job.found_count = len(collected)
            db.commit()
        for company in unique:
            score: int | None = None
            notes: str | None = None
            if mode == "audit_sites" and company["website"]:
                score, notes = await audit_site(company["website"])
            lead = Lead(
                job_id=job_id,
                source_id=company["source_id"],
                name=company["name"],
                category=company["category"],
                address=company["address"],
                phones=", ".join(company["phones"]),
                emails=", ".join(company["emails"]),
                messengers=", ".join(company["messengers"]),
                website=company["website"],
                rating=company["rating"],
                audit_score=score,
                audit_notes=notes,
                draft_text=draft_for(company, score, notes),
            )
            with SessionLocal() as db:
                db.add(lead)
                db.commit()
            if mode == "audit_sites":
                await asyncio.sleep(0.15)
        with SessionLocal() as db:
            job = db.get(Job, job_id)
            if job:
                job.status = "done"
                job.saved_count = len(unique)
                job.finished_at = datetime.now(timezone.utc)
                db.commit()
        await telegram_notify(f"✅ Задача #{job_id} завершена\nСохранено лидов: {len(unique)}\nГород: {city}\nКатегории: {', '.join(categories)}")
    except Exception as exc:
        message = str(exc)[:3000]
        with SessionLocal() as db:
            job = db.get(Job, job_id)
            if job:
                job.status = "failed"
                job.error = message
                job.finished_at = datetime.now(timezone.utc)
                db.commit()
        await telegram_notify(f"❌ Ошибка в задаче #{job_id}\n{message}")
