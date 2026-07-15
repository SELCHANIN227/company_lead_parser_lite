from __future__ import annotations

import asyncio
import html
import io
import ipaddress
import hashlib
import json
import os
import re
import secrets
import socket
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable
from urllib.parse import parse_qs, quote_plus, unquote, urljoin, urlparse, urlunparse

import httpx
from fastapi import BackgroundTasks, Depends, FastAPI, Form, HTTPException, Request
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse, StreamingResponse
from openpyxl import Workbook
from playwright.async_api import Page, Response, TimeoutError as PlaywrightTimeoutError, async_playwright
from sqlalchemy import DateTime, Float, ForeignKey, Integer, String, Text, create_engine, select
from sqlalchemy.orm import DeclarativeBase, Mapped, Session, mapped_column, relationship, sessionmaker
from starlette.middleware.sessions import SessionMiddleware

APP_TITLE = "Парсер компаний — Яндекс + 2ГИС"
TWOGIS_URL = "https://catalog.api.2gis.com/3.0/items"
YANDEX_MAPS_URL = "https://yandex.ru/maps/"
YANDEX_BROWSER_LOCK = asyncio.Lock()


def env(name: str, default: str = "") -> str:
    return os.getenv(name, default).strip()


def env_bool(name: str, default: bool = False) -> bool:
    value = env(name, "true" if default else "false").lower()
    return value in {"1", "true", "yes", "on", "да"}


DEBUG_DIR = Path(env("DEBUG_DIR", "/tmp/yandex-debug"))


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
    # Совместимость со старой БД: теперь в mode хранится source:filter, например yandex:no_site.
    mode: Mapped[str] = mapped_column(String(60), default="twogis:no_site")
    max_results: Mapped[int] = mapped_column(Integer, default=20)
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
    source_id: Mapped[str] = mapped_column(Text, default="")
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
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)


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
:root {{ color-scheme: dark; --bg:#0e1117; --card:#171b22; --line:#2b313b; --text:#f4f6f8; --muted:#aab2bf; --accent:#7c5cff; --ok:#37c978; --bad:#ff5d6c; --warn:#f6c85f; }}
* {{ box-sizing:border-box }} body {{ margin:0; background:var(--bg); color:var(--text); font:16px/1.45 system-ui,-apple-system,Segoe UI,sans-serif }}
a {{ color:#a997ff }} .wrap {{ max-width:1260px; margin:0 auto; padding:24px }}
header {{ display:flex; align-items:center; justify-content:space-between; gap:16px; margin-bottom:24px }}
.card {{ background:var(--card); border:1px solid var(--line); border-radius:16px; padding:20px; margin-bottom:18px }}
.grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:14px }}
label {{ display:block; color:var(--muted); margin-bottom:6px }} input,textarea,select {{ width:100%; background:#0c0f14; color:var(--text); border:1px solid var(--line); border-radius:10px; padding:11px }}
textarea {{ min-height:110px; resize:vertical }} button,.button {{ display:inline-block; border:0; border-radius:10px; background:var(--accent); color:white; padding:11px 16px; cursor:pointer; text-decoration:none; font-weight:650 }}
.secondary {{ background:#29303a }} table {{ width:100%; border-collapse:collapse; font-size:14px }} th,td {{ text-align:left; vertical-align:top; padding:10px; border-bottom:1px solid var(--line) }} th {{ color:var(--muted) }}
.badge {{ display:inline-block; padding:4px 8px; border-radius:999px; background:#29303a; font-size:12px }} .ok {{ color:var(--ok) }} .bad {{ color:var(--bad) }} .warn {{ color:var(--warn) }} .muted {{ color:var(--muted) }}
pre {{ white-space:pre-wrap; word-break:break-word; background:#0c0f14; padding:12px; border-radius:10px }}
.notice {{ border-left:4px solid var(--warn); padding-left:14px }}
@media(max-width:700px) {{ .wrap {{ padding:14px }} .table-wrap {{ overflow:auto }} table {{ min-width:1050px }} }}
</style></head><body><div class="wrap"><header><div><h1 style="margin:0">{esc(APP_TITLE)}</h1><div class="muted">Яндекс Карты (браузер) + 2ГИС API → Excel</div></div>{logout}</header>{body}</div></body></html>"""


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


def split_job_mode(raw_mode: str) -> tuple[str, str]:
    if ":" in raw_mode:
        source, mode = raw_mode.split(":", 1)
    else:
        source, mode = "twogis", raw_mode
    if source not in {"yandex", "twogis"}:
        source = "twogis"
    if mode not in {"all", "no_site", "audit_sites"}:
        mode = "no_site"
    return source, mode


def source_label(source: str) -> str:
    return "Яндекс Карты" if source == "yandex" else "2ГИС"


def mode_label(mode: str) -> str:
    return {
        "all": "Все карточки",
        "no_site": "Без сайта + есть контакт",
        "audit_sites": "Сайт есть — аудит",
    }.get(mode, mode)


def status_badge(status: str) -> str:
    labels = {"pending": "ожидает", "running": "работает", "done": "готово", "failed": "ошибка"}
    cls = "ok" if status == "done" else "bad" if status == "failed" else ""
    return f'<span class="badge {cls}">{esc(labels.get(status, status))}</span>'


def lead_source(lead: Lead) -> tuple[str, str]:
    if lead.source_id.startswith("yandex:"):
        candidate = lead.source_id[len("yandex:") :]
        return "Яндекс", candidate if candidate.startswith(("http://", "https://")) else ""
    if lead.source_id.startswith(("yandex-id:", "yandex-dom:")):
        return "Яндекс", ""
    if lead.source_id.startswith("2gis:"):
        return "2ГИС", ""
    return "2ГИС", ""


@app.get("/", response_class=HTMLResponse)
def dashboard(request: Request, db: Session = Depends(get_db)):
    require_login(request)
    jobs = list(db.scalars(select(Job).order_by(Job.id.desc()).limit(100)).all())
    rows_parts: list[str] = []
    for job in jobs:
        source, mode = split_job_mode(job.mode)
        rows_parts.append(
            f"<tr><td><a href='/jobs/{job.id}'>#{job.id}</a></td><td>{esc(job.city)}</td>"
            f"<td>{esc(job.categories)}</td><td>{esc(source_label(source))}</td><td>{esc(mode_label(mode))}</td>"
            f"<td>{status_badge(job.status)}</td><td>{job.saved_count}</td></tr>"
        )
    rows = "".join(rows_parts) or '<tr><td colspan="7" class="muted">Задач пока нет</td></tr>'
    settings = f"""<div class="card"><h2>Состояние подключения</h2><div class="grid">
<div>Яндекс без API: <b class="ok">готов</b> <span class="muted">(эксперимент)</span></div>
<div>Ключ 2ГИС: <b class="{'ok' if TWOGIS_API_KEY else 'bad'}">{'есть' if TWOGIS_API_KEY else 'не задан'}</b></div>
<div>Сохранение данных 2ГИС разрешено: <b class="{'ok' if DATA_LICENSE_ACKNOWLEDGED else 'bad'}">{'да' if DATA_LICENSE_ACKNOWLEDGED else 'нет'}</b></div>
<div>Telegram: <b class="{'ok' if TELEGRAM_BOT_TOKEN and TELEGRAM_ADMIN_CHAT_ID else 'muted'}">{'настроен' if TELEGRAM_BOT_TOKEN and TELEGRAM_ADMIN_CHAT_ID else 'не настроен'}</b></div>
</div></div>"""
    notice = """<div class="card notice"><b>Яндекс-режим работает браузером без API.</b>
Он собирает только видимые данные, не обходит CAPTCHA и остановится, если Яндекс запросит подтверждение. Для первого теста поставь одну категорию и максимум 5–10.</div>"""
    form = """<div class="card"><h2>Новый поиск</h2><form method="post" action="/jobs">
<div class="grid"><div><label>Город</label><input name="city" placeholder="Воронеж" required></div>
<div><label>Источник</label><select name="source"><option value="yandex" selected>Яндекс Карты — без API</option><option value="twogis">2ГИС API</option></select></div>
<div><label>Режим</label><select name="mode"><option value="all">Все карточки — для проверки</option><option value="no_site" selected>Без сайта, но с контактом</option><option value="audit_sites">Есть сайт — технический аудит</option></select></div>
<div><label>Максимум на категорию</label><input name="max_results" type="number" min="1" max="30" value="10"></div></div>
<div style="height:14px"></div><label>Категории — через запятую или каждую с новой строки</label>
<textarea name="categories" placeholder="стоматология&#10;автосервис&#10;салон красоты" required></textarea><div style="height:14px"></div><button>Запустить поиск</button></form></div>"""
    history = f"""<div class="card"><h2>История</h2><div class="table-wrap"><table><thead><tr><th>ID</th><th>Город</th><th>Категории</th><th>Источник</th><th>Режим</th><th>Статус</th><th>Лидов</th></tr></thead><tbody>{rows}</tbody></table></div></div>"""
    return layout("Панель", settings + notice + form + history, request)


@app.post("/jobs")
def create_job(
    request: Request,
    background_tasks: BackgroundTasks,
    city: str = Form(...),
    categories: str = Form(...),
    source: str = Form("yandex"),
    mode: str = Form("no_site"),
    max_results: int = Form(10),
    db: Session = Depends(get_db),
):
    require_login(request)
    category_list = [x.strip() for x in re.split(r"[,\n]+", categories) if x.strip()]
    if not category_list:
        raise HTTPException(400, "Укажи хотя бы одну категорию")
    source = "yandex" if source == "yandex" else "twogis"
    mode = mode if mode in {"all", "no_site", "audit_sites"} else "no_site"
    category_limit = 12 if source == "yandex" else 30
    category_list = list(dict.fromkeys(category_list))[:category_limit]
    per_category_limit = max(1, min(30 if source == "yandex" else 100, max_results))
    job = Job(
        city=city.strip(),
        categories="\n".join(category_list),
        mode=f"{source}:{mode}",
        max_results=per_category_limit,
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
    source, mode = split_job_mode(job.mode)
    leads = list(db.scalars(select(Lead).where(Lead.job_id == job_id).order_by(Lead.id)).all())
    rows = ""
    for lead in leads:
        lead_source_name, card_url = lead_source(lead)
        score = "—" if lead.audit_score is None else str(lead.audit_score)
        site = f'<a href="{esc(lead.website)}" target="_blank" rel="noopener">открыть</a>' if lead.website else "—"
        card = f'<a href="{esc(card_url)}" target="_blank" rel="noopener">карточка</a>' if card_url else "—"
        rows += f"""<tr><td>{esc(lead_source_name)}</td><td>{esc(lead.name)}</td><td>{esc(lead.category)}</td><td>{esc(lead.address)}</td>
<td>{esc(lead.phones)}</td><td>{esc(lead.emails)}</td><td>{site}</td><td>{card}</td><td>{esc(lead.rating)}</td><td>{score}</td><td>{esc(lead.audit_notes or '')}</td></tr>"""
    if not rows:
        rows = '<tr><td colspan="11" class="muted">Нет карточек, прошедших выбранный фильтр. Смотри диагностику выше.</td></tr>'
    diagnostic = ""
    if job.error:
        heading = "Ошибка" if job.status == "failed" else "Диагностика"
        cls = "bad" if job.status == "failed" else "warn"
        debug_links: list[str] = []
        for path in sorted(DEBUG_DIR.glob(f"job-{job.id}-category-*.*")):
            if path.suffix not in {".png", ".html"}:
                continue
            label = "Скриншот" if path.suffix == ".png" else "HTML страницы"
            debug_links.append(f'<a class="button secondary" href="/jobs/{job.id}/debug/{esc(path.name)}" target="_blank">{label}</a>')
        links_html = ("<div style=\"display:flex;gap:10px;flex-wrap:wrap;margin-top:12px\">" + "".join(debug_links) + "</div>") if debug_links else ""
        diagnostic = f'<div class="card"><h3 class="{cls}">{heading}</h3><pre>{esc(job.error)}</pre>{links_html}</div>'
    refresh = 5 if job.status in {"pending", "running"} else None
    body = f"""<div class="card"><a href="/">← Назад</a><h2>Задача #{job.id}</h2>
<div class="grid"><div>Город: <b>{esc(job.city)}</b></div><div>Источник: <b>{esc(source_label(source))}</b></div><div>Режим: <b>{esc(mode_label(mode))}</b></div><div>Статус: {status_badge(job.status)}</div><div>Получено карточек: <b>{job.found_count}</b></div><div>Сохранено: <b>{job.saved_count}</b></div></div>
<p class="muted">{esc(job.categories).replace(chr(10), ', ')}</p>
<a class="button" href="/jobs/{job.id}/export.xlsx">Скачать Excel</a></div>{diagnostic}
<div class="card"><div class="table-wrap"><table><thead><tr><th>Источник</th><th>Компания</th><th>Категория</th><th>Адрес</th><th>Телефоны</th><th>Email</th><th>Сайт</th><th>Карточка</th><th>Рейтинг</th><th>Оценка</th><th>Что не так</th></tr></thead><tbody>{rows}</tbody></table></div></div>"""
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
    headers = ["Источник", "Компания", "Категория", "Адрес", "Телефоны", "Email", "Мессенджеры", "Сайт", "Карточка на картах", "Рейтинг", "Оценка сайта", "Проблемы", "Черновик обращения"]
    ws.append(headers)
    for lead in leads:
        source_name, card_url = lead_source(lead)
        ws.append([source_name, lead.name, lead.category, lead.address, lead.phones, lead.emails, lead.messengers, lead.website or "", card_url, lead.rating, lead.audit_score, lead.audit_notes or "", lead.draft_text or ""])
    for col in ws.columns:
        width = min(55, max(12, max(len(str(cell.value or "")) for cell in col) + 2))
        ws.column_dimensions[col[0].column_letter].width = width
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f'attachment; filename="leads-{job_id}.xlsx"'})


@app.get("/jobs/{job_id}/debug/{filename}")
def job_debug_file(job_id: int, filename: str, request: Request):
    require_login(request)
    if not re.fullmatch(rf"job-{job_id}-category-\d+\.(?:png|html)", filename):
        raise HTTPException(404, "Файл диагностики не найден")
    path = DEBUG_DIR / filename
    if not path.is_file():
        raise HTTPException(404, "Файл диагностики не найден")
    media_type = "image/png" if path.suffix == ".png" else "text/html"
    return FileResponse(path, media_type=media_type, filename=filename)


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
    return "+" + digits if len(digits) >= 10 else ""


def normalize_url(value: str) -> str:
    value = value.strip()
    if not value:
        return ""
    if not re.match(r"^https?://", value, re.I):
        value = "https://" + value
    return value


def parse_company(item: dict[str, Any], city: str, fallback_category: str) -> dict[str, Any]:
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
        "source": "2gis",
        "source_id": "2gis:" + str(item.get("id") or ""),
        "name": str(item.get("name") or "Без названия"),
        "category": category,
        "address": str(item.get("full_address_name") or item.get("address_name") or city),
        "phones": clean_list(phones),
        "emails": clean_list(emails),
        "messengers": clean_list(messengers),
        "website": next(iter(clean_list(websites)), None),
        "rating": rating,
    }


async def search_twogis(city: str, category: str, has_site: bool | None, limit: int) -> list[dict[str, Any]]:
    page_size = min(50, limit)
    results: list[dict[str, Any]] = []
    async with httpx.AsyncClient(timeout=25) as client:
        page = 1
        while len(results) < limit:
            params: dict[str, Any] = {
                "q": f"{category} {city}",
                "type": "branch",
                "page": page,
                "page_size": page_size,
                "fields": "items.contact_groups,items.rubrics,items.reviews,items.full_address_name",
                "key": TWOGIS_API_KEY,
            }
            if has_site is not None:
                params["has_site"] = str(has_site).lower()
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


def yandex_captcha_text(text: str) -> bool:
    lowered = text.lower()
    return any(
        marker in lowered
        for marker in [
            "проверка, что вы не робот",
            "подтвердите, что запросы отправляли вы",
            "введите символы",
            "showcaptcha",
            "captcha",
        ]
    )


async def page_has_captcha(page: Page) -> bool:
    if "captcha" in page.url.lower() or "showcaptcha" in page.url.lower():
        return True
    try:
        text = await page.locator("body").inner_text(timeout=3000)
    except Exception:
        return False
    return yandex_captcha_text(text[:12000])


async def close_yandex_popups(page: Page) -> None:
    for label in ["Принять", "Хорошо", "Понятно", "Закрыть", "Не сейчас", "Разрешить"]:
        try:
            button = page.get_by_role("button", name=re.compile(label, re.I)).first
            if await button.count() and await button.is_visible():
                await button.click(timeout=1500)
                await page.wait_for_timeout(300)
        except Exception:
            pass


def canonical_yandex_org_url(base_url: str, href: str) -> str:
    href = unquote(str(href or "").strip())
    if not href:
        return ""
    if href.startswith("ymapsbm1://"):
        parsed_custom = urlparse(href)
        oid = (parse_qs(parsed_custom.query).get("oid") or [""])[0]
        return f"https://yandex.ru/maps/org/{oid}" if oid else ""
    absolute = urljoin(base_url, href)
    parsed = urlparse(absolute)
    if "/org/" in parsed.path:
        return urlunparse((parsed.scheme or "https", parsed.netloc or "yandex.ru", parsed.path.rstrip("/"), "", "", ""))
    query = parse_qs(parsed.query)
    oid = (query.get("oid") or query.get("orgpage") or [""])[0]
    if oid:
        return f"https://yandex.ru/maps/org/{oid}"
    return ""


def walk_json(value: Any) -> Iterable[Any]:
    yield value
    if isinstance(value, dict):
        for nested in value.values():
            yield from walk_json(nested)
    elif isinstance(value, list):
        for nested in value:
            yield from walk_json(nested)


def clean_text(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "")).strip()


def yandex_card_url_from_meta(meta: dict[str, Any], node: dict[str, Any]) -> str:
    candidates = [
        meta.get("uri"),
        node.get("uri"),
        (node.get("properties") or {}).get("uri") if isinstance(node.get("properties"), dict) else None,
    ]
    for candidate in candidates:
        url = canonical_yandex_org_url(YANDEX_MAPS_URL, str(candidate or ""))
        if url:
            return url
    oid = clean_text(meta.get("id") or meta.get("oid") or node.get("id"))
    if oid and re.search(r"\d", oid):
        return f"https://yandex.ru/maps/org/{oid}"
    return ""


def values_from_objects(value: Any, keys: set[str]) -> list[str]:
    result: list[str] = []
    if isinstance(value, dict):
        for key, nested in value.items():
            if str(key).lower() in keys and isinstance(nested, (str, int, float)):
                result.append(clean_text(nested))
            result.extend(values_from_objects(nested, keys))
    elif isinstance(value, list):
        for nested in value:
            result.extend(values_from_objects(nested, keys))
    return clean_list(result)


def yandex_company_from_node(node: dict[str, Any], fallback_category: str, city: str) -> dict[str, Any] | None:
    properties = node.get("properties") if isinstance(node.get("properties"), dict) else {}
    meta = properties.get("CompanyMetaData") if isinstance(properties.get("CompanyMetaData"), dict) else None
    if meta is None and isinstance(node.get("CompanyMetaData"), dict):
        meta = node.get("CompanyMetaData")
    if meta is None:
        # Некоторые ответы кладут карточку без обёртки CompanyMetaData.
        probable_name = clean_text(node.get("name") or properties.get("name") or node.get("title"))
        probable_address = clean_text(node.get("address") or properties.get("description") or properties.get("address"))
        has_business_fields = any(key in node for key in ["phones", "Phones", "categories", "Categories", "ratingData", "workingTime"])
        if not probable_name or not (probable_address or has_business_fields):
            return None
        meta = node

    name = clean_text(meta.get("name") or properties.get("name") or node.get("name") or node.get("title"))
    if not name or name.lower() in {"без названия", "яндекс карты", "карты"}:
        return None

    address_value = meta.get("address") or properties.get("description") or properties.get("address") or node.get("address")
    address = clean_text(address_value) or city

    categories_raw = meta.get("Categories") or meta.get("categories") or node.get("categories") or []
    categories: list[str] = []
    if isinstance(categories_raw, list):
        for item in categories_raw:
            if isinstance(item, dict):
                categories.append(clean_text(item.get("name") or item.get("class") or item.get("id")))
            else:
                categories.append(clean_text(item))
    category = ", ".join(clean_list(categories)) or fallback_category

    phone_values = values_from_objects(meta.get("Phones") or meta.get("phones") or [], {"formatted", "number", "value", "phone"})
    phones = clean_list(normalize_phone(value) for value in phone_values)
    email_values = values_from_objects(meta.get("Emails") or meta.get("emails") or [], {"value", "email", "address"})
    emails = clean_list(value.lower() for value in email_values if "@" in value)

    websites: list[str] = []
    for key in ["url", "website", "site"]:
        value = meta.get(key)
        if isinstance(value, str):
            target = external_target(value) or normalize_url(value)
            if target and "yandex." not in (urlparse(target).hostname or ""):
                websites.append(target)
    links_raw = meta.get("Links") or meta.get("links") or []
    if isinstance(links_raw, list):
        for item in links_raw:
            if isinstance(item, dict):
                href = clean_text(item.get("href") or item.get("url") or item.get("value"))
                target = external_target(href) or (normalize_url(href) if href else "")
                if target and "yandex." not in (urlparse(target).hostname or ""):
                    websites.append(target)

    messengers: list[str] = []
    for target in list(websites):
        host = (urlparse(target).hostname or "").lower()
        if host in {"t.me", "telegram.me", "wa.me", "api.whatsapp.com", "vk.com", "viber.com"}:
            messengers.append(target)
            websites.remove(target)

    rating: float | None = None
    rating_candidates = [
        properties.get("rating"),
        meta.get("rating"),
        (properties.get("ratingData") or {}).get("rating") if isinstance(properties.get("ratingData"), dict) else None,
        (meta.get("ratingData") or {}).get("rating") if isinstance(meta.get("ratingData"), dict) else None,
    ]
    for raw in rating_candidates:
        try:
            if raw is not None and str(raw).strip():
                rating = float(str(raw).replace(",", "."))
                break
        except ValueError:
            pass

    card_url = yandex_card_url_from_meta(meta, node)
    stable_id = clean_text(meta.get("id") or meta.get("oid") or node.get("id"))
    if card_url:
        source_id = "yandex:" + card_url
    elif stable_id:
        source_id = "yandex-id:" + stable_id
    else:
        digest = hashlib.sha1(f"{name}|{address}".encode("utf-8", errors="ignore")).hexdigest()[:16]
        source_id = "yandex-dom:" + digest

    return {
        "source": "yandex",
        "source_id": source_id,
        "name": name,
        "category": category,
        "address": address,
        "phones": phones,
        "emails": emails,
        "messengers": clean_list(messengers),
        "website": next(iter(clean_list(websites)), None),
        "rating": rating,
        "card_url": card_url,
    }


def parse_yandex_payload(payload: Any, fallback_category: str, city: str) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    seen: set[str] = set()
    for nested in walk_json(payload):
        if not isinstance(nested, dict):
            continue
        record = yandex_company_from_node(nested, fallback_category, city)
        if not record:
            continue
        key = record["source_id"] or f"{record['name'].lower()}|{record['address'].lower()}"
        if key in seen:
            continue
        seen.add(key)
        records.append(record)
    return records


class YandexResponseCapture:
    def __init__(self, fallback_category: str, city: str) -> None:
        self.fallback_category = fallback_category
        self.city = city
        self.records: list[dict[str, Any]] = []
        self.json_responses = 0
        self.total_responses = 0
        self.candidate_urls: list[str] = []
        self.tasks: list[asyncio.Task[Any]] = []

    def schedule(self, response: Response) -> None:
        self.total_responses += 1
        task = asyncio.create_task(self._consume(response))
        self.tasks.append(task)

    async def finish(self) -> None:
        if self.tasks:
            await asyncio.gather(*self.tasks, return_exceptions=True)

    async def _consume(self, response: Response) -> None:
        url = response.url
        if "yandex." not in url:
            return
        content_type = (response.headers.get("content-type") or "").lower()
        likely = any(token in url.lower() for token in ["search", "discovery", "business", "org", "maps/api"])
        if "json" not in content_type and not likely:
            return
        try:
            body = await response.body()
        except Exception:
            return
        if not body or len(body) > 8_000_000:
            return
        try:
            payload = json.loads(body)
        except (json.JSONDecodeError, UnicodeDecodeError):
            return
        self.json_responses += 1
        if likely and len(self.candidate_urls) < 20:
            self.candidate_urls.append(url[:500])
        self.records.extend(parse_yandex_payload(payload, self.fallback_category, self.city))


async def first_text(page: Page, selectors: list[str]) -> str:
    for selector in selectors:
        try:
            locator = page.locator(selector).first
            if await locator.count() and await locator.is_visible():
                text = (await locator.inner_text(timeout=1500)).strip()
                if text:
                    return re.sub(r"\s+", " ", text)
        except Exception:
            pass
    return ""


async def save_yandex_debug(page: Page, job_id: int, category_index: int) -> tuple[Path | None, Path | None]:
    DEBUG_DIR.mkdir(parents=True, exist_ok=True)
    stem = f"job-{job_id}-category-{category_index}"
    screenshot_path = DEBUG_DIR / f"{stem}.png"
    html_path = DEBUG_DIR / f"{stem}.html"
    try:
        await page.screenshot(path=str(screenshot_path), full_page=True)
    except Exception:
        screenshot_path = None
    try:
        source = await page.content()
        html_path.write_text(source[:5_000_000], encoding="utf-8")
    except Exception:
        html_path = None
    return screenshot_path, html_path


async def collect_yandex_org_links(page: Page, limit: int) -> list[str]:
    links: list[str] = []
    seen: set[str] = set()
    selectors = [
        'a[href*="/org/"]',
        'a[href*="oid="]',
        '[class*="search-business-snippet-view"] a[href]',
        '[class*="business-snippet-view"] a[href]',
        '[data-testid*="search-result"] a[href]',
    ]
    stable_rounds = 0
    for _ in range(24):
        before = len(links)
        for selector in selectors:
            try:
                hrefs = await page.locator(selector).evaluate_all(
                    "els => els.map(e => e.getAttribute('href') || e.href).filter(Boolean)"
                )
            except Exception:
                hrefs = []
            for href in hrefs:
                url = canonical_yandex_org_url(page.url, str(href))
                if url and url not in seen:
                    seen.add(url)
                    links.append(url)
                    if len(links) >= limit:
                        return links
        stable_rounds = stable_rounds + 1 if len(links) == before else 0
        if stable_rounds >= 4:
            break
        scrolled = False
        for selector in [
            '[class*="scroll__container"]',
            '[class*="search-list-view"]',
            '[class*="search-list-view__list"]',
            '[class*="sidebar-panel"]',
        ]:
            try:
                locator = page.locator(selector).first
                if await locator.count() and await locator.is_visible():
                    await locator.evaluate("el => { el.scrollTop = Math.min(el.scrollHeight, el.scrollTop + 1400); }")
                    scrolled = True
                    break
            except Exception:
                pass
        if not scrolled:
            await page.mouse.wheel(0, 1400)
        await page.wait_for_timeout(1000)
        if await page_has_captcha(page):
            raise RuntimeError("Яндекс показал CAPTCHA. Автоматический обход отключён; повтори позже.")
    return links[:limit]


async def collect_yandex_dom_cards(page: Page, fallback_category: str, city: str, limit: int) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    seen: set[str] = set()
    card_selectors = [
        '[class*="search-business-snippet-view"]',
        '[class*="business-snippet-view"]',
        '[class*="search-snippet-view"]',
        '[data-testid*="search-result"]',
    ]
    for selector in card_selectors:
        try:
            cards = page.locator(selector)
            count = min(await cards.count(), limit * 3)
        except Exception:
            continue
        for index in range(count):
            card = cards.nth(index)
            try:
                if not await card.is_visible():
                    continue
                text = clean_text(await card.inner_text(timeout=1500))
            except Exception:
                continue
            if not text or len(text) < 3:
                continue
            try:
                href = await card.locator("a[href]").first.get_attribute("href", timeout=1000)
            except Exception:
                href = ""
            card_url = canonical_yandex_org_url(page.url, href or "")
            try:
                name = clean_text(await card.locator('[class*="title"], h2, h3, a').first.inner_text(timeout=1000))
            except Exception:
                name = text.split(" ", 12)[0:6]
                name = " ".join(name)
            if not name:
                continue
            phone_matches = re.findall(r"(?:\+7|8)[\s\-(]*\d{3}[\s\-)]*\d{3}[\s\-]*\d{2}[\s\-]*\d{2}", text)
            phones = clean_list(normalize_phone(value) for value in phone_matches)
            rating: float | None = None
            rating_match = re.search(r"(?<!\d)([1-5][\.,]\d)(?!\d)", text)
            if rating_match:
                try:
                    rating = float(rating_match.group(1).replace(",", "."))
                except ValueError:
                    pass
            source_id = "yandex:" + card_url if card_url else "yandex-dom:" + hashlib.sha1(text.encode("utf-8", errors="ignore")).hexdigest()[:16]
            if source_id in seen:
                continue
            seen.add(source_id)
            records.append(
                {
                    "source": "yandex",
                    "source_id": source_id,
                    "name": name,
                    "category": fallback_category,
                    "address": city,
                    "phones": phones,
                    "emails": [],
                    "messengers": [],
                    "website": None,
                    "rating": rating,
                    "card_url": card_url,
                }
            )
            if len(records) >= limit:
                return records
    return records


def merge_company_records(base: dict[str, Any], extra: dict[str, Any]) -> dict[str, Any]:
    result = dict(base)
    for key in ["name", "category", "address", "website", "rating", "card_url"]:
        if not result.get(key) and extra.get(key):
            result[key] = extra[key]
    for key in ["phones", "emails", "messengers"]:
        result[key] = clean_list([*(result.get(key) or []), *(extra.get(key) or [])])
    if str(result.get("name", "")).startswith("Карточка не") and extra.get("name"):
        result["name"] = extra["name"]
    return result


def deduplicate_company_dicts(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    merged: dict[str, dict[str, Any]] = {}
    for record in records:
        key = str(record.get("source_id") or "")
        if not key:
            key = f"{clean_text(record.get('name')).lower()}|{clean_text(record.get('address')).lower()}"
        if key in merged:
            merged[key] = merge_company_records(merged[key], record)
        else:
            merged[key] = record
    return list(merged.values())


def jsonld_org_data(values: list[Any]) -> dict[str, Any]:
    candidates: list[dict[str, Any]] = []
    for value in values:
        for nested in walk_json(value):
            if not isinstance(nested, dict):
                continue
            kind = nested.get("@type")
            kinds = [str(x) for x in kind] if isinstance(kind, list) else [str(kind or "")]
            if any("organization" in k.lower() or "business" in k.lower() or "store" in k.lower() or "restaurant" in k.lower() for k in kinds):
                candidates.append(nested)
    return max(candidates, key=lambda x: len(x), default={})


def format_address(value: Any) -> str:
    if isinstance(value, str):
        return value.strip()
    if not isinstance(value, dict):
        return ""
    fields = [value.get("addressLocality"), value.get("streetAddress"), value.get("postalCode")]
    return ", ".join(str(x).strip() for x in fields if x)


def external_target(href: str) -> str:
    if not href:
        return ""
    absolute = unquote(href)
    parsed = urlparse(absolute)
    if parsed.netloc.endswith("yandex.ru") or parsed.netloc.endswith("yandex.com"):
        query = parse_qs(parsed.query)
        target = (query.get("url") or query.get("target") or [""])[0]
        if target:
            absolute = unquote(target)
            parsed = urlparse(absolute)
    if parsed.scheme not in {"http", "https"}:
        return ""
    host = (parsed.hostname or "").lower()
    if not host or host.endswith("yandex.ru") or host.endswith("yandex.com") or host.endswith("ya.ru"):
        return ""
    return absolute


async def extract_yandex_card(page: Page, card_url: str, fallback_category: str, city: str) -> dict[str, Any]:
    await page.goto(card_url, wait_until="domcontentloaded", timeout=60000)
    try:
        await page.wait_for_load_state("networkidle", timeout=10000)
    except PlaywrightTimeoutError:
        pass
    await page.wait_for_timeout(1200)
    await close_yandex_popups(page)
    if await page_has_captcha(page):
        raise RuntimeError("Яндекс показал CAPTCHA при открытии карточки. Обход CAPTCHA отключён.")

    for selector in ['button:has-text("Показать телефон")', '[role="button"]:has-text("Показать телефон")', '[class*="card-phones-view__more"]']:
        try:
            button = page.locator(selector).first
            if await button.count() and await button.is_visible():
                await button.click(timeout=1800)
                await page.wait_for_timeout(500)
                break
        except Exception:
            pass

    json_values: list[Any] = []
    try:
        scripts = await page.locator('script[type="application/ld+json"]').all_text_contents()
        for raw in scripts:
            try:
                json_values.append(json.loads(raw))
            except json.JSONDecodeError:
                pass
    except Exception:
        pass
    structured = jsonld_org_data(json_values)

    name = clean_text(structured.get("name"))
    if not name:
        name = await first_text(page, ['h1', '[class*="orgpage-header-view__header"]', '[class*="business-card-title-view__title"]'])
    if not name:
        title = await page.title()
        name = re.split(r"\s+[—–-]\s+", title, maxsplit=1)[0].strip()

    address = format_address(structured.get("address"))
    if not address:
        address = await first_text(page, ['[class*="business-contacts-view__address-link"]', '[class*="orgpage-header-view__address"]', '[class*="business-card-title-view__address"]', '[itemprop="address"]'])
    category = await first_text(page, ['[class*="business-card-title-view__categories"]', '[class*="orgpage-header-view__category"]', '[class*="orgpage-categories-info-view"]']) or fallback_category

    phones: list[str] = []
    telephone = structured.get("telephone")
    if isinstance(telephone, list):
        phones.extend(normalize_phone(str(x)) for x in telephone)
    elif telephone:
        phones.append(normalize_phone(str(telephone)))
    try:
        tel_hrefs = await page.locator('a[href^="tel:"]').evaluate_all("els => els.map(e => e.getAttribute('href')).filter(Boolean)")
        phones.extend(normalize_phone(str(x).removeprefix("tel:")) for x in tel_hrefs)
    except Exception:
        pass

    emails: list[str] = []
    email = structured.get("email")
    if isinstance(email, list):
        emails.extend(str(x).lower().strip() for x in email)
    elif email:
        emails.append(str(email).lower().strip())
    try:
        mail_hrefs = await page.locator('a[href^="mailto:"]').evaluate_all("els => els.map(e => e.getAttribute('href')).filter(Boolean)")
        emails.extend(str(x).removeprefix("mailto:").split("?", 1)[0].lower() for x in mail_hrefs)
    except Exception:
        pass

    websites: list[str] = []
    structured_url = structured.get("url")
    if isinstance(structured_url, str):
        target = external_target(structured_url)
        if target:
            websites.append(target)
    for selector in ['a[class*="business-urls-view__link"]', 'a[class*="business-contacts-view__website"]', 'a[data-type="website"]', 'a[aria-label*="сайт" i]']:
        try:
            hrefs = await page.locator(selector).evaluate_all("els => els.map(e => e.href || e.getAttribute('href')).filter(Boolean)")
            for href in hrefs:
                target = external_target(str(href))
                if target:
                    websites.append(target)
        except Exception:
            pass

    messengers: list[str] = []
    try:
        hrefs = await page.locator('a[href]').evaluate_all("els => els.map(e => e.href || '').filter(Boolean)")
        for href in hrefs:
            target = external_target(str(href))
            host = (urlparse(target).hostname or "").lower() if target else ""
            if host in {"t.me", "telegram.me", "wa.me", "api.whatsapp.com", "vk.com", "viber.com"}:
                messengers.append(target)
    except Exception:
        pass

    rating: float | None = None
    rating_text = await first_text(page, ['[class*="business-summary-rating-badge-view__rating-text"]', '[class*="business-rating-badge-view__rating-text"]', '[aria-label*="Рейтинг"]'])
    match = re.search(r"\d+[\.,]\d+|\d+", rating_text)
    if match:
        try:
            rating = float(match.group(0).replace(",", "."))
        except ValueError:
            pass
    if rating is None:
        aggregate = structured.get("aggregateRating")
        if isinstance(aggregate, dict):
            try:
                rating = float(str(aggregate.get("ratingValue") or "").replace(",", "."))
            except ValueError:
                pass

    return {
        "source": "yandex",
        "source_id": "yandex:" + card_url,
        "name": name or "Без названия",
        "category": category,
        "address": address or city,
        "phones": clean_list(phones),
        "emails": clean_list(emails),
        "messengers": clean_list(messengers),
        "website": next(iter(clean_list(websites)), None),
        "rating": rating,
        "card_url": card_url,
    }


async def search_yandex_categories(job_id: int, city: str, categories: list[str], limit: int) -> tuple[list[dict[str, Any]], str]:
    collected: list[dict[str, Any]] = []
    diagnostic_lines: list[str] = []
    async with YANDEX_BROWSER_LOCK:
        async with async_playwright() as playwright:
            browser = await playwright.chromium.launch(
                headless=True,
                args=["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage", "--disable-gpu"],
            )
            context = await browser.new_context(
                locale="ru-RU",
                timezone_id="Europe/Moscow",
                viewport={"width": 1440, "height": 1000},
                user_agent="Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
                extra_http_headers={"Accept-Language": "ru-RU,ru;q=0.9,en;q=0.7"},
            )
            search_page = await context.new_page()
            detail_page = await context.new_page()
            try:
                for category_index, category in enumerate(categories, start=1):
                    capture = YandexResponseCapture(category, city)
                    search_page.on("response", capture.schedule)
                    search_url = f"{YANDEX_MAPS_URL}?mode=search&text={quote_plus(category + ' ' + city)}"
                    await search_page.goto(search_url, wait_until="domcontentloaded", timeout=60000)
                    try:
                        await search_page.wait_for_load_state("networkidle", timeout=12000)
                    except PlaywrightTimeoutError:
                        pass
                    await search_page.wait_for_timeout(1800)
                    await close_yandex_popups(search_page)
                    if await page_has_captcha(search_page):
                        screenshot_path, html_path = await save_yandex_debug(search_page, job_id, category_index)
                        raise RuntimeError(
                            "Яндекс показал CAPTCHA. Обход отключён. "
                            f"Диагностика: {screenshot_path or '-'}, {html_path or '-'}"
                        )

                    links = await collect_yandex_org_links(search_page, limit)
                    dom_records = await collect_yandex_dom_cards(search_page, category, city, limit)
                    await capture.finish()
                    response_records = deduplicate_company_dicts(capture.records)
                    category_records = deduplicate_company_dicts([*response_records, *dom_records])

                    # Детальные карточки обогащают телефон и сайт. Открываем сначала найденные ссылки,
                    # затем ссылки, полученные из JSON-ответов.
                    detail_links = list(links)
                    for record in category_records:
                        card_url = str(record.get("card_url") or "")
                        if card_url and card_url not in detail_links:
                            detail_links.append(card_url)
                    enriched: list[dict[str, Any]] = []
                    for card_url in detail_links[:limit]:
                        try:
                            enriched.append(await extract_yandex_card(detail_page, card_url, category, city))
                        except PlaywrightTimeoutError:
                            enriched.append(
                                {
                                    "source": "yandex",
                                    "source_id": "yandex:" + card_url,
                                    "name": "Карточка не загрузилась",
                                    "category": category,
                                    "address": city,
                                    "phones": [],
                                    "emails": [],
                                    "messengers": [],
                                    "website": None,
                                    "rating": None,
                                    "card_url": card_url,
                                }
                            )
                        await detail_page.wait_for_timeout(350)
                    category_records = deduplicate_company_dicts([*category_records, *enriched])[:limit]

                    title = clean_text(await search_page.title())
                    body_preview = ""
                    try:
                        body_preview = clean_text((await search_page.locator("body").inner_text(timeout=3000))[:1500])
                    except Exception:
                        pass
                    diagnostic_lines.extend(
                        [
                            f"Категория: {category}",
                            f"URL: {search_page.url}",
                            f"Заголовок: {title or '-'}",
                            f"JSON-ответов: {capture.json_responses} из {capture.total_responses}",
                            f"Карточек из JSON: {len(response_records)}",
                            f"Карточек из DOM: {len(dom_records)}",
                            f"Ссылок на /org/: {len(links)}",
                            f"Итог по категории: {len(category_records)}",
                        ]
                    )
                    if not category_records:
                        screenshot_path, html_path = await save_yandex_debug(search_page, job_id, category_index)
                        diagnostic_lines.extend(
                            [
                                f"Скриншот: /jobs/{job_id}/debug/{screenshot_path.name}" if screenshot_path else "Скриншот: не сохранён",
                                f"HTML: /jobs/{job_id}/debug/{html_path.name}" if html_path else "HTML: не сохранён",
                                f"Текст страницы: {body_preview or '-'}",
                            ]
                        )
                    collected.extend(category_records)
                    search_page.remove_listener("response", capture.schedule)
            finally:
                await context.close()
                await browser.close()

    collected = deduplicate_company_dicts(collected)
    diagnostic = "\n".join(diagnostic_lines)
    if not collected:
        raise RuntimeError(
            "Яндекс не вернул ни одной распознаваемой карточки. Это не означает, что компаний нет.\n\n"
            + diagnostic
        )
    return collected, diagnostic


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
        if 'name="viewport"' not in lower and "name='viewport'" not in lower:
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


def draft_for(company: dict[str, Any], score: int | None, notes: str | None) -> str:
    source_name = "Яндекс Картах" if company.get("source") == "yandex" else "2ГИС"
    if company.get("website"):
        problem = notes or "есть возможности для улучшения сайта"
        return f"Здравствуйте! Нашёл компанию «{company['name']}» в {source_name} и посмотрел сайт. Заметил: {problem}. Могу предложить короткий план обновления без обязательств. Кому можно его отправить?"
    return f"Здравствуйте! Нашёл компанию «{company['name']}» в {source_name}. Отдельный сайт в карточке не указан. Могу бесплатно набросать структуру простой страницы с услугами и контактами. Кому можно отправить пример?"


async def telegram_notify(text: str) -> None:
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_ADMIN_CHAT_ID:
        return
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    try:
        async with httpx.AsyncClient(timeout=15) as client:
            await client.post(url, json={"chat_id": TELEGRAM_ADMIN_CHAT_ID, "text": text[:3900]})
    except Exception:
        pass


def deduplicate_and_filter(collected: list[dict[str, Any]], mode: str) -> tuple[list[dict[str, Any]], str]:
    seen: set[str] = set()
    unique_all: list[dict[str, Any]] = []
    for company in collected:
        key = company.get("source_id") or (str(company.get("name", "")).lower() + "|" + (company.get("phones") or [company.get("address", "")])[0])
        if key in seen:
            continue
        seen.add(key)
        unique_all.append(company)

    with_site = sum(bool(company.get("website")) for company in unique_all)
    with_contact = sum(bool(company.get("phones") or company.get("emails") or company.get("messengers")) for company in unique_all)
    no_site_with_contact = sum(not company.get("website") and bool(company.get("phones") or company.get("emails") or company.get("messengers")) for company in unique_all)

    if mode == "all":
        selected = unique_all
    elif mode == "audit_sites":
        selected = [company for company in unique_all if company.get("website")]
    else:
        selected = [company for company in unique_all if not company.get("website") and bool(company.get("phones") or company.get("emails") or company.get("messengers"))]

    diagnostic = (
        f"Получено карточек: {len(collected)}\n"
        f"Уникальных: {len(unique_all)}\n"
        f"С доступным сайтом: {with_site}\n"
        f"С телефоном/email/мессенджером: {with_contact}\n"
        f"Без сайта, но с контактом: {no_site_with_contact}\n"
        f"Прошло выбранный фильтр: {len(selected)}"
    )
    if not selected and unique_all:
        diagnostic += "\n\nСовет: запусти режим «Все карточки — для проверки», чтобы увидеть, какие поля реально удалось получить."
    return selected, diagnostic


async def run_job(job_id: int) -> None:
    with SessionLocal() as db:
        job = db.get(Job, job_id)
        if not job:
            return
        job.status = "running"
        job.error = None
        db.commit()
    try:
        with SessionLocal() as db:
            job = db.get(Job, job_id)
            if not job:
                return
            city = job.city
            categories = [x.strip() for x in job.categories.splitlines() if x.strip()]
            source, mode = split_job_mode(job.mode)
            limit = job.max_results

        collected: list[dict[str, Any]] = []
        if source == "twogis":
            if not TWOGIS_API_KEY:
                raise RuntimeError("Не задан TWOGIS_API_KEY в настройках Render")
            if not DATA_LICENSE_ACKNOWLEDGED:
                raise RuntimeError("DATA_LICENSE_ACKNOWLEDGED=false. Поставь true только если твой доступ 2ГИС разрешает сохранять данные")
            has_site = True if mode == "audit_sites" else False if mode == "no_site" else None
            for category in categories:
                collected.extend(await search_twogis(city, category, has_site=has_site, limit=limit))
        else:
            collected, source_diagnostic = await search_yandex_categories(job_id, city, categories, limit=min(30, limit))

        selected, filter_diagnostic = deduplicate_and_filter(collected, mode)
        diagnostic = filter_diagnostic if source == "twogis" else source_diagnostic + "\n\n--- Фильтр ---\n" + filter_diagnostic
        with SessionLocal() as db:
            job = db.get(Job, job_id)
            if not job:
                return
            job.found_count = len(collected)
            job.error = diagnostic
            db.commit()

        for company in selected:
            score: int | None = None
            notes: str | None = None
            if mode == "audit_sites" and company.get("website"):
                score, notes = await audit_site(str(company["website"]))
            lead = Lead(
                job_id=job_id,
                source_id=str(company.get("source_id") or ""),
                name=str(company.get("name") or "Без названия"),
                category=str(company.get("category") or ""),
                address=str(company.get("address") or ""),
                phones=", ".join(company.get("phones") or []),
                emails=", ".join(company.get("emails") or []),
                messengers=", ".join(company.get("messengers") or []),
                website=company.get("website"),
                rating=company.get("rating"),
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
                job.saved_count = len(selected)
                job.finished_at = datetime.now(timezone.utc)
                db.commit()
        await telegram_notify(
            f"✅ Задача #{job_id} завершена\nИсточник: {source_label(source)}\nПолучено карточек: {len(collected)}\nСохранено: {len(selected)}\nГород: {city}\nКатегории: {', '.join(categories)}"
        )
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
