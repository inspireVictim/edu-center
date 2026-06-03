"""Обработка формы обратной связи.

Алгоритм POST /api/feedback:
    1. Honeypot-проверка: поле website должно быть пустым; если заполнено — бот.
    2. Pydantic-валидация (см. FeedbackIn): длина имени, формат телефона/email.
    3. Проверка существования курса (если course_id передан).
    4. Запись в feedback_requests со status='new'.
    5. Триггер БД блокирует повторный INSERT с того же email в течение 60 секунд.
    6. Возврат успешного JSON-ответа.
"""
from __future__ import annotations

from fastapi import APIRouter, HTTPException, Request

from ..database import db_cursor
from ..schemas import FeedbackAck, FeedbackIn, FeedbackOut

router = APIRouter(prefix="/api/feedback", tags=["feedback"])


@router.post("", response_model=FeedbackAck, status_code=201)
def submit_feedback(payload: FeedbackIn, request: Request):
    # 1) Honeypot: бот заполнит скрытое поле "website"
    if payload.website:
        # Возвращаем 200 как будто всё ок, чтобы не подсказывать боту, что обнаружен
        raise HTTPException(status_code=400, detail="Заявка отклонена системой защиты")

    # 2) Проверка существования курса (если указан)
    if payload.course_id is not None:
        with db_cursor() as cur:
            cur.execute("SELECT id FROM courses WHERE id = ? AND is_active = 1",
                        (payload.course_id,))
            if cur.fetchone() is None:
                raise HTTPException(400, "Указанный курс не найден или неактивен")

    # 3) Получаем id статуса "Новая"
    with db_cursor() as cur:
        cur.execute("SELECT id FROM request_statuses WHERE code = 'new'")
        row = cur.fetchone()
        if row is None:
            raise HTTPException(500, "Справочник статусов не инициализирован")
        new_status_id = row["id"]

    # 4) Технические поля
    ip = request.client.host if request.client else None
    ua = request.headers.get("user-agent")

    # 5) Запись с обработкой ошибки триггера антиспама
    try:
        with db_cursor(commit=True) as cur:
            cur.execute(
                """INSERT INTO feedback_requests
                   (full_name, phone, email, course_id, message, status_id,
                    source_page, ip_address, user_agent)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (payload.full_name, payload.phone, payload.email,
                 payload.course_id, payload.message, new_status_id,
                 payload.source_page, ip, ua),
            )
            new_id = cur.lastrowid
    except Exception as exc:
        msg = str(exc)
        if "Повторная заявка" in msg:
            raise HTTPException(429, "Повторная заявка с этого email отклонена. Подождите минуту.")
        raise HTTPException(400, f"Не удалось сохранить заявку: {exc}")

    return FeedbackAck(request_id=new_id)


@router.get("", response_model=list[FeedbackOut])
def list_feedback(limit: int = 50):
    """Просмотр последних заявок (для админ-панели; в учебной версии без авторизации)."""
    with db_cursor() as cur:
        cur.execute(
            """SELECT
                   f.id, f.full_name, f.phone, f.email,
                   f.course_id, c.title AS course_title,
                   f.message,
                   s.code AS status_code, s.name AS status_name, s.color AS status_color,
                   f.created_at
               FROM feedback_requests f
               JOIN request_statuses s ON s.id = f.status_id
               LEFT JOIN courses c ON c.id = f.course_id
               ORDER BY f.created_at DESC
               LIMIT ?""",
            (limit,),
        )
        return [dict(r) for r in cur.fetchall()]
