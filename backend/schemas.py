"""Pydantic-схемы валидации."""
from __future__ import annotations

import datetime as dt
import re
from typing import Optional, List

from pydantic import BaseModel, EmailStr, Field, field_validator

PHONE_RE = re.compile(r"^\+[0-9\s\-()]{9,19}$")


# ---------- Категории ----------

class CategoryOut(BaseModel):
    id: int
    name: str
    slug: str
    icon: str
    color: str
    description: Optional[str]


# ---------- Преподаватели ----------

class TeacherBrief(BaseModel):
    id: int
    full_name: str
    position: str


class TeacherOut(BaseModel):
    id: int
    full_name: str
    position: str
    bio: Optional[str]
    photo_url: Optional[str]
    email: Optional[str]
    phone: Optional[str]
    experience_years: int


# ---------- Курсы ----------

class CourseOut(BaseModel):
    id: int
    category_id: int
    category_name: str
    category_icon: str
    category_color: str
    title: str
    short_summary: str
    description: Optional[str]
    level: str
    duration_hours: int
    price: float
    teachers: List[TeacherBrief]


# ---------- Заявка с формы ----------

class FeedbackIn(BaseModel):
    full_name: str = Field(min_length=2, max_length=120)
    phone: str = Field(min_length=10, max_length=20)
    email: EmailStr
    course_id: Optional[int] = None
    message: Optional[str] = Field(default=None, max_length=2000)
    source_page: Optional[str] = Field(default=None, max_length=200)
    # honeypot — невидимое поле, заполненное → бот
    website: Optional[str] = Field(default=None, max_length=200)

    @field_validator("full_name")
    @classmethod
    def full_name_clean(cls, v: str) -> str:
        v = " ".join(v.split())
        if not v or len(v) < 2:
            raise ValueError("Имя слишком короткое")
        return v

    @field_validator("phone")
    @classmethod
    def phone_format(cls, v: str) -> str:
        v = v.strip()
        if not PHONE_RE.match(v):
            raise ValueError("Телефон должен начинаться с '+' и содержать 10–20 цифр")
        return v

    @field_validator("message")
    @classmethod
    def message_clean(cls, v: Optional[str]) -> Optional[str]:
        if v is None:
            return v
        v = v.strip()
        return v or None


class FeedbackOut(BaseModel):
    id: int
    full_name: str
    phone: str
    email: str
    course_id: Optional[int]
    course_title: Optional[str]
    message: Optional[str]
    status_code: str
    status_name: str
    status_color: str
    created_at: dt.datetime


class FeedbackAck(BaseModel):
    """Ответ клиенту после успешной отправки формы."""
    success: bool = True
    request_id: int
    message: str = "Заявка принята. Менеджер свяжется с вами в течение рабочего дня."
