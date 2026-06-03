"""Преподавательский состав."""
from __future__ import annotations

from fastapi import APIRouter

from ..database import db_cursor
from ..schemas import TeacherOut

router = APIRouter(prefix="/api/teachers", tags=["teachers"])


@router.get("", response_model=list[TeacherOut])
def list_teachers():
    with db_cursor() as cur:
        cur.execute(
            "SELECT * FROM teachers WHERE is_active = 1 ORDER BY full_name"
        )
        return [dict(r) for r in cur.fetchall()]
