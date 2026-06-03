"""Каталог курсов и направлений."""
from __future__ import annotations

from typing import Optional

from fastapi import APIRouter

from ..database import db_cursor
from ..schemas import CategoryOut, CourseOut, TeacherBrief

router = APIRouter(prefix="/api", tags=["courses"])


@router.get("/categories", response_model=list[CategoryOut])
def list_categories():
    with db_cursor() as cur:
        cur.execute("SELECT * FROM categories ORDER BY name")
        return [dict(r) for r in cur.fetchall()]


@router.get("/courses", response_model=list[CourseOut])
def list_courses(category: Optional[str] = None):
    sql = """
        SELECT
            c.id, c.category_id, cat.name AS category_name,
            cat.icon AS category_icon, cat.color AS category_color,
            c.title, c.short_summary, c.description,
            c.level, c.duration_hours, c.price
        FROM courses c
        JOIN categories cat ON cat.id = c.category_id
        WHERE c.is_active = 1
    """
    params: list = []
    if category:
        sql += " AND cat.slug = ?"
        params.append(category)
    sql += " ORDER BY cat.name, c.title"

    with db_cursor() as cur:
        cur.execute(sql, params)
        courses = [dict(r) for r in cur.fetchall()]
        for c in courses:
            cur.execute(
                """SELECT t.id, t.full_name, t.position
                     FROM course_teachers ct
                     JOIN teachers t ON t.id = ct.teacher_id
                    WHERE ct.course_id = ?
                    ORDER BY t.full_name""",
                (c["id"],),
            )
            c["teachers"] = [TeacherBrief(**dict(r)).model_dump() for r in cur.fetchall()]
        return courses
