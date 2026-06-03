"""Управление подключением к SQLite и инициализация схемы."""
from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR.parent / "data" / "edu_center.db"
SCHEMA_PATH = BASE_DIR / "schema.sql"


def get_connection() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(DB_PATH, detect_types=sqlite3.PARSE_DECLTYPES)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON;")
    return conn


@contextmanager
def db_cursor(commit: bool = False):
    conn = get_connection()
    try:
        cur = conn.cursor()
        yield cur
        if commit:
            conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_schema() -> None:
    sql = SCHEMA_PATH.read_text(encoding="utf-8")
    conn = get_connection()
    try:
        conn.executescript(sql)
        conn.commit()
    finally:
        conn.close()


def seed_data() -> None:
    """Заполнение справочников и демонстрационных данных."""
    conn = get_connection()
    try:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM categories")
        if cur.fetchone()[0] > 0:
            return

        cur.executemany(
            "INSERT INTO categories (name, slug, icon, color, description) VALUES (?, ?, ?, ?, ?)",
            [
                ("Программирование", "it",        "💻", "#3B82F6",
                 "Веб-разработка, мобильные приложения, базы данных"),
                ("Иностранные языки","languages", "🌐", "#10B981",
                 "Английский, китайский, турецкий — все уровни"),
                ("Дизайн",           "design",    "🎨", "#F59E0B",
                 "UX/UI, графический дизайн, motion"),
                ("Подготовка к ОРТ", "ort",       "🎓", "#EF4444",
                 "Подготовка к Общереспубликанскому тестированию"),
            ],
        )

        cur.executemany(
            "INSERT INTO request_statuses (code, name, color) VALUES (?, ?, ?)",
            [
                ("new",         "Новая",       "#3B82F6"),
                ("in_progress", "В обработке", "#F59E0B"),
                ("enrolled",    "Записан",     "#10B981"),
                ("rejected",    "Отказ",       "#EF4444"),
            ],
        )

        cur.executemany(
            """INSERT INTO teachers
               (full_name, position, bio, photo_url, email, phone, experience_years)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            [
                ("Иванов Сергей Викторович",  "Senior Backend Engineer",
                 "Эксперт по Python и системам высокой нагрузки. Преподаёт с 2018 года.",
                 "/static/img/teacher1.svg", "ivanov@edu.kg", "+996700110011", 9),
                ("Петрова Анна Сергеевна",    "Методист, преподаватель английского",
                 "CELTA, IELTS 8.5. Готовит к международным экзаменам IELTS, TOEFL.",
                 "/static/img/teacher2.svg", "petrova@edu.kg", "+996700220022", 11),
                ("Бекова Айдай Темировна",    "UX/UI-дизайнер",
                 "Ex-Designer в продуктовых компаниях. Спикер регулярных дизайн-митапов.",
                 "/static/img/teacher3.svg", "bekova@edu.kg", "+996700330033", 6),
                ("Турсунов Алмаз Жакыпович",  "Frontend Engineer, JavaScript",
                 "Эксперт по React и современному JS. Автор курсов по фронтенду.",
                 "/static/img/teacher4.svg", "tursunov@edu.kg", "+996700440044", 7),
                ("Эркинбаева Жаныл Маратовна","Преподаватель математики",
                 "Готовит к ОРТ и олимпиадам. Сотни выпускников в ведущих вузах КР.",
                 "/static/img/teacher5.svg", "erkinbaeva@edu.kg", "+996700550055", 14),
            ],
        )

        cur.executemany(
            """INSERT INTO courses
               (category_id, title, short_summary, description, level, duration_hours, price)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            [
                (1, "Python Backend с нуля",
                    "Освойте бэкенд-разработку на Python и FastAPI с трудоустройством.",
                    "Полный курс от основ Python до production-ready REST API: SQL, ORM, тестирование, Docker.",
                    "beginner", 144, 28000),
                (1, "JavaScript и React",
                    "Современный фронтенд: ES2022, React, TypeScript, тесты.",
                    "Глубокий курс по экосистеме React с реальными проектами.",
                    "intermediate", 120, 26000),
                (2, "Английский: General + IELTS",
                    "Подготовка с уровня A2 до сдачи IELTS на 7.0+.",
                    "6-месячная программа: speaking-клубы, mock-тесты, индивидуальный план.",
                    "intermediate", 96, 18000),
                (2, "Китайский для начинающих",
                    "HSK 1 за 4 месяца с носителем языка.",
                    "Иероглифика, тоны, грамматика, разговорные клубы.",
                    "beginner", 72, 14000),
                (3, "UX/UI-дизайн в Figma",
                    "Стань UX/UI-дизайнером и собери портфолио из 4 проектов.",
                    "Дизайн-мышление, UX-исследования, прототипирование, дизайн-системы.",
                    "beginner", 84, 22000),
                (4, "ОРТ: математика + русский",
                    "Интенсивный курс подготовки к Общереспубликанскому тестированию.",
                    "5-месячная программа, разбор всех типовых заданий, пробные тестирования.",
                    "beginner", 160, 16000),
            ],
        )

        cur.executemany(
            "INSERT INTO course_teachers (course_id, teacher_id) VALUES (?, ?)",
            [
                (1, 1),
                (2, 4), (2, 1),
                (3, 2),
                (4, 2),
                (5, 3),
                (6, 5),
            ],
        )

        cur.executemany(
            """INSERT INTO feedback_requests
               (full_name, phone, email, course_id, message, status_id, source_page)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            [
                ("Жапаров Нуртилек", "+996555111222", "nur@example.com", 1,
                 "Здравствуйте! Хочу записаться на Python, есть ли вечерние группы?", 1, "/#courses"),
                ("Кадырова Айгерим", "+996555222333", "aigerim@example.com", 3,
                 "Какой уровень английского нужен для IELTS-группы?", 2, "/#feedback"),
                ("Турсунов Айбек",   "+996555333444", "aibek@example.com", None,
                 "Расскажите, какие у вас есть курсы для подростков?", 1, "/"),
                ("Эркинбаева Айдай", "+996555444555", "aiday@example.com", 5,
                 "Готова начать с октября, есть ли скидки?", 3, "/#courses"),
            ],
        )

        conn.commit()
    finally:
        conn.close()


if __name__ == "__main__":
    init_schema()
    seed_data()
    print(f"База данных создана: {DB_PATH}")
