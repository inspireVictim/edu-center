-- =====================================================
-- БАЗА ДАННЫХ "Образовательный центр"
-- СУБД: SQLite 3.35+    Кодировка: UTF-8
-- Схема в 3НФ
-- =====================================================

PRAGMA foreign_keys = ON;
PRAGMA journal_mode = WAL;
PRAGMA encoding     = 'UTF-8';

-- -------- 1. СПРАВОЧНИКИ -----------------------------

CREATE TABLE IF NOT EXISTS categories (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT    NOT NULL UNIQUE,
    slug        TEXT    NOT NULL UNIQUE,
    icon        TEXT    NOT NULL DEFAULT '📘',
    color       TEXT    NOT NULL DEFAULT '#3B82F6',
    description TEXT
);

CREATE TABLE IF NOT EXISTS teachers (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    full_name        TEXT    NOT NULL,
    position         TEXT    NOT NULL,
    bio              TEXT,
    photo_url        TEXT,
    email            TEXT    UNIQUE
                     CHECK (email IS NULL OR
                            email LIKE '_%@_%._%'),
    phone            TEXT    CHECK (phone IS NULL OR
                                    phone GLOB '+[0-9]*'),
    experience_years INTEGER NOT NULL DEFAULT 0
                     CHECK (experience_years >= 0),
    is_active        INTEGER NOT NULL DEFAULT 1
                     CHECK (is_active IN (0, 1))
);

CREATE TABLE IF NOT EXISTS request_statuses (
    id    INTEGER PRIMARY KEY AUTOINCREMENT,
    code  TEXT    NOT NULL UNIQUE,
    name  TEXT    NOT NULL,
    color TEXT    NOT NULL DEFAULT '#64748B'
);

-- -------- 2. ОСНОВНЫЕ СУЩНОСТИ -----------------------

CREATE TABLE IF NOT EXISTS courses (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    category_id     INTEGER NOT NULL,
    title           TEXT    NOT NULL UNIQUE,
    short_summary   TEXT    NOT NULL,
    description     TEXT,
    level           TEXT    NOT NULL DEFAULT 'beginner'
                    CHECK (level IN ('beginner','intermediate','advanced')),
    duration_hours  INTEGER NOT NULL CHECK (duration_hours > 0),
    price           NUMERIC(10, 2) NOT NULL CHECK (price >= 0),
    is_active       INTEGER NOT NULL DEFAULT 1
                    CHECK (is_active IN (0, 1)),
    created_at      TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (category_id) REFERENCES categories (id) ON DELETE RESTRICT
);

CREATE TABLE IF NOT EXISTS course_teachers (
    course_id  INTEGER NOT NULL,
    teacher_id INTEGER NOT NULL,
    PRIMARY KEY (course_id, teacher_id),
    FOREIGN KEY (course_id)  REFERENCES courses  (id) ON DELETE CASCADE,
    FOREIGN KEY (teacher_id) REFERENCES teachers (id) ON DELETE CASCADE
);

-- -------- 3. ЗАЯВКИ С ФОРМЫ ОБРАТНОЙ СВЯЗИ -----------

CREATE TABLE IF NOT EXISTS feedback_requests (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    full_name       TEXT    NOT NULL CHECK (length(trim(full_name)) >= 2),
    phone           TEXT    NOT NULL
                    CHECK (phone GLOB '+[0-9]*'
                           AND length(phone) BETWEEN 10 AND 20),
    email           TEXT    NOT NULL
                    CHECK (email LIKE '_%@_%._%'),
    course_id       INTEGER,
    message         TEXT,
    status_id       INTEGER NOT NULL,
    source_page     TEXT,
    ip_address      TEXT,
    user_agent      TEXT,
    created_at      TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    processed_at    TIMESTAMP,
    processed_by    TEXT,
    admin_note      TEXT,
    CHECK (processed_at IS NULL OR processed_at >= created_at),
    FOREIGN KEY (course_id) REFERENCES courses          (id) ON DELETE SET NULL,
    FOREIGN KEY (status_id) REFERENCES request_statuses (id) ON DELETE RESTRICT
);

-- -------- 4. ИНДЕКСЫ ---------------------------------

CREATE INDEX IF NOT EXISTS idx_courses_category  ON courses           (category_id);
CREATE INDEX IF NOT EXISTS idx_requests_status   ON feedback_requests (status_id);
CREATE INDEX IF NOT EXISTS idx_requests_created  ON feedback_requests (created_at DESC);
CREATE INDEX IF NOT EXISTS idx_requests_email    ON feedback_requests (email);
CREATE INDEX IF NOT EXISTS idx_requests_phone    ON feedback_requests (phone);

-- -------- 5. ТРИГГЕР АНТИСПАМА -----------------------

DROP TRIGGER IF EXISTS trg_feedback_antispam;
CREATE TRIGGER trg_feedback_antispam
BEFORE INSERT ON feedback_requests
FOR EACH ROW
BEGIN
    SELECT CASE
        WHEN (
            SELECT COUNT(*) FROM feedback_requests
             WHERE email = NEW.email
               AND created_at > datetime('now', '-60 seconds')
        ) > 0
        THEN RAISE(ABORT, 'Повторная заявка с этого email слишком быстро')
    END;
END;
