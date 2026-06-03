# Сайт образовательного центра (ВКР)

Веб-сайт образовательного центра с каталогом курсов, преподавателями и
формой обратной связи. ВКР по теме «Создание сайта для образовательного
центра с формой обратной связи».

## Стек
- SQLite 3.35+ — хранилище курсов, преподавателей, заявок
- Python 3.10+ / FastAPI / Pydantic 2 — REST-API и валидация
- HTML5 / CSS3 / Vanilla JS (Fetch API) — клиент в стиле «Modern Edu»

## Запуск
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Инициализация БД
python -m backend.init_db

# Запуск сервера
uvicorn backend.main:app --port 8002 --reload
```

Откройте `http://localhost:8002/` в браузере.

## Сборка ПЗ
```bash
python scripts/generate_pz.py
```
Результат: `ПЗ_Сайт_Образовательного_Центра.docx`
