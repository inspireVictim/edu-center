/* ============================================================
   EduCenter — клиентская логика
   ============================================================ */

const STATE = { categories: [], courses: [], activeCategory: '' };

const LEVEL_LABELS = {
    beginner:     'Для начинающих',
    intermediate: 'Средний уровень',
    advanced:     'Продвинутый',
};

document.addEventListener('DOMContentLoaded', () => {
    loadCategories();
    loadCourses();
    loadTeachers();
    initFeedbackForm();
    initModal();
    initSmoothNav();
});

/* ---------- API ---------- */

async function apiGet(path) {
    const res = await fetch(path);
    if (!res.ok) throw new Error('HTTP ' + res.status);
    return res.json();
}

async function apiPost(path, body) {
    const res = await fetch(path, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) {
        const detail = data.detail || `HTTP ${res.status}`;
        const err = new Error(typeof detail === 'string' ? detail : JSON.stringify(detail));
        err.status = res.status;
        err.payload = data;
        throw err;
    }
    return data;
}

/* ---------- Categories ---------- */

async function loadCategories() {
    try {
        STATE.categories = await apiGet('/api/categories');
        renderPills();
        populateCourseSelect();
    } catch (err) {
        console.warn('Не удалось загрузить категории:', err);
    }
}

function renderPills() {
    const wrap = document.getElementById('filter-pills');
    wrap.querySelectorAll('.pill[data-cat]:not([data-cat=""])').forEach(n => n.remove());

    STATE.categories.forEach(c => {
        const btn = document.createElement('button');
        btn.className = 'pill';
        btn.dataset.cat = c.slug;
        btn.textContent = `${c.icon} ${c.name}`;
        wrap.appendChild(btn);
    });

    wrap.addEventListener('click', (e) => {
        const pill = e.target.closest('.pill');
        if (!pill) return;
        wrap.querySelectorAll('.pill').forEach(p => p.classList.toggle('is-active', p === pill));
        STATE.activeCategory = pill.dataset.cat;
        renderCourses();
    }, { once: false });
}

function populateCourseSelect() {
    const sel = document.getElementById('form-course-select');
    const opts = ['<option value="">— выберите направление —</option>'];
    STATE.courses.forEach(c => {
        opts.push(`<option value="${c.id}">${c.category_icon} ${c.title}</option>`);
    });
    sel.innerHTML = opts.join('');
}

/* ---------- Courses ---------- */

async function loadCourses() {
    try {
        STATE.courses = await apiGet('/api/courses');
        renderCourses();
        populateCourseSelect();
    } catch (err) {
        document.getElementById('courses-grid').innerHTML =
            '<div class="skeleton">Не удалось загрузить курсы. Проверьте подключение.</div>';
        console.error(err);
    }
}

function renderCourses() {
    const grid = document.getElementById('courses-grid');
    const filtered = STATE.activeCategory
        ? STATE.courses.filter(c => {
            const cat = STATE.categories.find(cat => cat.id === c.category_id);
            return cat && cat.slug === STATE.activeCategory;
        })
        : STATE.courses;

    if (filtered.length === 0) {
        grid.innerHTML = '<div class="skeleton">В этом направлении пока нет открытых наборов.</div>';
        return;
    }

    grid.innerHTML = filtered.map(c => courseCard(c)).join('');
    grid.querySelectorAll('[data-enroll]').forEach(btn => {
        btn.addEventListener('click', () => {
            document.getElementById('form-course-select').value = btn.dataset.enroll;
            document.getElementById('feedback').scrollIntoView({ behavior: 'smooth', block: 'start' });
            setTimeout(() => document.querySelector('#feedback-form [name="full_name"]').focus(), 600);
        });
    });
}

function courseCard(c) {
    const teachers = (c.teachers || []).map(t => t.full_name.split(' ').slice(0, 2).join(' ')).join(', ');
    return `
        <article class="course-card">
            <span class="course-card__badge" style="background: ${c.category_color}1A; color: ${c.category_color}">
                ${c.category_icon} ${c.category_name}
            </span>
            <h3 class="course-card__title">${c.title}</h3>
            <p class="course-card__summary">${c.short_summary}</p>
            <div class="course-card__meta">
                <div class="course-card__meta-item">
                    <strong>${c.duration_hours} ч.</strong>
                    Длительность
                </div>
                <div class="course-card__meta-item">
                    <strong>${LEVEL_LABELS[c.level] || c.level}</strong>
                    Уровень
                </div>
                ${teachers ? `<div class="course-card__meta-item" style="grid-column: 1 / -1">
                    <strong>${teachers}</strong>
                    Преподаватели
                </div>` : ''}
            </div>
            <div class="course-card__footer">
                <div class="course-card__price">
                    ${(+c.price).toLocaleString('ru-RU')} с
                    <small> / весь курс</small>
                </div>
                <button class="btn btn--ghost btn--small" data-enroll="${c.id}">
                    Записаться
                </button>
            </div>
        </article>
    `;
}

/* ---------- Teachers ---------- */

async function loadTeachers() {
    try {
        const teachers = await apiGet('/api/teachers');
        const grid = document.getElementById('teachers-grid');
        if (teachers.length === 0) {
            grid.innerHTML = '<div class="skeleton">Состав уточняется.</div>';
            return;
        }
        grid.innerHTML = teachers.map(t => `
            <article class="teacher">
                <div class="teacher__avatar">${initials(t.full_name)}</div>
                <div class="teacher__name">${t.full_name}</div>
                <div class="teacher__role">${t.position}</div>
                <p class="teacher__bio">${t.bio || ''}</p>
                <span class="teacher__years">${t.experience_years}+ лет опыта</span>
            </article>
        `).join('');
    } catch (err) {
        document.getElementById('teachers-grid').innerHTML =
            '<div class="skeleton">Не удалось загрузить состав.</div>';
    }
}

function initials(fio) {
    return fio.split(' ').filter(Boolean).slice(0, 2).map(p => p[0]).join('').toUpperCase();
}

/* ---------- Feedback form ---------- */

function initFeedbackForm() {
    const form = document.getElementById('feedback-form');
    const submitBtn = document.getElementById('submit-btn');

    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        clearErrors(form);

        const fd = new FormData(form);
        const data = {
            full_name:   (fd.get('full_name') || '').trim(),
            phone:       (fd.get('phone') || '').trim(),
            email:       (fd.get('email') || '').trim(),
            course_id:   fd.get('course_id') ? Number(fd.get('course_id')) : null,
            message:     (fd.get('message') || '').trim() || null,
            website:     fd.get('website') || '',
            source_page: window.location.pathname + window.location.hash,
        };

        // Клиентская валидация
        const errors = validateClient(data);
        if (Object.keys(errors).length > 0) {
            showErrors(form, errors);
            return;
        }

        submitBtn.classList.add('is-loading');
        submitBtn.disabled = true;
        try {
            const res = await apiPost('/api/feedback', data);
            openSuccessModal(res.message || 'Заявка принята. Скоро свяжемся.');
            form.reset();
        } catch (err) {
            handleServerError(form, err);
        } finally {
            submitBtn.classList.remove('is-loading');
            submitBtn.disabled = false;
        }
    });
}

function validateClient(d) {
    const errors = {};
    if (!d.full_name || d.full_name.length < 2) errors.full_name = 'Укажите имя (минимум 2 символа)';

    const phoneOk = /^\+[0-9\s\-()]{9,19}$/.test(d.phone);
    if (!d.phone)        errors.phone = 'Укажите телефон';
    else if (!phoneOk)   errors.phone = 'Телефон должен начинаться с «+» и содержать цифры';

    const emailOk = /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email);
    if (!d.email)        errors.email = 'Укажите email';
    else if (!emailOk)   errors.email = 'Введите корректный email';

    return errors;
}

function clearErrors(form) {
    form.querySelectorAll('.field.is-invalid').forEach(f => f.classList.remove('is-invalid'));
    form.querySelectorAll('.field__error').forEach(e => e.textContent = '');
}

function showErrors(form, errors) {
    Object.entries(errors).forEach(([name, msg]) => {
        const input = form.querySelector(`[name="${name}"]`);
        const field = input && input.closest('.field');
        if (field) {
            field.classList.add('is-invalid');
            const err = field.querySelector('.field__error');
            if (err) err.textContent = msg;
        }
    });
    const first = form.querySelector('.field.is-invalid input, .field.is-invalid select, .field.is-invalid textarea');
    if (first) first.focus();
}

function handleServerError(form, err) {
    // Pydantic-ошибки возвращаются массивом
    if (err.payload && Array.isArray(err.payload.detail)) {
        const errs = {};
        err.payload.detail.forEach(d => {
            const field = d.loc && d.loc[d.loc.length - 1];
            if (field) errs[field] = d.msg.replace(/^Value error,?\s*/i, '');
        });
        showErrors(form, errs);
        return;
    }
    // Антиспам / прочее — показываем под общей формой
    openSuccessModal(err.message || 'Не удалось отправить заявку. Попробуйте позже.', /*isError=*/ true);
}

/* ---------- Modal ---------- */

function initModal() {
    document.getElementById('modal-close').addEventListener('click', closeModal);
    document.getElementById('success-modal').addEventListener('click', (e) => {
        if (e.target.id === 'success-modal') closeModal();
    });
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') closeModal();
    });
}

function openSuccessModal(message, isError = false) {
    document.getElementById('success-message').textContent = message;
    const icon = document.querySelector('.modal__icon');
    if (isError) {
        icon.textContent = '!';
        icon.style.background = '#FEF2F2';
        icon.style.color = '#B91C1C';
    } else {
        icon.textContent = '✓';
        icon.style.background = '';
        icon.style.color = '';
    }
    document.getElementById('success-modal').hidden = false;
}

function closeModal() {
    document.getElementById('success-modal').hidden = true;
}

/* ---------- Smooth nav ---------- */

function initSmoothNav() {
    document.querySelectorAll('a[href^="#"]').forEach(a => {
        a.addEventListener('click', (e) => {
            const target = document.querySelector(a.getAttribute('href'));
            if (target) {
                e.preventDefault();
                target.scrollIntoView({ behavior: 'smooth', block: 'start' });
            }
        });
    });
}
