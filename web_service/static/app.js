const state = {
  dashboard: null,
  currentMonthFilter: 'latest',
  compareMonthFilter: 'previous',
};

const elements = {
  summary: document.getElementById('summary'),
  heavyCards: document.getElementById('heavyCards'),
  dayViews: document.getElementById('dayViews'),
  fullReport: document.getElementById('fullReport'),
  currentMonthFilter: document.getElementById('currentMonthFilter'),
  compareMonthFilter: document.getElementById('compareMonthFilter'),
  expandAllButton: document.getElementById('expandAllButton'),
  collapseAllButton: document.getElementById('collapseAllButton'),
  refreshButton: document.getElementById('refreshButton'),
};

elements.currentMonthFilter.addEventListener('change', (event) => {
  state.currentMonthFilter = event.target.value;
  render();
});

elements.compareMonthFilter.addEventListener('change', (event) => {
  state.compareMonthFilter = event.target.value;
  render();
});

elements.expandAllButton.addEventListener('click', () => {
  document.querySelectorAll('.report-details').forEach((item) => {
    item.open = true;
  });
});

elements.collapseAllButton.addEventListener('click', () => {
  document.querySelectorAll('.report-details').forEach((item) => {
    item.open = false;
  });
});

elements.refreshButton.addEventListener('click', async () => {
  await loadDashboard();
});

function formatNumber(value) {
  if (value === null || value === undefined || value === '') return '—';
  if (typeof value !== 'number') return String(value);
  return Number.isInteger(value) ? String(value) : value.toFixed(1).replace('.0', '');
}

function formatDelta(value) {
  if (value === null || value === undefined) return '—';
  if (value === 0) return '0';
  const rounded = typeof value === 'number' ? Number(value.toFixed(1)) : value;
  return `${rounded > 0 ? '+' : ''}${rounded}`;
}

function formatSets(sets) {
  const parts = sets
    .filter((item) => item.weight !== null && item.reps !== null)
    .map((item) => `${formatNumber(item.weight)}x${formatNumber(item.reps)}`);
  return parts.length ? parts.join(', ') : '—';
}

function loggedEntries(entries) {
  return entries.filter((entry) => entry.logged);
}

function getAllLoggedMonths(reports) {
  const months = new Set();
  Object.values(reports).forEach((sheet) => {
    sheet.exercises.forEach((exercise) => {
      loggedEntries(exercise.entries).forEach((entry) => {
        if (entry.month !== null && entry.month !== undefined && entry.month !== '') {
          months.add(Number(entry.month));
        }
      });
    });
  });
  return [...months].sort((a, b) => a - b);
}

function getMonthLabel(month) {
  if (!state.dashboard || !month) return `Месяц ${month}`;
  for (const sheet of Object.values(state.dashboard.reports)) {
    for (const exercise of sheet.exercises) {
      const entry = loggedEntries(exercise.entries).find((item) => Number(item.month) === Number(month));
      if (entry?.date) {
        return `Месяц ${month} · ${entry.date}`;
      }
    }
  }
  return `Месяц ${month}`;
}

function getSelectedCurrentMonth() {
  const months = getAllLoggedMonths(state.dashboard.reports);
  if (!months.length) return null;
  if (state.currentMonthFilter === 'latest') return months.at(-1);
  return Number(state.currentMonthFilter);
}

function getSelectedCompareMonth(currentMonth) {
  const months = getAllLoggedMonths(state.dashboard.reports);
  if (!months.length || !currentMonth) return null;

  if (state.compareMonthFilter === 'previous') {
    const previous = months.filter((month) => month < currentMonth);
    return previous.length ? previous.at(-1) : null;
  }

  if (state.compareMonthFilter === 'first') {
    return months[0] === currentMonth ? null : months[0];
  }

  const value = Number(state.compareMonthFilter);
  return value === currentMonth ? null : value;
}

function getEntryByMonth(exercise, month) {
  return loggedEntries(exercise.entries).find((entry) => Number(entry.month) === Number(month)) || null;
}

function percentDelta(currentValue, previousValue) {
  if (!currentValue || !previousValue) return null;
  return ((currentValue - previousValue) / previousValue) * 100;
}

function getProgressStatus(currentEntry, previousEntry) {
  if (!currentEntry) return { label: 'Нет отчета', tone: 'flat', score: 0 };
  if (!previousEntry) return { label: 'Первый замер', tone: 'flat', score: 0 };

  const currentStrength = currentEntry.estimated_1rm || 0;
  const previousStrength = previousEntry.estimated_1rm || 0;
  const diff = currentStrength - previousStrength;

  if (diff > 0.7) return { label: 'Прогресс', tone: 'up', score: 1 };
  if (diff < -0.7) return { label: 'Просадка', tone: 'down', score: -1 };
  return { label: 'Почти без изменений', tone: 'flat', score: 0 };
}

function statusChip(status) {
  return `<span class="status-chip status-chip--${status.tone}">${status.label}</span>`;
}

function createStatCard(label, value, note = '') {
  const node = document.createElement('article');
  node.className = 'stat-card';
  node.innerHTML = `
    <div class="stat-card__label">${label}</div>
    <div class="stat-card__value">${value}</div>
    <div class="muted">${note}</div>
  `;
  return node;
}

function buildSparkline(values) {
  if (values.length < 2) {
    return '<div class="empty-state">Пока есть только одна контрольная точка.</div>';
  }

  const width = 320;
  const height = 90;
  const padding = 12;
  const min = Math.min(...values);
  const max = Math.max(...values);
  const spread = max - min || 1;

  const points = values.map((value, index) => {
    const x = padding + (index * (width - padding * 2)) / (values.length - 1);
    const y = height - padding - ((value - min) / spread) * (height - padding * 2);
    return [x, y];
  });

  const path = points.map(([x, y], index) => `${index === 0 ? 'M' : 'L'}${x},${y}`).join(' ');
  const circles = points.map(([x, y]) => `<circle cx="${x}" cy="${y}" r="4"></circle>`).join('');

  return `
    <svg class="trend-line" viewBox="0 0 320 90" preserveAspectRatio="none">
      <path class="line" d="${path}"></path>
      ${circles}
    </svg>
  `;
}

function visibleExercises(sheet) {
  return sheet.exercises;
}

function populateMonthControls() {
  const months = getAllLoggedMonths(state.dashboard.reports);
  const previousCurrent = state.currentMonthFilter;
  const previousCompare = state.compareMonthFilter;

  elements.currentMonthFilter.innerHTML = '<option value="latest">Последний месяц</option>';
  months.forEach((month) => {
    const option = document.createElement('option');
    option.value = String(month);
    option.textContent = getMonthLabel(month);
    elements.currentMonthFilter.append(option);
  });

  const validCurrent = ['latest', ...months.map(String)];
  state.currentMonthFilter = validCurrent.includes(previousCurrent) ? previousCurrent : 'latest';
  elements.currentMonthFilter.value = state.currentMonthFilter;

  const currentMonth = getSelectedCurrentMonth();
  elements.compareMonthFilter.innerHTML = '<option value="previous">Предыдущим месяцем</option><option value="first">Первым месяцем</option>';
  months
    .filter((month) => month !== currentMonth)
    .forEach((month) => {
      const option = document.createElement('option');
      option.value = String(month);
      option.textContent = getMonthLabel(month);
      elements.compareMonthFilter.append(option);
    });

  const validCompare = ['previous', 'first', ...months.filter((month) => month !== currentMonth).map(String)];
  state.compareMonthFilter = validCompare.includes(previousCompare) ? previousCompare : 'previous';
  elements.compareMonthFilter.value = state.compareMonthFilter;
}

function renderSummary() {
  elements.summary.innerHTML = '';

  const workbook = state.dashboard.workbook;
  const currentMonth = getSelectedCurrentMonth();
  const compareMonth = getSelectedCompareMonth(currentMonth);
  const sheets = Object.values(state.dashboard.reports).map((sheet) => ({ ...sheet, exercises: visibleExercises(sheet) }));

  let up = 0;
  let flat = 0;
  let down = 0;
  let visibleCount = 0;

  sheets.forEach((sheet) => {
    sheet.exercises.forEach((exercise) => {
      const currentEntry = getEntryByMonth(exercise, currentMonth);
      if (!currentEntry) return;
      visibleCount += 1;
      const status = getProgressStatus(currentEntry, getEntryByMonth(exercise, compareMonth));
      if (status.score > 0) up += 1;
      else if (status.score < 0) down += 1;
      else flat += 1;
    });
  });

  elements.summary.append(
    createStatCard('Текущий отчет', currentMonth ? getMonthLabel(currentMonth) : '—', compareMonth ? `Сравнение с ${getMonthLabel(compareMonth)}` : 'Прошлого сравнения нет'),
    createStatCard('Упражнений в выборке', String(visibleCount), 'Все упражнения из отчетов'),
    createStatCard('Выросло', String(up), `Без изменений: ${flat}`),
    createStatCard('Просело', String(down), `Файл обновлен ${new Date(workbook.updated_at).toLocaleString('ru-RU')}`),
  );
}

function renderHeavyCards() {
  elements.heavyCards.innerHTML = '';
  const currentMonth = getSelectedCurrentMonth();
  const compareMonth = getSelectedCompareMonth(currentMonth);
  const heavyExercises = Object.values(state.dashboard.reports)
    .map((sheet) => ({
      day: sheet.title,
      exercise: sheet.exercises[0],
    }))
    .filter((item) => item.exercise);

  heavyExercises.forEach(({ day, exercise }) => {
    const allEntries = loggedEntries(exercise.entries);
    const currentEntry = getEntryByMonth(exercise, currentMonth);
    const previousEntry = getEntryByMonth(exercise, compareMonth);
    const status = getProgressStatus(currentEntry, previousEntry);

    const card = document.createElement('article');
    card.className = 'heavy-card';

    if (!currentEntry) {
      card.innerHTML = `
        <div class="heavy-card__headline">
          <div>
            <div class="heavy-card__label">${day}</div>
            <div class="heavy-card__title">${exercise.title}</div>
          </div>
          ${statusChip({ label: 'Нет данных', tone: 'flat' })}
        </div>
        <div class="empty-state">Для выбранного месяца по этому движению пока нет отчета.</div>
      `;
      elements.heavyCards.append(card);
      return;
    }

    const trendValues = allEntries
      .map((entry) => entry.estimated_1rm)
      .filter((value) => typeof value === 'number');

    const strengthDelta =
      currentEntry.estimated_1rm && previousEntry?.estimated_1rm
        ? currentEntry.estimated_1rm - previousEntry.estimated_1rm
        : null;

    const strengthPercent = percentDelta(currentEntry.estimated_1rm, previousEntry?.estimated_1rm);

    card.innerHTML = `
      <div class="heavy-card__headline">
        <div>
          <div class="heavy-card__label">${day}</div>
          <div class="heavy-card__title">${exercise.title}</div>
        </div>
        ${statusChip(status)}
      </div>
      <div class="compare-strip">
        <div class="compare-box">
          <div class="muted">Сравниваем с</div>
          <strong>${previousEntry ? previousEntry.summary || formatSets(previousEntry.sets) : '—'}</strong>
          <span class="muted">${previousEntry ? getMonthLabel(previousEntry.month) : 'Нет базового отчета'}</span>
        </div>
        <div class="compare-box compare-box--accent">
          <div class="muted">Текущий отчет</div>
          <strong>${currentEntry.summary || formatSets(currentEntry.sets)}</strong>
          <span class="muted">${getMonthLabel(currentEntry.month)}</span>
        </div>
      </div>
      ${buildSparkline(trendValues)}
      <div class="heavy-card__meta">
        <div class="meta-box">
          <div class="muted">Лучший e1RM</div>
          <strong>${formatNumber(currentEntry.estimated_1rm)}</strong>
        </div>
        <div class="meta-box">
          <div class="muted">Дельта силы</div>
          <strong>${formatDelta(strengthDelta)}</strong>
        </div>
        <div class="meta-box">
          <div class="muted">Дельта силы %</div>
          <strong>${strengthPercent === null ? '—' : `${formatDelta(strengthPercent)}%`}</strong>
        </div>
      </div>
    `;
    elements.heavyCards.append(card);
  });

  if (!heavyExercises.length) {
    elements.heavyCards.append(createEmpty('По текущим фильтрам тяжелые упражнения не найдены.'));
  }
}

function createEmpty(text) {
  const node = document.createElement('div');
  node.className = 'empty-state';
  node.textContent = text;
  return node;
}

function makeDetailedSetTable(entries) {
  const table = document.createElement('table');
  table.className = 'entry-table';
  table.innerHTML = `
    <tr>
      <th>Месяц</th>
      <th>Подход 1</th>
      <th>Подход 2</th>
      <th>Подход 3</th>
      <th>Подход 4</th>
      <th>Итог</th>
      <th>Заметки</th>
    </tr>
  `;

  entries.forEach((entry) => {
    const row = document.createElement('tr');
    const setCells = entry.sets.map((setItem) =>
      setItem.weight !== null && setItem.reps !== null
        ? `${formatNumber(setItem.weight)}x${formatNumber(setItem.reps)}${setItem.rir !== null && setItem.rir !== undefined ? `<br><span class="muted">RIR ${formatNumber(setItem.rir)}</span>` : ''}`
        : '—',
    );
    row.innerHTML = `
      <td>${getMonthLabel(entry.month)}</td>
      <td>${setCells[0] || '—'}</td>
      <td>${setCells[1] || '—'}</td>
      <td>${setCells[2] || '—'}</td>
      <td>${setCells[3] || '—'}</td>
      <td>${entry.summary || '—'}</td>
      <td>${entry.notes || '—'}</td>
    `;
    table.append(row);
  });

  return table;
}

function makeProgressTable(sheet, currentMonth, compareMonth) {
  const table = document.createElement('table');
  table.className = 'entry-table';
  table.innerHTML = `
    <tr>
      <th>Упражнение</th>
      <th>Статус</th>
      <th>База</th>
      <th>Текущий</th>
      <th>e1RM</th>
    </tr>
  `;

  visibleExercises(sheet).forEach((exercise) => {
    const currentEntry = getEntryByMonth(exercise, currentMonth);
    const previousEntry = getEntryByMonth(exercise, compareMonth);
    const status = getProgressStatus(currentEntry, previousEntry);

    const strengthDelta =
      currentEntry?.estimated_1rm && previousEntry?.estimated_1rm
        ? currentEntry.estimated_1rm - previousEntry.estimated_1rm
        : null;

    const row = document.createElement('tr');
    row.innerHTML = `
      <td>${exercise.title}</td>
      <td>${statusChip(status)}</td>
      <td>${previousEntry ? previousEntry.summary || formatSets(previousEntry.sets) : '—'}</td>
      <td>${currentEntry ? currentEntry.summary || formatSets(currentEntry.sets) : '—'}</td>
      <td>${currentEntry ? formatNumber(currentEntry.estimated_1rm) : '—'}${strengthDelta !== null ? `<br><span class="muted">${formatDelta(strengthDelta)}</span>` : ''}</td>
    `;
    table.append(row);
  });

  return table;
}

function renderDayViews() {
  elements.dayViews.innerHTML = '';
  const currentMonth = getSelectedCurrentMonth();
  const compareMonth = getSelectedCompareMonth(currentMonth);

  Object.values(state.dashboard.reports)
    .forEach((sheet) => {
      if (!visibleExercises(sheet).length) return;
      const wrapper = document.createElement('div');
      wrapper.className = 'table-card';
      wrapper.innerHTML = `
        <div class="table-card__header">
          <h2>${sheet.title}</h2>
          <p class="table-note">Сравнение ${getMonthLabel(currentMonth)} с ${compareMonth ? getMonthLabel(compareMonth) : 'отчетом недоступно'}.</p>
        </div>
      `;
      wrapper.append(makeProgressTable(sheet, currentMonth, compareMonth));
      elements.dayViews.append(wrapper);
    });

  if (!elements.dayViews.children.length) {
    elements.dayViews.append(createEmpty('По текущим фильтрам нет упражнений для сравнения.'));
  }
}

function renderFullReport() {
  elements.fullReport.innerHTML = '';
  const currentMonth = getSelectedCurrentMonth();

  Object.values(state.dashboard.reports)
    .forEach((sheet) => {
      const exercises = visibleExercises(sheet);
      if (!exercises.length) return;

      const dayBlock = document.createElement('div');
      dayBlock.className = 'full-report__day';

      const title = document.createElement('div');
      title.className = 'full-report__day-title';
      title.innerHTML = `<h3>${sheet.title}</h3><p class="table-note">Полная история отчетов по упражнениям.</p>`;
      dayBlock.append(title);

      exercises.forEach((exercise) => {
        const details = document.createElement('details');
        details.className = 'report-details';

        const entries = loggedEntries(exercise.entries);
        const visibleEntries =
          state.currentMonthFilter === 'latest' || state.currentMonthFilter === 'all'
            ? entries
            : entries.filter((entry) => Number(entry.month) === Number(currentMonth));

        const latest = entries.at(-1);
        const summary = document.createElement('summary');
        summary.className = 'report-details__summary';
        summary.innerHTML = `
          <div>
            <strong>${exercise.title}</strong>
            <span class="muted">${latest ? `Последний итог: ${latest.summary || formatSets(latest.sets)}` : 'Нет отчетов'}</span>
          </div>
          <span class="badge">${visibleEntries.length} записей</span>
        `;
        details.append(summary);

        const body = document.createElement('div');
        body.className = 'report-details__body';
        body.append(
          visibleEntries.length
            ? makeDetailedSetTable(visibleEntries)
            : createEmpty('Для выбранного фильтра по этому упражнению нет записей.'),
        );
        details.append(body);
        dayBlock.append(details);
      });

      elements.fullReport.append(dayBlock);
    });

  if (!elements.fullReport.children.length) {
    elements.fullReport.append(createEmpty('По текущим фильтрам полный отчет пустой.'));
  }
}

function render() {
  if (!state.dashboard) return;
  populateMonthControls();
  renderSummary();
  renderHeavyCards();
  renderDayViews();
  renderFullReport();
}

async function loadDashboard() {
  elements.refreshButton.disabled = true;
  elements.refreshButton.textContent = 'Читаю Excel...';
  try {
    const response = await fetch('/api/dashboard');
    if (!response.ok) throw new Error('Не удалось прочитать Excel');
    state.dashboard = await response.json();
    render();
  } catch (error) {
    console.error(error);
    elements.dayViews.innerHTML = '';
    elements.heavyCards.innerHTML = '';
    elements.summary.innerHTML = '';
    elements.fullReport.innerHTML = '';
    elements.dayViews.append(createEmpty('Не удалось загрузить данные из Excel. Проверь файл и перезапусти сервис.'));
  } finally {
    elements.refreshButton.disabled = false;
    elements.refreshButton.textContent = 'Обновить из Excel';
  }
}

loadDashboard();
