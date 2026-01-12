const form = document.getElementById("config-form");
const statusEl = document.getElementById("status");
const exportBtn = document.getElementById("export-btn");
const tableBody = document.querySelector("#data-table tbody");
const legendEl = document.getElementById("chart-legend");
const themeToggle = document.getElementById("theme-toggle");
const filtersSection = document.getElementById("filters-section");
const projectFilter = document.getElementById("project-filter");
const periodFilter = document.getElementById("period-filter");
const monthPicker = document.getElementById("month-picker");
const weekPicker = document.getElementById("week-picker");
const yearPicker = document.getElementById("year-picker");
const prevPeriodBtn = document.getElementById("prev-period");
const nextPeriodBtn = document.getElementById("next-period");

let chart;
let exportRows = [];
let allRows = [];
let boardsCache = [];
let currentPeriod = "all";
let currentDate = new Date();

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const mondayRequest = async (token, query, variables) => {
  const response = await fetch("https://api.monday.com/v2", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: token,
    },
    body: JSON.stringify({ query, variables }),
  });

  if (!response.ok) {
    throw new Error(`Monday API error: ${response.status}`);
  }

  const payload = await response.json();
  if (payload.errors) {
    throw new Error(payload.errors.map((error) => error.message).join(", "));
  }
  return payload.data;
};

const parseHours = (value) => {
  if (!value) return 0;
  try {
    const parsed = JSON.parse(value);
    if (!parsed || typeof parsed.duration !== "number") return 0;
    return parsed.duration / 3600000;
  } catch (error) {
    return 0;
  }
};

const parseStartDate = (value) => {
  if (!value) return null;
  try {
    const parsed = JSON.parse(value);
    if (!parsed || !parsed.startTime) return null;
    const date = new Date(parsed.startTime);
    if (Number.isNaN(date.getTime())) return null;
    return date;
  } catch (error) {
    return null;
  }
};

const parsePeople = (columnValue) => {
  if (!columnValue) return [];
  const names = columnValue.text
    ? columnValue.text.split(",").map((name) => name.trim()).filter(Boolean)
    : [];
  let ids = [];
  try {
    const parsed = columnValue.value ? JSON.parse(columnValue.value) : null;
    if (parsed && Array.isArray(parsed.personsAndTeams)) {
      ids = parsed.personsAndTeams
        .filter((person) => person.kind === "person")
        .map((person) => String(person.id));
    }
  } catch (error) {
    ids = [];
  }

  if (!names.length && !ids.length) {
    return [];
  }

  return names.map((name, index) => ({
    id: ids[index] || `unknown-${index}`,
    name,
  }));
};

const createLegend = (datasets) => {
  legendEl.innerHTML = "";
  datasets.forEach((dataset) => {
    const wrapper = document.createElement("span");
    const dot = document.createElement("span");
    dot.className = "legend-dot";
    dot.style.background = dataset.backgroundColor;
    wrapper.appendChild(dot);
    wrapper.appendChild(document.createTextNode(dataset.label));
    legendEl.appendChild(wrapper);
  });
};

const renderChart = (labels, datasets) => {
  const ctx = document.getElementById("hours-chart");
  if (chart) {
    chart.destroy();
  }

  chart = new Chart(ctx, {
    type: "bar",
    data: {
      labels,
      datasets,
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: {
          beginAtZero: true,
          title: {
            display: true,
            text: "Hours",
          },
        },
      },
      plugins: {
        legend: {
          display: false,
        },
      },
    },
  });

  createLegend(datasets);
};

const renderTable = (rows) => {
  tableBody.innerHTML = "";
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    [
      row.boardName,
      row.itemName,
      row.personName,
      row.hours.toFixed(2),
      row.dateLabel || "—",
    ].forEach((value) => {
      const td = document.createElement("td");
      td.textContent = value;
      tr.appendChild(td);
    });
    tableBody.appendChild(tr);
  });
};

const setStatus = (message, tone = "muted") => {
  statusEl.textContent = message;
  statusEl.dataset.tone = tone;
};

const fetchBoardItems = async (token, boardId) => {
  let cursor = null;
  let allItems = [];

  const boardQuery = `
    query ($boardId: ID!) {
      boards(ids: [$boardId]) {
        id
        name
        columns {
          id
          title
          type
        }
        items_page(limit: 200) {
          cursor
          items {
            id
            name
            column_values {
              id
              type
              text
              value
            }
          }
        }
      }
    }
  `;

  const initialData = await mondayRequest(token, boardQuery, { boardId });
  const board = initialData.boards[0];
  if (!board) {
    return null;
  }

  allItems = board.items_page.items;
  cursor = board.items_page.cursor;

  const nextQuery = `
    query ($cursor: String!) {
      next_items_page(limit: 200, cursor: $cursor) {
        cursor
        items {
          id
          name
          column_values {
            id
            type
            text
            value
          }
        }
      }
    }
  `;

  while (cursor) {
    await sleep(250);
    const nextData = await mondayRequest(token, nextQuery, { cursor });
    allItems = allItems.concat(nextData.next_items_page.items);
    cursor = nextData.next_items_page.cursor;
  }

  return {
    id: board.id,
    name: board.name,
    columns: board.columns,
    items: allItems,
  };
};

const loadBoards = async (token, boardIds) => {
  if (boardIds.length) {
    const results = [];
    for (const boardId of boardIds) {
      const board = await fetchBoardItems(token, boardId);
      if (board) {
        results.push(board);
      }
    }
    return results;
  }

  const boardsQuery = `
    query {
      boards(limit: 50) {
        id
      }
    }
  `;

  const data = await mondayRequest(token, boardsQuery, {});
  const results = [];
  for (const board of data.boards) {
    const fullBoard = await fetchBoardItems(token, board.id);
    if (fullBoard) {
      results.push(fullBoard);
    }
  }
  return results;
};

const buildDashboard = ({ boards, timeColumnInput, peopleColumnInput }) => {
  allRows = [];

  boards.forEach((board) => {
    const timeColumnId =
      timeColumnInput ||
      board.columns.find((column) => column.type === "time_tracking")?.id;
    const peopleColumnId =
      peopleColumnInput ||
      board.columns.find((column) => column.type === "people")?.id;

    if (!timeColumnId) {
      return;
    }

    board.items.forEach((item) => {
      const timeValue = item.column_values.find(
        (column) => column.id === timeColumnId
      );
      const hours = parseHours(timeValue?.value);
      const startDate = parseStartDate(timeValue?.value);
      if (!hours) return;

      const peopleValue = peopleColumnId
        ? item.column_values.find((column) => column.id === peopleColumnId)
        : null;
      const people = parsePeople(peopleValue);

      if (!people.length) {
        const row = {
          boardId: board.id,
          boardName: board.name,
          itemId: item.id,
          itemName: item.name,
          personId: "unassigned",
          personName: "Unassigned",
          hours,
          date: startDate,
          dateLabel: startDate ? startDate.toLocaleDateString() : "",
        };
        allRows.push(row);
        return;
      }

      const splitHours = hours / people.length;
      people.forEach((person) => {
        const row = {
          boardId: board.id,
          boardName: board.name,
          itemId: item.id,
          itemName: item.name,
          personId: person.id,
          personName: person.name,
          hours: splitHours,
          date: startDate,
          dateLabel: startDate ? startDate.toLocaleDateString() : "",
        };
        allRows.push(row);
      });
    });
  });

  boardsCache = boards.map((board) => ({ id: board.id, name: board.name }));
  populateProjectFilter(boardsCache);
  applyFilters();
};

const exportCsv = () => {
  if (!exportRows.length) return;
  const headers = [
    "board_id",
    "board_name",
    "item_id",
    "item_name",
    "person_id",
    "person_name",
    "hours",
    "date",
  ];
  const lines = [headers.join(",")];
  exportRows.forEach((row) => {
    const line = [
      row.boardId,
      `"${row.boardName.replace(/"/g, '""')}"`,
      row.itemId,
      `"${row.itemName.replace(/"/g, '""')}"`,
      row.personId,
      `"${row.personName.replace(/"/g, '""')}"`,
      row.hours.toFixed(2),
      row.date ? row.date.toISOString() : "",
    ];
    lines.push(line.join(","));
  });

  const blob = new Blob([lines.join("\n")], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "monday-hours-export.csv";
  link.click();
  URL.revokeObjectURL(url);
};

const populateProjectFilter = (boards) => {
  projectFilter.innerHTML = "";
  const allOption = document.createElement("option");
  allOption.value = "all";
  allOption.textContent = "All projects";
  projectFilter.appendChild(allOption);

  boards.forEach((board) => {
    const option = document.createElement("option");
    option.value = board.id;
    option.textContent = board.name;
    projectFilter.appendChild(option);
  });
};

const startOfWeek = (date) => {
  const copy = new Date(date);
  const day = copy.getDay();
  const diff = copy.getDate() - day + (day === 0 ? -6 : 1);
  copy.setDate(diff);
  copy.setHours(0, 0, 0, 0);
  return copy;
};

const endOfWeek = (date) => {
  const start = startOfWeek(date);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  end.setHours(23, 59, 59, 999);
  return end;
};

const getPeriodRange = (period, date) => {
  if (period === "month") {
    const start = new Date(date.getFullYear(), date.getMonth(), 1);
    const end = new Date(date.getFullYear(), date.getMonth() + 1, 0, 23, 59, 59);
    return { start, end };
  }
  if (period === "week") {
    return { start: startOfWeek(date), end: endOfWeek(date) };
  }
  if (period === "year") {
    const start = new Date(date.getFullYear(), 0, 1);
    const end = new Date(date.getFullYear(), 11, 31, 23, 59, 59);
    return { start, end };
  }
  return null;
};

const applyFilters = () => {
  const projectId = projectFilter.value;
  const period = currentPeriod;
  const range = getPeriodRange(period, currentDate);

  exportRows = allRows.filter((row) => {
    if (projectId !== "all" && row.boardId !== projectId) {
      return false;
    }
    if (!range) return true;
    if (!row.date) return false;
    return row.date >= range.start && row.date <= range.end;
  });

  const userTotals = new Map();
  const boardTotals = new Map();
  exportRows.forEach((row) => {
    const key = `${row.boardId}|${row.personId}`;
    boardTotals.set(key, (boardTotals.get(key) || 0) + row.hours);
    if (!userTotals.has(row.personId)) {
      userTotals.set(row.personId, row.personName);
    }
  });

  const users = Array.from(userTotals.entries()).map(([id, name]) => ({
    id,
    name,
  }));
  const boards = boardsCache.filter(
    (board) => projectId === "all" || board.id === projectId
  );

  const palette = [
    "#4d4dff",
    "#15a0ff",
    "#00c875",
    "#ffcb00",
    "#ff7a00",
    "#ff5ac4",
    "#9d50ff",
  ];

  const datasets = boards.map((board, index) => {
    const data = users.map((user) =>
      boardTotals.get(`${board.id}|${user.id}`) || 0
    );
    return {
      label: board.name,
      data,
      backgroundColor: palette[index % palette.length],
    };
  });

  renderChart(users.map((user) => user.name), datasets);
  renderTable(exportRows);
  exportBtn.disabled = exportRows.length === 0;
};

const updatePeriodInputs = () => {
  monthPicker.disabled = currentPeriod !== "month";
  weekPicker.disabled = currentPeriod !== "week";
  yearPicker.disabled = currentPeriod !== "year";

  if (currentPeriod === "month") {
    monthPicker.value = `${currentDate.getFullYear()}-${String(
      currentDate.getMonth() + 1
    ).padStart(2, "0")}`;
  }
  if (currentPeriod === "week") {
    const weekStart = startOfWeek(currentDate);
    weekPicker.value = weekStart.toISOString().split("T")[0];
  }
  if (currentPeriod === "year") {
    yearPicker.value = currentDate.getFullYear();
  }
  if (currentPeriod === "all") {
    monthPicker.value = "";
    weekPicker.value = "";
    yearPicker.value = "";
  }
};

const shiftPeriod = (direction) => {
  if (currentPeriod === "month") {
    currentDate = new Date(
      currentDate.getFullYear(),
      currentDate.getMonth() + direction,
      1
    );
  } else if (currentPeriod === "week") {
    currentDate = new Date(currentDate);
    currentDate.setDate(currentDate.getDate() + direction * 7);
  } else if (currentPeriod === "year") {
    currentDate = new Date(currentDate.getFullYear() + direction, 0, 1);
  }
  updatePeriodInputs();
  applyFilters();
};

const initTheme = () => {
  const body = document.body;
  const isDark = body.dataset.theme !== "light";
  body.dataset.theme = isDark ? "dark" : "light";
  themeToggle.textContent = isDark
    ? "Switch to light mode"
    : "Switch to dark mode";
};

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const formData = new FormData(form);
  const token = formData.get("token").trim();
  const boardInput = formData.get("boards").trim();
  const timeColumnInput = formData.get("timeColumn").trim();
  const peopleColumnInput = formData.get("peopleColumn").trim();

  if (!token) {
    setStatus("Please provide a Monday API token.", "error");
    return;
  }

  const boardIds = boardInput
    ? boardInput.split(",").map((id) => id.trim()).filter(Boolean)
    : [];

  exportBtn.disabled = true;
  setStatus("Loading boards and hours data...");

  try {
    const boards = await loadBoards(token, boardIds);
    if (!boards.length) {
      setStatus("No boards found. Check your access or IDs.", "warning");
      return;
    }

    buildDashboard({
      boards,
      timeColumnInput,
      peopleColumnInput,
    });

    setStatus("Dashboard ready. Export is available below.");
    filtersSection.classList.add("active");
    updatePeriodInputs();
  } catch (error) {
    setStatus(`Error: ${error.message}`, "error");
  }
});

exportBtn.addEventListener("click", exportCsv);

themeToggle.addEventListener("click", () => {
  const body = document.body;
  const nextTheme = body.dataset.theme === "dark" ? "light" : "dark";
  body.dataset.theme = nextTheme;
  themeToggle.textContent =
    nextTheme === "dark" ? "Switch to light mode" : "Switch to dark mode";
});

projectFilter.addEventListener("change", applyFilters);

periodFilter.addEventListener("change", (event) => {
  currentPeriod = event.target.value;
  updatePeriodInputs();
  applyFilters();
});

monthPicker.addEventListener("change", (event) => {
  if (!event.target.value) return;
  const [year, month] = event.target.value.split("-").map(Number);
  currentDate = new Date(year, month - 1, 1);
  applyFilters();
});

weekPicker.addEventListener("change", (event) => {
  if (!event.target.value) return;
  currentDate = new Date(event.target.value);
  applyFilters();
});

yearPicker.addEventListener("change", (event) => {
  const year = Number(event.target.value);
  if (!year) return;
  currentDate = new Date(year, 0, 1);
  applyFilters();
});

prevPeriodBtn.addEventListener("click", () => shiftPeriod(-1));
nextPeriodBtn.addEventListener("click", () => shiftPeriod(1));

initTheme();
