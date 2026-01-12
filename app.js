const form = document.getElementById("config-form");
const statusEl = document.getElementById("status");
const exportBtn = document.getElementById("export-btn");
const tableBody = document.querySelector("#data-table tbody");
const legendEl = document.getElementById("chart-legend");

let chart;
let exportRows = [];

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
    [row.boardName, row.itemName, row.personName, row.hours.toFixed(2)].forEach(
      (value) => {
        const td = document.createElement("td");
        td.textContent = value;
        tr.appendChild(td);
      }
    );
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
  const userTotals = new Map();
  const boardTotals = new Map();
  exportRows = [];

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
        };
        exportRows.push(row);
        const key = `${board.id}|${row.personId}`;
        const prev = boardTotals.get(key) || 0;
        boardTotals.set(key, prev + hours);
        if (!userTotals.has(row.personId)) {
          userTotals.set(row.personId, row.personName);
        }
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
        };
        exportRows.push(row);
        const key = `${board.id}|${person.id}`;
        const prev = boardTotals.get(key) || 0;
        boardTotals.set(key, prev + splitHours);
        if (!userTotals.has(person.id)) {
          userTotals.set(person.id, person.name);
        }
      });
    });
  });

  const users = Array.from(userTotals.entries()).map(([id, name]) => ({
    id,
    name,
  }));
  const boardsList = boards.map((board) => ({ id: board.id, name: board.name }));

  const palette = [
    "#4d4dff",
    "#15a0ff",
    "#00c875",
    "#ffcb00",
    "#ff7a00",
    "#ff5ac4",
    "#9d50ff",
  ];

  const datasets = boardsList.map((board, index) => {
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
  } catch (error) {
    setStatus(`Error: ${error.message}`, "error");
  }
});

exportBtn.addEventListener("click", exportCsv);
