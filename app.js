const form = document.getElementById("config-form");
const statusEl = document.getElementById("status");
const exportBtn = document.getElementById("export-btn");
const tableBody = document.querySelector("#data-table tbody");
const loadingIndicator = document.getElementById("loading-indicator");
const loadingOverlay = document.getElementById("loading-overlay");
const metricProjects = document.getElementById("metric-projects");
const metricWorkspaces = document.getElementById("metric-workspaces");
const metricUsers = document.getElementById("metric-users");
const metricHours = document.getElementById("metric-hours");
const metricRecords = document.getElementById("metric-records");
const themeToggle = document.getElementById("theme-toggle");
const filtersSection = document.getElementById("filters-section");
const projectFilter = document.getElementById("project-filter");
const userFilter = document.getElementById("user-filter");
const periodFilter = document.getElementById("period-filter");
const prevPeriodBtn = document.getElementById("prev-period");
const nextPeriodBtn = document.getElementById("next-period");
const cycleLabel = document.getElementById("cycle-label");
const jumpCurrentBtn = document.getElementById("jump-current");
const rawEntriesToggle = document.getElementById("raw-entries-only");
const allDataToggle = document.getElementById("all-data");
const startDateInput = document.getElementById("start-date");
const endDateInput = document.getElementById("end-date");
const workspaceFilter = document.getElementById("workspace-filter");
const boardFilter = document.getElementById("board-filter");
const loadWorkspacesBtn = document.getElementById("load-workspaces");
const debugOutput = document.getElementById("debug-output");
const copyDebugBtn = document.getElementById("copy-debug");
const clearDebugBtn = document.getElementById("clear-debug");
const userTotalsEl = document.getElementById("user-totals");
const debugPanel = document.getElementById("debug-panel");
const debugToggleBtn = document.getElementById("debug-toggle");
const lastLoadTimeEl = document.getElementById("last-load-time");
const lastLoadRequestsEl = document.getElementById("last-load-requests");
const tokenInput = form.querySelector('input[name="token"]');

const TOKEN_STORAGE_KEY = "mondayApiToken";
const THEME_STORAGE_KEY = "mondayTheme";

let exportRows = [];
let allRows = [];
let parentRows = [];
let boardsCache = [];
let currentPeriod = "month";
let currentDate = new Date();
let preloadedBoards = [];
let preloadedWorkspaces = [];
const usersCache = new Map();
const collapsedGroups = new Set();
let isMeasuringLoad = false;
let loadRequestCount = 0;
let loadStartTime = 0;

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const BOARD_FETCH_CONCURRENCY = 3;
const PAGE_THROTTLE_MS = 75;

const mapWithConcurrency = async (list, limit, iterator) => {
  const results = new Array(list.length);
  let nextIndex = 0;
  const workerCount = Math.min(limit, list.length);
  const workers = Array.from({ length: workerCount }, async () => {
    while (nextIndex < list.length) {
      const currentIndex = nextIndex;
      nextIndex += 1;
      results[currentIndex] = await iterator(list[currentIndex], currentIndex);
    }
  });
  await Promise.all(workers);
  return results;
};

const getSelectedValues = (selectEl) => {
  if (!selectEl) return [];
  return Array.from(selectEl.selectedOptions).map((option) => option.value);
};

const isAllSelected = (values) =>
  !values.length || values.includes("all");

const formatLoadDuration = (ms) => {
  if (!ms && ms !== 0) return "--";
  const totalSeconds = Math.max(1, Math.ceil(ms / 1000));
  if (totalSeconds < 60) {
    return `${totalSeconds}s`;
  }
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;
  return `${minutes}m ${String(seconds).padStart(2, "0")}s`;
};

const updateLastLoadMetrics = (durationMs, requestCount) => {
  if (lastLoadTimeEl) {
    lastLoadTimeEl.textContent = formatLoadDuration(durationMs);
  }
  if (lastLoadRequestsEl) {
    lastLoadRequestsEl.textContent =
      typeof requestCount === "number" ? requestCount : 0;
  }
};

const appendDebug = (label, data) => {
  if (!debugOutput) return;
  const timestamp = new Date().toISOString();
  let entry = `[${timestamp}] ${label}`;
  if (data !== undefined) {
    let payload = data;
    if (typeof payload !== "string") {
      try {
        payload = JSON.stringify(payload, null, 2);
      } catch (error) {
        payload = String(payload);
      }
    }
    entry += `\n${payload}`;
  }
  debugOutput.value += (debugOutput.value ? "\n\n" : "") + entry;
  debugOutput.scrollTop = debugOutput.scrollHeight;
};

const mondayRequest = async (token, query, variables) => {
  if (isMeasuringLoad) {
    loadRequestCount += 1;
  }
  appendDebug("REQUEST", { query, variables });
  const response = await fetch("https://api.monday.com/v2", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: token,
    },
    body: JSON.stringify({ query, variables }),
  });

  appendDebug("RESPONSE_STATUS", {
    status: response.status,
    ok: response.ok,
  });

  if (!response.ok) {
    appendDebug("ERROR", `HTTP ${response.status}`);
    throw new Error(`Monday API error: ${response.status}`);
  }

  const payload = await response.json();
  appendDebug("RESPONSE_BODY", payload);

  if (payload.errors) {
    appendDebug("ERROR", payload.errors);
    throw new Error(payload.errors.map((error) => error.message).join(", "));
  }
  return payload.data;
};

const parseHours = (value) => {
  if (!value) return 0;
  try {
    const parsed = JSON.parse(value);
    if (!parsed || typeof parsed.duration !== "number") return 0;
    const durationMs = parsed.duration < 1000000 ? parsed.duration * 1000 : parsed.duration;
    return durationMs / 3600000;
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

const parseDurationFromText = (text) => {
  if (!text) return 0;
  const hoursMatch = text.match(/(\d+(?:\.\d+)?)\s*h/i);
  const minutesMatch = text.match(/(\d+)\s*m/i);
  const hours = hoursMatch ? Number(hoursMatch[1]) : 0;
  const minutes = minutesMatch ? Number(minutesMatch[1]) : 0;
  if (!hours && !minutes) return 0;
  return hours + minutes / 60;
};

const normalizeDate = (value) => {
  if (!value) return null;
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return null;
  return date;
};

const normalizeDuration = (duration) => {
  if (typeof duration !== "number" || duration <= 0) return 0;
  if (duration < 1000000) return duration * 1000;
  return duration;
};

const parseTimeTrackingEntries = (columnValue) => {
  if (!columnValue) return [];
  if (!columnValue.value && !columnValue.duration && !columnValue.history) {
    const textHours = parseDurationFromText(columnValue.text);
    if (textHours) {
      return [{ hours: textHours, date: null, userId: null }];
    }
    return [];
  }

  let parsed = {};
  if (columnValue.value) {
    try {
      parsed = JSON.parse(columnValue.value);
    } catch (error) {
      parsed = {};
    }
  }

  const entries = [];
  const pushEntry = (
    durationMs,
    startValue,
    endValue,
    userId,
    startedUserId,
    endedUserId
  ) => {
    const normalizedDuration = normalizeDuration(durationMs);
    if (!normalizedDuration) return;
    const startDate = normalizeDate(startValue);
    const endDate = normalizeDate(endValue);
    const entryDate = startDate || endDate;
    entries.push({
      hours: normalizedDuration / 3600000,
      date: entryDate,
      startDate,
      endDate,
      userId: userId ? String(userId) : null,
      startedUserId: startedUserId ? String(startedUserId) : null,
      endedUserId: endedUserId ? String(endedUserId) : null,
    });
  };

  if (Array.isArray(columnValue.history) && columnValue.history.length) {
    columnValue.history.forEach((entry) => {
      if (!entry) return;
      const startValue = entry.started_at ?? entry.startedAt;
      const endValue = entry.ended_at ?? entry.endedAt;
      const startedUserId = entry.started_user_id ?? entry.startedUserId;
      const endedUserId = entry.ended_user_id ?? entry.endedUserId;
      const startDate = normalizeDate(startValue);
      const endDate = normalizeDate(endValue);
      if (!startDate || !endDate) return;
      const durationSeconds = (endDate.getTime() - startDate.getTime()) / 1000;
      pushEntry(
        durationSeconds,
        startDate,
        endDate,
        startedUserId ?? endedUserId,
        startedUserId,
        endedUserId
      );
    });
    if (entries.length) return entries;
  }

  if (typeof columnValue.duration === "number") {
    const secondsMs = columnValue.duration * 1000;
    pushEntry(secondsMs, null, null, null, null, null);
    if (entries.length) return entries;
  }

  const addFromEntry = (entry) => {
    if (!entry || typeof entry !== "object") return;
    const durationMs =
      entry.duration ??
      entry.duration_ms ??
      entry.time ??
      entry.total_time ??
      entry.totalTime;
    const startValue =
      entry.started_at ??
      entry.startedAt ??
      entry.startTime ??
      entry.start ??
      entry.date ??
      entry.created_at ??
      entry.createdAt;
    const endValue =
      entry.ended_at ?? entry.endedAt ?? entry.endTime ?? entry.end ?? entry.endAt;
    const explicitUserId = entry.user_id ?? entry.userId ?? entry.owner_id ?? entry.ownerId;
    const startedUserId =
      entry.started_user_id ??
      entry.startedUserId ??
      entry.started_by_id ??
      entry.startedById;
    const endedUserId =
      entry.ended_user_id ??
      entry.endedUserId ??
      entry.ended_by_id ??
      entry.endedById;
    const userId =
      explicitUserId ??
      entry.user_id ??
      entry.userId ??
      entry.started_user_id ??
      entry.ended_user_id ??
      entry.owner_id ??
      entry.ownerId;
    pushEntry(durationMs, startValue, endValue, userId, startedUserId, endedUserId);
  };

  const candidateArrays = [];
  if (Array.isArray(parsed.entries)) candidateArrays.push(parsed.entries);
  if (Array.isArray(parsed.history)) candidateArrays.push(parsed.history);
  if (Array.isArray(parsed.tracked)) candidateArrays.push(parsed.tracked);
  if (Array.isArray(parsed.sessions)) candidateArrays.push(parsed.sessions);
  if (parsed.additional_value) {
    if (Array.isArray(parsed.additional_value.entries)) {
      candidateArrays.push(parsed.additional_value.entries);
    }
    if (Array.isArray(parsed.additional_value.history)) {
      candidateArrays.push(parsed.additional_value.history);
    }
    if (Array.isArray(parsed.additional_value.tracked)) {
      candidateArrays.push(parsed.additional_value.tracked);
    }
    if (Array.isArray(parsed.additional_value.sessions)) {
      candidateArrays.push(parsed.additional_value.sessions);
    }
  }

  if (candidateArrays.length) {
    candidateArrays.forEach((list) => list.forEach(addFromEntry));
    return entries;
  }

  if (typeof parsed.duration === "number") {
    const userId =
      parsed.user_id ??
      parsed.userId ??
      parsed.started_user_id ??
      parsed.ended_user_id;
    const startValue = parsed.startTime ?? parsed.started_at ?? parsed.date;
    const endValue = parsed.endTime ?? parsed.ended_at ?? null;
    const startedUserId = parsed.started_user_id ?? parsed.startedUserId;
    const endedUserId = parsed.ended_user_id ?? parsed.endedUserId;
    pushEntry(
      parsed.duration,
      startValue,
      endValue,
      userId,
      startedUserId,
      endedUserId
    );
  }

  return entries;
};

const buildGroupId = (boardId, itemId, columnId) =>
  `${boardId}:${itemId}:${columnId}`;

const getTimeTrackingTotalHours = (columnValue, entries) => {
  if (!columnValue) return 0;
  if (typeof columnValue.duration === "number" && columnValue.duration > 0) {
    const normalizedDuration = normalizeDuration(columnValue.duration);
    return normalizedDuration ? normalizedDuration / 3600000 : 0;
  }
  if (columnValue.value) {
    try {
      const parsed = JSON.parse(columnValue.value);
      if (parsed && typeof parsed.duration === "number") {
        const normalizedDuration = normalizeDuration(parsed.duration);
        if (normalizedDuration) {
          return normalizedDuration / 3600000;
        }
      }
    } catch (error) {
      // Ignore invalid JSON in value.
    }
  }
  if (Array.isArray(entries) && entries.length) {
    return entries.reduce((total, entry) => total + entry.hours, 0);
  }
  return 0;
};

const chunkArray = (list, size) => {
  const chunks = [];
  for (let i = 0; i < list.length; i += size) {
    chunks.push(list.slice(i, i + size));
  }
  return chunks;
};

const resolveUserNames = async (token, userIds) => {
  const ids = Array.from(userIds).filter((id) => !usersCache.has(id));
  if (!ids.length) return;

  const query = `
    query ($ids: [ID!]) {
      users(ids: $ids) {
        id
        name
      }
    }
  `;

  for (const chunk of chunkArray(ids, 100)) {
    const data = await mondayRequest(token, query, { ids: chunk });
    const users = data.users || [];
    users.forEach((user) => {
      if (user?.id) {
        usersCache.set(String(user.id), user.name || `User ${user.id}`);
      }
    });
  }
};

const renderTable = (rows) => {
  tableBody.innerHTML = "";
  const hasParents = rows.some((row) => row.rowType === "parent");
  const createCell = (text, className) => {
    const td = document.createElement("td");
    if (className) td.classList.add(className);
    td.textContent = text || "";
    return td;
  };

  rows.forEach((row) => {
    const tr = document.createElement("tr");
    if (row.rowType === "parent") {
      tr.classList.add("row-parent");
      tr.appendChild(createCell(row.workspaceName || ""));
      tr.appendChild(createCell(row.boardName));

      const itemCell = document.createElement("td");
      const isCollapsed = collapsedGroups.has(row.groupId);
      const toggleBtn = document.createElement("button");
      toggleBtn.type = "button";
      toggleBtn.className = "row-toggle";
      toggleBtn.dataset.groupId = row.groupId;
      toggleBtn.textContent = isCollapsed ? ">" : "v";
      toggleBtn.setAttribute("aria-expanded", String(!isCollapsed));
      toggleBtn.setAttribute(
        "aria-label",
        isCollapsed ? "Expand item rows" : "Collapse item rows"
      );
      itemCell.appendChild(toggleBtn);
      const itemText = document.createElement("span");
      itemText.textContent = row.itemName || "";
      itemCell.appendChild(itemText);
      tr.appendChild(itemCell);

      tr.appendChild(createCell(row.personName || "Total"));
      tr.appendChild(createCell(formatHoursMinutes(row.hours)));
      tr.appendChild(createCell(formatDateTime(row.startDate || row.date)));
      tr.appendChild(createCell(formatDateTime(row.endDate)));
    } else {
      const showContext = !hasParents;
      tr.appendChild(createCell(showContext ? row.workspaceName : ""));
      tr.appendChild(createCell(showContext ? row.boardName : ""));
      tr.appendChild(createCell(showContext ? row.itemName : ""));
      const userCell = createCell(row.personName);
      if (!showContext) {
        userCell.classList.add("child-indent");
      }
      tr.appendChild(userCell);
      tr.appendChild(createCell(formatHoursMinutes(row.hours)));
      tr.appendChild(createCell(formatDateTime(row.startDate || row.date)));
      tr.appendChild(createCell(formatDateTime(row.endDate)));
    }
    tableBody.appendChild(tr);
  });
};

const formatHoursMinutes = (hours) => {
  if (typeof hours !== "number" || Number.isNaN(hours)) return "";
  const totalMinutes = Math.round(hours * 60);
  const wholeHours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;
  return `${wholeHours}:${String(minutes).padStart(2, "0")}`;
};

const formatDateTime = (value) => {
  if (!(value instanceof Date) || Number.isNaN(value.getTime())) return "";
  return value.toLocaleString();
};

const formatShortDate = (value) =>
  value.toLocaleString(undefined, { month: "short", day: "2-digit" });

const formatMonthYear = (value) =>
  value.toLocaleString(undefined, { month: "long", year: "numeric" });

const getInitials = (name) => {
  if (!name) return "??";
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return "??";
  if (parts.length === 1) {
    return parts[0].slice(0, 2).toUpperCase();
  }
  return (parts[0][0] + parts[1][0]).toUpperCase();
};

const getAvatarColor = (seed) => {
  const palette = [
    "#f76808",
    "#ffb020",
    "#e11d48",
    "#7c3aed",
    "#0ea5e9",
    "#10b981",
    "#f59e0b",
    "#2563eb",
    "#14b8a6",
    "#9333ea",
  ];
  if (!seed) return palette[0];
  let hash = 0;
  for (let i = 0; i < seed.length; i += 1) {
    hash = (hash * 31 + seed.charCodeAt(i)) % palette.length;
  }
  return palette[hash];
};

const renderUserTotals = (rows) => {
  if (!userTotalsEl) return;
  userTotalsEl.innerHTML = "";
  if (!rows.length) {
    const empty = document.createElement("p");
    empty.className = "subtitle";
    empty.textContent = "No user totals available yet.";
    userTotalsEl.appendChild(empty);
    return;
  }

  const totals = new Map();
  rows.forEach((row) => {
    const key = row.personId || "unknown";
    if (!totals.has(key)) {
      totals.set(key, { name: row.personName || "Unknown user", hours: 0 });
    }
    totals.get(key).hours += row.hours;
  });

  const sortedTotals = Array.from(totals.entries())
    .map(([id, data]) => ({ id, ...data }))
    .sort((a, b) => b.hours - a.hours);

  sortedTotals.forEach((user) => {
    const row = document.createElement("div");
    row.className = "user-total-row";

    const info = document.createElement("div");
    info.className = "user-total-info";

    const avatar = document.createElement("div");
    avatar.className = "user-avatar";
    avatar.style.background = getAvatarColor(user.name);
    avatar.textContent = getInitials(user.name);

    const name = document.createElement("span");
    name.className = "user-name";
    name.textContent = user.name;

    info.appendChild(avatar);
    info.appendChild(name);

    const hours = document.createElement("span");
    hours.className = "user-total-hours";
    hours.textContent = formatHoursMinutes(user.hours);

    row.appendChild(info);
    row.appendChild(hours);
    userTotalsEl.appendChild(row);
  });
};

const buildDisplayRows = (entries, showParents, useFilteredTotals) => {
  const entriesByGroup = new Map();
  entries.forEach((entry) => {
    const groupId = entry.groupId || "ungrouped";
    if (!entriesByGroup.has(groupId)) {
      entriesByGroup.set(groupId, []);
    }
    entriesByGroup.get(groupId).push(entry);
  });

  const rows = [];
  const usedGroups = new Set();
  const sortEntries = (list) => {
    list.sort((a, b) => {
      const aTime = (a.startDate || a.date)?.getTime?.() || 0;
      const bTime = (b.startDate || b.date)?.getTime?.() || 0;
      return aTime - bTime;
    });
  };

  parentRows.forEach((parent) => {
    const groupEntries = entriesByGroup.get(parent.groupId);
    if (!groupEntries || !groupEntries.length) return;
    usedGroups.add(parent.groupId);
    if (showParents) {
      const entrySum = groupEntries.reduce(
        (total, entry) => total + entry.hours,
        0
      );
      const displayHours = useFilteredTotals
        ? entrySum
        : parent.totalHours ?? parent.hours;
      rows.push({ ...parent, hours: displayHours });
    }
    if (!showParents || !collapsedGroups.has(parent.groupId)) {
      sortEntries(groupEntries);
      groupEntries.forEach((entry) => rows.push(entry));
    }
  });

  entriesByGroup.forEach((groupEntries, groupId) => {
    if (usedGroups.has(groupId)) return;
    sortEntries(groupEntries);
    groupEntries.forEach((entry) => rows.push(entry));
  });

  return rows;
};

const setStatus = (message, tone = "muted") => {
  statusEl.textContent = message;
  statusEl.dataset.tone = tone;
};

const setLoading = (isLoading) => {
  if (loadingOverlay) {
    loadingOverlay.classList.toggle("active", isLoading);
  } else if (loadingIndicator) {
    loadingIndicator.classList.toggle("active", isLoading);
  }
};

const setFiltersEnabled = (isEnabled) => {
  if (!filtersSection) return;
  filtersSection.classList.toggle("disabled", !isEnabled);
  const fields = filtersSection.querySelectorAll("input, select, button");
  fields.forEach((field) => {
    field.disabled = !isEnabled;
  });
};

const persistToken = (token) => {
  if (!token) {
    localStorage.removeItem(TOKEN_STORAGE_KEY);
    return;
  }
  localStorage.setItem(TOKEN_STORAGE_KEY, token);
};

const updateMetrics = (rows) => {
  const uniqueBoards = new Set();
  const uniqueWorkspaces = new Set();
  const uniqueUsers = new Set();
  let totalHours = 0;

  rows.forEach((row) => {
    uniqueBoards.add(row.boardId);
    if (row.workspaceName) {
      uniqueWorkspaces.add(row.workspaceName);
    }
    uniqueUsers.add(row.personId);
    totalHours += row.hours;
  });

  if (metricWorkspaces) {
    metricWorkspaces.textContent = uniqueWorkspaces.size;
  }
  metricProjects.textContent = uniqueBoards.size;
  metricUsers.textContent = uniqueUsers.size;
  metricHours.textContent = totalHours.toFixed(2);
  metricRecords.textContent = rows.length;
};

const fetchBoardItems = async (token, boardId) => {
  let cursor = null;
  let allItems = [];

  const boardQuery = `
    query ($boardId: ID!) {
      boards(ids: [$boardId]) {
        id
        name
        workspace {
          id
          name
        }
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
            created_at
            column_values {
              id
              type
              text
              value
              ... on TimeTrackingValue {
                duration
                history {
                  id
                  started_at
                  ended_at
                  started_user_id
                  ended_user_id
                  status
                }
              }
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
          created_at
          column_values {
            id
            type
            text
            value
            ... on TimeTrackingValue {
              duration
              history {
                id
                started_at
                ended_at
                started_user_id
                ended_user_id
                status
              }
            }
          }
        }
      }
    }
  `;

  while (cursor) {
    if (PAGE_THROTTLE_MS) {
      await sleep(PAGE_THROTTLE_MS);
    }
    const nextData = await mondayRequest(token, nextQuery, { cursor });
    allItems = allItems.concat(nextData.next_items_page.items);
    cursor = nextData.next_items_page.cursor;
  }

  return {
    id: board.id,
    name: board.name,
    workspaceName: board.workspace?.name || "",
    columns: board.columns,
    items: allItems,
  };
};

const loadBoards = async (token, boardIds) => {
  if (boardIds.length) {
    const results = await mapWithConcurrency(
      boardIds,
      BOARD_FETCH_CONCURRENCY,
      async (boardId) => fetchBoardItems(token, boardId)
    );
    return results.filter(Boolean);
  }

  const boardsQuery = `
    query ($page: Int!) {
      boards(limit: 500, page: $page) {
        id
      }
    }
  `;

  const boardIdsToFetch = [];
  let page = 1;

  while (true) {
    const data = await mondayRequest(token, boardsQuery, { page });
    const boards = data.boards || [];
    if (!boards.length) break;
    boards.forEach((board) => boardIdsToFetch.push(board.id));
    page += 1;
  }

  const results = await mapWithConcurrency(
    boardIdsToFetch,
    BOARD_FETCH_CONCURRENCY,
    async (boardId) => fetchBoardItems(token, boardId)
  );
  return results.filter(Boolean);
};

const loadPreloadOptions = async (token) => {
  const workspacesQuery = `
    query {
      workspaces {
        id
        name
      }
    }
  `;
  const boardsQuery = `
    query ($page: Int!) {
      boards(limit: 500, page: $page) {
        id
        name
        workspace {
          id
          name
        }
      }
    }
  `;

  const workspaceData = await mondayRequest(token, workspacesQuery);
  preloadedWorkspaces = (workspaceData.workspaces || []).map((workspace) => ({
    id: String(workspace.id),
    name: workspace.name || `Workspace ${workspace.id}`,
  }));

  preloadedBoards = [];
  let page = 1;
  while (true) {
    const data = await mondayRequest(token, boardsQuery, { page });
    const boards = data.boards || [];
    if (!boards.length) break;
    boards.forEach((board) => {
      preloadedBoards.push({
        id: String(board.id),
        name: board.name,
        workspaceId: board.workspace?.id ? String(board.workspace.id) : "",
        workspaceName: board.workspace?.name || "",
      });
    });
    page += 1;
  }

  const workspaceMap = new Map(
    preloadedWorkspaces.map((workspace) => [workspace.id, workspace.name])
  );
  preloadedBoards.forEach((board) => {
    if (board.workspaceId && !workspaceMap.has(board.workspaceId)) {
      workspaceMap.set(
        board.workspaceId,
        board.workspaceName || `Workspace ${board.workspaceId}`
      );
    }
  });
  preloadedWorkspaces = Array.from(workspaceMap.entries()).map(
    ([id, name]) => ({ id, name })
  );

  populateWorkspaceFilter(preloadedWorkspaces);
  populateBoardFilter(preloadedBoards, getSelectedValues(workspaceFilter));
};

const buildDashboard = async ({ token, boards, timeColumnIdsInput, dateRange }) => {
  allRows = [];
  parentRows = [];
  collapsedGroups.clear();
  const userIds = new Set();

  boards.forEach((board) => {
    const timeColumnIds = timeColumnIdsInput.length
      ? timeColumnIdsInput
      : board.columns
          .filter((column) => column.type === "time_tracking")
          .map((column) => column.id);

    if (!timeColumnIds.length) {
      return;
    }

    board.items.forEach((item) => {
      const timeValues = timeColumnIds
        .map((columnId) =>
          item.column_values.find((column) => column.id === columnId)
        )
        .filter(Boolean);

      timeValues.forEach((timeValue) => {
        const groupId = buildGroupId(board.id, item.id, timeValue.id);
        let entries = parseTimeTrackingEntries(timeValue);
        if (!entries.length) {
          const fallbackHours = parseHours(timeValue.value);
          if (!fallbackHours) return;
          entries = [
            {
              hours: fallbackHours,
              date: parseStartDate(timeValue.value),
              userId: null,
            },
          ];
        }

        const filteredEntries = [];
        entries.forEach((entry) => {
          const entryDate =
            entry.date || (item.created_at ? new Date(item.created_at) : null);
          if (dateRange) {
            if (!entryDate) return;
            if (entryDate < dateRange.start || entryDate > dateRange.end) {
              return;
            }
          }
          filteredEntries.push({ entry, entryDate });
        });

        const totalFromEntries = filteredEntries.reduce(
          (total, current) => total + current.entry.hours,
          0
        );
        const totalFromAllEntries = entries.reduce(
          (total, entry) => total + entry.hours,
          0
        );
        const totalHours = dateRange ? totalFromEntries : totalFromAllEntries;

        if (totalHours || filteredEntries.length) {
          parentRows.push({
            rowType: "parent",
            groupId,
            boardId: board.id,
            boardName: board.name,
            workspaceName: board.workspaceName,
            itemId: item.id,
            itemName: item.name,
            personId: "total",
            personName: "Total",
            hours: totalHours,
            totalHours,
            date: null,
            startDate: null,
            endDate: null,
            dateLabel: "",
          });
        }

        filteredEntries.forEach(({ entry, entryDate }) => {
          const personId = entry.userId ? String(entry.userId) : "unknown";
          if (entry.userId) {
            userIds.add(String(entry.userId));
          }

          allRows.push({
            rowType: "entry",
            groupId,
            boardId: board.id,
            boardName: board.name,
            workspaceName: board.workspaceName,
            itemId: item.id,
            itemName: item.name,
            personId,
            personName:
              personId === "unknown"
                ? "Unknown user"
                : usersCache.get(personId) || `User ${personId}`,
            hours: entry.hours,
            date: entryDate,
            startDate: entry.startDate || null,
            endDate: entry.endDate || null,
            dateLabel: entryDate ? entryDate.toLocaleDateString() : "",
          });
        });
      });
    });
  });

  await resolveUserNames(token, userIds);

  allRows = allRows.map((row) => {
    if (row.personId === "unknown") return row;
    const resolvedName = usersCache.get(row.personId);
    if (!resolvedName) return row;
    return { ...row, personName: resolvedName };
  });

  boardsCache = boards.map((board) => ({
    id: board.id,
    name: board.name,
    workspaceName: board.workspaceName,
  }));
  populateProjectFilter(boardsCache);
  populateUserFilter(allRows);
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
  const current = projectFilter.value || "all";
  projectFilter.innerHTML = "";
  const allOption = document.createElement("option");
  allOption.value = "all";
  allOption.textContent = "All projects";
  projectFilter.appendChild(allOption);

  const sortedBoards = boards
    .map((board) => {
      const workspacePrefix = board.workspaceName
        ? `${board.workspaceName} - `
        : "";
      return {
        id: board.id,
        label: `${workspacePrefix}${board.name}`,
      };
    })
    .sort((a, b) => a.label.localeCompare(b.label));

  sortedBoards.forEach((board) => {
    const option = document.createElement("option");
    option.value = board.id;
    option.textContent = board.label;
    projectFilter.appendChild(option);
  });

  if (sortedBoards.some((board) => board.id === current)) {
    projectFilter.value = current;
  }
};

const populateWorkspaceFilter = (workspaces) => {
  const selectedValues = getSelectedValues(workspaceFilter);
  workspaceFilter.innerHTML = "";
  const allOption = document.createElement("option");
  allOption.value = "all";
  allOption.textContent = "All workspaces";
  workspaceFilter.appendChild(allOption);

  const sortedWorkspaces = [...workspaces].sort((a, b) =>
    a.name.localeCompare(b.name)
  );
  sortedWorkspaces.forEach((workspace) => {
    const option = document.createElement("option");
    option.value = workspace.id;
    option.textContent = workspace.name;
    workspaceFilter.appendChild(option);
  });

  if (isAllSelected(selectedValues)) {
    allOption.selected = true;
  } else {
    Array.from(workspaceFilter.options).forEach((option) => {
      if (selectedValues.includes(option.value)) {
        option.selected = true;
      }
    });
  }
};

const populateBoardFilter = (boards, workspaceIds) => {
  const selectedValues = getSelectedValues(boardFilter);
  boardFilter.innerHTML = "";
  const allOption = document.createElement("option");
  allOption.value = "all";
  allOption.textContent = "All boards";
  boardFilter.appendChild(allOption);

  const filteredBoards = boards.filter((board) => {
    if (isAllSelected(workspaceIds)) return true;
    return workspaceIds.includes(board.workspaceId);
  });

  const options = filteredBoards
    .map((board) => {
      const showPrefix = isAllSelected(workspaceIds) || workspaceIds.length > 1;
      const workspacePrefix =
        showPrefix && board.workspaceName ? `${board.workspaceName} - ` : "";
      return {
        id: board.id,
        label: `${workspacePrefix}${board.name}`,
      };
    })
    .sort((a, b) => a.label.localeCompare(b.label));

  options.forEach((board) => {
    const option = document.createElement("option");
    option.value = board.id;
    option.textContent = board.label;
    boardFilter.appendChild(option);
  });

  if (isAllSelected(selectedValues)) {
    allOption.selected = true;
  } else {
    Array.from(boardFilter.options).forEach((option) => {
      if (selectedValues.includes(option.value)) {
        option.selected = true;
      }
    });
  }
};

const populateUserFilter = (rows) => {
  if (!userFilter) return;
  const current = userFilter.value || "all";
  const users = new Map();
  rows.forEach((row) => {
    users.set(row.personId, row.personName);
  });

  userFilter.innerHTML = "";
  const allOption = document.createElement("option");
  allOption.value = "all";
  allOption.textContent = "All users";
  userFilter.appendChild(allOption);

  const sortedUsers = Array.from(users.entries()).sort((a, b) =>
    a[1].localeCompare(b[1])
  );
  sortedUsers.forEach(([id, name]) => {
    const option = document.createElement("option");
    option.value = id;
    option.textContent = name;
    userFilter.appendChild(option);
  });

  if (users.has(current)) {
    userFilter.value = current;
  }
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
  if (period === "day") {
    const start = new Date(date);
    start.setHours(0, 0, 0, 0);
    const end = new Date(date);
    end.setHours(23, 59, 59, 999);
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
  const selectedUser = userFilter ? userFilter.value : "all";
  const period = currentPeriod;
  const range = getPeriodRange(period, currentDate);
  const showParents = rawEntriesToggle ? !rawEntriesToggle.checked : true;
  const useFilteredTotals =
    selectedUser !== "all" || currentPeriod !== "all";

  const filteredEntries = allRows.filter((row) => {
    if (projectId !== "all" && row.boardId !== projectId) {
      return false;
    }
    if (selectedUser !== "all" && row.personId !== selectedUser) {
      return false;
    }
    if (!range) return true;
    if (!row.date) return false;
    return row.date >= range.start && row.date <= range.end;
  });

  exportRows = filteredEntries;
  updateMetrics(filteredEntries);
  renderUserTotals(filteredEntries);
  const tableRows = buildDisplayRows(
    filteredEntries,
    showParents,
    useFilteredTotals
  );
  renderTable(tableRows);
  exportBtn.disabled = exportRows.length === 0;
};

const updatePeriodInputs = () => {
  if (periodFilter && periodFilter.value !== currentPeriod) {
    periodFilter.value = currentPeriod;
  }

  if (cycleLabel) {
    if (currentPeriod === "month") {
      cycleLabel.textContent = formatMonthYear(currentDate);
    } else if (currentPeriod === "week") {
      const weekStart = startOfWeek(currentDate);
      const weekEnd = endOfWeek(currentDate);
      cycleLabel.textContent = `${formatShortDate(weekStart)} - ${formatShortDate(
        weekEnd
      )}`;
    } else if (currentPeriod === "day") {
      cycleLabel.textContent = formatShortDate(currentDate);
    } else if (currentPeriod === "year") {
      cycleLabel.textContent = String(currentDate.getFullYear());
    } else {
      cycleLabel.textContent = "All time";
    }
  }

  if (jumpCurrentBtn) {
    if (currentPeriod === "month") {
      jumpCurrentBtn.textContent = "Jump to current month";
      jumpCurrentBtn.disabled = false;
      jumpCurrentBtn.style.display = "";
    } else if (currentPeriod === "week") {
      jumpCurrentBtn.textContent = "Jump to current week";
      jumpCurrentBtn.disabled = false;
      jumpCurrentBtn.style.display = "";
    } else if (currentPeriod === "day") {
      jumpCurrentBtn.textContent = "Jump to current day";
      jumpCurrentBtn.disabled = false;
      jumpCurrentBtn.style.display = "";
    } else if (currentPeriod === "year") {
      jumpCurrentBtn.textContent = "Jump to current year";
      jumpCurrentBtn.disabled = false;
      jumpCurrentBtn.style.display = "";
    } else {
      jumpCurrentBtn.textContent = "Jump to current period";
      jumpCurrentBtn.disabled = true;
      jumpCurrentBtn.style.display = "none";
    }
  }

  if (prevPeriodBtn && nextPeriodBtn) {
    const canCycle = currentPeriod !== "all";
    prevPeriodBtn.disabled = !canCycle;
    nextPeriodBtn.disabled = !canCycle;
  }
};

const shiftPeriod = (direction) => {
  if (currentPeriod === "month") {
    currentDate = new Date(
      currentDate.getFullYear(),
      currentDate.getMonth() + direction,
      1
    );
  } else if (currentPeriod === "day") {
    currentDate = new Date(currentDate);
    currentDate.setDate(currentDate.getDate() + direction);
  } else if (currentPeriod === "week") {
    currentDate = new Date(currentDate);
    currentDate.setDate(currentDate.getDate() + direction * 7);
  } else if (currentPeriod === "year") {
    currentDate = new Date(currentDate.getFullYear() + direction, 0, 1);
  }
  updatePeriodInputs();
  applyFilters();
};

const toggleDateRangeInputs = () => {
  if (!allDataToggle || !startDateInput || !endDateInput) return;
  const isAllData = allDataToggle.checked;
  startDateInput.disabled = isAllData;
  endDateInput.disabled = isAllData;
};

const initTheme = () => {
  const body = document.body;
  const savedTheme = localStorage.getItem(THEME_STORAGE_KEY);
  const resolvedTheme =
    savedTheme === "light" || savedTheme === "dark"
      ? savedTheme
      : body.dataset.theme === "light"
        ? "light"
        : "dark";
  body.dataset.theme = resolvedTheme;
  themeToggle.textContent =
    resolvedTheme === "dark"
    ? "Switch to light mode"
    : "Switch to dark mode";
};

loadWorkspacesBtn.addEventListener("click", async () => {
  const token = new FormData(form).get("token").trim();
  if (!token) {
    setStatus("Please provide a Monday API token.", "error");
    return;
  }

  persistToken(token);
  setStatus("Loading workspaces and boards...");
  setLoading(true);
  try {
    await loadPreloadOptions(token);
    setStatus("Workspaces and boards loaded.");
  } catch (error) {
    setStatus(`Error: ${error.message}`, "error");
  } finally {
    setLoading(false);
  }
});

workspaceFilter.addEventListener("change", () => {
  populateBoardFilter(preloadedBoards, getSelectedValues(workspaceFilter));
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const formData = new FormData(form);
  const token = formData.get("token").trim();
  const timeColumnIdsInput = [];
  const selectedWorkspaces = workspaceFilter
    ? getSelectedValues(workspaceFilter)
    : ["all"];

  if (!token) {
    setStatus("Please provide a Monday API token.", "error");
    setLoading(false);
    return;
  }

  persistToken(token);
  const isAllData = allDataToggle ? allDataToggle.checked : true;
  let dateRange = null;

  if (!isAllData) {
    if (!startDateInput.value || !endDateInput.value) {
      setStatus("Select a start and end date or choose All data.", "error");
      return;
    }

    const start = new Date(startDateInput.value);
    const end = new Date(endDateInput.value);
    if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) {
      setStatus("Enter a valid date range.", "error");
      return;
    }
    if (start > end) {
      setStatus("Start date must be before the end date.", "error");
      return;
    }
    start.setHours(0, 0, 0, 0);
    end.setHours(23, 59, 59, 999);
    dateRange = { start, end };
  }

  let boardIds = [];
  if (boardFilter) {
    const selectedBoards = getSelectedValues(boardFilter);
    boardIds = selectedBoards.filter((value) => value !== "all");
  }

  exportBtn.disabled = true;
  setStatus("Loading boards and hours data...");
  setLoading(true);
  loadRequestCount = 0;
  loadStartTime = Date.now();
  isMeasuringLoad = true;

  try {
    if (!isAllSelected(selectedWorkspaces) && preloadedBoards.length === 0) {
      await loadPreloadOptions(token);
    }

    if (!isAllSelected(selectedWorkspaces)) {
      const allowedBoards = preloadedBoards.filter(
        (board) => selectedWorkspaces.includes(board.workspaceId)
      );
      if (boardIds.length) {
        const allowedIds = new Set(allowedBoards.map((board) => board.id));
        boardIds = boardIds.filter((id) => allowedIds.has(id));
        if (!boardIds.length) {
          setStatus("Selected board is not in the chosen workspace.", "warning");
          return;
        }
      } else {
        boardIds = allowedBoards.map((board) => board.id);
      }
    }

    const boards = await loadBoards(token, boardIds);
    if (!boards.length) {
      setStatus("No boards found. Check your access or IDs.", "warning");
      return;
    }

    await buildDashboard({
      token,
      boards,
      timeColumnIdsInput,
      dateRange,
    });

    setStatus("Dashboard ready. Export is available below.");
    setFiltersEnabled(true);
    updatePeriodInputs();
  } catch (error) {
    setStatus(`Error: ${error.message}`, "error");
  } finally {
    isMeasuringLoad = false;
    updateLastLoadMetrics(Date.now() - loadStartTime, loadRequestCount);
    setLoading(false);
  }
});

copyDebugBtn.addEventListener("click", async () => {
  if (!debugOutput) return;
  try {
    await navigator.clipboard.writeText(debugOutput.value);
  } catch (error) {
    debugOutput.select();
    document.execCommand("copy");
  }
  setStatus("Debug log copied.");
});

clearDebugBtn.addEventListener("click", () => {
  if (!debugOutput) return;
  debugOutput.value = "";
  setStatus("Debug log cleared.");
});

if (debugToggleBtn && debugPanel) {
  const setDebugCollapsed = (collapsed) => {
    debugPanel.classList.toggle("collapsed", collapsed);
    const caret = debugToggleBtn.querySelector(".debug-toggle-caret");
    if (caret) {
      caret.textContent = collapsed ? "<" : ">";
    }
    debugToggleBtn.setAttribute("aria-expanded", String(!collapsed));
  };

  debugToggleBtn.addEventListener("click", () => {
    const isCollapsed = debugPanel.classList.contains("collapsed");
    setDebugCollapsed(!isCollapsed);
  });

  setDebugCollapsed(true);
}

exportBtn.addEventListener("click", exportCsv);

tableBody.addEventListener("click", (event) => {
  const target = event.target.closest("button[data-group-id]");
  if (!target) return;
  const groupId = target.dataset.groupId;
  if (!groupId) return;
  if (collapsedGroups.has(groupId)) {
    collapsedGroups.delete(groupId);
  } else {
    collapsedGroups.add(groupId);
  }
  applyFilters();
});

allDataToggle.addEventListener("change", toggleDateRangeInputs);

themeToggle.addEventListener("click", () => {
  const body = document.body;
  const nextTheme = body.dataset.theme === "dark" ? "light" : "dark";
  body.dataset.theme = nextTheme;
  themeToggle.textContent =
    nextTheme === "dark" ? "Switch to light mode" : "Switch to dark mode";
  localStorage.setItem(THEME_STORAGE_KEY, nextTheme);
});

projectFilter.addEventListener("change", applyFilters);
userFilter.addEventListener("change", applyFilters);
if (rawEntriesToggle) {
  rawEntriesToggle.addEventListener("change", applyFilters);
}

periodFilter.addEventListener("change", (event) => {
  currentPeriod = event.target.value;
  updatePeriodInputs();
  applyFilters();
});

prevPeriodBtn.addEventListener("click", () => shiftPeriod(-1));
nextPeriodBtn.addEventListener("click", () => shiftPeriod(1));

if (jumpCurrentBtn) {
  jumpCurrentBtn.addEventListener("click", () => {
    currentDate = new Date();
    updatePeriodInputs();
    applyFilters();
  });
}

toggleDateRangeInputs();
updatePeriodInputs();
initTheme();
updateLastLoadMetrics(null, 0);

if (tokenInput) {
  tokenInput.addEventListener("input", () => {
    const token = tokenInput.value.trim();
    persistToken(token);
  });
}

setFiltersEnabled(false);

const preloadFromStoredToken = async () => {
  if (!tokenInput) return;
  const savedToken = localStorage.getItem(TOKEN_STORAGE_KEY);
  if (!savedToken) return;
  tokenInput.value = savedToken;
  setStatus("Saved token found. Loading workspaces and boards...");
  setLoading(true);
  try {
    await loadPreloadOptions(savedToken);
    setStatus("Workspaces and boards loaded.");
  } catch (error) {
    setStatus(`Error: ${error.message}`, "error");
  } finally {
    setLoading(false);
  }
};

preloadFromStoredToken();
