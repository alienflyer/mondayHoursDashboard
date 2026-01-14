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
const showAllUsersToggle = document.getElementById("show-all-users");
const underThresholdOnly = document.getElementById("under-threshold-only");
const rawEntriesToggle = document.getElementById("raw-entries-only");
const ptoOnlyToggle = document.getElementById("pto-only");
const allDataToggle = document.getElementById("all-data");
const startDateInput = document.getElementById("start-date");
const endDateInput = document.getElementById("end-date");
const workspaceFilter = document.getElementById("workspace-filter");
const boardFilter = document.getElementById("board-filter");
const loadWorkspacesBtn = document.getElementById("load-workspaces");
const debugOutput = document.getElementById("debug-output");
const copyDebugBtn = document.getElementById("copy-debug");
const clearDebugBtn = document.getElementById("clear-debug");
const loadSummaryEl = document.getElementById("load-summary");
const loadSummaryEmpty = document.getElementById("load-summary-empty");
const loadSummarySearchInput = document.getElementById("load-summary-search");
const loadSummarySearchBtn = document.getElementById("load-summary-search-btn");
const loadSummaryPrevBtn = document.getElementById("load-summary-prev");
const loadSummaryNextBtn = document.getElementById("load-summary-next");
const summaryLoadTimeEl = document.getElementById("summary-load-time");
const summaryItemsEl = document.getElementById("summary-items");
const summaryWorkspacesEl = document.getElementById("summary-workspaces");
const summaryBoardsEl = document.getElementById("summary-boards");
const userTotalsEl = document.getElementById("user-totals");
const userTotalsExportBtn = document.getElementById("user-totals-export");
const teamFilter = document.getElementById("team-filter");
const adminToggleBtn = document.getElementById("admin-toggle");
const adminModal = document.getElementById("admin-modal");
const adminForm = document.getElementById("admin-form");
const adminStatus = document.getElementById("admin-status");
const adminMondayTokenInput = document.getElementById("admin-monday-token");
const adminTokenVisibilityBtn = document.getElementById("admin-token-visibility");
const ptoJsonUrlInput = document.getElementById("pto-json-url");
const ptoCsvUploadInput = document.getElementById("pto-csv-upload");
const ptoRememberUploadToggle = document.getElementById("pto-remember-upload");
const ptoAutoLoadToggle = document.getElementById("pto-auto-load");
const adminLoadPtoBtn = document.getElementById("admin-load-pto");
const debugPanel = document.getElementById("debug-panel");
const debugToggleBtn = document.getElementById("debug-toggle");
const lastLoadTimeEl = document.getElementById("last-load-time");
const lastLoadRequestsEl = document.getElementById("last-load-requests");
const tokenInput = form.querySelector('input[name="token"]');
const thresholdSettingsBtn = document.getElementById("user-total-settings-btn");
const thresholdModal = document.getElementById("threshold-modal");
const thresholdForm = document.getElementById("threshold-form");
const thresholdDayInput = document.getElementById("threshold-day");
const thresholdWeekInput = document.getElementById("threshold-week");
const thresholdMonthInput = document.getElementById("threshold-month");

const TOKEN_STORAGE_KEY = "mondayApiToken";
const THEME_STORAGE_KEY = "mondayTheme";
const THRESHOLD_STORAGE_KEY = "mondayThresholds";
const ADMIN_STORAGE_KEY = "mondayAdminSettings";
const PTO_CACHE_STORAGE_KEY = "mondayPtoCache";
const LOAD_SUMMARY_LIMIT = 2000;

let exportRows = [];
let allRows = [];
let parentRows = [];
let lastUserTotalsRows = [];
let boardsCache = [];
let currentPeriod = "week";
let currentDate = new Date();
let preloadedBoards = [];
let preloadedWorkspaces = [];
const usersCache = new Map();
const userTeamsCache = new Map();
let allUsersCache = [];
let allUsersLoadPromise = null;
let allTeamsCache = [];
const multiSelectInstances = new Map();
let multiSelectEventsAttached = false;
let ptoRows = [];
const loadSummaryItems = [];
const loadSummaryMatches = [];
let loadSummaryMatchIndex = -1;
const loadSummaryMetrics = {
  loadTime: null,
  items: 0,
  workspaces: 0,
  boards: 0,
};
const collapsedGroups = new Set();
let isMeasuringLoad = false;
let loadRequestCount = 0;
let loadStartTime = 0;

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const yieldToBrowser = () =>
  new Promise((resolve) => {
    if (typeof requestAnimationFrame === "function") {
      requestAnimationFrame(() => resolve());
    } else {
      setTimeout(resolve, 0);
    }
  });
const BOARD_FETCH_CONCURRENCY = 5;
const BOARD_PAGE_CONCURRENCY = 5;
const BOARD_PAGE_LIMIT = 500;
const USER_PAGE_CONCURRENCY = 4;
const USER_PAGE_LIMIT = 100;
const PAGE_THROTTLE_MS = 25;
const MAX_REQUEST_RETRIES = 4;
const BASE_RETRY_DELAY_MS = 600;
const MAX_RETRY_DELAY_MS = 8000;
const REQUEST_DELAY_DECAY = 0.85;
const RETRY_JITTER_MS = 200;
const MONDAY_API_ENDPOINT = "/api/monday";
const PTO_PROXY_ENDPOINT = "/api/pto";
const UNASSIGNED_TEAM = { id: "unassigned", name: "Unassigned" };
const DEFAULT_THRESHOLDS = { day: 8, week: 40, month: 160 };
const THERMO_COLORS = {
  red: "#ff6b6b",
  yellow: "#f59e0b",
  green: "#22c55e",
};
const THERMO_STOP_WARN = 0.4;
const THERMO_STOP_GOOD = 0.75;
let thresholds = null;
let adaptiveRequestDelayMs = 0;

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

const withJitter = (ms) => ms + Math.floor(Math.random() * RETRY_JITTER_MS);

const getRetryAfterMs = (response) => {
  const retryAfter = response.headers.get("Retry-After");
  if (!retryAfter) return null;
  const seconds = Number(retryAfter);
  if (Number.isFinite(seconds) && seconds > 0) {
    return seconds * 1000;
  }
  const parsedDate = Date.parse(retryAfter);
  if (!Number.isNaN(parsedDate)) {
    const delta = parsedDate - Date.now();
    return delta > 0 ? delta : null;
  }
  return null;
};

const parseResponsePayload = async (response) => {
  const text = await response.text();
  if (!text) return { json: null, text: "" };
  try {
    return { json: JSON.parse(text), text };
  } catch (error) {
    return { json: null, text };
  }
};

const isRateLimitError = (errors) =>
  errors.some((error) => {
    const message = error?.message || "";
    const code = error?.extensions?.code || "";
    return /rate limit|throttl|complexity/i.test(message) ||
      /rate|complexity/i.test(code);
  });

const increaseRequestDelay = (delayMs) => {
  adaptiveRequestDelayMs = Math.min(
    MAX_RETRY_DELAY_MS,
    Math.max(adaptiveRequestDelayMs, delayMs)
  );
};

const decayRequestDelay = () => {
  if (adaptiveRequestDelayMs <= 0) return;
  adaptiveRequestDelayMs = Math.max(
    0,
    Math.floor(adaptiveRequestDelayMs * REQUEST_DELAY_DECAY)
  );
};

const fetchBoardsPage = async (token, query, page) => {
  const data = await mondayRequest(token, query, { page });
  return data.boards || [];
};

const fetchBoardsPaged = async (token, query) => {
  let page = 1;
  const allBoards = [];
  let hasMore = true;

  while (hasMore) {
    const pages = Array.from(
      { length: BOARD_PAGE_CONCURRENCY },
      (_, index) => page + index
    );
    const results = await mapWithConcurrency(
      pages,
      BOARD_PAGE_CONCURRENCY,
      (pageNumber) => fetchBoardsPage(token, query, pageNumber)
    );

    results.forEach((boards) => {
      if (Array.isArray(boards) && boards.length) {
        allBoards.push(...boards);
      }
    });

    hasMore = results.some(
      (boards) => Array.isArray(boards) && boards.length === BOARD_PAGE_LIMIT
    );
    page += BOARD_PAGE_CONCURRENCY;
  }

  return allBoards;
};

const fetchUsersPage = async (token, query, page) => {
  const data = await mondayRequest(token, query, { page });
  return data.users || [];
};

const fetchUsersPaged = async (token, query) => {
  let page = 1;
  const allUsers = [];
  let hasMore = true;

  while (hasMore) {
    const pages = Array.from(
      { length: USER_PAGE_CONCURRENCY },
      (_, index) => page + index
    );
    const results = await mapWithConcurrency(
      pages,
      USER_PAGE_CONCURRENCY,
      (pageNumber) => fetchUsersPage(token, query, pageNumber)
    );

    results.forEach((users) => {
      if (Array.isArray(users) && users.length) {
        allUsers.push(...users);
      }
    });

    hasMore = results.some(
      (users) => Array.isArray(users) && users.length === USER_PAGE_LIMIT
    );
    page += USER_PAGE_CONCURRENCY;
  }

  return allUsers;
};

const normalizeTeams = (teams) => {
  if (!Array.isArray(teams)) return [];
  const seen = new Set();
  return teams.reduce((acc, team) => {
    if (!team?.id) return acc;
    const id = String(team.id);
    if (seen.has(id)) return acc;
    seen.add(id);
    acc.push({ id, name: team.name || `Team ${team.id}` });
    return acc;
  }, []);
};

const loadAllUsers = async (token) => {
  const usersQuery = `
    query ($page: Int!) {
      users(limit: ${USER_PAGE_LIMIT}, page: $page) {
        id
        name
        teams {
          id
          name
        }
      }
    }
  `;

  const fallbackQuery = `
    query ($page: Int!) {
      users(limit: ${USER_PAGE_LIMIT}, page: $page) {
        id
        name
      }
    }
  `;

  let users = [];
  try {
    users = await fetchUsersPaged(token, usersQuery);
  } catch (error) {
    if (String(error?.message || "").toLowerCase().includes("teams")) {
      appendDebug("ERROR", "User teams unavailable; loading users without teams.");
      users = await fetchUsersPaged(token, fallbackQuery);
    } else {
      throw error;
    }
  }

  userTeamsCache.clear();
  const userMap = new Map();
  users.forEach((user) => {
    if (!user?.id) return;
    const id = String(user.id);
    const name = user.name || `User ${user.id}`;
    const teams = normalizeTeams(user.teams);
    userMap.set(id, { name, teams });
    usersCache.set(id, name);
    userTeamsCache.set(id, teams);
  });

  allUsersCache = Array.from(userMap.entries()).map(([id, data]) => ({
    id,
    ...data,
  }));
  rebuildTeamsCache();
  populateTeamFilter();
  refreshPtoUserMapping();

  if (showAllUsersToggle?.checked && allRows.length) {
    applyFilters();
  }
};

const ensureAllUsersLoaded = async (token) => {
  if (allUsersCache.length) return;
  if (!allUsersLoadPromise) {
    allUsersLoadPromise = loadAllUsers(token).finally(() => {
      allUsersLoadPromise = null;
    });
  }
  return allUsersLoadPromise;
};

const getSelectedValues = (selectEl) => {
  if (!selectEl) return [];
  return Array.from(selectEl.selectedOptions).map((option) => option.value);
};

const isAllSelected = (values) =>
  !values.length || values.includes("all");

const setAllSelected = (selectEl) => {
  if (!selectEl) return;
  const options = Array.from(selectEl.options);
  let hasAll = false;
  options.forEach((option) => {
    if (option.value === "all") {
      option.selected = true;
      hasAll = true;
    } else {
      option.selected = false;
    }
  });
  if (!hasAll) {
    options.forEach((option) => {
      option.selected = false;
    });
  }
};

const ensureAllSelected = (selectEl) => {
  if (!selectEl) return;
  const options = Array.from(selectEl.options);
  const allOption = options.find((option) => option.value === "all");
  if (!allOption) return;
  const hasSelected = options.some(
    (option) => option.value !== "all" && option.selected
  );
  if (!hasSelected && !allOption.selected) {
    setAllSelected(selectEl);
  }
};

const syncMultiSelectState = (selectEl) => {
  const instance = multiSelectInstances.get(selectEl);
  if (!instance) return;
  const isDisabled = Boolean(selectEl.disabled);
  instance.wrapper.classList.toggle("is-disabled", isDisabled);
  instance.input.disabled = isDisabled;
  if (instance.searchInput) {
    instance.searchInput.disabled = isDisabled;
  }
  if (isDisabled) {
    instance.wrapper.classList.remove("is-open");
    instance.input.setAttribute("aria-expanded", "false");
  }
};

const setMultiSelectValue = (selectEl, value, forceSelected = null) => {
  if (!selectEl) return;
  const instance = multiSelectInstances.get(selectEl);
  const options = Array.from(selectEl.options);
  const target = options.find((option) => option.value === value);
  if (!target || target.disabled) return;

  if (!selectEl.multiple) {
    options.forEach((option) => {
      option.selected = option === target;
    });
  } else if (value === "all") {
    setAllSelected(selectEl);
  } else {
    const nextSelected =
      typeof forceSelected === "boolean" ? forceSelected : !target.selected;
    target.selected = nextSelected;
    const allOption = options.find((option) => option.value === "all");
    if (allOption) {
      allOption.selected = false;
    }
    const hasSelected = options.some(
      (option) => option.value !== "all" && option.selected
    );
    if (!hasSelected && allOption) {
      allOption.selected = true;
    }
  }

  selectEl.dispatchEvent(new Event("change", { bubbles: true }));
  syncMultiSelect(selectEl);
  if (!selectEl.multiple && instance) {
    instance.wrapper.classList.remove("is-open");
    instance.input.setAttribute("aria-expanded", "false");
    if (instance.searchInput) {
      instance.searchInput.value = "";
    }
  }
};

const closeAllMultiSelects = (exceptSelect = null) => {
  multiSelectInstances.forEach((instance, selectEl) => {
    if (selectEl === exceptSelect) return;
    instance.wrapper.classList.remove("is-open");
    instance.input.setAttribute("aria-expanded", "false");
  });
};

const toggleMultiSelect = (selectEl) => {
  const instance = multiSelectInstances.get(selectEl);
  if (!instance || selectEl.disabled) return;
  const isOpen = instance.wrapper.classList.contains("is-open");
  if (isOpen) {
    instance.wrapper.classList.remove("is-open");
    instance.input.setAttribute("aria-expanded", "false");
    return;
  }
  closeAllMultiSelects(selectEl);
  instance.wrapper.classList.add("is-open");
  instance.input.setAttribute("aria-expanded", "true");
  if (instance.searchInput) {
    instance.searchInput.focus();
    instance.searchInput.select();
  }
};

const syncMultiSelect = (selectEl) => {
  const instance = multiSelectInstances.get(selectEl);
  if (!instance) return;

  if (selectEl.multiple) {
    ensureAllSelected(selectEl);
  }
  const options = Array.from(selectEl.options);
  const allOption = options.find((option) => option.value === "all");
  const selectedOptions = options.filter((option) => option.selected);
  const hasAll =
    Boolean(allOption?.selected) ||
    selectedOptions.every((option) => option.value === "all");
  const displayOptions = selectEl.multiple
    ? hasAll && allOption
      ? [allOption]
      : selectedOptions.filter((option) => option.value !== "all")
    : selectedOptions.length
      ? [selectedOptions[0]]
      : allOption
        ? [allOption]
        : [];

  instance.chips.innerHTML = "";
  if (!options.length) {
    const placeholder = document.createElement("span");
    placeholder.className = "multi-select-placeholder";
    placeholder.textContent = "No options available";
    instance.chips.appendChild(placeholder);
  } else if (!displayOptions.length && allOption) {
    const placeholder = document.createElement("span");
    placeholder.className = "multi-select-placeholder";
    placeholder.textContent = allOption.textContent || "All";
    instance.chips.appendChild(placeholder);
  } else {
    displayOptions.forEach((option) => {
      const chip = document.createElement("span");
      chip.className = "multi-select-chip";

      const chipLabel = document.createElement("span");
      chipLabel.textContent = option.textContent || option.value;
      chip.appendChild(chipLabel);

      if (selectEl.multiple && option.value !== "all") {
        const removeBtn = document.createElement("button");
        removeBtn.type = "button";
        removeBtn.className = "multi-select-chip-remove";
        removeBtn.textContent = "x";
        removeBtn.addEventListener("click", (event) => {
          event.stopPropagation();
          setMultiSelectValue(selectEl, option.value, false);
        });
        chip.appendChild(removeBtn);
      }

      instance.chips.appendChild(chip);
    });
  }

  instance.optionsWrap.innerHTML = "";
  const searchValue = instance.searchInput
    ? normalizeKey(instance.searchInput.value)
    : "";
  let visibleCount = 0;

  options.forEach((option) => {
    const labelText = option.textContent || option.value;
    const normalizedLabel = normalizeKey(labelText);
    const matchesSearch =
      !searchValue || normalizedLabel.includes(searchValue);
    const shouldRender =
      option.value === "all" || option.selected || matchesSearch;
    if (!shouldRender) return;
    visibleCount += 1;

    const optionEl = document.createElement("button");
    optionEl.type = "button";
    optionEl.className = "multi-select-option";
    optionEl.dataset.value = option.value;
    optionEl.disabled = option.disabled;
    optionEl.setAttribute("aria-selected", String(option.selected));
    if (option.selected) {
      optionEl.classList.add("is-selected");
    }

    const label = document.createElement("span");
    label.className = "multi-select-option-label";
    label.textContent = labelText;
    optionEl.appendChild(label);

    const check = document.createElement("span");
    check.className = "multi-select-check";
    optionEl.appendChild(check);

    instance.optionsWrap.appendChild(optionEl);
  });

  if (!visibleCount) {
    const empty = document.createElement("div");
    empty.className = "multi-select-empty";
    empty.textContent = "No matches found.";
    instance.optionsWrap.appendChild(empty);
  }

  syncMultiSelectState(selectEl);
};

const buildMultiSelect = (selectEl) => {
  if (!selectEl || multiSelectInstances.has(selectEl)) return;

  const wrapper = document.createElement("div");
  wrapper.className = "multi-select";

  const input = document.createElement("button");
  input.type = "button";
  input.className = "multi-select-input";
  input.setAttribute("aria-haspopup", "listbox");
  input.setAttribute("aria-expanded", "false");

  const chips = document.createElement("div");
  chips.className = "multi-select-chips";

  const caret = document.createElement("span");
  caret.className = "multi-select-caret";
  caret.setAttribute("aria-hidden", "true");

  input.appendChild(chips);
  input.appendChild(caret);

  const dropdown = document.createElement("div");
  dropdown.className = "multi-select-dropdown";
  dropdown.setAttribute("role", "listbox");
  let searchInput = null;
  const optionsWrap = document.createElement("div");
  optionsWrap.className = "multi-select-options";

  if (selectEl.dataset.searchable === "true") {
    const searchWrap = document.createElement("div");
    searchWrap.className = "multi-select-search";
    searchInput = document.createElement("input");
    searchInput.type = "search";
    searchInput.placeholder =
      selectEl.dataset.searchPlaceholder || "Search";
    searchInput.autocomplete = "off";
    searchWrap.appendChild(searchInput);
    dropdown.appendChild(searchWrap);
  }
  dropdown.appendChild(optionsWrap);

  const parent = selectEl.parentNode;
  parent.insertBefore(wrapper, selectEl);
  wrapper.appendChild(input);
  wrapper.appendChild(selectEl);
  wrapper.appendChild(dropdown);

  selectEl.classList.add("multi-select-native");
  selectEl.tabIndex = -1;
  selectEl.setAttribute("aria-hidden", "true");

  const instance = {
    wrapper,
    input,
    chips,
    dropdown,
    optionsWrap,
    searchInput,
  };
  multiSelectInstances.set(selectEl, instance);

  input.addEventListener("click", (event) => {
    event.preventDefault();
    toggleMultiSelect(selectEl);
  });

  dropdown.addEventListener("click", (event) => {
    const optionEl = event.target.closest(".multi-select-option");
    if (!optionEl || optionEl.disabled) return;
    event.preventDefault();
    setMultiSelectValue(selectEl, optionEl.dataset.value);
  });

  if (searchInput) {
    searchInput.addEventListener("input", () => {
      syncMultiSelect(selectEl);
    });
    searchInput.addEventListener("click", (event) => {
      event.stopPropagation();
    });
  }

  syncMultiSelect(selectEl);

  if (!multiSelectEventsAttached) {
    multiSelectEventsAttached = true;
    document.addEventListener("click", (event) => {
      let clickedInside = false;
      multiSelectInstances.forEach((instance) => {
        if (instance.wrapper.contains(event.target)) {
          clickedInside = true;
        }
      });
      if (!clickedInside) {
        closeAllMultiSelects();
      }
    });
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        closeAllMultiSelects();
      }
    });
  }
};

const initMultiSelects = () => {
  document
    .querySelectorAll("select[multiple], select[data-searchable=\"true\"]")
    .forEach((selectEl) => {
      buildMultiSelect(selectEl);
    });
};

const rebuildTeamsCache = () => {
  const teamMap = new Map();
  let hasUnassigned = false;
  allUsersCache.forEach((user) => {
    if (Array.isArray(user.teams) && user.teams.length) {
      user.teams.forEach((team) => {
        if (!team?.id) return;
        teamMap.set(team.id, team.name || `Team ${team.id}`);
      });
    } else {
      hasUnassigned = true;
    }
  });
  if (hasUnassigned) {
    teamMap.set(UNASSIGNED_TEAM.id, UNASSIGNED_TEAM.name);
  }
  allTeamsCache = Array.from(teamMap.entries())
    .map(([id, name]) => ({ id, name }))
    .sort((a, b) => a.name.localeCompare(b.name));
};

const populateTeamFilter = () => {
  if (!teamFilter) return;
  const previousSelections = new Set(getSelectedValues(teamFilter));
  const selections = new Set();
  if (previousSelections.size && !previousSelections.has("all")) {
    allTeamsCache.forEach((team) => {
      if (previousSelections.has(team.id)) {
        selections.add(team.id);
      }
    });
  }
  if (!selections.size) {
    selections.add("all");
  }

  teamFilter.innerHTML = "";
  const allOption = document.createElement("option");
  allOption.value = "all";
  allOption.textContent = "No grouping";
  teamFilter.appendChild(allOption);

  allTeamsCache.forEach((team) => {
    const option = document.createElement("option");
    option.value = team.id;
    option.textContent = team.name;
    teamFilter.appendChild(option);
  });

  Array.from(teamFilter.options).forEach((option) => {
    option.selected = selections.has(option.value);
  });

  syncMultiSelect(teamFilter);
};

const getUserTeams = (userId) => {
  const teams = userTeamsCache.get(userId);
  if (Array.isArray(teams) && teams.length) {
    return teams;
  }
  return [UNASSIGNED_TEAM];
};

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

const formatInputDateValue = (date) => {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  const offsetMs = date.getTimezoneOffset() * 60000;
  const localDate = new Date(date.getTime() - offsetMs);
  return localDate.toISOString().split("T")[0];
};

const formatBoardSummaryLabel = (board) => {
  if (!board) return "Board";
  const workspacePrefix = board.workspaceName ? `${board.workspaceName} - ` : "";
  const boardName = board.name || board.id || "Board";
  return `${workspacePrefix}${boardName}`;
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

const updateLoadSummaryMetrics = (updates) => {
  if (!updates) return;
  Object.assign(loadSummaryMetrics, updates);
  if (summaryLoadTimeEl) {
    summaryLoadTimeEl.textContent = formatLoadDuration(loadSummaryMetrics.loadTime);
  }
  if (summaryItemsEl) {
    summaryItemsEl.textContent =
      Number.isFinite(loadSummaryMetrics.items)
        ? loadSummaryMetrics.items.toLocaleString()
        : "0";
  }
  if (summaryWorkspacesEl) {
    summaryWorkspacesEl.textContent =
      Number.isFinite(loadSummaryMetrics.workspaces)
        ? loadSummaryMetrics.workspaces.toLocaleString()
        : "0";
  }
  if (summaryBoardsEl) {
    summaryBoardsEl.textContent =
      Number.isFinite(loadSummaryMetrics.boards)
        ? loadSummaryMetrics.boards.toLocaleString()
        : "0";
  }
};

const resetLoadSummaryMetrics = () => {
  updateLoadSummaryMetrics({
    loadTime: null,
    items: 0,
    workspaces: 0,
    boards: 0,
  });
};

const clearLoadSummaryMatches = () => {
  loadSummaryMatches.forEach((match) => {
    match.classList.remove("is-match", "is-current");
  });
  loadSummaryMatches.length = 0;
  loadSummaryMatchIndex = -1;
  if (loadSummaryPrevBtn) loadSummaryPrevBtn.disabled = true;
  if (loadSummaryNextBtn) loadSummaryNextBtn.disabled = true;
};

const setCurrentLoadSummaryMatch = (index) => {
  if (!loadSummaryMatches.length) return;
  const count = loadSummaryMatches.length;
  loadSummaryMatchIndex = ((index % count) + count) % count;
  loadSummaryMatches.forEach((match) => match.classList.remove("is-current"));
  const current = loadSummaryMatches[loadSummaryMatchIndex];
  current.classList.add("is-current");
  current.scrollIntoView({ block: "nearest" });
  if (loadSummaryPrevBtn) loadSummaryPrevBtn.disabled = false;
  if (loadSummaryNextBtn) loadSummaryNextBtn.disabled = false;
};

const updateLoadSummarySearch = (query, keepIndex = false) => {
  clearLoadSummaryMatches();
  const normalized = (query || "").trim().toLowerCase();
  if (!normalized) return;
  loadSummaryItems.forEach((item) => {
    if (item.textContent.toLowerCase().includes(normalized)) {
      item.classList.add("is-match");
      loadSummaryMatches.push(item);
    }
  });
  if (loadSummaryMatches.length) {
    const nextIndex = keepIndex
      ? Math.min(Math.max(loadSummaryMatchIndex, 0), loadSummaryMatches.length - 1)
      : 0;
    setCurrentLoadSummaryMatch(nextIndex);
  }
};

const appendLoadSummary = (message) => {
  if (!loadSummaryEl || !message) return;
  const entry = document.createElement("div");
  entry.className = "load-summary-item";
  entry.textContent = message;
  loadSummaryEl.appendChild(entry);
  loadSummaryItems.push(entry);

  if (loadSummaryItems.length > LOAD_SUMMARY_LIMIT) {
    const removed = loadSummaryItems.shift();
    removed?.remove();
  }

  if (loadSummaryEmpty) {
    loadSummaryEmpty.hidden = loadSummaryItems.length > 0;
  }

  loadSummaryEl.scrollTop = loadSummaryEl.scrollHeight;
  if (loadSummarySearchInput?.value) {
    updateLoadSummarySearch(loadSummarySearchInput.value, true);
  }
};

const clearLoadSummary = () => {
  if (!loadSummaryEl) return;
  loadSummaryItems.forEach((entry) => entry.remove());
  loadSummaryItems.length = 0;
  if (loadSummaryEmpty) {
    loadSummaryEmpty.hidden = false;
  }
  clearLoadSummaryMatches();
  resetLoadSummaryMetrics();
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
  let lastError = null;

  for (let attempt = 0; attempt <= MAX_REQUEST_RETRIES; attempt += 1) {
    if (adaptiveRequestDelayMs > 0) {
      await sleep(adaptiveRequestDelayMs);
    }

    if (isMeasuringLoad) {
      loadRequestCount += 1;
    }
    appendDebug("REQUEST", { query, variables, attempt: attempt + 1 });

    const response = await fetch(MONDAY_API_ENDPOINT, {
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
      attempt: attempt + 1,
    });

    const { json, text } = await parseResponsePayload(response);
    if (json) {
      appendDebug("RESPONSE_BODY", json);
    } else if (text) {
      appendDebug("RESPONSE_BODY", text);
    }

    const errors = Array.isArray(json?.errors) ? json.errors : null;
    const rateLimited =
      response.status === 429 ||
      (errors && errors.length && isRateLimitError(errors));

    if (rateLimited) {
      const retryAfterMs = getRetryAfterMs(response);
      const baseDelay =
        retryAfterMs ?? BASE_RETRY_DELAY_MS * 2 ** attempt;
      const waitMs = withJitter(baseDelay);
      lastError = new Error("Monday API rate limit reached.");

      if (attempt < MAX_REQUEST_RETRIES) {
        increaseRequestDelay(Math.min(waitMs, MAX_RETRY_DELAY_MS));
        appendDebug("RATE_LIMIT", { waitMs, attempt: attempt + 1 });
        await sleep(waitMs);
        continue;
      }
    }

    if (!response.ok) {
      const errorMessage = errors?.map((error) => error.message).join(", ");
      const message = errorMessage || `HTTP ${response.status}`;
      appendDebug("ERROR", message);
      throw new Error(`Monday API error: ${message}`);
    }

    if (!json) {
      lastError = new Error("Monday API returned an invalid response.");
      appendDebug("ERROR", lastError.message);
      throw lastError;
    }

    if (errors && errors.length) {
      const message = errors.map((error) => error.message).join(", ");
      appendDebug("ERROR", errors);
      throw new Error(message);
    }

    decayRequestDelay();
    return json.data;
  }

  throw lastError || new Error("Monday API request failed after retries.");
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

const parseProjectorDate = (value) => {
  if (!value) return null;
  if (value instanceof Date) return value;
  const text = String(value).trim();
  if (!text) return null;
  const match = text.match(/(\d{4})[\/-](\d{2})[\/-](\d{2})/);
  if (!match) {
    return normalizeDate(text);
  }
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (!year || !month || !day) return null;
  return new Date(year, month - 1, day);
};

const normalizeDuration = (duration) => {
  if (typeof duration !== "number" || duration <= 0) return 0;
  if (duration < 1000000) return duration * 1000;
  return duration;
};

const normalizeKey = (value) =>
  String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[_\s]+/g, " ")
    .replace(/[^a-z0-9 ]+/g, "")
    .trim();

const cleanPtoName = (value) =>
  String(value || "")
    .replace(/\s*\(\d+\)\s*$/, "")
    .trim();

const cleanPtoTeam = (value) =>
  String(value || "")
    .replace(/\s*\(\d+\)\s*$/, "")
    .trim();

const parsePtoHours = (value) => {
  if (value === null || value === undefined) return 0;
  const text = String(value).trim();
  if (!text) return 0;
  if (text.includes(":")) {
    const [hours, minutes] = text.split(":").map((part) => Number(part));
    if (Number.isFinite(hours) && Number.isFinite(minutes)) {
      return hours + minutes / 60;
    }
  }
  const numeric = Number(text.replace(/[^0-9.-]/g, ""));
  return Number.isFinite(numeric) ? numeric : 0;
};

const parseCsv = (text) => {
  const rows = [];
  let row = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    if (inQuotes) {
      if (char === "\"") {
        if (text[i + 1] === "\"") {
          current += "\"";
          i += 1;
        } else {
          inQuotes = false;
        }
      } else {
        current += char;
      }
    } else if (char === "\"") {
      inQuotes = true;
    } else if (char === ",") {
      row.push(current);
      current = "";
    } else if (char === "\n") {
      row.push(current);
      rows.push(row);
      row = [];
      current = "";
    } else if (char !== "\r") {
      current += char;
    }
  }

  if (current.length || row.length) {
    row.push(current);
    rows.push(row);
  }

  return rows.filter((cells) =>
    cells.some((cell) => String(cell).trim() !== "")
  );
};

const PTO_FIELD_HINTS = {
  name: ["user", "user name", "resource", "employee", "person", "name"],
  email: ["email", "email address"],
  hours: ["hours", "pto hours", "time", "duration"],
  date: ["date", "day", "work date", "start date"],
  reason: [
    "reason",
    "time off reason",
    "timeoffreason",
    "pto reason",
    "holiday",
  ],
  description: [
    "time off description",
    "timeoffdescription",
    "description",
    "details",
    "notes",
  ],
  team: [
    "team",
    "department",
    "group",
    "practice",
    "business unit",
    "business units",
    "resource business units",
    "resourcebusinessunits",
    "businessunits",
  ],
};

const findHeaderIndex = (headers, hints) => {
  const normalized = headers.map(normalizeKey);
  for (const hint of hints) {
    const target = normalizeKey(hint);
    const index = normalized.findIndex(
      (header) => header === target || header.includes(target)
    );
    if (index >= 0) return index;
  }
  return -1;
};

const extractJsonRows = (data) => {
  if (Array.isArray(data)) return data;
  if (Array.isArray(data?.rows)) return data.rows;
  if (Array.isArray(data?.data)) return data.data;
  if (Array.isArray(data?.Results)) return data.Results;
  if (Array.isArray(data?.results)) return data.results;
  if (data && typeof data === "object") {
    for (const value of Object.values(data)) {
      if (Array.isArray(value) && value.some((item) => item && typeof item === "object")) {
        return value;
      }
    }
  }
  return [];
};

const findJsonField = (row, hints) => {
  if (!row || typeof row !== "object") return null;
  const keys = Object.keys(row);
  const normalizedKeys = keys.map(normalizeKey);
  for (const hint of hints) {
    const target = normalizeKey(hint);
    const index = normalizedKeys.findIndex(
      (key) => key === target || key.includes(target)
    );
    if (index >= 0) {
      return row[keys[index]];
    }
  }
  return null;
};

const buildUserLookup = () => {
  const lookup = new Map();
  const addUser = (id, name) => {
    if (!id || !name) return;
    const key = normalizeKey(name);
    if (!key || lookup.has(key)) return;
    lookup.set(key, { id: String(id), name });
  };

  if (allUsersCache.length) {
    allUsersCache.forEach((user) => addUser(user.id, user.name));
  } else {
    usersCache.forEach((name, id) => addUser(id, name));
  }

  return lookup;
};

const buildPtoRows = (entries) => {
  const lookup = buildUserLookup();
  return entries.map((entry) => {
    const normalizedName = normalizeKey(entry.name);
    const match = normalizedName ? lookup.get(normalizedName) : null;
    const personId = match?.id || `pto:${normalizedName || "unknown"}`;
    const personName = match?.name || entry.name || "Unknown user";
    const reason = String(entry.reason || "").trim();
    const description = String(entry.description || "").trim();
    const teamLabel = entry.team ? ` (${entry.team})` : "";
    const itemName = reason ? `${reason}${teamLabel}` : `PTO${teamLabel}`;
    return {
      source: "pto",
      rowType: "pto",
      groupId: "pto",
      boardId: "pto",
      boardName: "PTO",
      workspaceName: "PTO",
      itemId: "pto",
      itemName,
      ptoDescription: description,
      personId,
      personName,
      personKey: normalizedName,
      hours: entry.hours,
      date: entry.date,
      startDate: entry.date,
      endDate: entry.date,
      dateLabel: entry.date ? entry.date.toLocaleDateString() : "",
    };
  });
};

const refreshPtoUserMapping = () => {
  if (!ptoRows.length) return;
  const lookup = buildUserLookup();
  let updated = false;
  ptoRows = ptoRows.map((row) => {
    if (!row.personKey) return row;
    const match = lookup.get(row.personKey);
    if (!match) return row;
    if (row.personId !== match.id || row.personName !== match.name) {
      updated = true;
      return { ...row, personId: match.id, personName: match.name };
    }
    return row;
  });
  if (updated) {
    applyFilters();
  }
};

const getUserNameById = (userId) => {
  if (!userId || userId === "all") return "";
  const cached = usersCache.get(userId);
  if (cached) return cached;
  const user = allUsersCache.find((entry) => entry.id === userId);
  return user?.name || "";
};

const parsePtoCsvData = (text) => {
  const rows = parseCsv(text);
  if (!rows.length) {
    return { entries: [], error: "CSV file is empty." };
  }
  const headers = rows.shift();
  const nameIndex = findHeaderIndex(headers, PTO_FIELD_HINTS.name);
  const hoursIndex = findHeaderIndex(headers, PTO_FIELD_HINTS.hours);
  const dateIndex = findHeaderIndex(headers, PTO_FIELD_HINTS.date);
  const reasonIndex = findHeaderIndex(headers, PTO_FIELD_HINTS.reason);
  const descriptionIndex = findHeaderIndex(headers, PTO_FIELD_HINTS.description);
  const teamIndex = findHeaderIndex(headers, PTO_FIELD_HINTS.team);
  const emailIndex = findHeaderIndex(headers, PTO_FIELD_HINTS.email);

  if (nameIndex === -1 || hoursIndex === -1 || dateIndex === -1) {
    return { entries: [], error: "CSV must include name, hours, and date columns." };
  }

  const entries = [];
  rows.forEach((row) => {
    const rawName = row[nameIndex];
    const name = cleanPtoName(rawName);
    const hours = parsePtoHours(row[hoursIndex]);
    const dateValue = row[dateIndex];
    const date = parseProjectorDate(dateValue);
    if (!name || !hours || !date) return;
    const teamValue = teamIndex >= 0 ? row[teamIndex] : "";
    entries.push({
      name,
      hours,
      date,
      reason: reasonIndex >= 0 ? String(row[reasonIndex] || "").trim() : "",
      description:
        descriptionIndex >= 0 ? String(row[descriptionIndex] || "").trim() : "",
      team: cleanPtoTeam(teamValue),
      email: emailIndex >= 0 ? String(row[emailIndex] || "").trim() : "",
    });
  });
  return { entries, error: null };
};

const parsePtoJsonData = (data) => {
  const rows = extractJsonRows(data);
  if (!rows.length) {
    return { entries: [], error: "JSON response did not include rows." };
  }

  const entries = [];
  rows.forEach((row) => {
    const rawName = findJsonField(row, PTO_FIELD_HINTS.name);
    const name = cleanPtoName(rawName);
    const hours = parsePtoHours(findJsonField(row, PTO_FIELD_HINTS.hours));
    const dateValue = findJsonField(row, PTO_FIELD_HINTS.date);
    const date = parseProjectorDate(dateValue);
    if (!name || !hours || !date) return;
    const teamValue = findJsonField(row, PTO_FIELD_HINTS.team);
    const reasonValue = findJsonField(row, PTO_FIELD_HINTS.reason);
    const descriptionValue = findJsonField(row, PTO_FIELD_HINTS.description);
    entries.push({
      name,
      hours,
      date,
      reason: String(reasonValue || "").trim(),
      description: String(descriptionValue || "").trim(),
      team: cleanPtoTeam(teamValue),
      email: String(findJsonField(row, PTO_FIELD_HINTS.email) || "").trim(),
    });
  });

  return { entries, error: null };
};

const setAdminStatus = (message, tone = "muted") => {
  if (!adminStatus) return;
  adminStatus.textContent = message;
  adminStatus.dataset.tone = tone;
};

const setPtoRows = (entries, sourceLabel) => {
  ptoRows = buildPtoRows(entries);
  setAdminStatus(
    `Loaded ${ptoRows.length} PTO entries from ${sourceLabel || "PTO"}.`,
    "muted"
  );
  applyFilters();
};

const loadPtoFromText = (text, sourceLabel) => {
  const { entries, error } = parsePtoCsvData(text);
  if (error) {
    setAdminStatus(error, "error");
    return false;
  }
  setPtoRows(entries, sourceLabel);
  return true;
};

const loadPtoFromJson = (data, sourceLabel) => {
  const { entries, error } = parsePtoJsonData(data);
  if (error) {
    setAdminStatus(error, "error");
    return false;
  }
  setPtoRows(entries, sourceLabel);
  return true;
};

const loadPtoFromUrl = async (url) => {
  if (!url) {
    setAdminStatus("Enter a PTO URL first.", "error");
    return false;
  }
  const proxyUrl = `${PTO_PROXY_ENDPOINT}?url=${encodeURIComponent(url)}`;
  setAdminStatus("Loading PTO data...", "muted");
  try {
    const response = await fetch(proxyUrl);
    if (!response.ok) {
      if (response.status === 404) {
        setAdminStatus(
          "PTO proxy not found. Run `python3 dev_server.py`.",
          "error"
        );
        return false;
      }
      throw new Error(`HTTP ${response.status}`);
    }
    const contentType = response.headers.get("Content-Type") || "";
    if (contentType.includes("application/json") || url.includes("format=json")) {
      const data = await response.json();
      return loadPtoFromJson(data, "Projector JSON");
    }
    const text = await response.text();
    return loadPtoFromText(text, "Projector CSV");
  } catch (error) {
    setAdminStatus(`Failed to load PTO data: ${error.message}`, "error");
    return false;
  }
};

const loadPtoFromCache = () => {
  const cached = localStorage.getItem(PTO_CACHE_STORAGE_KEY);
  if (!cached) return false;
  return loadPtoFromText(cached, "Saved CSV");
};

const savePtoCache = (text) => {
  localStorage.setItem(PTO_CACHE_STORAGE_KEY, text);
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
  const createUserCell = (row, showContext) => {
    const td = document.createElement("td");
    if (!showContext) {
      td.classList.add("child-indent");
    }
    if (row?.source === "pto" && row?.ptoDescription) {
      td.classList.add("pto-user-cell");
      const desc = document.createElement("span");
      desc.className = "pto-description";
      desc.textContent = row.ptoDescription;
      const name = document.createElement("span");
      name.className = "pto-user-name";
      name.textContent = row.personName || "";
      td.appendChild(desc);
      td.appendChild(name);
    } else {
      td.textContent = row.personName || "";
    }
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
      tr.appendChild(createUserCell(row, showContext));
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

const formatHoursValue = (hours) => {
  if (typeof hours !== "number" || Number.isNaN(hours)) return "0";
  const value = Math.round(hours * 10) / 10;
  return String(value).replace(/\.0$/, "");
};

const sanitizeThresholdValue = (value, fallback) => {
  const numberValue = Number(value);
  if (Number.isFinite(numberValue) && numberValue >= 0) {
    return numberValue;
  }
  return fallback;
};

const loadThresholds = () => {
  let stored = {};
  try {
    stored = JSON.parse(localStorage.getItem(THRESHOLD_STORAGE_KEY) || "{}");
  } catch (error) {
    stored = {};
  }
  return {
    day: sanitizeThresholdValue(stored.day, DEFAULT_THRESHOLDS.day),
    week: sanitizeThresholdValue(stored.week, DEFAULT_THRESHOLDS.week),
    month: sanitizeThresholdValue(stored.month, DEFAULT_THRESHOLDS.month),
  };
};

const persistThresholds = (values) => {
  thresholds = { ...values };
  localStorage.setItem(THRESHOLD_STORAGE_KEY, JSON.stringify(thresholds));
};

const loadAdminSettings = () => {
  try {
    return JSON.parse(localStorage.getItem(ADMIN_STORAGE_KEY) || "{}");
  } catch (error) {
    return {};
  }
};

const persistAdminSettings = (settings) => {
  localStorage.setItem(ADMIN_STORAGE_KEY, JSON.stringify(settings));
};

const populateAdminForm = (settings) => {
  if (adminMondayTokenInput) {
    adminMondayTokenInput.value = localStorage.getItem(TOKEN_STORAGE_KEY) || "";
  }
  if (ptoJsonUrlInput) ptoJsonUrlInput.value = settings.ptoJsonUrl || "";
  if (ptoRememberUploadToggle) {
    ptoRememberUploadToggle.checked = Boolean(settings.ptoRememberUpload);
  }
  if (ptoAutoLoadToggle) {
    ptoAutoLoadToggle.checked = Boolean(settings.ptoAutoLoad);
  }
};

const openAdminModal = () => {
  if (!adminModal) return;
  const settings = loadAdminSettings();
  populateAdminForm(settings);
  if (adminTokenVisibilityBtn && adminMondayTokenInput) {
    setAdminTokenVisibility(false);
  }
  adminModal.classList.add("active");
  adminModal.setAttribute("aria-hidden", "false");
};

const closeAdminModal = () => {
  if (!adminModal) return;
  adminModal.classList.remove("active");
  adminModal.setAttribute("aria-hidden", "true");
};

const loadPtoFromSettings = async (settings) => {
  if (settings?.ptoJsonUrl) {
    const loaded = await loadPtoFromUrl(settings.ptoJsonUrl);
    if (loaded) return true;
    if (settings?.ptoRememberUpload) {
      return loadPtoFromCache();
    }
    return false;
  }
  if (settings?.ptoRememberUpload) {
    return loadPtoFromCache();
  }
  setAdminStatus("No PTO source configured.", "warning");
  return false;
};

const getThresholdForPeriod = (period) => {
  if (period === "day" || period === "week" || period === "month") {
    return thresholds?.[period] ?? DEFAULT_THRESHOLDS[period];
  }
  return null;
};

const getThermoFillGradient = (progress) => {
  const normalized = Math.min(Math.max(progress, 0), 1);
  if (normalized <= THERMO_STOP_WARN) {
    return THERMO_COLORS.red;
  }
  if (normalized <= THERMO_STOP_GOOD) {
    const redStop = Math.round((THERMO_STOP_WARN / normalized) * 100);
    return `linear-gradient(90deg, ${THERMO_COLORS.red} 0%, ${THERMO_COLORS.red} ${redStop}%, ${THERMO_COLORS.yellow} 100%)`;
  }
  const redStop = Math.round((THERMO_STOP_WARN / normalized) * 100);
  const yellowStop = Math.round((THERMO_STOP_GOOD / normalized) * 100);
  return `linear-gradient(90deg, ${THERMO_COLORS.red} 0%, ${THERMO_COLORS.red} ${redStop}%, ${THERMO_COLORS.yellow} ${yellowStop}%, ${THERMO_COLORS.green} 100%)`;
};

const setThresholdInputs = (values) => {
  if (!thresholdDayInput || !thresholdWeekInput || !thresholdMonthInput) return;
  thresholdDayInput.value = values.day;
  thresholdWeekInput.value = values.week;
  thresholdMonthInput.value = values.month;
};

const openThresholdModal = () => {
  if (!thresholdModal) return;
  setThresholdInputs(thresholds || loadThresholds());
  thresholdModal.classList.add("active");
  thresholdModal.setAttribute("aria-hidden", "false");
};

const closeThresholdModal = () => {
  if (!thresholdModal) return;
  thresholdModal.classList.remove("active");
  thresholdModal.setAttribute("aria-hidden", "true");
};

thresholds = loadThresholds();

const buildUserTotalsDisplay = (rows) => {
  const wantsAllUsers = Boolean(showAllUsersToggle?.checked);
  const totals = new Map();
  rows.forEach((row) => {
    const key = row.personId || "unknown";
    if (!totals.has(key)) {
      totals.set(key, { name: row.personName || "Unknown user", hours: 0 });
    }
    totals.get(key).hours += row.hours;
  });

  let displayTotals = [];
  if (wantsAllUsers) {
    if (!allUsersCache.length) {
      return { message: "User list not loaded yet." };
    }
    const seen = new Set();
    displayTotals = allUsersCache.map((user) => {
      const cached = totals.get(user.id);
      const name = cached?.name || user.name || `User ${user.id}`;
      const hours = cached?.hours ?? 0;
      seen.add(user.id);
      return { id: user.id, name, hours };
    });
    totals.forEach((data, id) => {
      if (!seen.has(id) && data.hours > 0) {
        displayTotals.push({ id, ...data });
      }
    });
  } else {
    displayTotals = Array.from(totals.entries()).map(([id, data]) => ({
      id,
      ...data,
    }));
  }

  displayTotals.sort((a, b) => {
    if (b.hours !== a.hours) return b.hours - a.hours;
    return (a.name || "").localeCompare(b.name || "");
  });
  displayTotals = displayTotals.map((user) => ({
    ...user,
    teams: getUserTeams(user.id),
  }));

  const threshold = getThresholdForPeriod(currentPeriod);
  const hasThreshold = typeof threshold === "number" && threshold > 0;
  const filterUnderThreshold = Boolean(underThresholdOnly?.checked);

  if (filterUnderThreshold && !hasThreshold) {
    return { message: "No threshold available for this period." };
  }

  if (filterUnderThreshold && hasThreshold) {
    displayTotals = displayTotals.filter((user) => user.hours < threshold);
  }

  if (!displayTotals.length) {
    return {
      message: filterUnderThreshold
        ? "No users below the threshold."
        : "No user totals available yet.",
    };
  }

  const selectedTeams = teamFilter ? getSelectedValues(teamFilter) : ["all"];
  const filterTeams = teamFilter && !isAllSelected(selectedTeams);
  if (!filterTeams) {
    return { displayTotals, threshold, hasThreshold, filterTeams: false };
  }

  const selectedTeamSet = new Set(selectedTeams);
  const teamGroups = new Map();

  displayTotals.forEach((user) => {
    const teams = Array.isArray(user.teams) && user.teams.length
      ? user.teams
      : [UNASSIGNED_TEAM];
    const visibleTeams = teams.filter((team) => selectedTeamSet.has(team.id));
    if (!visibleTeams.length) return;
    visibleTeams.forEach((team) => {
      if (!teamGroups.has(team.id)) {
        teamGroups.set(team.id, { team, users: [] });
      }
      teamGroups.get(team.id).users.push(user);
    });
  });

  if (!teamGroups.size) {
    return { message: "No users match the selected teams." };
  }

  const sortedGroups = Array.from(teamGroups.values()).sort((a, b) =>
    (a.team.name || "").localeCompare(b.team.name || "")
  );

  return {
    displayTotals,
    threshold,
    hasThreshold,
    filterTeams: true,
    teamGroups: sortedGroups,
  };
};

const renderUserTotals = (rows) => {
  if (!userTotalsEl) return;
  userTotalsEl.innerHTML = "";
  lastUserTotalsRows = rows;
  const displayState = buildUserTotalsDisplay(rows);
  if (displayState.message) {
    const empty = document.createElement("p");
    empty.className = "subtitle";
    empty.textContent = displayState.message;
    userTotalsEl.appendChild(empty);
    return;
  }

  const {
    displayTotals,
    threshold,
    hasThreshold,
    filterTeams,
    teamGroups,
  } = displayState;

  const buildUserTotalRow = (user) => {
    const row = document.createElement("div");
    row.className = "user-total-row";

    const info = document.createElement("div");
    info.className = "user-total-info";

    const avatar = document.createElement("div");
    avatar.className = "user-avatar";
    avatar.style.background = getAvatarColor(user.name);
    avatar.textContent = getInitials(user.name);

    const textWrap = document.createElement("div");
    textWrap.className = "user-total-text";

    const name = document.createElement("span");
    name.className = "user-name";
    name.textContent = user.name;

    const thermometer = document.createElement("div");
    thermometer.className = "user-thermo";
    if (!hasThreshold) {
      thermometer.classList.add("is-disabled");
    }

    const bulb = document.createElement("div");
    bulb.className = "user-thermo-bulb";

    const track = document.createElement("div");
    track.className = "user-thermo-track";

    const fill = document.createElement("div");
    fill.className = "user-thermo-fill";

    const label = document.createElement("span");
    label.className = "user-thermo-label";

    if (hasThreshold) {
      const progress = Math.min(user.hours / threshold, 1);
      fill.style.width = `${progress * 100}%`;
      fill.style.background = getThermoFillGradient(progress);
      label.textContent = `${formatHoursValue(user.hours)} / ${formatHoursValue(
        threshold
      )} hrs`;
      if (user.hours >= threshold) {
        row.classList.add("over-threshold");
      }
    } else {
      fill.style.width = "0%";
      fill.style.background = "";
      label.textContent = "No threshold for this period";
    }

    track.appendChild(fill);
    thermometer.appendChild(bulb);
    thermometer.appendChild(track);

    textWrap.appendChild(name);
    textWrap.appendChild(thermometer);
    textWrap.appendChild(label);

    info.appendChild(avatar);
    info.appendChild(textWrap);

    const hours = document.createElement("span");
    hours.className = "user-total-hours";
    hours.textContent = formatHoursMinutes(user.hours);

    row.appendChild(info);
    row.appendChild(hours);
    return row;
  };

  if (!filterTeams) {
    displayTotals.forEach((user) => {
      userTotalsEl.appendChild(buildUserTotalRow(user));
    });
    return;
  }
  teamGroups.forEach((group) => {
    const groupEl = document.createElement("div");
    groupEl.className = "team-group";

    const groupHeader = document.createElement("div");
    groupHeader.className = "team-group-header";

    const groupName = document.createElement("span");
    groupName.className = "team-group-name";
    groupName.textContent = group.team.name;

    const groupCount = document.createElement("span");
    groupCount.className = "team-group-count";
    groupCount.textContent = `${group.users.length} ${
      group.users.length === 1 ? "user" : "users"
    }`;

    groupHeader.appendChild(groupName);
    groupHeader.appendChild(groupCount);

    const groupRows = document.createElement("div");
    groupRows.className = "team-group-rows";
    group.users.forEach((user) => {
      groupRows.appendChild(buildUserTotalRow(user));
    });

    groupEl.appendChild(groupHeader);
    groupEl.appendChild(groupRows);
    userTotalsEl.appendChild(groupEl);
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

const buildPtoDisplayRows = (entries, showParents) => {
  if (!entries.length) return [];
  if (!showParents) {
    return entries.map((entry) => ({
      ...entry,
      rowType: "entry",
      groupId: "pto",
    }));
  }

  const totalHours = entries.reduce((sum, entry) => sum + entry.hours, 0);
  const dates = entries
    .map((entry) => entry.date)
    .filter((date) => date instanceof Date && !Number.isNaN(date.getTime()))
    .sort((a, b) => a - b);
  const startDate = dates[0] || null;
  const endDate = dates[dates.length - 1] || null;

  const rows = [
    {
      rowType: "parent",
      groupId: "pto",
      boardId: "pto",
      boardName: "PTO",
      workspaceName: "PTO",
      itemId: "pto",
      itemName: "PTO",
      personId: "total",
      personName: "Total",
      hours: totalHours,
      totalHours,
      date: null,
      startDate,
      endDate,
      dateLabel: "",
    },
  ];

  if (!collapsedGroups.has("pto")) {
    const sortedEntries = [...entries].sort((a, b) => {
      const aTime = a.date?.getTime?.() || 0;
      const bTime = b.date?.getTime?.() || 0;
      return aTime - bTime;
    });
    sortedEntries.forEach((entry) => {
      rows.push({
        ...entry,
        rowType: "entry",
        groupId: "pto",
      });
    });
  }

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
  syncMultiSelectState(workspaceFilter);
  syncMultiSelectState(boardFilter);
  syncMultiSelectState(teamFilter);
  syncMultiSelectState(userFilter);
  syncMultiSelectState(projectFilter);
};

const persistToken = (token) => {
  if (!token) {
    localStorage.removeItem(TOKEN_STORAGE_KEY);
    return;
  }
  localStorage.setItem(TOKEN_STORAGE_KEY, token);
};

const getApiToken = () => {
  const directToken = tokenInput?.value?.trim();
  if (directToken) return directToken;
  return localStorage.getItem(TOKEN_STORAGE_KEY) || "";
};

const updateMetrics = (mondayRows, ptoRowsInput = []) => {
  const uniqueBoards = new Set();
  const uniqueWorkspaces = new Set();
  const uniqueUsers = new Set();
  let totalHours = 0;
  const combinedRows = [...mondayRows, ...ptoRowsInput];

  mondayRows.forEach((row) => {
    uniqueBoards.add(row.boardId);
    if (row.workspaceName) {
      uniqueWorkspaces.add(row.workspaceName);
    }
  });

  combinedRows.forEach((row) => {
    uniqueUsers.add(row.personId);
    totalHours += row.hours;
  });

  if (metricWorkspaces) {
    metricWorkspaces.textContent = uniqueWorkspaces.size;
  }
  metricProjects.textContent = uniqueBoards.size;
  metricUsers.textContent = uniqueUsers.size;
  metricHours.textContent = totalHours.toFixed(2);
  metricRecords.textContent = combinedRows.length;
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
          type
        }
      }
    }
  `;

  const initialData = await mondayRequest(token, boardQuery, { boardId });
  const board = initialData.boards[0];
  if (!board) {
    return null;
  }

  const timeColumnIds = board.columns
    .filter((column) => column.type === "time_tracking")
    .map((column) => column.id);

  if (!timeColumnIds.length) {
    return {
      id: board.id,
      name: board.name,
      workspaceName: board.workspace?.name || "",
      columns: board.columns,
      items: [],
    };
  }

  const itemsQuery = `
    query ($boardId: ID!, $columnIds: [String!]) {
      boards(ids: [$boardId]) {
        items_page(limit: 200) {
          cursor
          items {
            id
            name
            created_at
            column_values(ids: $columnIds) {
              id
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

  const itemsData = await mondayRequest(token, itemsQuery, {
    boardId,
    columnIds: timeColumnIds,
  });
  const itemsBoard = itemsData.boards[0];
  if (!itemsBoard) {
    return {
      id: board.id,
      name: board.name,
      workspaceName: board.workspace?.name || "",
      columns: board.columns,
      items: [],
    };
  }

  const itemsPage = itemsBoard.items_page || { items: [], cursor: null };
  allItems = itemsPage.items || [];
  cursor = itemsPage.cursor;

  const nextQuery = `
    query ($cursor: String!, $columnIds: [String!]) {
      next_items_page(limit: 200, cursor: $cursor) {
        cursor
        items {
          id
          name
          created_at
          column_values(ids: $columnIds) {
            id
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
    const nextData = await mondayRequest(token, nextQuery, {
      cursor,
      columnIds: timeColumnIds,
    });
    const nextPage = nextData?.next_items_page;
    allItems = allItems.concat(nextPage?.items || []);
    cursor = nextPage?.cursor || null;
  }

  return {
    id: board.id,
    name: board.name,
    workspaceName: board.workspace?.name || "",
    columns: board.columns,
    items: allItems,
  };
};

const loadBoards = async (token, boardIds, onBoardLoaded) => {
  let boardIdsToFetch = Array.isArray(boardIds) ? boardIds : [];

  if (!boardIdsToFetch.length && preloadedBoards.length) {
    boardIdsToFetch = preloadedBoards.map((board) => board.id);
  }

  if (!boardIdsToFetch.length) {
    const boardsQuery = `
      query ($page: Int!) {
        boards(limit: ${BOARD_PAGE_LIMIT}, page: $page) {
          id
        }
      }
    `;
    const boards = await fetchBoardsPaged(token, boardsQuery);
    boardIdsToFetch = boards.map((board) => board.id);
  }

  if (boardIdsToFetch.length) {
    const results = await mapWithConcurrency(
      boardIdsToFetch,
      BOARD_FETCH_CONCURRENCY,
      async (boardId) => {
        const board = await fetchBoardItems(token, boardId);
        if (board && typeof onBoardLoaded === "function") {
          onBoardLoaded(board);
        }
        return board;
      }
    );
    return results.filter(Boolean);
  }

  return [];
};

const loadPreloadOptions = async (token) => {
  const preloadStart = Date.now();
  appendLoadSummary("Loading workspaces, boards, and users...");
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
      boards(limit: ${BOARD_PAGE_LIMIT}, page: $page) {
        id
        name
        workspace {
          id
          name
        }
      }
    }
  `;

  const usersPromise = ensureAllUsersLoaded(token).catch((error) => {
    appendDebug("ERROR", `User preload failed: ${error.message}`);
    return [];
  });

  const [workspaceData, boards] = await Promise.all([
    mondayRequest(token, workspacesQuery),
    fetchBoardsPaged(token, boardsQuery),
    usersPromise,
  ]);

  preloadedWorkspaces = (workspaceData.workspaces || []).map((workspace) => ({
    id: String(workspace.id),
    name: workspace.name || `Workspace ${workspace.id}`,
  }));

  preloadedBoards = (boards || []).map((board) => ({
    id: String(board.id),
    name: board.name,
    workspaceId: board.workspace?.id ? String(board.workspace.id) : "",
    workspaceName: board.workspace?.name || "",
  }));

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

  const sortedWorkspaces = [...preloadedWorkspaces].sort((a, b) =>
    a.name.localeCompare(b.name)
  );
  const sortedBoards = [...preloadedBoards].sort((a, b) =>
    formatBoardSummaryLabel(a).localeCompare(formatBoardSummaryLabel(b))
  );

  if (sortedWorkspaces.length) {
    appendLoadSummary(`Loaded ${sortedWorkspaces.length} workspaces.`);
    let workspaceIndex = 0;
    for (const workspace of sortedWorkspaces) {
      appendLoadSummary(`Workspace: ${workspace.name}`);
      workspaceIndex += 1;
      if (workspaceIndex % 25 === 0) {
        await yieldToBrowser();
      }
    }
  } else {
    appendLoadSummary("Loaded 0 workspaces.");
  }

  if (sortedBoards.length) {
    appendLoadSummary(`Loaded ${sortedBoards.length} boards.`);
    let boardIndex = 0;
    for (const board of sortedBoards) {
      appendLoadSummary(`Board: ${formatBoardSummaryLabel(board)}`);
      boardIndex += 1;
      if (boardIndex % 25 === 0) {
        await yieldToBrowser();
      }
    }
  } else {
    appendLoadSummary("Loaded 0 boards.");
  }

  updateLoadSummaryMetrics({
    loadTime: Date.now() - preloadStart,
    workspaces: preloadedWorkspaces.length,
    boards: preloadedBoards.length,
  });
};

const buildDashboard = async ({ token, boards, timeColumnIdsInput, dateRange }) => {
  allRows = [];
  parentRows = [];
  collapsedGroups.clear();
  const userIds = new Set();
  let totalRecordCount = 0;

  for (const board of boards) {
    const boardLabel = formatBoardSummaryLabel(board);
    let boardRecordCount = 0;
    const timeColumnIds = timeColumnIdsInput.length
      ? timeColumnIdsInput
      : board.columns
          .filter((column) => column.type === "time_tracking")
          .map((column) => column.id);

    if (!timeColumnIds.length) {
      appendLoadSummary(`${boardLabel}: 0 records loaded`);
      updateLoadSummaryMetrics({ items: totalRecordCount });
      await yieldToBrowser();
      continue;
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

          boardRecordCount += 1;
          allRows.push({
            source: "monday",
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

    const recordLabel = boardRecordCount === 1 ? "record" : "records";
    appendLoadSummary(`${boardLabel}: ${boardRecordCount} ${recordLabel} loaded`);
    totalRecordCount += boardRecordCount;
    updateLoadSummaryMetrics({ items: totalRecordCount });
    await yieldToBrowser();
  }

  await resolveUserNames(token, userIds);

  allRows = allRows.map((row) => {
    if (row.personId === "unknown") return row;
    const resolvedName = usersCache.get(row.personId);
    if (!resolvedName) return row;
    return { ...row, personName: resolvedName };
  });

  refreshPtoUserMapping();

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
    const boardName = row.boardName || "";
    const itemName = row.itemName || "";
    const personName = row.personName || "";
    const line = [
      row.boardId,
      `"${boardName.replace(/"/g, '""')}"`,
      row.itemId,
      `"${itemName.replace(/"/g, '""')}"`,
      row.personId,
      `"${personName.replace(/"/g, '""')}"`,
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

const exportUserTotals = () => {
  const displayState = buildUserTotalsDisplay(lastUserTotalsRows);
  if (displayState.message) {
    setStatus(displayState.message, "warning");
    return;
  }

  const { displayTotals, filterTeams, teamGroups, threshold, hasThreshold } =
    displayState;
  const periodLabel = currentPeriod || "all";
  const rowsToExport = [];

  if (filterTeams) {
    teamGroups.forEach((group) => {
      group.users.forEach((user) => {
        rowsToExport.push({ user, team: group.team.name || "" });
      });
    });
  } else {
    displayTotals.forEach((user) => {
      rowsToExport.push({ user, team: "" });
    });
  }

  if (!rowsToExport.length) {
    setStatus("No user totals to export.", "warning");
    return;
  }

  const escapeCsv = (value) =>
    `"${String(value ?? "").replace(/"/g, '""')}"`;

  const headers = ["user", "hours", "period", "threshold", "team"];
  const lines = [headers.join(",")];
  rowsToExport.forEach(({ user, team }) => {
    const line = [
      escapeCsv(user.name || ""),
      Number.isFinite(user.hours) ? user.hours.toFixed(2) : "0.00",
      escapeCsv(periodLabel),
      hasThreshold && Number.isFinite(threshold) ? threshold.toFixed(2) : "",
      escapeCsv(team),
    ];
    lines.push(line.join(","));
  });

  const blob = new Blob([lines.join("\n")], { type: "text/csv" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "user-totals-export.csv";
  link.click();
  URL.revokeObjectURL(url);
  setStatus("User totals exported.");
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
  syncMultiSelect(projectFilter);
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

  syncMultiSelect(workspaceFilter);
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

  syncMultiSelect(boardFilter);
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
  syncMultiSelect(userFilter);
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

const filterPtoEntries = (selectedUser, range) =>
  ptoRows.filter((row) => {
    if (selectedUser !== "all" && row.personId !== selectedUser) {
      const selectedName = normalizeKey(getUserNameById(selectedUser));
      if (!selectedName || row.personKey !== selectedName) {
        return false;
      }
    }
    if (!range) return true;
    if (!row.date) return false;
    return row.date >= range.start && row.date <= range.end;
  });

const applyFilters = () => {
  const projectId = projectFilter.value;
  const selectedUser = userFilter ? userFilter.value : "all";
  const period = currentPeriod;
  const range = getPeriodRange(period, currentDate);
  const showPtoOnly = Boolean(ptoOnlyToggle?.checked);
  const showParents = !showPtoOnly && (rawEntriesToggle ? !rawEntriesToggle.checked : true);
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

  const ptoFilteredEntries = filterPtoEntries(selectedUser, range);
  const combinedEntries = filteredEntries.concat(ptoFilteredEntries);
  exportRows = showPtoOnly ? ptoFilteredEntries : combinedEntries;
  updateMetrics(filteredEntries, ptoFilteredEntries);
  renderUserTotals(combinedEntries);
  let tableRows = [];
  if (showPtoOnly) {
    tableRows = buildPtoDisplayRows(ptoFilteredEntries, false);
  } else {
    const mondayRows = buildDisplayRows(
      filteredEntries,
      showParents,
      useFilteredTotals
    );
    const ptoRowsDisplay = buildPtoDisplayRows(ptoFilteredEntries, showParents);
    tableRows = mondayRows.concat(ptoRowsDisplay);
  }
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
  if (!isAllData && !endDateInput.value) {
    endDateInput.value = formatInputDateValue(new Date());
  }
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
  const token = getApiToken();
  if (!token) {
    setStatus("Set a Monday API token in Admin first.", "error");
    return;
  }

  clearLoadSummary();
  persistToken(token);
  setStatus("Loading workspaces, boards, and users...");
  setLoading(true);
  try {
    await loadPreloadOptions(token);
    setStatus("Workspaces, boards, and users loaded.");
  } catch (error) {
    setStatus(`Error: ${error.message}`, "error");
    appendLoadSummary(`Workspace preload failed: ${error.message}`);
  } finally {
    setLoading(false);
  }
});

workspaceFilter.addEventListener("change", () => {
  populateBoardFilter(preloadedBoards, getSelectedValues(workspaceFilter));
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  const token = getApiToken();
  const timeColumnIdsInput = [];
  const selectedWorkspaces = workspaceFilter
    ? getSelectedValues(workspaceFilter)
    : ["all"];

  if (!token) {
    setStatus("Set a Monday API token in Admin first.", "error");
    setLoading(false);
    return;
  }

  persistToken(token);
  ensureAllUsersLoaded(token).catch((error) => {
    appendDebug("ERROR", `User preload failed: ${error.message}`);
  });
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
  appendLoadSummary("Loading hours data...");
  updateLoadSummaryMetrics({ items: 0, loadTime: null, boards: 0, workspaces: 0 });
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

    const loadSummaryWorkspaceNames = new Set();
    let loadedBoardCount = 0;
    const boards = await loadBoards(token, boardIds, (board) => {
      loadedBoardCount += 1;
      if (board.workspaceName) {
        loadSummaryWorkspaceNames.add(board.workspaceName);
      }
      const itemCount = Array.isArray(board.items) ? board.items.length : 0;
      appendLoadSummary(
        `Fetched ${formatBoardSummaryLabel(board)} (${itemCount} items)`
      );
      updateLoadSummaryMetrics({
        boards: loadedBoardCount,
        workspaces: loadSummaryWorkspaceNames.size,
      });
    });
    if (!boards.length) {
      setStatus("No boards found. Check your access or IDs.", "warning");
      return;
    }

    const workspaceNames = boards
      .map((board) => board.workspaceName || "")
      .filter(Boolean);
    const workspaceCount = new Set(workspaceNames).size;
    updateLoadSummaryMetrics({
      boards: boards.length,
      workspaces: workspaceCount,
    });

    await buildDashboard({
      token,
      boards,
      timeColumnIdsInput,
      dateRange,
    });

    setStatus("Dashboard ready. Export is available below.");
    appendLoadSummary("Hours data load complete.");
    setFiltersEnabled(true);
    updatePeriodInputs();
  } catch (error) {
    setStatus(`Error: ${error.message}`, "error");
    appendLoadSummary(`Hours load failed: ${error.message}`);
  } finally {
    isMeasuringLoad = false;
    const loadDuration = Date.now() - loadStartTime;
    updateLastLoadMetrics(loadDuration, loadRequestCount);
    updateLoadSummaryMetrics({ loadTime: loadDuration });
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

if (loadSummarySearchInput && loadSummarySearchBtn) {
  const runLoadSummarySearch = () => {
    updateLoadSummarySearch(loadSummarySearchInput.value);
  };
  loadSummarySearchBtn.addEventListener("click", runLoadSummarySearch);
  loadSummarySearchInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      runLoadSummarySearch();
    }
  });
  loadSummarySearchInput.addEventListener("input", () => {
    updateLoadSummarySearch(loadSummarySearchInput.value);
  });
}

if (loadSummaryPrevBtn) {
  loadSummaryPrevBtn.addEventListener("click", () => {
    if (!loadSummaryMatches.length) return;
    setCurrentLoadSummaryMatch(loadSummaryMatchIndex - 1);
  });
  loadSummaryPrevBtn.disabled = true;
}

if (loadSummaryNextBtn) {
  loadSummaryNextBtn.addEventListener("click", () => {
    if (!loadSummaryMatches.length) return;
    setCurrentLoadSummaryMatch(loadSummaryMatchIndex + 1);
  });
  loadSummaryNextBtn.disabled = true;
}

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

if (adminToggleBtn) {
  adminToggleBtn.addEventListener("click", openAdminModal);
}

if (adminModal) {
  adminModal.addEventListener("click", (event) => {
    const target = event.target;
    if (target?.dataset?.close) {
      closeAdminModal();
    }
  });
}

if (adminForm) {
  adminForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const mondayToken = adminMondayTokenInput?.value.trim() || "";
    persistToken(mondayToken);
    if (tokenInput) {
      tokenInput.value = mondayToken;
    }

    const settings = {
      ptoJsonUrl: ptoJsonUrlInput?.value.trim() || "",
      ptoRememberUpload: Boolean(ptoRememberUploadToggle?.checked),
      ptoAutoLoad: Boolean(ptoAutoLoadToggle?.checked),
    };

    persistAdminSettings(settings);
    if (!settings.ptoRememberUpload) {
      localStorage.removeItem(PTO_CACHE_STORAGE_KEY);
    }
    setAdminStatus("Settings saved.", "muted");
    if (settings.ptoAutoLoad) {
      loadPtoFromSettings(settings);
    }
  });
}

if (adminLoadPtoBtn) {
  adminLoadPtoBtn.addEventListener("click", () => {
    const settings = loadAdminSettings();
    loadPtoFromSettings(settings);
  });
}

if (ptoCsvUploadInput) {
  ptoCsvUploadInput.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const text = await file.text();
      const loaded = loadPtoFromText(text, `Upload: ${file.name}`);
      if (loaded && ptoRememberUploadToggle?.checked) {
        savePtoCache(text);
      }
    } catch (error) {
      setAdminStatus(`Failed to read CSV: ${error.message}`, "error");
    }
  });
}

if (thresholdSettingsBtn) {
  thresholdSettingsBtn.addEventListener("click", openThresholdModal);
}

if (thresholdModal) {
  thresholdModal.addEventListener("click", (event) => {
    const target = event.target;
    if (target && target.closest("[data-close]")) {
      closeThresholdModal();
    }
  });
}

if (thresholdForm) {
  thresholdForm.addEventListener("submit", (event) => {
    event.preventDefault();
    const values = {
      day: sanitizeThresholdValue(
        thresholdDayInput?.value,
        DEFAULT_THRESHOLDS.day
      ),
      week: sanitizeThresholdValue(
        thresholdWeekInput?.value,
        DEFAULT_THRESHOLDS.week
      ),
      month: sanitizeThresholdValue(
        thresholdMonthInput?.value,
        DEFAULT_THRESHOLDS.month
      ),
    };
    persistThresholds(values);
    closeThresholdModal();
    applyFilters();
  });
}

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && thresholdModal?.classList.contains("active")) {
    closeThresholdModal();
  }
});

exportBtn.addEventListener("click", exportCsv);
if (userTotalsExportBtn) {
  userTotalsExportBtn.addEventListener("click", exportUserTotals);
}

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
if (teamFilter) {
  teamFilter.addEventListener("change", applyFilters);
}
if (rawEntriesToggle) {
  rawEntriesToggle.addEventListener("change", applyFilters);
}
if (ptoOnlyToggle) {
  ptoOnlyToggle.addEventListener("change", () => {
    if (ptoOnlyToggle.checked && rawEntriesToggle) {
      rawEntriesToggle.checked = true;
    }
    applyFilters();
  });
}
if (showAllUsersToggle) {
  showAllUsersToggle.addEventListener("change", applyFilters);
}
if (underThresholdOnly) {
  underThresholdOnly.addEventListener("change", applyFilters);
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
resetLoadSummaryMetrics();
initMultiSelects();

const setAdminTokenVisibility = (isVisible) => {
  if (!adminMondayTokenInput || !adminTokenVisibilityBtn) return;
  adminMondayTokenInput.type = isVisible ? "text" : "password";
  adminTokenVisibilityBtn.classList.toggle("is-visible", isVisible);
  adminTokenVisibilityBtn.setAttribute("aria-pressed", String(isVisible));
  adminTokenVisibilityBtn.setAttribute(
    "aria-label",
    isVisible ? "Hide API token" : "Show API token"
  );
};

if (adminTokenVisibilityBtn && adminMondayTokenInput) {
  adminTokenVisibilityBtn.addEventListener("click", () => {
    setAdminTokenVisibility(adminMondayTokenInput.type === "password");
    adminMondayTokenInput.focus({ preventScroll: true });
  });
  setAdminTokenVisibility(false);
}

setFiltersEnabled(false);

const preloadFromStoredToken = async () => {
  const savedToken = localStorage.getItem(TOKEN_STORAGE_KEY);
  if (!savedToken) return;
  if (tokenInput) {
    tokenInput.value = savedToken;
  }
  clearLoadSummary();
  setStatus("Saved token found. Loading workspaces, boards, and users...");
  setLoading(true);
  try {
    await loadPreloadOptions(savedToken);
    setStatus("Workspaces, boards, and users loaded.");
  } catch (error) {
    setStatus(`Error: ${error.message}`, "error");
    appendLoadSummary(`Workspace preload failed: ${error.message}`);
  } finally {
    setLoading(false);
  }
};

preloadFromStoredToken();

const initAdminSettings = () => {
  const settings = loadAdminSettings();
  if (settings.ptoAutoLoad) {
    loadPtoFromSettings(settings);
    return;
  }
  if (settings.ptoRememberUpload) {
    const loaded = loadPtoFromCache();
    if (loaded) return;
  }
  setAdminStatus("PTO data is not loaded.", "muted");
};

initAdminSettings();
