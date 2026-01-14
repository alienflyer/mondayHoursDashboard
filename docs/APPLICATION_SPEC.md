# Monday Hours Dashboard - Application Specification

## Purpose and goals
This application loads time tracking data from Monday.com, combines it with PTO and holiday time from Projector (JSON report), and presents a dashboard for filtering, exporting, and analyzing hours. It is designed to run locally in a browser with a lightweight local proxy for API requests.

Key goals:
- Show hours per user and per board with flexible time period filters.
- Merge PTO and holiday time into user totals for accurate weekly/monthly analysis.
- Provide an admin page for configuring tokens and PTO sources without editing code.
- Remain local-only with no external hosting required.

## Architecture overview
- Frontend: static HTML/CSS/JavaScript (no frameworks).
- Local proxy: `dev_server.py` handles Monday API requests and Projector PTO JSON proxying.
- Storage: browser `localStorage` for settings, tokens, and optional PTO CSV cache.

Diagrams:
- `docs/diagrams/architecture.mmd` + `docs/diagrams/architecture.png`
- `docs/diagrams/data-flow.mmd` + `docs/diagrams/data-flow.png`
- `docs/diagrams/ui-map.mmd` + `docs/diagrams/ui-map.png`

## File structure
- `index.html`: application markup and modal dialogs.
- `styles.css`: global styles, layout, cards, tables, custom multi-select UI.
- `app.js`: all application logic (data fetching, filtering, rendering, PTO merge).
- `dev_server.py`: local proxy for Monday and Projector requests.
- `docs/APPLICATION_SPEC.md`: this document.
- `docs/diagrams/*`: mermaid diagrams and PNG exports.

## Running the app
1. Start the local proxy server:
   - `python3 dev_server.py`
2. Open the app in a browser:
   - `http://localhost:8000`

Notes:
- Do not use `python3 -m http.server`. It does not support POST and will return 501.
- The proxy is required for CORS when accessing Monday and Projector.

## Data sources
### Monday.com
- Endpoint: `https://api.monday.com/v2`
- Accessed via local proxy at `/api/monday`.
- Authorization: Monday API token provided by the user.
- Query patterns:
  - Workspaces list.
  - Boards list (paged).
  - Users list (paged, includes teams when available).
  - Board items with time tracking entries.

### Projector PTO JSON
- Endpoint: Projector report JSON URL.
- Accessed via local proxy at `/api/pto?url=...`.
- Report fields used:
  - `Resource`: user name (includes numeric ID in parentheses).
  - `PersonHours`: hours amount.
  - `Day`: date string in `YYYY/MM/DD` format.
  - `ResourceBusinessUnits`: team label (includes numeric ID).
  - `TimeOffReason`: holiday or PTO reason.

### PTO CSV fallback
- User can upload a CSV file when Projector URL is unavailable.
- Optional cached CSV is stored in `localStorage` and auto-loaded at startup when enabled.

## Local proxy details
### `dev_server.py`
- Serves static assets and provides:
  - `POST /api/monday` to forward Monday GraphQL requests.
  - `GET /api/pto?url=...` to fetch Projector JSON and return it with CORS headers.

CORS:
- Adds `Access-Control-Allow-Origin: *`.
- Supports OPTIONS for `/api/monday`.

## UI layout and components
### Header
- Title and subtitle.
- Local-only badge.
- Theme toggle button.
- Admin button (opens admin modal).
- Debug toggle button (shows/hides debug panel).

### Connect to Monday card
- Monday API token field (masked with show/hide button).
- Workspace multi-select.
- Board multi-select.
- Date range controls (All dates toggle, Start/End date). When All dates is unchecked, End date auto-fills with today.
- Status message.
- Actions:
  - Load workspaces.
  - Load hours data.
  - Export CSV.

### Load summary card
- Placed beside Connect to Monday (50/50 split) as its own card.
- Metrics:
  - Load time (last preload or hours load).
  - Loaded items (hours records).
  - Loaded workspaces.
  - Loaded boards.
- Search:
  - Search input with a magnifying-glass button.
  - Previous/next match buttons (left/right arrows).
- Summary list:
  - Shows high-level progress for workspace/board preload and hours loads.
  - Lists each workspace and board as they are loaded.
  - During hours load, shows per-board record counts (e.g., `Workspace - Board: 3 records loaded`).
  - Updates incrementally as boards are fetched and processed.
  - Separate from the verbose debug log.

### Filters card
- User single-select with searchable dropdown.
- Boards single-select with searchable dropdown.
- Time period select (all/day/week/month/year), default weekly.
- Cycle navigation with previous/next and jump-to-current period.
- Checkbox row:
  - Show raw time entries only (no parent rows).
  - Show PTO entries only (only PTO in table).

### Metrics + Hours table
- Metrics: total workspaces, total boards, total users, total hours, records count.
- Table columns: Workspace, Board, Item, User, Duration, Start, End.
- Parent rows represent aggregated items and are collapsible.
- PTO adds a parent row labeled "PTO" with entry rows beneath it.

### User totals card
- Teams filter (multi-select, default "No grouping").
- Show all users toggle.
- Below threshold only toggle.
- Threshold settings button.
- Totals grouped by team when enabled.
- Thermometer bars indicate progress against thresholds with dynamic color progress.

### Debug panel
- Request/response logs.
- Copy and clear buttons.
- Last load metrics.

### Modals
- Admin modal:
  - Monday API token field.
  - Projector PTO JSON URL field.
  - PTO CSV upload and cache options.
  - Auto-load PTO on startup.
  - Load PTO now.
- Threshold modal:
  - Day/week/month hour thresholds.

## Custom multi-select component
- Applied to:
  - Multi-selects: workspace filter, board filter, teams filter.
  - Searchable single selects: user filter and boards filter.
- Behavior:
  - Chips display selected values.
  - "All" option acts as default for multi-selects.
  - Search box appears inside dropdown when `data-searchable="true"`.
  - Selected options remain visible in the dropdown.

## Data model
### Monday entry row
Fields (example):
- `source`: "monday"
- `rowType`: "entry"
- `groupId`: `boardId:itemId:columnId`
- `boardId`, `boardName`, `workspaceName`
- `itemId`, `itemName`
- `personId`, `personName`
- `hours`
- `date`, `startDate`, `endDate`

### Parent row (Monday)
Fields:
- `rowType`: "parent"
- `groupId`: same as entry group
- `personName`: "Total"
- `hours`: aggregated

### PTO row
Fields:
- `source`: "pto"
- `rowType`: "pto" (rendered as entry)
- `groupId`: "pto"
- `boardId`: "pto"
- `boardName`: "PTO"
- `workspaceName`: "PTO"
- `itemName`: includes `TimeOffReason` and team label
- `personId`, `personName`
- `personKey`: normalized name for matching
- `hours`
- `date`, `startDate`, `endDate`

## Key workflows
### Load workspaces, boards, users
1. User enters Monday token.
2. Click "Load workspaces".
3. App fetches:
   - Workspaces
   - Boards (paged)
   - Users (paged, with teams)
4. Filters are populated.

### Load hours data
1. User submits form.
2. Boards are fetched (respecting workspace/board selection).
3. Items and time tracking entries are fetched.
4. Entries are normalized into `allRows` and `parentRows`.
5. Filters are applied and the dashboard renders.

### Load PTO JSON
1. Admin "Load PTO now".
2. App calls `/api/pto?url=...` via local proxy.
3. JSON rows are parsed using Projector field names.
4. PTO rows are built and merged with Monday entries.

### CSV fallback
1. Admin uploads a CSV.
2. CSV is parsed; if "Remember uploaded CSV" is checked, it is cached.
3. If PTO JSON fails and cached CSV exists, app uses the cache.

## Filtering and rendering rules
- Board filter: single select (All boards or one board).
- User filter: single select (All users or one user).
- Period filter: all/day/week/month/year. Date range is computed per period.
- Default period: weekly.
- Raw entries only: hides parent rows, shows entry rows.
- PTO entries only: table shows only PTO entries, no Monday rows.
- Team grouping in user totals:
  - Default "No grouping" shows flat list.
  - Selecting teams groups totals by team and filters visible users.

## User totals and thresholds
- Totals are computed from combined Monday + PTO entries.
- "Show all users" includes users with zero hours.
- Thresholds per period (day/week/month) stored in `localStorage`.
- Thermometer colors:
  - 0-40%: red
  - 40-75%: yellow
  - 75-100%: green
  - Gradient only reveals colors once progress crosses each threshold.

## Export
- Exports current filtered rows to CSV.
- Includes PTO rows when PTO is loaded (unless PTO-only filter is active).

## Local storage keys
- `mondayApiToken`: Monday token.
- `mondayTheme`: dark/light theme.
- `mondayThresholds`: thresholds for day/week/month.
- `mondayAdminSettings`: admin configuration for PTO.
- `mondayPtoCache`: cached PTO CSV (optional).

## Error handling
- Monday requests have retry/backoff and rate limit detection.
- PTO load errors display a status in the Admin modal.
- If PTO JSON fails and cache is available, cache is used.

## Security considerations
- Tokens are stored in browser `localStorage` (not encrypted).
- Projector URL should not be exposed in shared environments.
- Proxy allows any URL by default; restrict to Projector domain if needed.

## Rebuild from scratch (high-level steps)
1. Create base HTML with cards, tables, filters, modals, and debug panel.
2. Implement CSS for layout, cards, forms, tables, and custom selects.
3. Implement Monday API loader with pagination and rate limit handling.
4. Normalize Monday time tracking data into rows and parent rows.
5. Build filters, period navigation, and table rendering with collapsible groups.
6. Implement user totals with thresholds and team grouping.
7. Add PTO JSON parsing with Projector field mapping.
8. Merge PTO into totals and add PTO table parent row.
9. Add Admin modal with local settings and PTO load controls.
10. Add local proxy server to handle CORS for Monday and Projector.
11. Add diagrams and update documentation.
