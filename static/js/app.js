(() => {
  "use strict";

  const PREFERRED_CITIES = new Set(["bellevue", "kirkland", "seattle"]);
  const VALID_VIEWS = new Set([
    "dashboard",
    "all",
    "new",
    "preferred",
    "removed",
    "spreadsheets",
  ]);
  const VIEW_COPY = {
    dashboard: ["Gig dashboard", "Current opportunities across every loaded month."],
    all: ["All current gigs", "Every active opportunity across the loaded spreadsheets."],
    new: ["New gigs", "Opportunities newly added in the latest spreadsheet update."],
    preferred: ["Preferred cities", "Gigs in Bellevue, Kirkland, and Seattle."],
    removed: ["Removed gigs", "Opportunities no longer present in the latest calendar scrape."],
    spreadsheets: ["Spreadsheets", "Open or download the source-of-truth Excel workbooks."],
  };
  const REFRESH_POLL_INTERVAL = 1500;
  const REFRESH_NOTICE_KEY = "us-event-refresh-notice-v1";
  const REFRESH_WATCH_KEY = "us-event-refresh-watching-v1";
  const REFRESH_STATES = new Set(["idle", "running", "success", "error"]);
  const refreshState = {
    activeStatus: "idle",
    autoDismissTimer: null,
    dismissedFor: "",
    failureCount: 0,
    generation: 0,
    pollTimer: null,
    reloadScheduled: false,
    startPending: false,
    watching: false,
  };
  const recruiterState = {
    selectedGigIds: new Set(),
    isOpen: false,
    returnFocus: null,
    previousBodyOverflow: "",
    closeTimer: null,
    copyTimer: null,
  };

  const dom = {
    body: document.body,
    sidebar: document.querySelector("#sidebar"),
    sidebarToggle: document.querySelector("#mobile-menu-toggle"),
    sidebarOverlay: document.querySelector("#sidebar-overlay"),
    navLinks: [...document.querySelectorAll("[data-view]")],
    monthLinks: [...document.querySelectorAll("[data-month]")],
    refresh: document.querySelector("#refresh-data"),
    refreshButtonLabel: document.querySelector("#refresh-button-label"),
    refreshJobStatus: document.querySelector("#refresh-job-status"),
    refreshIndicator: document.querySelector(".refresh-status__indicator"),
    lastRefreshed: document.querySelector("#last-refreshed"),
    notification: document.querySelector("#refresh-notification"),
    notificationTitle: document.querySelector("#refresh-notification-title"),
    notificationMessage: document.querySelector("#refresh-notification-message"),
    notificationDismiss: document.querySelector("#refresh-notification-dismiss"),
    detailsToggle: document.querySelector("#refresh-details-toggle"),
    errorDetails: document.querySelector("#refresh-error-details"),
    errorOutput: document.querySelector("#refresh-error-output"),
    filters: document.querySelector("#gig-filters"),
    search: document.querySelector("#global-search"),
    month: document.querySelector("#month-filter"),
    date: document.querySelector("#date-filter"),
    city: document.querySelector("#city-filter"),
    status: document.querySelector("#status-filter"),
    account: document.querySelector("#account-filter, #store-filter"),
    time: document.querySelector("#time-filter"),
    preferred: document.querySelector("#preferred-filter"),
    resetButtons: [...document.querySelectorAll("#reset-filters, #clear-filters, [data-reset-filters]")],
    cityQuick: [...document.querySelectorAll("[data-city-quick]")],
    table: document.querySelector("#gig-table, #gigs-table"),
    tbody: document.querySelector("#gig-table-body, #gigs-table-body"),
    empty: document.querySelector("#empty-state"),
    resultsCount: document.querySelector("#results-count"),
    pageSize: document.querySelector("#page-size"),
    previousPage: document.querySelector("#prev-page"),
    nextPage: document.querySelector("#next-page"),
    pageInfo: document.querySelector("#page-info, #page-status"),
    pagination: document.querySelector("#pagination"),
    sortButtons: [...document.querySelectorAll(".sort-button[data-sort]")],
    spreadsheets: document.querySelector("#spreadsheets"),
    gigsView: document.querySelector("#gigs-view, [data-view-panel='gigs']"),
    viewTitle: document.querySelector("#view-title"),
    viewDescription: document.querySelector("#view-description"),
    generateRecruiterText: document.querySelector("#generate-recruiter-text"),
    selectedGigCount: document.querySelector("#selected-gig-count"),
    recruiterDrawer: document.querySelector("#recruiter-drawer"),
    recruiterDrawerBackdrop: document.querySelector("#recruiter-drawer-backdrop"),
    recruiterCloseControls: [...document.querySelectorAll("[data-recruiter-close]")],
    recruiterDrawerCount: document.querySelector("#recruiter-drawer-count"),
    recruiterMessage: document.querySelector("#recruiter-message"),
    copyRecruiterText: document.querySelector("#copy-recruiter-text"),
    copyRecruiterTextLabel: document.querySelector("#copy-recruiter-text-label"),
    recruiterCopyFeedback: document.querySelector("#recruiter-copy-feedback"),
  };

  setupRefreshControl();

  if (!dom.tbody) {
    setupRecruiterTextUnavailable();
    setupSidebarOnly();
    return;
  }

  const query = new URLSearchParams(window.location.search);
  const rawRows = [...dom.tbody.querySelectorAll("tr.gig-row")];
  const rowData = rawRows.map((element, index) => createRowRecord(element, index));
  const rowsByGigId = new Map(
    rowData
      .filter((row) => row.gigId && row.status !== "REMOVED" && row.selectCheckbox)
      .map((row) => [row.gigId, row]),
  );
  const state = {
    view: normalizeView(query.get("view") || dom.body.dataset.initialView || "dashboard"),
    quickCity: normalizeQuickCity(query.get("quick") || "all"),
    sortKey: normalizeSortKey(query.get("sort") || "default"),
    sortDirection: query.get("dir") === "desc" ? "desc" : "asc",
    page: positiveInteger(query.get("page"), 1),
    pageSize: positiveInteger(query.get("page_size"), positiveInteger(dom.pageSize?.value, 25)),
  };

  hydrateSelects();
  initializeControlsFromUrl(query);
  bindEvents();
  setupRecruiterText();
  setupSidebarOnly();
  applyDashboardState({ updateUrl: false });

  function createRowRecord(element, index) {
    const value = (...keys) => {
      for (const key of keys) {
        const raw = element.dataset[key];
        if (raw !== undefined && raw !== null && String(raw).trim() !== "") {
          return String(raw).trim();
        }
      }
      return "";
    };
    const cell = (...fields) => {
      for (const field of fields) {
        const found = element.querySelector(`[data-field="${field}"]`);
        if (found) return found.textContent.trim();
      }
      return "";
    };

    const city = value("city") || cell("city");
    const status = (value("status") || cell("status")).toUpperCase();
    const preferredRaw = value("preferred", "preferredArea") || cell("preferred", "preferred-area");
    const preferred = isTruthy(preferredRaw) || PREFERRED_CITIES.has(city.toLowerCase());
    const date = normalizeDate(value("date") || cell("date"));
    const time = value("time", "startTime") || cell("time", "start-time");
    const timeSort = normalizeTime(value("timeSort", "startTimeSort") || time);
    const store = value("store", "account", "storeAccount") || cell("store", "account");
    const month = normalizeMonth(value("month")) || date.slice(0, 7);
    const priorityRaw = Number.parseInt(value("priority", "locationPriority"), 10);
    const priority = Number.isFinite(priorityRaw)
      ? priorityRaw
      : ({ bellevue: 1, kirkland: 2, seattle: 3 }[city.toLowerCase()] || 9);
    const search = (value("search") || element.textContent).replace(/\s+/g, " ").trim().toLowerCase();
    const selectCheckbox = element.querySelector(".gig-select-checkbox");

    return {
      element,
      gigId: value("gigId"),
      selectCheckbox,
      index: positiveInteger(value("index"), index + 1) - 1,
      month,
      date,
      day: value("day") || cell("day"),
      time,
      timeSort,
      city,
      cityKey: city.toLowerCase(),
      status,
      store,
      storeNumber: value("storeNumber") || cell("store-number"),
      preferred,
      priority,
      raw: value("raw", "rawListing") || cell("raw", "raw-listing"),
      firstSeen: value("firstSeen") || cell("first-seen"),
      lastSeen: value("lastSeen") || cell("last-seen"),
      search,
    };
  }

  function bindEvents() {
    let searchTimer;
    dom.filters?.addEventListener("submit", (event) => event.preventDefault());

    dom.search?.addEventListener("input", () => {
      window.clearTimeout(searchTimer);
      searchTimer = window.setTimeout(() => {
        state.page = 1;
        applyDashboardState();
      }, 120);
    });

    [dom.month, dom.date, dom.status, dom.account, dom.time, dom.preferred].forEach((control) => {
      control?.addEventListener("change", () => {
        state.page = 1;
        applyDashboardState();
      });
    });

    dom.city?.addEventListener("change", () => {
      state.quickCity = "all";
      state.page = 1;
      applyDashboardState();
    });

    dom.pageSize?.addEventListener("change", () => {
      state.pageSize = positiveInteger(dom.pageSize.value, 25);
      state.page = 1;
      applyDashboardState();
    });

    dom.resetButtons.forEach((button) => {
      button.addEventListener("click", (event) => {
        event.preventDefault();
        clearFilterControls();
        state.quickCity = "all";
        state.sortKey = "default";
        state.sortDirection = "asc";
        state.page = 1;
        applyDashboardState();
        dom.search?.focus();
      });
    });

    dom.cityQuick.forEach((button) => {
      button.addEventListener("click", (event) => {
        event.preventDefault();
        state.quickCity = normalizeQuickCity(button.dataset.cityQuick || "all");
        setControlValue(dom.city, "");
        state.page = 1;
        applyDashboardState();
      });
    });

    dom.navLinks.forEach((link) => {
      link.addEventListener("click", (event) => {
        if (link.dataset.month) return;
        const requested = link.dataset.view;
        if (!VALID_VIEWS.has(requested)) return;
        event.preventDefault();
        clearFilterControls();
        state.quickCity = "all";
        state.view = requested;
        state.page = 1;
        applyDashboardState();
        closeSidebar();
      });
    });

    dom.monthLinks.forEach((link) => {
      link.addEventListener("click", (event) => {
        const month = normalizeMonth(link.dataset.month || "");
        if (!month) return;
        event.preventDefault();
        clearFilterControls();
        state.quickCity = "all";
        setControlValue(dom.month, month);
        state.view = normalizeView(link.dataset.view || "all");
        state.page = 1;
        applyDashboardState();
        closeSidebar();
      });
    });

    dom.sortButtons.forEach((button) => {
      button.addEventListener("click", () => {
        const sortKey = normalizeSortKey(button.dataset.sort);
        if (sortKey === state.sortKey && sortKey !== "default") {
          state.sortDirection = state.sortDirection === "asc" ? "desc" : "asc";
        } else {
          state.sortKey = sortKey;
          state.sortDirection = "asc";
        }
        state.page = 1;
        applyDashboardState();
      });
    });

    dom.previousPage?.addEventListener("click", () => {
      if (state.page <= 1) return;
      state.page -= 1;
      applyDashboardState();
      scrollTableIntoView();
    });

    dom.nextPage?.addEventListener("click", () => {
      state.page += 1;
      applyDashboardState();
      scrollTableIntoView();
    });

    window.addEventListener("popstate", () => {
      initializeControlsFromUrl(new URLSearchParams(window.location.search));
      applyDashboardState({ updateUrl: false });
    });

    document.addEventListener("keydown", (event) => {
      if (event.key !== "/" || event.metaKey || event.ctrlKey || event.altKey) return;
      const target = event.target;
      if (target instanceof HTMLElement && target.matches("input, select, textarea, [contenteditable='true']")) return;
      event.preventDefault();
      dom.search?.focus();
    });
  }

  function applyDashboardState({ updateUrl = true } = {}) {
    updateViewPresentation();

    if (state.view === "spreadsheets") {
      rawRows.forEach((row) => { row.hidden = true; });
      updateActiveNavigation();
      updateUrl && writeUrl();
      return;
    }

    const filters = currentFilters();
    const matching = rowData.filter((row) => matches(row, filters));
    const ordered = [...matching].sort(buildComparator());
    const totalPages = Math.max(1, Math.ceil(ordered.length / state.pageSize));
    state.page = Math.min(Math.max(state.page, 1), totalPages);
    const pageStart = (state.page - 1) * state.pageSize;
    const visible = new Set(ordered.slice(pageStart, pageStart + state.pageSize));
    const matchingSet = new Set(matching);

    const fragment = document.createDocumentFragment();
    ordered.forEach((row) => fragment.appendChild(row.element));
    rowData.filter((row) => !matchingSet.has(row)).forEach((row) => fragment.appendChild(row.element));
    dom.tbody.appendChild(fragment);
    rowData.forEach((row) => { row.element.hidden = !visible.has(row); });

    const first = ordered.length ? pageStart + 1 : 0;
    const last = Math.min(pageStart + state.pageSize, ordered.length);
    if (dom.resultsCount) {
      dom.resultsCount.textContent = ordered.length
        ? `Showing ${number(first)}–${number(last)} of ${number(ordered.length)} gigs`
        : "0 gigs match these filters";
      dom.resultsCount.setAttribute("aria-live", "polite");
    }
    if (dom.pageInfo) dom.pageInfo.textContent = `Page ${number(state.page)} of ${number(totalPages)}`;
    if (dom.previousPage) dom.previousPage.disabled = state.page <= 1;
    if (dom.nextPage) dom.nextPage.disabled = state.page >= totalPages || ordered.length === 0;
    if (dom.pagination) dom.pagination.hidden = ordered.length === 0;
    if (dom.empty) dom.empty.hidden = ordered.length !== 0;
    if (dom.table) dom.table.classList.toggle("has-no-results", ordered.length === 0);

    updateSortIndicators();
    updateQuickCities();
    updateActiveNavigation();
    updateUrl && writeUrl();
  }

  function currentFilters() {
    return {
      search: normalizeText(dom.search?.value),
      month: normalizeMonth(dom.month?.value || ""),
      date: normalizeDate(dom.date?.value || ""),
      city: normalizeText(dom.city?.value),
      status: normalizeText(dom.status?.value).toUpperCase(),
      account: normalizeText(dom.account?.value),
      time: normalizeText(dom.time?.value),
      preferred: normalizeText(dom.preferred?.value),
    };
  }

  function matches(row, filters) {
    if ((state.view === "dashboard" || state.view === "all") && row.status === "REMOVED") return false;
    if (state.view === "new" && row.status !== "NEW") return false;
    if (state.view === "preferred" && (!row.preferred || row.status === "REMOVED")) return false;
    if (state.view === "removed" && row.status !== "REMOVED") return false;

    if (filters.search && !row.search.includes(filters.search)) return false;
    if (filters.month && row.month !== filters.month) return false;
    if (filters.date && row.date !== filters.date) return false;
    if (filters.city && row.cityKey !== filters.city) return false;
    if (filters.status === "CURRENT" && row.status === "REMOVED") return false;
    if (filters.status && filters.status !== "CURRENT" && row.status !== filters.status) return false;
    if (filters.account && !row.store.toLowerCase().includes(filters.account)) return false;
    if (filters.time && !sameTime(row, filters.time)) return false;
    if (filters.preferred && filters.preferred !== "all") {
      const expectsPreferred = isTruthy(filters.preferred);
      if (row.preferred !== expectsPreferred) return false;
    }

    if (state.quickCity === "other" && row.preferred) return false;
    if (state.quickCity !== "all" && state.quickCity !== "other" && row.cityKey !== state.quickCity) return false;
    return true;
  }

  function sameTime(row, filter) {
    const normalized = normalizeTime(filter);
    return row.time.toLowerCase() === filter || row.timeSort === normalized;
  }

  function buildComparator() {
    const defaultCompare = (a, b) => (
      compare(a.date, b.date)
      || compare(a.timeSort, b.timeSort)
      || compareNumber(a.priority, b.priority)
      || compare(a.cityKey, b.cityKey)
      || compare(a.store.toLowerCase(), b.store.toLowerCase())
      || compareNumber(a.index, b.index)
    );

    if (state.sortKey === "default") return defaultCompare;
    const direction = state.sortDirection === "desc" ? -1 : 1;
    return (a, b) => {
      const first = compare(sortValue(a, state.sortKey), sortValue(b, state.sortKey));
      return first ? first * direction : defaultCompare(a, b);
    };
  }

  function sortValue(row, key) {
    const values = {
      date: row.date,
      day: row.day.toLowerCase(),
      time: row.timeSort,
      city: row.cityKey,
      store: row.store.toLowerCase(),
      account: row.store.toLowerCase(),
      status: row.status,
      preferred: String(row.priority).padStart(2, "0"),
      raw: row.raw.toLowerCase(),
      "first-seen": row.firstSeen,
      firstSeen: row.firstSeen,
      "last-seen": row.lastSeen,
      lastSeen: row.lastSeen,
    };
    return values[key] ?? "";
  }

  function updateViewPresentation() {
    const spreadsheetsMode = state.view === "spreadsheets";
    if (dom.spreadsheets) dom.spreadsheets.hidden = !spreadsheetsMode;
    if (dom.gigsView) dom.gigsView.hidden = spreadsheetsMode;

    const copy = VIEW_COPY[state.view] || VIEW_COPY.dashboard;
    if (dom.viewTitle) dom.viewTitle.textContent = copy[0];
    if (dom.viewDescription) dom.viewDescription.textContent = copy[1];
  }

  function updateActiveNavigation() {
    dom.navLinks.forEach((link) => {
      const active = link.dataset.view === state.view;
      link.classList.toggle("is-active", active);
      active ? link.setAttribute("aria-current", "page") : link.removeAttribute("aria-current");
    });
    const selectedMonth = normalizeMonth(dom.month?.value || "");
    dom.monthLinks.forEach((link) => {
      const active = Boolean(selectedMonth) && normalizeMonth(link.dataset.month || "") === selectedMonth;
      link.classList.toggle("is-active", active);
      active ? link.setAttribute("aria-current", "true") : link.removeAttribute("aria-current");
    });
  }

  function updateQuickCities() {
    dom.cityQuick.forEach((button) => {
      const active = normalizeQuickCity(button.dataset.cityQuick || "all") === state.quickCity;
      button.classList.toggle("is-active", active);
      button.setAttribute("aria-pressed", String(active));
    });
  }

  function updateSortIndicators() {
    dom.sortButtons.forEach((button) => {
      const active = normalizeSortKey(button.dataset.sort) === state.sortKey;
      button.classList.toggle("is-active", active);
      button.dataset.direction = active ? state.sortDirection : "none";
      const header = button.closest("th");
      if (header) header.setAttribute("aria-sort", active ? `${state.sortDirection}ending` : "none");
    });
  }

  function hydrateSelects() {
    populateSelect(dom.month, unique(rowData.map((row) => row.month)).sort(), monthLabel);
    if (dom.date?.tagName === "SELECT") {
      populateSelect(dom.date, unique(rowData.map((row) => row.date)).sort(), dateLabel);
    }
    populateSelect(dom.city, unique(rowData.map((row) => row.city)).sort(compare), identity);
    populateSelect(dom.status, orderedStatuses(unique(rowData.map((row) => row.status)), identity));
    populateSelect(dom.account, unique(rowData.map((row) => row.store)).sort(compare), identity);
    populateSelect(
      dom.time,
      unique(rowData.map((row) => `${row.timeSort}|${row.time}`))
        .sort(compare)
        .map((item) => item.split("|")[1]),
      identity,
    );
  }

  function populateSelect(select, values, labeler) {
    if (!select || select.tagName !== "SELECT") return;
    const known = new Set([...select.options].map((option) => option.value));
    values.filter(Boolean).forEach((value) => {
      if (known.has(value)) return;
      const option = document.createElement("option");
      option.value = value;
      option.textContent = labeler(value);
      select.appendChild(option);
      known.add(value);
    });
  }

  function initializeControlsFromUrl(params) {
    state.view = normalizeView(params.get("view") || dom.body.dataset.initialView || state.view);
    state.quickCity = normalizeQuickCity(params.get("quick") || "all");
    state.sortKey = normalizeSortKey(params.get("sort") || "default");
    state.sortDirection = params.get("dir") === "desc" ? "desc" : "asc";
    state.page = positiveInteger(params.get("page"), 1);
    state.pageSize = positiveInteger(params.get("page_size"), positiveInteger(dom.pageSize?.value, 25));

    setControlValue(dom.search, params.get("q") || "");
    setControlValue(dom.month, params.get("month") || dom.body.dataset.initialMonth || "");
    setControlValue(dom.date, params.get("date") || "");
    setControlValue(dom.city, params.get("city") || "");
    setControlValue(dom.status, params.get("status") || "");
    setControlValue(dom.account, params.get("account") || params.get("store") || "");
    setControlValue(dom.time, params.get("time") || "");
    setControlValue(dom.preferred, params.get("preferred") || "");
    setControlValue(dom.pageSize, String(state.pageSize));
  }

  function writeUrl() {
    const params = new URLSearchParams();
    const filters = currentFilters();
    if (state.view !== "dashboard") params.set("view", state.view);
    if (filters.search) params.set("q", dom.search.value.trim());
    if (filters.month) params.set("month", filters.month);
    if (filters.date) params.set("date", filters.date);
    if (filters.city) params.set("city", dom.city.value);
    if (filters.status) params.set("status", dom.status.value);
    if (filters.account) params.set("account", dom.account.value);
    if (filters.time) params.set("time", dom.time.value);
    if (filters.preferred && filters.preferred !== "all") params.set("preferred", dom.preferred.value);
    if (state.quickCity !== "all") params.set("quick", state.quickCity);
    if (state.sortKey !== "default") params.set("sort", state.sortKey);
    if (state.sortDirection !== "asc") params.set("dir", state.sortDirection);
    if (state.page > 1) params.set("page", String(state.page));
    if (state.pageSize !== 25) params.set("page_size", String(state.pageSize));
    const search = params.toString();
    const nextUrl = `${window.location.pathname}${search ? `?${search}` : ""}${window.location.hash}`;
    window.history.replaceState(null, "", nextUrl);
  }

  function clearFilterControls() {
    [dom.search, dom.month, dom.date, dom.city, dom.status, dom.account, dom.time, dom.preferred]
      .forEach((control) => setControlValue(control, ""));
  }

  function setupRecruiterText() {
    if (!dom.generateRecruiterText || !dom.recruiterDrawer || !dom.recruiterMessage) return;

    rowData.forEach((row) => {
      if (!row.selectCheckbox || !row.gigId || row.status === "REMOVED") return;
      row.selectCheckbox.checked = false;

      // Selection is local UI state. Keep the checkbox click from reaching any
      // current or future row/table navigation handlers while preserving the
      // checkbox's native toggle (so do not preventDefault here).
      row.selectCheckbox.addEventListener("click", (event) => {
        event.stopPropagation();
      });
      row.selectCheckbox.addEventListener("change", (event) => {
        event.stopPropagation();
        if (row.selectCheckbox.checked) {
          recruiterState.selectedGigIds.add(row.gigId);
        } else {
          recruiterState.selectedGigIds.delete(row.gigId);
        }
        updateRecruiterSelectionPresentation();
      });
    });

    dom.generateRecruiterText.addEventListener("click", openRecruiterDrawer);
    dom.recruiterDrawerBackdrop?.addEventListener("click", closeRecruiterDrawer);
    dom.recruiterCloseControls.forEach((control) => control.addEventListener("click", closeRecruiterDrawer));
    dom.copyRecruiterText?.addEventListener("click", () => void copyRecruiterMessage());
    document.addEventListener("keydown", handleRecruiterDrawerKeydown);
    updateRecruiterSelectionPresentation();
  }

  function setupRecruiterTextUnavailable() {
    if (dom.generateRecruiterText) {
      dom.generateRecruiterText.disabled = true;
      dom.generateRecruiterText.setAttribute("aria-disabled", "true");
    }
    if (dom.selectedGigCount) dom.selectedGigCount.textContent = "0 gigs selected";
  }

  function selectedRecruiterRows() {
    return [...recruiterState.selectedGigIds]
      .map((gigId) => rowsByGigId.get(gigId))
      .filter((row) => row && row.status !== "REMOVED")
      .sort(compareRecruiterRows);
  }

  function compareRecruiterRows(a, b) {
    return (
      compare(validRecruiterDate(a.date) ? a.date : "9999-12-31", validRecruiterDate(b.date) ? b.date : "9999-12-31")
      || compare(validRecruiterTime(a.timeSort) ? a.timeSort : "99:99", validRecruiterTime(b.timeSort) ? b.timeSort : "99:99")
      || compareNumber(a.priority, b.priority)
      || compare(a.cityKey, b.cityKey)
      || compare(a.raw.toLowerCase(), b.raw.toLowerCase())
      || compare(a.gigId, b.gigId)
    );
  }

  function updateRecruiterSelectionPresentation() {
    for (const gigId of [...recruiterState.selectedGigIds]) {
      if (!rowsByGigId.has(gigId)) recruiterState.selectedGigIds.delete(gigId);
    }

    rowData.forEach((row) => {
      const selected = Boolean(row.gigId) && recruiterState.selectedGigIds.has(row.gigId);
      if (row.selectCheckbox) row.selectCheckbox.checked = selected;
      row.element.classList.toggle("is-selected", selected);
    });

    const count = recruiterState.selectedGigIds.size;
    const countText = recruiterGigCountLabel(count);
    if (dom.selectedGigCount) dom.selectedGigCount.textContent = countText;
    if (dom.generateRecruiterText) {
      dom.generateRecruiterText.disabled = count === 0;
      dom.generateRecruiterText.setAttribute("aria-disabled", String(count === 0));
    }
    if (dom.recruiterDrawerCount && recruiterState.isOpen) {
      dom.recruiterDrawerCount.textContent = countText;
    }
  }

  function recruiterGigCountLabel(count) {
    return `${number(count)} ${count === 1 ? "gig" : "gigs"} selected`;
  }

  function openRecruiterDrawer() {
    const selectedRows = selectedRecruiterRows();
    if (!selectedRows.length || !dom.recruiterDrawer || !dom.recruiterMessage) return;

    window.clearTimeout(recruiterState.closeTimer);
    window.clearTimeout(recruiterState.copyTimer);
    recruiterState.returnFocus = document.activeElement instanceof HTMLElement
      ? document.activeElement
      : dom.generateRecruiterText;
    recruiterState.isOpen = true;
    recruiterState.previousBodyOverflow = dom.body.style.overflow;
    dom.recruiterMessage.value = generateRecruiterMessage(selectedRows);
    if (dom.recruiterDrawerCount) {
      dom.recruiterDrawerCount.textContent = recruiterGigCountLabel(selectedRows.length);
    }
    resetRecruiterCopyFeedback();
    dom.recruiterDrawer.hidden = false;
    if (dom.recruiterDrawerBackdrop) dom.recruiterDrawerBackdrop.hidden = false;
    dom.body.classList.add("recruiter-drawer-open");
    dom.body.style.overflow = "hidden";
    dom.generateRecruiterText.setAttribute("aria-expanded", "true");

    window.requestAnimationFrame(() => {
      dom.recruiterDrawer.classList.add("is-open");
      dom.recruiterDrawerBackdrop?.classList.add("is-open");
      window.requestAnimationFrame(() => dom.recruiterMessage.focus());
    });
  }

  function closeRecruiterDrawer() {
    if (!recruiterState.isOpen || !dom.recruiterDrawer) return;
    recruiterState.isOpen = false;
    window.clearTimeout(recruiterState.copyTimer);
    dom.recruiterDrawer.classList.remove("is-open");
    dom.recruiterDrawerBackdrop?.classList.remove("is-open");
    dom.body.classList.remove("recruiter-drawer-open");
    dom.body.style.overflow = recruiterState.previousBodyOverflow;
    dom.generateRecruiterText?.setAttribute("aria-expanded", "false");
    resetRecruiterCopyFeedback();

    const returnFocus = recruiterState.returnFocus;
    recruiterState.closeTimer = window.setTimeout(() => {
      dom.recruiterDrawer.hidden = true;
      if (dom.recruiterDrawerBackdrop) dom.recruiterDrawerBackdrop.hidden = true;
      if (returnFocus?.isConnected) returnFocus.focus();
    }, prefersReducedMotion() ? 0 : 260);
  }

  function handleRecruiterDrawerKeydown(event) {
    if (!recruiterState.isOpen || !dom.recruiterDrawer) return;
    if (event.key === "Escape") {
      event.preventDefault();
      closeRecruiterDrawer();
      return;
    }
    if (event.key !== "Tab") return;

    const focusable = recruiterDrawerFocusableElements();
    if (!focusable.length) {
      event.preventDefault();
      dom.recruiterDrawer.focus();
      return;
    }
    const first = focusable[0];
    const last = focusable[focusable.length - 1];
    if (event.shiftKey && (document.activeElement === first || document.activeElement === dom.recruiterDrawer)) {
      event.preventDefault();
      last.focus();
    } else if (!event.shiftKey && document.activeElement === last) {
      event.preventDefault();
      first.focus();
    }
  }

  function recruiterDrawerFocusableElements() {
    if (!dom.recruiterDrawer) return [];
    return [...dom.recruiterDrawer.querySelectorAll(
      "button:not([disabled]), textarea:not([disabled]), input:not([disabled]), select:not([disabled]), a[href], [tabindex]:not([tabindex='-1'])",
    )].filter((element) => !element.hidden && element.getAttribute("aria-hidden") !== "true");
  }

  function clearRecruiterState() {
    recruiterState.selectedGigIds.clear();
    if (recruiterState.isOpen) closeRecruiterDrawer();
    if (dom.recruiterMessage) dom.recruiterMessage.value = "";
    if (!dom.tbody) {
      setupRecruiterTextUnavailable();
      return;
    }
    updateRecruiterSelectionPresentation();
  }

  function generateRecruiterMessage(rows) {
    const grouped = new Map();
    rows.forEach((row) => {
      if (!grouped.has(row.date)) grouped.set(row.date, []);
      grouped.get(row.date).push(row);
    });

    const count = rows.length;
    if (grouped.size === 1) {
      const [date, dateRows] = grouped.entries().next().value;
      const opening = count === 1
        ? `Hey Adam, I’d like to pick up this gig for ${naturalRecruiterDate(date, dateRows[0]?.day)} if it’s still available:`
        : count === 2
          ? `Hey Adam, I’d like to pick up these two gigs for ${naturalRecruiterDate(date, dateRows[0]?.day)} if they’re still available:`
          : `Hey Adam, I’d like to pick up these gigs for ${naturalRecruiterDate(date, dateRows[0]?.day)} if they’re still available:`;
      return `${opening}\n\n${recruiterGigBullets(dateRows)}\n\n${recruiterClosing(count)}`;
    }

    const dateSections = [...grouped.entries()].map(([date, dateRows]) => (
      `${naturalRecruiterDate(date, dateRows[0]?.day)}\n${recruiterGigBullets(dateRows)}`
    ));
    return `Hey Adam, I’d like to pick up these gigs if they’re still available:\n\n${dateSections.join("\n\n")}\n\n${recruiterClosing(count)}`;
  }

  function recruiterGigBullets(rows) {
    return rows.map((row) => `• ${cleanRecruiterGigName(row)}, ${displayRecruiterTime(row)}`).join("\n");
  }

  function cleanRecruiterGigName(row) {
    const clean = (value) => String(value || "").replace(/\s+/g, " ").trim();
    const city = clean(row.city);
    const storeNumber = clean(row.storeNumber);
    let store = clean(row.store).replace(/^WA\|\|\s*/i, "");

    if (!store) {
      store = clean(row.raw)
        .replace(/^\s*\d{1,2}(?::\d{2})?\s*[ap](?:\.?m\.?)?\s*/i, "")
        .replace(/^WA\|\|\s*/i, "");
    }
    if (storeNumber && !containsStandaloneText(store, storeNumber)) store = `${store} ${storeNumber}`.trim();
    if (city && !containsCitySuffix(store, city) && !containsStandaloneText(store, city)) {
      store = `${store} (${city})`.trim();
    }
    return store || city || "Gig";
  }

  function containsStandaloneText(value, search) {
    const escaped = escapeRegExp(search);
    return new RegExp(`(^|[^A-Za-z0-9])${escaped}([^A-Za-z0-9]|$)`, "i").test(value);
  }

  function containsCitySuffix(value, city) {
    return new RegExp(`\\(\\s*${escapeRegExp(city)}\\s*\\)`, "i").test(value);
  }

  function displayRecruiterTime(row) {
    const time = String(row.time || "").trim();
    if (time && time !== "—") return time;
    const normalized = String(row.timeSort || "").match(/^(\d{1,2}):(\d{2})$/);
    if (!normalized) return "Time not specified";
    const hours = Number(normalized[1]);
    const minutes = normalized[2];
    const period = hours >= 12 ? "PM" : "AM";
    return `${hours % 12 || 12}:${minutes} ${period}`;
  }

  function naturalRecruiterDate(value, dayFallback = "") {
    const match = String(value || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) return [dayFallback, value].filter(Boolean).join(", ") || "the selected date";
    const date = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
    return new Intl.DateTimeFormat("en-US", {
      weekday: "long",
      month: "long",
      day: "numeric",
      timeZone: "UTC",
    }).format(date);
  }

  function validRecruiterDate(value) {
    const match = String(value || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) return false;
    const year = Number(match[1]);
    const month = Number(match[2]);
    const day = Number(match[3]);
    const date = new Date(Date.UTC(year, month - 1, day));
    return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
  }

  function validRecruiterTime(value) {
    const match = String(value || "").match(/^(\d{2}):(\d{2})$/);
    return Boolean(match) && Number(match[1]) < 24 && Number(match[2]) < 60;
  }

  function recruiterClosing(count) {
    if (count === 1) return "Please let me know if I can be added. Thanks!";
    if (count === 2) return "Please let me know if I can be added to both. Thanks!";
    return "Please let me know if I can be added to these. Thanks!";
  }

  async function copyRecruiterMessage() {
    if (!dom.recruiterMessage || !dom.copyRecruiterText) return;
    const text = dom.recruiterMessage.value;
    if (!text) {
      showRecruiterCopyFailure("There isn’t any text to copy yet.");
      return;
    }

    try {
      if (!navigator.clipboard?.writeText) throw new Error("Clipboard API unavailable");
      await navigator.clipboard.writeText(text);
      showRecruiterCopySuccess();
      return;
    } catch (_clipboardError) {
      try {
        if (!fallbackCopyRecruiterMessage()) throw new Error("Fallback copy failed");
        showRecruiterCopySuccess();
      } catch (_fallbackError) {
        showRecruiterCopyFailure("Couldn’t copy automatically. Select the text and copy it manually.");
      }
    }
  }

  function fallbackCopyRecruiterMessage() {
    if (!dom.recruiterMessage || typeof document.execCommand !== "function") return false;
    const selectionStart = dom.recruiterMessage.selectionStart;
    const selectionEnd = dom.recruiterMessage.selectionEnd;
    dom.recruiterMessage.focus();
    dom.recruiterMessage.select();
    let copied = false;
    try {
      copied = document.execCommand("copy") === true;
    } finally {
      dom.recruiterMessage.setSelectionRange(selectionStart, selectionEnd);
    }
    return copied;
  }

  function showRecruiterCopySuccess() {
    window.clearTimeout(recruiterState.copyTimer);
    dom.copyRecruiterText?.classList.add("is-copied");
    if (dom.copyRecruiterTextLabel) dom.copyRecruiterTextLabel.textContent = "Copied!";
    if (dom.recruiterCopyFeedback) {
      dom.recruiterCopyFeedback.textContent = "Copied to clipboard.";
      dom.recruiterCopyFeedback.classList.remove("is-error");
    }
    recruiterState.copyTimer = window.setTimeout(resetRecruiterCopyFeedback, 1500);
  }

  function showRecruiterCopyFailure(message) {
    window.clearTimeout(recruiterState.copyTimer);
    dom.copyRecruiterText?.classList.remove("is-copied");
    if (dom.copyRecruiterTextLabel) dom.copyRecruiterTextLabel.textContent = "Copy Text";
    if (dom.recruiterCopyFeedback) {
      dom.recruiterCopyFeedback.textContent = message;
      dom.recruiterCopyFeedback.classList.add("is-error");
    }
    dom.recruiterMessage?.focus();
  }

  function resetRecruiterCopyFeedback() {
    window.clearTimeout(recruiterState.copyTimer);
    dom.copyRecruiterText?.classList.remove("is-copied");
    if (dom.copyRecruiterTextLabel) dom.copyRecruiterTextLabel.textContent = "Copy Text";
    if (dom.recruiterCopyFeedback) {
      dom.recruiterCopyFeedback.textContent = "";
      dom.recruiterCopyFeedback.classList.remove("is-error");
    }
  }

  function escapeRegExp(value) {
    return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }

  function prefersReducedMotion() {
    return window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  }

  function setupRefreshControl() {
    if (!dom.refresh) return;

    dom.refresh.addEventListener("click", startGigRefresh);
    dom.notificationDismiss?.addEventListener("click", () => {
      refreshState.dismissedFor = dom.notification?.dataset.jobKey || refreshState.activeStatus;
      hideRefreshNotification();
    });
    dom.detailsToggle?.addEventListener("click", () => {
      const expanded = dom.detailsToggle.getAttribute("aria-expanded") === "true";
      dom.detailsToggle.setAttribute("aria-expanded", String(!expanded));
      if (dom.errorDetails) dom.errorDetails.hidden = expanded;
      if (!expanded) dom.errorOutput?.focus?.();
    });

    refreshState.watching = restoreRefreshWatch();
    restoreRefreshNotice();
    void checkRefreshStatus({ initial: true });
  }

  async function startGigRefresh() {
    if (refreshState.activeStatus === "running" || refreshState.startPending) return;

    // Any status request issued before this click is stale and must not be
    // allowed to re-enable the button after the POST has started a job.
    refreshState.generation += 1;
    refreshState.startPending = true;
    refreshState.watching = true;
    markRefreshWatching(true);
    refreshState.reloadScheduled = false;
    refreshState.dismissedFor = "";
    renderRefreshRunning({
      status: "running",
      message: "Starting the calendar refresh...",
      jobKey: "starting",
    });

    try {
      const response = await fetch("/api/refresh", {
        method: "POST",
        headers: {
          Accept: "application/json",
          "X-US-Event-Refresh": "1",
        },
        credentials: "same-origin",
        cache: "no-store",
      });
      const payload = await readJsonResponse(response);
      refreshState.startPending = false;

      if (!response.ok && response.status !== 409) {
        throw refreshRequestError(response, payload);
      }

      const parsed = parseRefreshPayload(payload);
      if (response.status === 409 && parsed.status !== "running") {
        parsed.status = "running";
        parsed.message = parsed.message || "A gig data refresh is already running.";
      }
      handleRefreshPayload(parsed, { source: "start" });
    } catch (error) {
      refreshState.startPending = false;
      refreshState.watching = false;
      renderRefreshError({
        status: "error",
        message: "Your previous gig data is still available.",
        details: sanitizeDetails(error?.message || "The refresh request could not be started."),
        jobKey: "start-error",
      });
    }
  }

  async function checkRefreshStatus({ initial = false } = {}) {
    const requestGeneration = refreshState.generation;
    try {
      const response = await fetch("/api/refresh/status", {
        method: "GET",
        headers: { Accept: "application/json" },
        credentials: "same-origin",
        cache: "no-store",
      });
      const payload = await readJsonResponse(response);
      if (requestGeneration !== refreshState.generation) return;
      if (!response.ok) throw refreshRequestError(response, payload);
      refreshState.failureCount = 0;
      handleRefreshPayload(parseRefreshPayload(payload), { initial, source: "poll" });
    } catch (error) {
      if (requestGeneration !== refreshState.generation) return;
      refreshState.failureCount += 1;
      if (refreshState.watching || refreshState.activeStatus === "running") {
        setRefreshButton("running");
        announceRefreshStatus("Refreshing gig data. Waiting to reconnect to the dashboard status service.");
        showRefreshNotification({
          kind: "running",
          title: "Refreshing gig data",
          message: "The refresh is still being monitored. Waiting to reconnect...",
          jobKey: "status-retry",
        });
        scheduleRefreshPoll(Math.min(5000, REFRESH_POLL_INTERVAL * (refreshState.failureCount + 1)));
      } else if (!initial) {
        renderRefreshError({
          status: "error",
          message: "The dashboard could not check refresh status. Your existing gig data is still available.",
          details: sanitizeDetails(error?.message || "Status request failed."),
          jobKey: "status-error",
        });
      }
    }
  }

  function handleRefreshPayload(payload, { initial = false } = {}) {
    if (initial && refreshState.startPending && payload.status !== "running") return;

    if (payload.lastSuccessDisplay && dom.lastRefreshed) {
      dom.lastRefreshed.textContent = payload.lastSuccessDisplay;
    }

    if (payload.status === "running") {
      refreshState.watching = true;
      markRefreshWatching(true);
      renderRefreshRunning(payload);
      scheduleRefreshPoll();
      return;
    }

    clearRefreshPoll();

    if (payload.status === "success") {
      if (refreshState.watching) {
        completeGigRefresh(payload);
      } else {
        refreshState.activeStatus = "success";
        if (!dom.notification || dom.notification.hidden) setRefreshButton("idle");
      }
      return;
    }

    if (payload.status === "error") {
      refreshState.watching = false;
      markRefreshWatching(false);
      renderRefreshError(payload);
      return;
    }

    if (refreshState.startPending) return;
    refreshState.watching = false;
    markRefreshWatching(false);
    refreshState.activeStatus = "idle";
    setRefreshButton("idle");
    if (!initial && dom.notification?.dataset.kind === "running") hideRefreshNotification();
  }

  function renderRefreshRunning(payload) {
    refreshState.activeStatus = "running";
    setRefreshButton("running");
    const message = runningMessage(payload);
    announceRefreshStatus(`Refreshing gig data. ${message}`);
    showRefreshNotification({
      kind: "running",
      title: "Refreshing gig data",
      message,
      jobKey: payload.jobKey,
    });
  }

  function completeGigRefresh(payload) {
    if (refreshState.reloadScheduled) return;
    // The refreshed spreadsheets replace the source dataset. Clear all UI-only
    // recruiter state before the existing reload so no stale gig IDs survive.
    clearRecruiterState();
    refreshState.watching = false;
    markRefreshWatching(false);
    refreshState.activeStatus = "success";
    refreshState.reloadScheduled = true;
    clearRefreshPoll();
    setRefreshButton("success");

    const countMessage = payload.hasNewCount
      ? `${formatCount(payload.newCount)} new ${payload.newCount === 1 ? "gig" : "gigs"} found.`
      : "";
    const message = ["Gig data refreshed successfully.", countMessage].filter(Boolean).join(" ");
    announceRefreshStatus(message);
    showRefreshNotification({
      kind: "success",
      title: "Refresh complete",
      message,
      jobKey: payload.jobKey,
    });
    saveRefreshNotice({ message, newCount: payload.newCount, hasNewCount: payload.hasNewCount });

    window.setTimeout(() => {
      const target = new URL(window.location.href);
      if (payload.hasNewCount && payload.newCount > 0) {
        target.search = "";
        target.searchParams.set("view", "new");
      }
      window.location.assign(target.toString());
    }, 900);
  }

  function renderRefreshError(payload) {
    refreshState.activeStatus = "error";
    refreshState.watching = false;
    markRefreshWatching(false);
    clearRefreshPoll();
    setRefreshButton("idle");
    const details = sanitizeDetails(payload.details || payload.message || "No additional diagnostic details were returned.");
    announceRefreshStatus("Refresh failed. Your previous gig data is still available.");
    showRefreshNotification({
      kind: "error",
      title: "Refresh failed",
      message: "Your previous gig data is still available.",
      details,
      jobKey: payload.jobKey || "refresh-error",
    });
  }

  function setRefreshButton(mode) {
    if (!dom.refresh) return;
    const running = mode === "running";
    const success = mode === "success";
    dom.refresh.disabled = running || success;
    dom.refresh.classList.toggle("is-refreshing", running);
    dom.refresh.classList.toggle("is-updated", success);
    dom.refresh.setAttribute("aria-busy", String(running));
    if (dom.refreshButtonLabel) {
      dom.refreshButtonLabel.textContent = running ? "Refreshing..." : success ? "Updated" : "Refresh Gig Data";
    }
    dom.refreshIndicator?.classList.toggle("is-refreshing", running);
    dom.refreshIndicator?.classList.toggle("is-updated", success);
  }

  function announceRefreshStatus(message) {
    if (dom.refreshJobStatus) dom.refreshJobStatus.textContent = safeUiText(message, 400);
  }

  function showRefreshNotification({ kind, title, message, details = "", jobKey = "" }) {
    if (!dom.notification) return;
    const normalizedKind = ["running", "success", "error"].includes(kind) ? kind : "running";
    if (normalizedKind === "running" && refreshState.dismissedFor === jobKey) return;

    window.clearTimeout(refreshState.autoDismissTimer);
    dom.notification.dataset.kind = normalizedKind;
    dom.notification.dataset.jobKey = jobKey;
    dom.notification.className = `refresh-notification is-${normalizedKind}`;
    dom.notification.setAttribute("role", normalizedKind === "error" ? "alert" : "status");
    dom.notificationTitle.textContent = safeUiText(title, 120);
    dom.notificationMessage.textContent = safeUiText(message, 500);

    const safeDetails = sanitizeDetails(details);
    if (dom.detailsToggle) {
      dom.detailsToggle.hidden = !safeDetails;
      dom.detailsToggle.setAttribute("aria-expanded", "false");
    }
    if (dom.errorDetails) dom.errorDetails.hidden = true;
    if (dom.errorOutput) dom.errorOutput.textContent = safeDetails;
    dom.notification.hidden = false;

    if (normalizedKind === "success") {
      refreshState.autoDismissTimer = window.setTimeout(hideRefreshNotification, 7000);
    }
  }

  function hideRefreshNotification() {
    window.clearTimeout(refreshState.autoDismissTimer);
    if (dom.notification) dom.notification.hidden = true;
    if (dom.detailsToggle) dom.detailsToggle.setAttribute("aria-expanded", "false");
    if (dom.errorDetails) dom.errorDetails.hidden = true;
  }

  function scheduleRefreshPoll(delay = REFRESH_POLL_INTERVAL) {
    clearRefreshPoll();
    refreshState.pollTimer = window.setTimeout(() => void checkRefreshStatus(), delay);
  }

  function clearRefreshPoll() {
    window.clearTimeout(refreshState.pollTimer);
    refreshState.pollTimer = null;
  }

  function parseRefreshPayload(payload) {
    const source = payload && typeof payload === "object" ? payload : {};
    const nested = source.job && typeof source.job === "object" ? source.job : {};
    const statusCandidate = normalizeText(source.status || source.state || nested.status || nested.state || "idle");
    const status = REFRESH_STATES.has(statusCandidate) ? statusCandidate : "idle";
    const newCountCandidates = [
      source.new_count,
      source.new_gigs,
      source.new_gigs_found,
      nested.new_count,
      source.result?.new_count,
    ];
    const newCountValue = newCountCandidates.find((value) => value !== null && value !== undefined && value !== "" && Number.isFinite(Number(value)));
    const newCount = newCountValue === undefined ? 0 : Math.max(0, Number.parseInt(newCountValue, 10));
    const jobId = safeUiText(source.job_id || nested.job_id || nested.id || "", 120);
    const startedAt = safeUiText(source.started_at || nested.started_at || "", 120);
    const finishedAt = safeUiText(source.finished_at || nested.finished_at || "", 120);
    const message = safeUiText(source.message || nested.message || "", 500);
    const rawDetails = source.details ?? source.error_details ?? source.error ?? nested.details ?? nested.error ?? "";

    return {
      status,
      message,
      details: sanitizeDetails(Array.isArray(rawDetails) ? rawDetails.join("\n") : rawDetails),
      jobId,
      startedAt,
      finishedAt,
      jobKey: jobId || startedAt || finishedAt || status,
      newCount,
      hasNewCount: newCountValue !== undefined,
      lastSuccessDisplay: safeUiText(source.last_success_display || nested.last_success_display || "", 160),
      stage: normalizeText(source.stage || nested.stage || ""),
    };
  }

  function runningMessage(payload) {
    if (payload.message) return payload.message;
    const stageMessages = {
      connecting: "Connecting to US Event Management...",
      calendar: "Reading the calendar...",
      scraping: "Reading the calendar...",
      comparing: "Updating gig data...",
      exporting: "Updating spreadsheets...",
      spreadsheets: "Updating spreadsheets...",
      starting: "Starting the calendar refresh...",
    };
    return stageMessages[payload.stage] || "The calendar scraper is running. This may take a few minutes.";
  }

  async function readJsonResponse(response) {
    const text = await response.text();
    if (!text) return {};
    try {
      return JSON.parse(text);
    } catch (_error) {
      return { details: safeUiText(text, 2000) };
    }
  }

  function refreshRequestError(response, payload) {
    const parsed = parseRefreshPayload(payload);
    const message = parsed.details || parsed.message || `Refresh request failed with HTTP ${response.status}.`;
    return new Error(sanitizeDetails(message));
  }

  function sanitizeDetails(value) {
    let text = String(value ?? "")
      .replace(/\r\n?/g, "\n")
      .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/g, "");
    text = text
      .replace(/\b(Bearer)\s+[A-Za-z0-9._~+\/-]+=*/gi, "$1 [REDACTED]")
      .replace(/((?:authorization|password|passwd|secret|token|cookie|session(?:_?id)?|auth(?:_?state)?)[^:=\n]{0,24}[:=]\s*)(?:\"[^\"]*\"|'[^']*'|[^\s,;]+)/gi, "$1[REDACTED]")
      .replace(/([?&](?:password|token|secret|cookie|session|auth)[^=&#\s]*=)[^&#\s]+/gi, "$1[REDACTED]")
      .replace(/^([A-Z][A-Z0-9_]*(?:EMAIL|USERNAME|PASSWORD|PASS|TOKEN|SECRET|COOKIE|SESSION|AUTH)[A-Z0-9_]*\s*=).*$/gim, "$1[REDACTED]");
    const trimmed = text.trim();
    return trimmed.length > 8000 ? `...${trimmed.slice(-8000)}` : trimmed;
  }

  function safeUiText(value, maxLength = 500) {
    const clean = String(value ?? "")
      .replace(/[\u0000-\u001f\u007f]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    return clean.slice(0, maxLength);
  }

  function saveRefreshNotice(notice) {
    try {
      window.sessionStorage.setItem(REFRESH_NOTICE_KEY, JSON.stringify(notice));
    } catch (_error) {
      // The refresh still completes when private browsing blocks session storage.
    }
  }

  function markRefreshWatching(watching) {
    try {
      if (watching) {
        window.sessionStorage.setItem(REFRESH_WATCH_KEY, "1");
      } else {
        window.sessionStorage.removeItem(REFRESH_WATCH_KEY);
      }
    } catch (_error) {
      // In-memory polling still works when session storage is unavailable.
    }
  }

  function restoreRefreshWatch() {
    try {
      return window.sessionStorage.getItem(REFRESH_WATCH_KEY) === "1";
    } catch (_error) {
      return false;
    }
  }

  function restoreRefreshNotice() {
    let stored = "";
    try {
      stored = window.sessionStorage.getItem(REFRESH_NOTICE_KEY) || "";
      window.sessionStorage.removeItem(REFRESH_NOTICE_KEY);
    } catch (_error) {
      return;
    }
    if (!stored) return;
    try {
      const notice = JSON.parse(stored);
      const count = Number.isFinite(Number(notice.newCount)) ? Math.max(0, Number.parseInt(notice.newCount, 10)) : 0;
      const message = safeUiText(notice.message || "Gig data refreshed successfully.", 500);
      setRefreshButton("success");
      announceRefreshStatus(message);
      showRefreshNotification({
        kind: "success",
        title: "Refresh complete",
        message,
        jobKey: "completed",
      });
      window.setTimeout(() => setRefreshButton("idle"), 1600);
      if (notice.hasNewCount && count > 0) refreshState.activeStatus = "success";
    } catch (_error) {
      // Ignore malformed local state; the live endpoint remains authoritative.
    }
  }

  function formatCount(value) {
    return new Intl.NumberFormat().format(value);
  }

  function setControlValue(control, value) {
    if (!control) return;
    if (control.tagName === "SELECT" && value && ![...control.options].some((option) => option.value === value)) {
      return;
    }
    control.value = value;
  }

  function setupSidebarOnly() {
    dom.sidebarToggle?.addEventListener("click", () => {
      const opening = !dom.body.classList.contains("sidebar-open");
      dom.body.classList.toggle("sidebar-open", opening);
      dom.sidebarToggle.setAttribute("aria-expanded", String(opening));
      dom.sidebarOverlay?.setAttribute("aria-hidden", String(!opening));
    });
    dom.sidebarOverlay?.addEventListener("click", closeSidebar);
    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape") closeSidebar();
    });
    window.matchMedia("(min-width: 1025px)").addEventListener?.("change", (event) => {
      if (event.matches) closeSidebar();
    });
  }

  function closeSidebar() {
    dom.body.classList.remove("sidebar-open");
    dom.sidebarToggle?.setAttribute("aria-expanded", "false");
    dom.sidebarOverlay?.setAttribute("aria-hidden", "true");
  }

  function scrollTableIntoView() {
    const target = dom.filters || dom.table;
    if (!target) return;
    const top = target.getBoundingClientRect().top + window.scrollY - 92;
    window.scrollTo({ top, behavior: window.matchMedia("(prefers-reduced-motion: reduce)").matches ? "auto" : "smooth" });
  }

  function normalizeView(value) {
    const normalized = normalizeText(value);
    return VALID_VIEWS.has(normalized) ? normalized : "dashboard";
  }

  function normalizeQuickCity(value) {
    const normalized = normalizeText(value);
    return ["all", "other", ...PREFERRED_CITIES].includes(normalized) ? normalized : "all";
  }

  function normalizeSortKey(value) {
    const normalized = String(value || "default").trim();
    const allowed = new Set([
      "default", "date", "day", "time", "city", "store", "account", "status", "preferred",
      "raw", "first-seen", "firstSeen", "last-seen", "lastSeen",
    ]);
    return allowed.has(normalized) ? normalized : "default";
  }

  function normalizeMonth(value) {
    const match = String(value || "").match(/^(\d{4})-(\d{2})/);
    return match ? `${match[1]}-${match[2]}` : "";
  }

  function normalizeDate(value) {
    const text = String(value || "").trim();
    const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
    const parsed = new Date(text);
    if (Number.isNaN(parsed.getTime())) return text;
    return [parsed.getFullYear(), String(parsed.getMonth() + 1).padStart(2, "0"), String(parsed.getDate()).padStart(2, "0")].join("-");
  }

  function normalizeTime(value) {
    const text = String(value || "").trim().toLowerCase();
    const twentyFour = text.match(/^(\d{1,2}):(\d{2})$/);
    if (twentyFour) return `${String(Number(twentyFour[1])).padStart(2, "0")}:${twentyFour[2]}`;
    const twelveHour = text.match(/^(\d{1,2})(?::(\d{2}))?\s*([ap])(?:\.?m\.?)?$/);
    if (!twelveHour) return text;
    let hours = Number(twelveHour[1]) % 12;
    if (twelveHour[3] === "p") hours += 12;
    return `${String(hours).padStart(2, "0")}:${twelveHour[2] || "00"}`;
  }

  function normalizeText(value) {
    return String(value || "").trim().toLowerCase();
  }

  function isTruthy(value) {
    return ["yes", "y", "true", "1", "preferred"].includes(normalizeText(value));
  }

  function unique(values) {
    return [...new Set(values.filter(Boolean))];
  }

  function orderedStatuses(values) {
    const order = { NEW: 1, EXISTING: 2, REMOVED: 3 };
    return [...values].sort((a, b) => (order[a] || 99) - (order[b] || 99) || compare(a, b));
  }

  function compare(a, b) {
    return String(a || "").localeCompare(String(b || ""), undefined, { numeric: true, sensitivity: "base" });
  }

  function compareNumber(a, b) {
    return Number(a) - Number(b);
  }

  function positiveInteger(value, fallback) {
    const numberValue = Number.parseInt(value, 10);
    return Number.isFinite(numberValue) && numberValue > 0 ? numberValue : fallback;
  }

  function monthLabel(value) {
    const [year, month] = value.split("-").map(Number);
    return new Intl.DateTimeFormat(undefined, { month: "long", year: "numeric", timeZone: "UTC" })
      .format(new Date(Date.UTC(year, month - 1, 1)));
  }

  function dateLabel(value) {
    const [year, month, day] = value.split("-").map(Number);
    return new Intl.DateTimeFormat(undefined, { month: "short", day: "numeric", year: "numeric", timeZone: "UTC" })
      .format(new Date(Date.UTC(year, month - 1, day)));
  }

  function number(value) {
    return new Intl.NumberFormat().format(value);
  }

  function identity(value) {
    return value;
  }
})();
