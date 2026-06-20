(function (window) {
  const STORAGE_KEYS = {
    users: "users",
    serviceApplications: "serviceApplications",
    eventApplications: "eventApplications",
    projectActivity: "projectActivity",
    reviewChecklistState: "reviewChecklistState",
    reviewNotes: "reviewNotes",
    projectProfile: "projectProfile",
    plannerTasks: "plannerTasks",
    plannerBudget: "plannerBudget",
    plannerMilestones: "plannerMilestones",
    plannerRisks: "plannerRisks"
  };

  function readJson(key, fallback) {
    try {
      const raw = window.localStorage.getItem(key);
      if (raw === null) return fallback;
      const parsed = JSON.parse(raw);
      return parsed ?? fallback;
    } catch {
      return fallback;
    }
  }

  function writeJson(key, value) {
    window.localStorage.setItem(key, JSON.stringify(value));
    return value;
  }

  function readArray(key) {
    const value = readJson(key, []);
    return Array.isArray(value) ? value : [];
  }

  function readObject(key) {
    const value = readJson(key, {});
    return value && typeof value === "object" && !Array.isArray(value) ? value : {};
  }

  function currentDateValue(offsetDays) {
    const date = new Date();
    if (Number.isFinite(offsetDays)) {
      date.setDate(date.getDate() + offsetDays);
    }
    return date.toISOString().slice(0, 10);
  }

  function getCurrentUser() {
    return window.sessionStorage.getItem("currentUser") || "Guest";
  }

  function getProjectProfile() {
    const defaults = {
      projectName: "Campus Event Excellence",
      clientName: "College Review Board",
      venue: "Main Auditorium",
      eventDate: currentDateValue(7),
      owner: getCurrentUser(),
      budget: "250000",
      priority: "high",
      notes: "Focus the review on end-to-end workflow, demo readiness, and reporting proof."
    };

    return Object.assign(defaults, readObject(STORAGE_KEYS.projectProfile));
  }

  function saveProjectProfile(profile) {
    return writeJson(STORAGE_KEYS.projectProfile, Object.assign(getProjectProfile(), profile || {}));
  }

  function addActivity(message, type) {
    const logs = readArray(STORAGE_KEYS.projectActivity);
    logs.unshift({
      time: new Date().toLocaleString(),
      msg: message,
      type: type || "info"
    });
    writeJson(STORAGE_KEYS.projectActivity, logs.slice(0, 40));
  }

  function getChecklistPercent(totalItems) {
    const state = readObject(STORAGE_KEYS.reviewChecklistState);
    const completed = Object.values(state).filter(Boolean).length;
    const divisor = totalItems || Math.max(Object.keys(state).length, 0);
    if (!divisor) return 0;
    return Math.min(100, Math.round((completed / divisor) * 100));
  }

  function formatCurrency(value) {
    return "INR " + Number(value || 0).toLocaleString(undefined, { maximumFractionDigits: 2 });
  }

  function getOverview(totalChecklistItems) {
    const users = readArray(STORAGE_KEYS.users);
    const applications = readArray(STORAGE_KEYS.serviceApplications);
    const eventApplications = readArray(STORAGE_KEYS.eventApplications);
    const activity = readArray(STORAGE_KEYS.projectActivity);
    const tasks = readArray(STORAGE_KEYS.plannerTasks);
    const budget = readArray(STORAGE_KEYS.plannerBudget);
    const milestones = readArray(STORAGE_KEYS.plannerMilestones);
    const risks = readArray(STORAGE_KEYS.plannerRisks);

    const completedTasks = tasks.filter((item) => Boolean(item.done)).length;
    const openTasks = Math.max(tasks.length - completedTasks, 0);
    const completedMilestones = milestones.filter((item) => item.status === "done").length;
    const checklistPercent = getChecklistPercent(totalChecklistItems);
    const taskHealth = tasks.length ? Math.round((completedTasks / tasks.length) * 100) : 0;
    const milestoneHealth = milestones.length ? Math.round((completedMilestones / milestones.length) * 100) : 0;
    const plannedBudget = budget.reduce((sum, item) => sum + Number(item.planned || 0), 0);
    const actualBudget = budget.reduce((sum, item) => sum + Number(item.actual || 0), 0);
    const budgetVariance = actualBudget - plannedBudget;
    const openRisks = risks.filter((item) => item.status !== "closed");
    const highRisks = openRisks.filter((item) => item.level === "high");
    const coverageSignals = [
      users.length > 0 ? 100 : 0,
      applications.length > 0 ? 100 : 0,
      eventApplications.length > 0 ? 100 : 0,
      activity.length > 0 ? 100 : 0
    ];
    const coverage = Math.round(coverageSignals.reduce((sum, item) => sum + item, 0) / coverageSignals.length);
    const readinessScore = Math.round(
      checklistPercent * 0.35 +
      taskHealth * 0.25 +
      milestoneHealth * 0.2 +
      coverage * 0.2
    );

    return {
      usersCount: users.length,
      applicationsCount: applications.length,
      eventCount: eventApplications.length,
      activityCount: activity.length,
      tasksCount: tasks.length,
      completedTasks,
      openTasks,
      milestonesCount: milestones.length,
      completedMilestones,
      checklistPercent,
      taskHealth,
      milestoneHealth,
      plannedBudget,
      actualBudget,
      budgetVariance,
      riskCount: openRisks.length,
      highRiskCount: highRisks.length,
      readinessScore
    };
  }

  function getReadinessActions(totalChecklistItems) {
    const overview = getOverview(totalChecklistItems);
    const profile = getProjectProfile();
    const actions = [];

    if (overview.checklistPercent < 100) {
      actions.push("Finish the final review checklist to strengthen demo readiness.");
    }
    if (!profile.projectName || !profile.clientName || !profile.venue) {
      actions.push("Complete the project brief with project name, client, and venue details.");
    }
    if (overview.usersCount === 0) {
      actions.push("Register at least one demo user so authentication appears production-ready.");
    }
    if (overview.applicationsCount === 0) {
      actions.push("Submit a sample service application to prove the service workflow.");
    }
    if (overview.eventCount === 0) {
      actions.push("Save a stage design from the AI Event Designer to populate event records.");
    }
    if (overview.openTasks > 0) {
      actions.push("Close or re-prioritize the " + overview.openTasks + " open planner tasks.");
    }
    if (overview.highRiskCount > 0) {
      actions.push("Mitigate the " + overview.highRiskCount + " high-severity planning risks.");
    }
    if (overview.budgetVariance > 0) {
      actions.push("Budget is over target by " + formatCurrency(overview.budgetVariance) + ". Adjust allocations before review.");
    }
    if (!actions.length) {
      actions.push("Platform is review-ready. Focus on storytelling, screenshots, and a clean demo sequence.");
    }

    return actions.slice(0, 5);
  }

  function exportCsv(fileName, rows) {
    if (!Array.isArray(rows) || !rows.length) return;

    const headers = Object.keys(rows[0]);
    const lines = [headers.join(",")].concat(
      rows.map((row) => headers.map((header) => {
        const value = row[header] == null ? "" : String(row[header]);
        return '"' + value.replace(/"/g, '""') + '"';
      }).join(","))
    );

    const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.click();
    window.URL.revokeObjectURL(url);
  }

  const translationTextStore = new WeakMap();
  const translationCache = new Map();

  function hasEnglishLetters(text) {
    return /[A-Za-z]/.test(String(text || "").trim());
  }

  async function translateText(text, language) {
    if (!text || !String(text).trim() || language === "en") return text;
    if (!hasEnglishLetters(text)) return text;

    const key = language + "::" + text;
    if (translationCache.has(key)) return translationCache.get(key);

    try {
      const response = await window.fetch(
        "https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=" +
          encodeURIComponent(language) +
          "&dt=t&q=" +
          encodeURIComponent(text)
      );
      const data = await response.json();
      const translated = Array.isArray(data && data[0])
        ? data[0].map(function (chunk) { return chunk[0]; }).join("")
        : text;
      const finalText = translated || text;
      translationCache.set(key, finalText);
      return finalText;
    } catch {
      return text;
    }
  }

  async function translateUniqueTexts(texts, language, parallelLimit) {
    const unique = Array.from(new Set(texts.filter(function (item) {
      return item && String(item).trim() && hasEnglishLetters(item);
    })));

    const result = {};
    if (!unique.length || language === "en") return result;

    const limit = Math.max(1, Number(parallelLimit) || 10);
    let index = 0;

    async function worker() {
      while (index < unique.length) {
        const current = unique[index++];
        result[current] = await translateText(current, language);
      }
    }

    const workers = [];
    for (let i = 0; i < Math.min(limit, unique.length); i += 1) {
      workers.push(worker());
    }
    await Promise.all(workers);
    return result;
  }

  function collectTextNodes(root) {
    const textNodes = [];
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, {
      acceptNode: function (node) {
        if (!node || !node.nodeValue) return NodeFilter.FILTER_REJECT;
        const value = node.nodeValue.trim();
        if (!value || !hasEnglishLetters(value)) return NodeFilter.FILTER_REJECT;
        const parent = node.parentElement;
        if (!parent) return NodeFilter.FILTER_REJECT;
        const tag = parent.tagName;
        if (tag === "SCRIPT" || tag === "STYLE" || tag === "NOSCRIPT") return NodeFilter.FILTER_REJECT;
        return NodeFilter.FILTER_ACCEPT;
      }
    });

    let node;
    while ((node = walker.nextNode())) {
      textNodes.push(node);
    }

    return textNodes;
  }

  async function autoTranslatePage(language, options) {
    const opts = options || {};
    const root = opts.root || document.body;
    const parallel = opts.parallelLimit || 12;

    if (!root || language === "en") return;

    const textNodes = collectTextNodes(root);
    const inputs = root.querySelectorAll("input[placeholder], textarea[placeholder]");
    const selectOptions = root.querySelectorAll("select option");

    const sourceTexts = [];

    textNodes.forEach(function (node) {
      const source = translationTextStore.get(node) || node.nodeValue;
      if (!translationTextStore.has(node)) translationTextStore.set(node, source);
      sourceTexts.push(source);
    });

    inputs.forEach(function (input) {
      if (!input.dataset.originalPlaceholder) {
        input.dataset.originalPlaceholder = input.getAttribute("placeholder") || "";
      }
      sourceTexts.push(input.dataset.originalPlaceholder);
    });

    selectOptions.forEach(function (option) {
      if (!option.dataset.originalText) {
        option.dataset.originalText = option.textContent;
      }
      sourceTexts.push(option.dataset.originalText);
    });

    const dictionary = await translateUniqueTexts(sourceTexts, language, parallel);

    textNodes.forEach(function (node) {
      const source = translationTextStore.get(node) || node.nodeValue;
      if (dictionary[source]) node.nodeValue = dictionary[source];
    });

    inputs.forEach(function (input) {
      const source = input.dataset.originalPlaceholder || "";
      if (dictionary[source]) input.setAttribute("placeholder", dictionary[source]);
    });

    selectOptions.forEach(function (option) {
      const source = option.dataset.originalText || option.textContent;
      if (dictionary[source]) option.textContent = dictionary[source];
    });
  }

  // Approval workflow helper functions for Event Designer
  function getEventApplications() {
    return readArray(STORAGE_KEYS.eventApplications);
  }

  function saveEventApplications(applications) {
    return writeJson(STORAGE_KEYS.eventApplications, applications);
  }

  function updateEventApplicationStatus(eventIndex, status, notes) {
    const apps = getEventApplications();
    if (apps[eventIndex]) {
      apps[eventIndex].status = status;
      apps[eventIndex].statusUpdatedAt = new Date().toISOString();
      if (notes) {
        apps[eventIndex].adminNotes = notes;
      }
      saveEventApplications(apps);
      return apps[eventIndex];
    }
    return null;
  }

  function getEventApplicationStats() {
    const apps = getEventApplications();
    const total = apps.length;
    const approved = apps.filter(e => e.status === "approved").length;
    const disapproved = apps.filter(e => e.status === "disapproved").length;
    const pending = apps.filter(e => !e.status || e.status === "pending").length;
    
    return {
      total: total,
      approved: approved,
      disapproved: disapproved,
      pending: pending,
      approvalRate: total > 0 ? Math.round((approved / total) * 100) : 0
    };
  }

  function getUserEventApplications(userId) {
    const apps = getEventApplications();
    const user = userId || getCurrentUser() || "guest";
    return apps.filter(e => (e.userId || "guest") === user);
  }

  window.AppShared = {
    STORAGE_KEYS,
    readArray,
    readObject,
    writeJson,
    currentDateValue,
    getCurrentUser,
    getProjectProfile,
    saveProjectProfile,
    addActivity,
    getChecklistPercent,
    getOverview,
    getReadinessActions,
    formatCurrency,
    exportCsv,
    translateText,
    translateUniqueTexts,
    autoTranslatePage,
    getEventApplications,
    saveEventApplications,
    updateEventApplicationStatus,
    getEventApplicationStats,
    getUserEventApplications
  };
})(window);