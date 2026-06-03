// ============================================================
// CS COMMAND CENTER v2 — Google Apps Script
// Deako Customer Support real-time dashboard
//
// Combines Zendesk email status + Aircall phone status.
// Branded to Deako design standards (Air Blue, Inter font).
//
// SETUP:
// 1. Open a new Google Sheet
// 2. Extensions > Apps Script
// 3. Paste this entire file into Code.gs (replace everything)
// 4. Go to Project Settings (gear icon) > Script Properties
//    Add these properties:
//      ZENDESK_TOKEN     = your_email/token:your_api_token
//      AIRCALL_API_ID    = your_aircall_api_id
//      AIRCALL_API_TOKEN = your_aircall_api_token
//      NICEREPLY_TOKEN   = your_email:your_nicereply_api_key (or just the key)
//      META_IG_TOKEN     = your_instagram_login_access_token (for IG DMs — see README)
//      ANTHROPIC_API_KEY  = your_anthropic_api_key (optional — for Queue Intelligence spike analysis)
//      SPIKE_MULTIPLIER   = 2.0 (optional — multiplier for spike detection, default 2x)
// 5. Select initializeSheet from dropdown, click Run
// 6. Select setupTrigger from dropdown, click Run
// ============================================================

// --- CONFIGURATION ---
const CONFIG = {
  zendesk: {
    subdomain: "deako",
    viewId: "40216633935767",
  },
  aircall: {
    baseUrl: "https://api.aircall.io/v1",
    // Only count calls on these CS support lines (digits format from Aircall API)
    supportNumbers: [
      "+18449030847",   // nonpro support
      "+18442033196",   // pro support
      "+14259880376",   // distributor support
    ],
    answeringServiceNumber: "+10000000000",  // external call answer service
  },
  agents: ["Agent A", "Agent B", "Agent C"],
  // Agents who use Zendesk but are NOT on the support team — exclude from all dashboard stats
  excludeAgents: ["Excluded"],
  // Agents to exclude from phone CSAT (PostCall surveys) — not on CS team
  excludePostCallAgents: ["AgentD"],
  // Aircall lines to exclude from SMS activity tracking
  excludeSMSLines: ["nonpro sales (post close)", "Agent D", "Agent E", "Agent F", "Agent G", "Agent H", "Agent I", "Agent J", "Agent K", "Agent L", "Agent M", "Agent N"],
  // Business hours for phone metrics (calls outside these hours excluded from answer rate)
  businessHours: {
    timezone: "America/Los_Angeles",  // Pacific
    startHour: 6,   // 6:00 AM
    endHour: 17,    // 5:00 PM
    workDays: [1, 2, 3, 4, 5],  // Mon=1 through Fri=5
  },
  // SLA TARGETS — loaded from Script Properties at runtime (see loadThresholds below)
  // Fallback defaults are used if properties aren't set.
  thresholds: null,  // populated by loadThresholds()
};

// Default thresholds — used when Script Properties aren't set
const DEFAULT_THRESHOLDS = {
  oldestUnanswered: { green: 12, yellow: 24 },     // hours — 12h SLA, 24h = critical
  openBacklog:      { green: 30, yellow: 50 },     // ticket count
  phoneAnswerRate:  { green: 75, yellow: 60 },     // % — Goal: 75%+ answer rate
  medianFRT:        { green: 12, yellow: 24 },     // hours — 12h = Green per SLA
  avgWaitTime:      { green: 30, yellow: 60 },     // seconds
  socialResponseTime: { green: 120, yellow: 360 }, // minutes — 2h = Healthy, 6h = At Risk
};

// Load thresholds from Script Properties with fallback to defaults.
// Script Properties (all optional):
//   SLA_EMAIL_GREEN=12         SLA_EMAIL_YELLOW=24
//   SLA_BACKLOG_GREEN=30       SLA_BACKLOG_YELLOW=50
//   SLA_PHONE_GREEN=75         SLA_PHONE_YELLOW=60
//   SLA_FRT_GREEN=12           SLA_FRT_YELLOW=24
//   SLA_WAIT_GREEN=30          SLA_WAIT_YELLOW=60
function loadThresholds() {
  const props = PropertiesService.getScriptProperties();
  const d = DEFAULT_THRESHOLDS;

  function num(key, fallback) {
    const val = props.getProperty(key);
    if (val === null || val === "") return fallback;
    const parsed = Number(val);
    return isNaN(parsed) ? fallback : parsed;
  }

  CONFIG.thresholds = {
    oldestUnanswered: {
      green:  num("SLA_EMAIL_GREEN",   d.oldestUnanswered.green),
      yellow: num("SLA_EMAIL_YELLOW",  d.oldestUnanswered.yellow),
    },
    openBacklog: {
      green:  num("SLA_BACKLOG_GREEN",  d.openBacklog.green),
      yellow: num("SLA_BACKLOG_YELLOW", d.openBacklog.yellow),
    },
    phoneAnswerRate: {
      green:  num("SLA_PHONE_GREEN",  d.phoneAnswerRate.green),
      yellow: num("SLA_PHONE_YELLOW", d.phoneAnswerRate.yellow),
    },
    medianFRT: {
      green:  num("SLA_FRT_GREEN",  d.medianFRT.green),
      yellow: num("SLA_FRT_YELLOW", d.medianFRT.yellow),
    },
    avgWaitTime: {
      green:  num("SLA_WAIT_GREEN",  d.avgWaitTime.green),
      yellow: num("SLA_WAIT_YELLOW", d.avgWaitTime.yellow),
    },
    socialResponseTime: {
      green:  num("SLA_SOCIAL_GREEN",  d.socialResponseTime.green),
      yellow: num("SLA_SOCIAL_YELLOW", d.socialResponseTime.yellow),
    },
  };
}

// --- DEAKO BRAND COLORS (from Logo Usage Guidelines 2025) ---
const BRAND = {
  // Primary
  white:          "#FAFAFA",
  black:          "#1D1D1D",
  beigeMedium:    "#CCC6C0",
  beigeLight:     "#E1DFDD",
  airBlueMedium:  "#7597A0",
  airBlueLight:   "#C3D3D7",
  // Secondary
  beigeDark:      "#523823",
  airBlueDark:    "#1B3747",
  mossGreen:      "#889578",
  mossGreenLight: "#BCC7B0",
  ashGray:        "#9AA19B",
  ashGrayLight:   "#BEC6BF",
  terracotta:     "#BA866A",
  terracottaLight:"#DEAC90",
  roseQuartz:     "#B692A1",
  roseQuartzLight:"#D6BDC8",
};

// Helper: get set of hidden IG sender IDs and usernames (filtered from dashboard + webhooks)
// Operators can paste either a sender ID (column A) or a username (column B) into the
// "Hidden IG Senders" sheet - both are checked. No Apps Script auth needed to edit the sheet.
function getHiddenIGSenders(ssOverride) {
  const ss = ssOverride || SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Hidden IG Senders");
  const hidden = new Set();
  if (sheet && sheet.getLastRow() > 1) {
    const cols = Math.min(sheet.getLastColumn(), 2); // columns A (ID) and B (Username)
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, cols).getValues();
    data.forEach(row => {
      if (row[0]) hidden.add(String(row[0]).trim());               // Sender ID
      if (cols >= 2 && row[1]) hidden.add(String(row[1]).trim());   // Username
    });
  }
  return hidden;
}

// --- MAIN REFRESH FUNCTION ---
function refreshDashboard() {
  // Skip refresh outside business hours to conserve Apps Script execution quota
  const bh = CONFIG.businessHours;
  const nowCheck = new Date();
  const pacificStr = nowCheck.toLocaleString("en-US", { timeZone: bh.timezone });
  const pacificNow = new Date(pacificStr);
  const dow = pacificNow.getDay();
  const hour = pacificNow.getHours();
  if (!bh.workDays.includes(dow) || hour < bh.startHour || hour >= bh.endHour) {
    return;  // outside Mon–Fri 6am–5pm Pacific — skip silently
  }

  loadThresholds();  // read SLA targets from Script Properties (falls back to defaults)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const runLog = getOrCreateSheet(ss, "Run Log");
  const startTime = new Date();

  try {
    const zendeskData = fetchZendeskStatus();
    const aircallData = fetchAircallStatus();
    const csatData = fetchNicereplyCSAT();
    const postCallData = readPostCallCSAT();
    const smsData = readSMSActivity();
    const metaData = fetchMetaStatus();

    // Queue Intelligence — volume spike detection + throughput (v2)
    let qiData = null;
    try {
      qiData = fetchQueueIntelligence(zendeskData);
    } catch (e) {
      Logger.log("Queue Intelligence fetch failed (non-fatal): " + e.toString());
    }

    writeZendeskRaw(ss, zendeskData);
    writeAircallRaw(ss, aircallData);
    writeDashboard(ss, zendeskData, aircallData, csatData, postCallData, smsData, metaData, qiData);

    // Update individual agent performance dashboards (separate spreadsheets)
    try { updateAgentDashboards(zendeskData, aircallData, csatData, postCallData); } catch (e) {
      Logger.log("Agent dashboards update failed (non-fatal): " + e.toString());
    }

    logRun(runLog, startTime, "SUCCESS", "");
  } catch (error) {
    logRun(runLog, startTime, "ERROR", error.toString());
    Logger.log("Dashboard refresh failed: " + error.toString());
  }
}

// --- ZENDESK API (search-based — no view dependency) ---
function fetchZendeskStatus() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("ZENDESK_TOKEN");
  if (!token) throw new Error("ZENDESK_TOKEN not set in Script Properties");

  const subdomain = CONFIG.zendesk.subdomain;
  const authHeader = "Basic " + Utilities.base64Encode(token);
  const headers = { "Authorization": authHeader, "Content-Type": "application/json" };
  const fetchOpts = { method: "get", headers: headers, muteHttpExceptions: true };

  // Helper to run a search/count query
  function zendeskSearchCount(query) {
    try {
      const searchUrl = `https://${subdomain}.zendesk.com/api/v2/search/count.json?query=${encodeURIComponent(query)}`;
      const resp = UrlFetchApp.fetch(searchUrl, fetchOpts);
      if (resp.getResponseCode() === 200) {
        return JSON.parse(resp.getContentText()).count || 0;
      }
    } catch (e) {
      Logger.log("Zendesk search failed for: " + query + " — " + e.toString());
    }
    return 0;
  }

  // Helper to run a full-object search query (returns ticket objects)
  function zendeskSearch(query, perPage) {
    try {
      const url = `https://${subdomain}.zendesk.com/api/v2/search.json?query=${encodeURIComponent(query)}&per_page=${perPage || 50}&sort_by=created_at&sort_order=desc`;
      const resp = UrlFetchApp.fetch(url, fetchOpts);
      if (resp.getResponseCode() === 200) return JSON.parse(resp.getContentText()).results || [];
    } catch (e) { Logger.log("Zendesk search failed: " + query + " -- " + e.toString()); }
    return [];
  }

  // Step 1: Search for all new + open tickets (customers waiting for a response)
  // Replaces the old view-based approach — no dependency on view configuration.
  let tickets = [];
  let searchPage = 1;
  let hasMore = true;
  while (hasMore && searchPage <= 5) {
    const query = "type:ticket status<pending";
    const searchUrl = `https://${subdomain}.zendesk.com/api/v2/search.json?query=${encodeURIComponent(query)}&per_page=100&page=${searchPage}&sort_by=created_at&sort_order=desc`;
    const searchResp = UrlFetchApp.fetch(searchUrl, fetchOpts);
    if (searchResp.getResponseCode() !== 200) {
      throw new Error(`Zendesk search API returned ${searchResp.getResponseCode()}: ${searchResp.getContentText().substring(0, 300)}`);
    }
    const searchData = JSON.parse(searchResp.getContentText());
    tickets = tickets.concat(searchData.results || []);
    hasMore = searchData.next_page;
    searchPage++;
  }

  // Filter out AI agent tickets (bot-only conversations that can't be edited)
  const aiAgentCount = tickets.filter(t => t.support_type === "ai_agent").length;
  tickets = tickets.filter(t => t.support_type !== "ai_agent");
  if (aiAgentCount > 0) {
    Logger.log("Excluded " + aiAgentCount + " AI agent tickets from dashboard");
  }

  // Step 2: Resolve user IDs to names/emails
  const userIds = new Set();
  tickets.forEach(t => {
    if (t.requester_id) userIds.add(t.requester_id);
    if (t.assignee_id) userIds.add(t.assignee_id);
  });

  const userMap = {};
  const userEmailMap = {};
  if (userIds.size > 0) {
    // Batch fetch users (show_many supports up to 100 IDs per call)
    const idArray = [...userIds];
    for (let i = 0; i < idArray.length; i += 100) {
      const batch = idArray.slice(i, i + 100).join(",");
      const usersUrl = `https://${subdomain}.zendesk.com/api/v2/users/show_many.json?ids=${batch}`;
      const usersResp = UrlFetchApp.fetch(usersUrl, fetchOpts);
      if (usersResp.getResponseCode() === 200) {
        const usersData = JSON.parse(usersResp.getContentText());
        (usersData.users || []).forEach(u => {
          userMap[u.id] = u.name || u.email || "Unknown";
          userEmailMap[u.id] = u.email || "";
        });
      }
    }
  }

  // Step 3: Fetch metric_sets via show_many (for SLA wait time calculation)
  const metricMap = {};
  if (tickets.length > 0) {
    // show_many supports up to 100 IDs per call
    for (let i = 0; i < tickets.length; i += 100) {
      const batch = tickets.slice(i, i + 100).map(t => t.id).join(",");
      const metricsUrl = `https://${subdomain}.zendesk.com/api/v2/tickets/show_many.json?ids=${batch}&include=metric_sets`;
      const metricsResp = UrlFetchApp.fetch(metricsUrl, fetchOpts);
      if (metricsResp.getResponseCode() === 200) {
        const metricsData = JSON.parse(metricsResp.getContentText());
        if (metricsData.metric_sets) {
          metricsData.metric_sets.forEach(ms => { metricMap[ms.ticket_id] = ms; });
        }
        if (metricsData.tickets) {
          metricsData.tickets.forEach(t => {
            if (t.metric_set && !metricMap[t.id]) metricMap[t.id] = t.metric_set;
          });
        }
      }
    }
  }

  const now = new Date();
  const slaMinutes = CONFIG.thresholds.oldestUnanswered.green * 60; // 12h = 720 min

  // Step 4: Process tickets with metrics
  const processed = tickets.map(ticket => {
    const created = new Date(ticket.created_at);
    const updated = new Date(ticket.updated_at);
    const ageHours = (now - created) / (1000 * 60 * 60);
    const ms = metricMap[ticket.id] || {};

    const assigneeName = ticket.assignee_id
      ? (userMap[ticket.assignee_id] || "Agent #" + ticket.assignee_id)
      : "Unassigned";
    const requesterName = ticket.requester_id
      ? (userMap[ticket.requester_id] || "User #" + ticket.requester_id)
      : "Unknown";
    const requesterEmail = ticket.requester_id
      ? (userEmailMap[ticket.requester_id] || "") : "";

    // Determine when the customer started waiting for THIS response:
    // - "new" tickets: customer has been waiting since created_at
    // - "open" tickets: customer replied back; use requester_updated_at
    let waitingSince;
    if (ticket.status === "new") {
      waitingSince = created;
    } else {
      const reqUpdated = ms.requester_updated_at
        ? new Date(ms.requester_updated_at)
        : updated;
      waitingSince = reqUpdated;
    }

    const waitBizMin = calcBusinessMinutes(waitingSince, now);
    const pastSla = waitBizMin > slaMinutes;

    return {
      id: ticket.id,
      subject: ticket.subject || "(no subject)",
      requester: requesterName,
      assignee: assigneeName,
      status: ticket.status || "unknown",
      priority: ticket.priority || "normal",
      created: created,
      updated: updated,
      ageHours: ageHours,
      tags: ticket.tags || [],
      requesterEmail: requesterEmail,
      waitingSince: waitingSince,
      waitBizMin: waitBizMin,
      pastSla: pastSla,
    };
  });

  // Sort by wait time descending (longest waiting first)
  processed.sort((a, b) => b.waitBizMin - a.waitBizMin);

  // Filter out excluded agents (people who use Zendesk but aren't on the support team)
  const excludeLower = (CONFIG.excludeAgents || []).map(n => n.toLowerCase());
  const filtered = processed.filter(t => {
    if (!t.assignee || t.assignee === "Unassigned") return true;
    const aLower = t.assignee.toLowerCase();
    return !excludeLower.some(ex => aLower.includes(ex));
  });

  // Step 5: Calculate metrics from the filtered ticket list
  const totalOpen = filtered.length;
  const slaHours = CONFIG.thresholds.oldestUnanswered.green;

  // Queue counts — derived from the same search results (no separate API calls needed)
  const openQueueCount = filtered.filter(t => t.status === "open").length;
  const onHoldQueueCount = zendeskSearchCount("type:ticket status:hold");
  const unassigned = filtered.filter(t => t.assignee === "Unassigned").length;
  const pastSlaTickets = filtered.filter(t => t.pastSla);
  const totalBreached = pastSlaTickets.length;

  // "No Reply 12h+" — tickets that have NEVER received a first agent response and are past SLA
  // In Zendesk, status "new" means no agent has replied yet
  const noReplyBreached = filtered.filter(t => t.status === "new" && t.pastSla).length;

  // SAS tickets — from the call answer service (belt-and-suspenders: 3 detection methods)
  // SAS tickets may not appear in the monitored view, so we search independently.
  // Method 1: sas_flex tag (added by automation)
  // Method 2: subject line match
  // Method 3: requester email (SAS always sends from notifications@sasdesk.com)
  const sasByTag = zendeskSearchCount('type:ticket status:new tags:sas_flex');
  const sasBySubject = zendeskSearchCount('type:ticket status:new subject:"You have a new call from SAS Flex"');
  const sasByRequester = zendeskSearchCount('type:ticket status:new requester:notifications@sasdesk.com');
  const sasTicketsView = filtered.filter(t =>
    t.status === "new" && (
      (t.subject && t.subject.toLowerCase().includes("you have a new call from sas flex"))
      || (t.requesterEmail && t.requesterEmail.toLowerCase() === "notifications@sasdesk.com")
      || (t.tags && t.tags.includes("sas_flex"))
    )
  ).length;
  const sasTickets = Math.max(sasByTag, sasBySubject, sasByRequester, sasTicketsView);

  // "Emails handled today" -- tickets solved or closed today, per agent
  // Excludes:
  //   - aircall: call logs already counted in phone activity section
  //   - internal__testing: Gleap QA test submissions
  //   - auto_close: Deako Main voicemails that auto-close immediately (no human work)
  //   - assignee "AI Agent": bot-handled tickets, not human work
  const solvedExclusions = `-tags:aircall -tags:internal__testing -tags:auto_close -assignee:"AI Agent"`;
  const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const handledToday = {};
  CONFIG.agents.forEach(agent => {
    const query = `type:ticket solved>=${todayStr} assignee:"${agent}" ${solvedExclusions}`;
    handledToday[agent] = zendeskSearchCount(query);
  });
  // Also count "Other" and total (same exclusions)
  const totalHandledQuery = `type:ticket solved>=${todayStr} ${solvedExclusions}`;
  const totalHandledToday = zendeskSearchCount(totalHandledQuery);
  const knownHandled = Object.values(handledToday).reduce((a, b) => a + b, 0);
  handledToday["Other"] = Math.max(0, totalHandledToday - knownHandled);

  // Count tickets created today (for daily metrics log)
  const createdQuery = `type:ticket created>=${todayStr} ${solvedExclusions}`;
  const ticketsCreatedToday = zendeskSearchCount(createdQuery);

  // Find who the "Other" solvers are so we can label the row (e.g. "Other (Manager)")
  const otherSolverNames = new Set();
  if (handledToday["Other"] > 0) {
    // Fetch a few "Other" solved tickets to extract assignee names
    const otherQuery = `type:ticket solved>=${todayStr} ${solvedExclusions}`
      + CONFIG.agents.map(a => ` -assignee:"${a}"`).join("");
    const otherTickets = zendeskSearch(otherQuery, 20);
    const userIds = [...new Set(otherTickets.map(t => t.assignee_id).filter(Boolean))];
    userIds.forEach(uid => {
      try {
        const uResp = UrlFetchApp.fetch(
          `https://${subdomain}.zendesk.com/api/v2/users/${uid}.json`, fetchOpts);
        if (uResp.getResponseCode() === 200) {
          const u = JSON.parse(uResp.getContentText()).user;
          if (u && u.name && u.name !== "AI Agent") {
            otherSolverNames.add(u.name.split(" ")[0]); // first name only
          }
        }
      } catch (e) { /* skip */ }
    });
  }

  // High priority / tagged tickets for visibility (builder warranty, etc.)
  const flaggedTickets = filtered.filter(t =>
    t.priority === "high" || t.priority === "urgent"
    || (t.tags && (
      t.tags.includes("builder_warranty")
      || t.tags.includes("warranty")
      || t.tags.includes("escalated")
      || t.tags.includes("vip")
    ))
  );

  // Per-agent breakdown
  const agentCounts = {};
  CONFIG.agents.forEach(a => agentCounts[a] = { assigned: 0, pastSla: 0, longestWaitMin: 0, handledToday: handledToday[a] || 0 });
  agentCounts["Other"] = { assigned: 0, pastSla: 0, longestWaitMin: 0, handledToday: handledToday["Other"] || 0 };
  const otherAgentNames = new Set();

  filtered.forEach(ticket => {
    const agent = ticket.assignee;
    if (agent === "Unassigned") return;
    const matched = CONFIG.agents.find(ca => {
      const parts = ca.toLowerCase().split(/\s+/);
      const agentLower = agent.toLowerCase();
      return ca === agent || parts.some(p => p.length > 1 && agentLower.split(/\s+/).some(ap => ap === p));
    });
    const bucket = matched ? agentCounts[matched] : agentCounts["Other"];
    if (!matched) otherAgentNames.add(agent);
    bucket.assigned++;
    if (ticket.pastSla) bucket.pastSla++;
    if (ticket.waitBizMin > bucket.longestWaitMin) {
      bucket.longestWaitMin = ticket.waitBizMin;
    }
  });

  // Count open voicemail tickets on support lines (pro + nonpro only, excludes Deako Main)
  const openVoicemails = zendeskSearchCount('subject:"Voicemail on pro support" status<solved')
    + zendeskSearchCount('subject:"Voicemail on nonpro support" status<solved');

  return {
    totalOpen,
    totalBreached,
    noReplyBreached,
    sasTickets,
    slaHours,
    otherAgentNames: [...new Set([...otherAgentNames, ...otherSolverNames])],
    unassigned,
    openQueueCount,
    onHoldQueueCount,
    openVoicemails,
    aiAgentCount,
    agentCounts,
    tickets: filtered,
    flaggedTickets,
    totalHandledToday,
    ticketsCreatedToday,
    // Top 10 longest-waiting tickets for the detail table
    longest10: filtered.slice(0, 10),
  };
}

// --- QUEUE INTELLIGENCE (v2) ---
// Computes inflow/outflow rates, estimated queue clear time, and detects volume spikes.
// Spike detection: if the last-hour ticket creation rate exceeds the 24h rolling average
// by more than SPIKE_MULTIPLIER (default 2x), a spike is flagged and optionally analyzed
// by the Anthropic Claude API for common themes.
function fetchQueueIntelligence(zendeskData) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("ZENDESK_TOKEN");
  if (!token) return null;

  const subdomain = CONFIG.zendesk.subdomain;
  const authHeader = "Basic " + Utilities.base64Encode(token);
  const headers = { "Authorization": authHeader, "Content-Type": "application/json" };
  const fetchOpts = { method: "get", headers: headers, muteHttpExceptions: true };

  function zendeskSearchCount(query) {
    try {
      const url = `https://${subdomain}.zendesk.com/api/v2/search/count.json?query=${encodeURIComponent(query)}`;
      const resp = UrlFetchApp.fetch(url, fetchOpts);
      if (resp.getResponseCode() === 200) return JSON.parse(resp.getContentText()).count || 0;
    } catch (e) { Logger.log("QI search failed: " + query + " -- " + e.toString()); }
    return 0;
  }

  function zendeskSearch(query, perPage) {
    try {
      const url = `https://${subdomain}.zendesk.com/api/v2/search.json?query=${encodeURIComponent(query)}&per_page=${perPage || 50}&sort_by=created_at&sort_order=desc`;
      const resp = UrlFetchApp.fetch(url, fetchOpts);
      if (resp.getResponseCode() === 200) return JSON.parse(resp.getContentText()).results || [];
    } catch (e) { Logger.log("QI search failed: " + query + " -- " + e.toString()); }
    return [];
  }

  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");

  // --- CREATED TODAY ---
  // Fetch full ticket objects for today so we can client-side filter out noise:
  //   1. AI agent tickets (support_type === "ai_agent") -- bot-handled, not real queue items
  //   2. Aircall call logs (tagged "aircall") -- already tracked in phone activity section
  //   3. Gleap internal testing (tagged "internal__testing") -- QA test submissions
  //   4. Auto-closed voicemails (tagged "auto_close") -- Deako Main line VMs that never enter queue
  // Real Gleap in-app submissions (tagged "gleap" but NOT "internal__testing") are kept.
  //
  // We fetch up to 100 tickets for filtering. If there are more than 100 today, we apply the
  // filter ratio from the sample to the full count (same approach as before, just daily window).
  const rawTodayCount = zendeskSearchCount(`type:ticket created>=${todayStr}`);
  const todayTicketsRaw = zendeskSearch(`type:ticket created>=${todayStr}`, 100);
  const todayTickets = todayTicketsRaw.filter(t => {
    if (t.support_type === "ai_agent") return false;
    if ((t.tags || []).includes("aircall")) return false;
    if ((t.tags || []).includes("internal__testing")) return false;
    if ((t.tags || []).includes("auto_close")) return false;
    return true;
  });
  const filterRatio = todayTicketsRaw.length > 0
    ? todayTickets.length / todayTicketsRaw.length : 1;
  const createdToday = rawTodayCount > 100
    ? Math.round(rawTodayCount * filterRatio) : todayTickets.length;
  Logger.log("QI filter: " + rawTodayCount + " raw today -> " + createdToday +
    " after filtering (ratio " + filterRatio.toFixed(2) + ")");

  // --- 7-DAY BASELINE ---
  // Count tickets created in the last 7 calendar days for a daily average baseline.
  // Uses count endpoint (cheap) with the same filter ratio applied.
  const sevenDaysAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  const sevenDaysAgoStr = Utilities.formatDate(sevenDaysAgo, tz, "yyyy-MM-dd");
  const raw7d = zendeskSearchCount(`type:ticket created>=${sevenDaysAgoStr}`);
  const created7d = Math.round(raw7d * filterRatio);
  const avgDailyCreated = Math.round((created7d / 7) * 10) / 10;

  // --- OUTFLOW ---
  const solvedToday = zendeskData ? (zendeskData.totalHandledToday || 0) : 0;
  const currentQueue = zendeskData ? zendeskData.totalOpen : 0;

  // Solve rate: solved today spread across business hours elapsed so far
  const pacificStr = now.toLocaleString("en-US", { timeZone: "America/Los_Angeles" });
  const pacificNow = new Date(pacificStr);
  const bizHoursElapsed = Math.max(1, pacificNow.getHours() - CONFIG.businessHours.startHour);
  const solveRatePerHour = Math.round((solvedToday / bizHoursElapsed) * 10) / 10;

  // Net today (positive = more created than solved, queue growing)
  const netToday = createdToday - solvedToday;

  // Estimated queue clear time (hours) based on current solve rate
  let estClearHours = null;
  if (solveRatePerHour > 0 && currentQueue > 0) {
    estClearHours = currentQueue / solveRatePerHour;
  }

  // --- SPIKE DETECTION (daily grain) ---
  // Spike = today's created count already exceeds the 7-day daily average * multiplier.
  // Use case: app issues, outages, or product problems that cause sustained elevated volume.
  // Min threshold prevents false positives on low-volume days.
  const spikeMultiplier = parseFloat(props.getProperty("SPIKE_MULTIPLIER")) || 2.0;
  const spikeThreshold = avgDailyCreated * spikeMultiplier;
  const isSpike = createdToday > spikeThreshold && createdToday >= 10;  // minimum 10 tickets to trigger

  // If spike detected, analyze today's tickets for common themes
  let spikeAnalysis = null;
  if (isSpike) {
    try {
      spikeAnalysis = analyzeSpikeTrend(todayTickets);
    } catch (e) {
      Logger.log("Spike analysis failed (non-fatal): " + e.toString());
      spikeAnalysis = "Analysis unavailable";
    }
  }

  return {
    createdToday,
    avgDailyCreated,
    created7d,
    solvedToday,
    solveRatePerHour,
    netToday,
    currentQueue,
    estClearHours: estClearHours !== null ? Math.round(estClearHours * 10) / 10 : null,
    isSpike,
    spikeMultiplier,
    spikeThreshold: Math.round(spikeThreshold * 10) / 10,
    spikeAnalysis,
    spikeTicketCount: todayTickets.length,
  };
}

// Analyze spike-window tickets using Anthropic Claude API to identify common themes.
// Returns a brief summary string (1-2 sentences) or null if no API key is set.
function analyzeSpikeTrend(tickets) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty("ANTHROPIC_API_KEY");
  if (!apiKey || tickets.length === 0) return null;

  // Build a condensed ticket summary for the prompt (minimize tokens)
  const ticketSummaries = tickets.slice(0, 30).map(t => {
    const tags = (t.tags || []).slice(0, 5).join(", ");
    const channel = t.via && t.via.channel ? t.via.channel : "unknown";
    const priority = t.priority || "normal";
    const desc = (t.description || "").substring(0, 200);
    return `- [${priority}] ${t.subject || "(no subject)"} (via ${channel}${tags ? ", tags: " + tags : ""}) -- ${desc}`;
  }).join("\n");

  const prompt = `You are analyzing a support ticket volume spike for Deako, a smart lighting company. ${tickets.length} tickets were created today, which is significantly above the daily average.

Here are the most recent tickets:
${ticketSummaries}

In 1-2 short sentences, identify the dominant theme(s) or common patterns. If there is an obvious product issue or app bug, name it. Be specific and actionable. Do not use bullet points.`;

  try {
    const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method: "post",
      headers: {
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "Content-Type": "application/json",
      },
      payload: JSON.stringify({
        model: "claude-haiku-4-5-20251001",
        max_tokens: 150,
        messages: [{ role: "user", content: prompt }],
      }),
      muteHttpExceptions: true,
    });

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data.content && data.content[0] && data.content[0].text) {
        return data.content[0].text.trim();
      }
    } else {
      Logger.log("Anthropic API error: " + response.getResponseCode() + " " + response.getContentText().substring(0, 200));
    }
  } catch (e) {
    Logger.log("Anthropic API call failed: " + e.toString());
  }
  return null;
}

// Analyze all of today's tickets + call intents to identify the top 3 support themes.
// Called by sendDailyRecap() to include AI-powered theme analysis in the end-of-day email.
// Returns an HTML string with the 3 themes, or null if no API key or no tickets.
function analyzeDailyThemes() {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty("ANTHROPIC_API_KEY");
  if (!apiKey) return null;

  const token = props.getProperty("ZENDESK_TOKEN");
  if (!token) return null;

  const subdomain = CONFIG.zendesk.subdomain;
  const authHeader = "Basic " + Utilities.base64Encode(token);
  const headers = { "Authorization": authHeader, "Content-Type": "application/json" };
  const fetchOpts = { method: "get", headers: headers, muteHttpExceptions: true };

  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  // Fetch today's tickets (paginate to get all, up to 300 max for token budget)
  let allTickets = [];
  let page = 1;
  let hasMore = true;
  while (hasMore && page <= 3) {
    try {
      const url = `https://${subdomain}.zendesk.com/api/v2/search.json?query=${encodeURIComponent("type:ticket created>=" + todayStr)}&per_page=100&page=${page}&sort_by=created_at&sort_order=desc`;
      const resp = UrlFetchApp.fetch(url, fetchOpts);
      if (resp.getResponseCode() === 200) {
        const data = JSON.parse(resp.getContentText());
        const results = data.results || [];
        allTickets = allTickets.concat(results);
        hasMore = results.length === 100 && allTickets.length < (data.count || 0);
        page++;
      } else { hasMore = false; }
    } catch (e) { hasMore = false; }
  }

  // Apply standard filters
  const tickets = allTickets.filter(t => {
    if (t.support_type === "ai_agent") return false;
    if ((t.tags || []).includes("internal__testing")) return false;
    if ((t.tags || []).includes("auto_close")) return false;
    return true;
  });

  if (tickets.length === 0) return null;

  // Separate email/web/messaging tickets from Aircall call tickets
  // For email/web: include full description (the customer's actual message)
  // For Aircall: description is just "Call Initiated" (useless), but ci- tags capture call intent
  const emailTickets = tickets.filter(t => !(t.tags || []).includes("aircall"));
  const callTickets = tickets.filter(t => (t.tags || []).includes("aircall"));

  // Build email/web ticket summaries with full descriptions and ticket IDs
  // Strip broken surrogate pairs and non-BMP chars that break JSON serialization
  const sanitizeText = (str) => (str || "").replace(/[\uD800-\uDFFF]/g, "");

  const emailSummaries = emailTickets.map(t => {
    const tags = (t.tags || []).filter(tag => !tag.match(/^\d+$/)).slice(0, 8).join(", ");
    const channel = t.via && t.via.channel ? t.via.channel : "unknown";
    const desc = sanitizeText((t.description || "").replace(/\n{3,}/g, "\n\n").substring(0, 800));
    const subj = sanitizeText(t.subject || "(no subject)");
    return `[#${t.id}] [${channel}] ${subj}${tags ? " (tags: " + tags + ")" : ""}\n${desc}`;
  });

  // Build call summaries from ci- tags (these are AI-generated call intents)
  const callIntents = {};
  callTickets.forEach(t => {
    (t.tags || []).filter(tag => tag.startsWith("ci-")).forEach(tag => {
      const label = tag.replace("ci-", "").replace(/-/g, " ");
      callIntents[label] = (callIntents[label] || 0) + 1;
    });
  });
  const callIntentSummary = Object.entries(callIntents)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 30)
    .map(([label, count]) => `  ${label}: ${count}`)
    .join("\n");

  // Build call ticket ID list for citation
  const callTicketIds = callTickets.map(t => t.id);

  // Build the prompt -- asks for ticket IDs so we can persist them
  // System prompt with taxonomy vocabulary (cached across calls)
  const systemPrompt = `You are analyzing daily customer support activity for Deako, a smart lighting company that makes smart switches, dimmers, backplates, faceplates, outlets, and a mobile app.

Use the following taxonomy vocabulary when naming themes. Themes should use these exact terms so they are consistent and aggregatable over time.

SYMPTOM CATEGORIES: Connectivity > WiFi (Cannot Connect, Intermittent Drops, Slow Response, Won't Reconnect After Change), Connectivity > Bluetooth (Cannot Pair, Lost Connection, Range Issues), Connectivity > Integration Link (Device Not Discovered, Commands Unresponsive, Shows Offline, State Out of Sync) -- ONLY for smart products that support integrations: Deako App, Smart Switch, Smart Switch Gen 2, Smart Switch Multiway, Smart Dimmer variants, Smart Plug. Simple products (Motion Switch, Rocker, Timer, Nightlight, Fan Speed Controller) do NOT support integrations and should NOT be tagged with Integration Link, Connectivity > Cloud / Remote Access (Commands Failed, State Out of Sync, Cloud Unreachable), Hardware > Electrical (No Power, Flickering, Buzzing/Humming, Breaker Tripping, Overheating/Thermal Shutoff, Dimming Issues), Hardware > Physical (Doesn't Fit, Button Stuck, Cosmetic Damage, Dead on Arrival), App (Crash/Freeze, UI/Display Issue, Firmware/App Update, Feature Not Working > Scheduling/Scenes/Notifications/Sharing)
REQUEST TYPES: Technical Issue, How-To/Education, Warranty Claim, Return/Exchange, Order Cancellation, Order Inquiry, Account Update, Compatibility Question, Purchase Inquiry, Feature Request, Scheduling/Dispatch, Missed Call/Voicemail
PRODUCT FAMILIES: Smart Switch, Smart Switch Gen 2, Smart Switch Multiway, Smart Dimmer (Single Pole, Master, Remote), Smart Plug, Simple Switches (Rocker, 3-Way, Multiway), Simple Dimmers, Specialty (Motion, Fan Speed, Timer, Nightlight), Backplates, Faceplates, Outlets, Deako App
NOTE on simple/specialty products: Motion Switch, Rocker, Timer, Nightlight, Fan Speed Controller are NOT smart products. They have NO app connectivity, NO WiFi, NO integrations, NO scheduling. If these products have issues, classify as Hardware symptoms only -- never App, never Integration Link, never Cloud/Remote Access.
PARTNERS: Safe Haven/ADT, D.R. Horton
CUSTOMER SENTIMENT — classify based ONLY on the customer's actual words and tone, not on the severity of the issue:
- Positive: customer says thank you, gives compliments, expresses satisfaction
- Neutral: customer describes what happened without emotional language. THIS IS THE DEFAULT. Most tickets are Neutral. Reporting a broken product, describing a failure, or listing symptoms is Neutral unless the customer explicitly expresses emotion. "My lights don't work" = Neutral. "All my lights stopped working after the update" = Neutral. "I can't connect" = Neutral.
- Frustrated: customer explicitly expresses frustration, uses words like "again," "still," "keeps happening," "I've tried everything," "this is ridiculous," or describes repeated failed attempts with exasperation. The emotion must be in their words, not inferred from the problem.
- Angry/Escalation Risk: customer threatens action (BBB, social media, legal), demands to speak to a manager, uses hostile or aggressive language

For each analysis, identify exactly 3 top themes ranked by ticket count (theme 1 = most tickets). For each theme:
1. Name it using the SPECIFIC leaf-level taxonomy term (e.g. "Connectivity > WiFi > Won't Reconnect After Change" not "Connectivity > WiFi" and not "Hardware > Electrical"). Do NOT group different symptoms under a parent category -- "No Power" and "Dimming Issues" are different themes, not one theme.
2. List the ticket IDs that relate to this theme
3. Write 1 sentence explaining the pattern
4. For each ticket in the theme, classify the customer's sentiment individually based on their actual words
5. Pick 1-2 representative tickets and quote the customer's own words (1 short line each) that best illustrate the theme
6. For any ticket classified as Frustrated or Angry, quote the specific phrase that shows the emotion

You MUST respond with valid JSON only, no other text. Use this exact format:
[
  {"theme": "Taxonomy-based Theme Name", "ticket_ids": [12345, 12346, 12347], "ticket_sentiments": {"12345": "Frustrated", "12346": "Neutral", "12347": "Frustrated"}, "summary": "One sentence explanation.", "customer_quotes": [{"ticket_id": 12345, "quote": "Customer's exact words here"}, {"ticket_id": 12346, "quote": "Customer's exact words here"}], "sentiment_evidence": [{"ticket_id": 12345, "sentiment": "Frustrated", "phrase": "The exact phrase showing frustration"}]},
  {"theme": "Taxonomy-based Theme Name", "ticket_ids": [12348, 12349], "ticket_sentiments": {"12348": "Neutral", "12349": "Neutral"}, "summary": "One sentence explanation.", "customer_quotes": [{"ticket_id": 12348, "quote": "Customer's exact words here"}], "sentiment_evidence": []},
  {"theme": "Taxonomy-based Theme Name", "ticket_ids": [12350, 12351], "ticket_sentiments": {"12350": "Frustrated", "12351": "Angry/Escalation Risk"}, "summary": "One sentence explanation.", "customer_quotes": [{"ticket_id": 12350, "quote": "Customer's exact words here"}], "sentiment_evidence": [{"ticket_id": 12350, "sentiment": "Frustrated", "phrase": "Exact phrase"}, {"ticket_id": 12351, "sentiment": "Angry/Escalation Risk", "phrase": "Exact phrase"}]}
]`;

  // User message with today's ticket data (changes every call)
  const userMessage = `TODAY'S VOLUME: ${tickets.length} total tickets (${emailTickets.length} email/web/messaging, ${callTickets.length} phone calls)

EMAIL/WEB/MESSAGING TICKETS (${emailTickets.length} tickets with full descriptions):
${emailSummaries.slice(0, 50).join("\n---\n")}

PHONE CALL TOPICS (${callTickets.length} calls, AI-tagged intents ranked by frequency):
${callIntentSummary || "(no call intent data)"}
Phone call ticket IDs: ${callTicketIds.join(", ")}

Identify the top 3 themes from today's activity.`;

  try {
    const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method: "post",
      headers: {
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "anthropic-beta": "prompt-caching-2024-07-31",
        "Content-Type": "application/json",
      },
      payload: JSON.stringify({
        model: "claude-haiku-4-5-20251001",
        max_tokens: 1500,
        system: [{ type: "text", text: systemPrompt, cache_control: { type: "ephemeral" } }],
        messages: [{ role: "user", content: userMessage }],
      }),
      muteHttpExceptions: true,
    });

    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data.content && data.content[0] && data.content[0].text) {
        const raw = data.content[0].text.trim();
        Logger.log("Daily themes analysis (raw): " + raw);

        // Parse JSON response
        let themes = [];
        try {
          themes = JSON.parse(raw);
        } catch (parseErr) {
          // Fallback: try to extract JSON from the response if wrapped in text
          const jsonMatch = raw.match(/\[[\s\S]*\]/);
          if (jsonMatch) {
            themes = JSON.parse(jsonMatch[0]);
          } else {
            Logger.log("Could not parse themes as JSON, falling back to text");
            const escaped = raw.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
            return { html: escaped.replace(/\n/g, "<br>"), themes: [] };
          }
        }

        // Sort themes by ticket count descending (stack rank: #1 = most tickets)
        themes.sort((a, b) => (b.ticket_ids || []).length - (a.ticket_ids || []).length);

        // Log themes with ticket IDs to the Daily Themes Log sheet
        try {
          logDailyThemes(todayStr, themes);
        } catch (logErr) {
          Logger.log("Theme logging failed (non-fatal): " + logErr.toString());
        }

        // Build HTML for email (human-readable format)
        const subdomain2 = CONFIG.zendesk.subdomain;
        let themesHtml = "";
        themes.forEach((t, i) => {
          const ticketCount = (t.ticket_ids || []).length;
          const escapedTheme = (t.theme || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
          const escapedSummary = (t.summary || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

          // Build clickable ticket links
          const ticketLinks = (t.ticket_ids || []).map(id =>
            `<a href="https://${subdomain2}.zendesk.com/agent/tickets/${id}" style="color:#7597A0;text-decoration:none;">#${id}</a>`
          ).join(", ");

          // Build sentiment sentence from per-ticket sentiments
          let sentimentSentence = "";
          const sentiments = t.ticket_sentiments || {};
          const sentVals = Object.values(sentiments);
          if (sentVals.length > 0) {
            const frustrated = sentVals.filter(s => s === "Frustrated").length;
            const angry = sentVals.filter(s => s === "Angry/Escalation Risk" || s === "Angry / Escalation Risk").length;
            const negative = frustrated + angry;
            const total = sentVals.length;
            if (negative === 0) {
              sentimentSentence = "Tone is neutral.";
            } else if (negative === total) {
              sentimentSentence = total === 1
                ? "Customer is " + (angry > 0 ? "angry." : "frustrated.")
                : "All " + total + " customers expressed frustration.";
            } else {
              const parts = [];
              if (frustrated > 0) parts.push(frustrated + " frustrated");
              if (angry > 0) parts.push(angry + " angry");
              sentimentSentence = parts.join(", ") + " of " + total + " customers.";
            }
          }

          const sentColor = sentimentSentence.includes("angry") ? "#C62828"
            : sentimentSentence.includes("frustrated") || sentimentSentence.includes("frustration") ? "#E65100"
            : "#888";

          // Theme header + summary
          themesHtml += `${i + 1}. <strong>${escapedTheme}</strong><br>`;
          themesHtml += `${ticketCount} tickets/calls (${ticketLinks}). ${escapedSummary}`;
          if (sentimentSentence) themesHtml += ` <span style="color:${sentColor};font-size:12px;">${sentimentSentence}</span>`;
          themesHtml += `<br>`;

          // Customer quotes
          const quotes = t.customer_quotes || [];
          if (quotes.length > 0) {
            quotes.forEach(q => {
              const escapedQuote = (q.quote || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
              const qLink = `<a href="https://${subdomain2}.zendesk.com/agent/tickets/${q.ticket_id}" style="color:#7597A0;text-decoration:none;">#${q.ticket_id}</a>`;
              themesHtml += `<div style="margin:4px 0 2px 16px;font-size:12px;color:#555;border-left:2px solid #CCC6C0;padding-left:8px;">"${escapedQuote}" <span style="color:#999;">${qLink}</span></div>`;
            });
          }

          // Sentiment evidence for frustrated/angry
          const sentEvidence = t.sentiment_evidence || [];
          if (sentEvidence.length > 0) {
            sentEvidence.forEach(se => {
              const escapedPhrase = (se.phrase || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
              const seColor = (se.sentiment || "").includes("Angry") ? "#C62828" : "#E65100";
              const seLink = `<a href="https://${subdomain2}.zendesk.com/agent/tickets/${se.ticket_id}" style="color:#7597A0;text-decoration:none;">#${se.ticket_id}</a>`;
              themesHtml += `<div style="margin:2px 0 2px 16px;font-size:11px;color:${seColor};border-left:2px solid ${seColor};padding-left:8px;">${se.sentiment}: "${escapedPhrase}" ${seLink}</div>`;
            });
          }

          if (i < themes.length - 1) themesHtml += `<br>`;
        });

        return { html: themesHtml, themes: themes };
      }
    } else {
      Logger.log("Daily themes API error: " + response.getResponseCode() + " " + response.getContentText().substring(0, 200));
    }
  } catch (e) {
    Logger.log("Daily themes API call failed: " + e.toString());
  }
  return null;
}

// ─── BACKFILL DAILY THEMES ───
// Runs theme analysis for a specific historical date and logs to Daily Themes Log.
// Uses THEME_BACKFILL_START and THEME_BACKFILL_END Script Properties.
// Run setupThemeBackfillTrigger() to start, or call backfillThemesOneDay() manually.
function backfillThemesOneDay() {
  loadThresholds();
  const props = PropertiesService.getScriptProperties();
  const tz = CONFIG.businessHours.timezone;
  const startDate = props.getProperty("THEME_BACKFILL_START");
  const endDate = props.getProperty("THEME_BACKFILL_END");
  if (!startDate || !endDate) {
    Logger.log("THEME_BACKFILL_START or THEME_BACKFILL_END not set — nothing to backfill");
    cleanupThemeBackfillTrigger();
    return;
  }

  const nextDate = props.getProperty("THEME_BACKFILL_NEXT") || startDate;
  if (nextDate > endDate) {
    Logger.log("Theme backfill complete through " + endDate);
    props.deleteProperty("THEME_BACKFILL_NEXT");
    props.deleteProperty("THEME_BACKFILL_START");
    props.deleteProperty("THEME_BACKFILL_END");
    cleanupThemeBackfillTrigger();
    return;
  }

  // Theme backfill runs on ALL days including weekends/holidays
  // (tickets still get created on non-working days)
  const dateObj = new Date(nextDate + "T12:00:00");

  // Check if already logged
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Daily Themes Log");
  if (logSheet) {
    const data = logSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === nextDate) {
        Logger.log("Already have theme data for " + nextDate + " — skipping");
        props.setProperty("THEME_BACKFILL_NEXT", advanceDateStr(nextDate));
        return;
      }
    }
  }

  Logger.log("Backfilling themes for: " + nextDate);

  const apiKey = props.getProperty("ANTHROPIC_API_KEY");
  if (!apiKey) { Logger.log("No ANTHROPIC_API_KEY"); return; }

  const token = props.getProperty("ZENDESK_TOKEN");
  if (!token) { Logger.log("No ZENDESK_TOKEN"); return; }

  const subdomain = CONFIG.zendesk.subdomain;
  const authHeader = "Basic " + Utilities.base64Encode(token);
  const headers = { "Authorization": authHeader, "Content-Type": "application/json" };
  const fetchOpts = { method: "get", headers: headers, muteHttpExceptions: true };

  const nextDatePlusOne = advanceDateStr(nextDate);

  // Fetch tickets created on that date
  let allTickets = [];
  let page = 1;
  let hasMore = true;
  while (hasMore && page <= 3) {
    try {
      const url = `https://${subdomain}.zendesk.com/api/v2/search.json?query=${encodeURIComponent("type:ticket created>=" + nextDate + " created<" + nextDatePlusOne)}&per_page=100&page=${page}&sort_by=created_at&sort_order=desc`;
      const resp = UrlFetchApp.fetch(url, fetchOpts);
      if (resp.getResponseCode() === 200) {
        const data = JSON.parse(resp.getContentText());
        const results = data.results || [];
        allTickets = allTickets.concat(results);
        hasMore = results.length === 100 && allTickets.length < (data.count || 0);
        page++;
      } else { hasMore = false; }
    } catch (e) { hasMore = false; }
  }

  // Apply standard filters
  const tickets = allTickets.filter(t => {
    if (t.support_type === "ai_agent") return false;
    if ((t.tags || []).includes("internal__testing")) return false;
    if ((t.tags || []).includes("auto_close")) return false;
    return true;
  });

  if (tickets.length === 0) {
    Logger.log("No tickets for " + nextDate + " — skipping");
    props.setProperty("THEME_BACKFILL_NEXT", nextDatePlusOne);
    return;
  }

  // Build summaries (same logic as analyzeDailyThemes)
  const emailTickets = tickets.filter(t => !(t.tags || []).includes("aircall"));
  const callTickets = tickets.filter(t => (t.tags || []).includes("aircall"));

  // Strip broken surrogate pairs and non-BMP chars that break JSON serialization
  const sanitizeText = (str) => (str || "").replace(/[\uD800-\uDFFF]/g, "");

  const emailSummaries = emailTickets.map(t => {
    const tags = (t.tags || []).filter(tag => !tag.match(/^\d+$/)).slice(0, 8).join(", ");
    const channel = t.via && t.via.channel ? t.via.channel : "unknown";
    const desc = sanitizeText((t.description || "").replace(/\n{3,}/g, "\n\n").substring(0, 800));
    const subj = sanitizeText(t.subject || "(no subject)");
    return `[#${t.id}] [${channel}] ${subj}${tags ? " (tags: " + tags + ")" : ""}\n${desc}`;
  });

  const callIntents = {};
  callTickets.forEach(t => {
    (t.tags || []).filter(tag => tag.startsWith("ci-")).forEach(tag => {
      const label = tag.replace("ci-", "").replace(/-/g, " ");
      callIntents[label] = (callIntents[label] || 0) + 1;
    });
  });
  const callIntentSummary = Object.entries(callIntents)
    .sort((a, b) => b[1] - a[1]).slice(0, 30)
    .map(([label, count]) => `  ${label}: ${count}`).join("\n");
  const callTicketIds = callTickets.map(t => t.id);

  // Use the same system prompt (cached)
  const systemPrompt = `You are analyzing daily customer support activity for Deako, a smart lighting company that makes smart switches, dimmers, backplates, faceplates, outlets, and a mobile app.

Use the following taxonomy vocabulary when naming themes. Themes should use these exact terms so they are consistent and aggregatable over time.

SYMPTOM CATEGORIES: Connectivity > WiFi (Cannot Connect, Intermittent Drops, Slow Response, Won't Reconnect After Change), Connectivity > Bluetooth (Cannot Pair, Lost Connection, Range Issues), Connectivity > Integration Link (Device Not Discovered, Commands Unresponsive, Shows Offline, State Out of Sync) -- ONLY for smart products that support integrations: Deako App, Smart Switch, Smart Switch Gen 2, Smart Switch Multiway, Smart Dimmer variants, Smart Plug. Simple products (Motion Switch, Rocker, Timer, Nightlight, Fan Speed Controller) do NOT support integrations and should NOT be tagged with Integration Link, Connectivity > Cloud / Remote Access (Commands Failed, State Out of Sync, Cloud Unreachable), Hardware > Electrical (No Power, Flickering, Buzzing/Humming, Breaker Tripping, Overheating/Thermal Shutoff, Dimming Issues), Hardware > Physical (Doesn't Fit, Button Stuck, Cosmetic Damage, Dead on Arrival), App (Crash/Freeze, UI/Display Issue, Firmware/App Update, Feature Not Working > Scheduling/Scenes/Notifications/Sharing)
REQUEST TYPES: Technical Issue, How-To/Education, Warranty Claim, Return/Exchange, Order Cancellation, Order Inquiry, Account Update, Compatibility Question, Purchase Inquiry, Feature Request, Scheduling/Dispatch, Missed Call/Voicemail
PRODUCT FAMILIES: Smart Switch, Smart Switch Gen 2, Smart Switch Multiway, Smart Dimmer (Single Pole, Master, Remote), Smart Plug, Simple Switches (Rocker, 3-Way, Multiway), Simple Dimmers, Specialty (Motion, Fan Speed, Timer, Nightlight), Backplates, Faceplates, Outlets, Deako App
NOTE on simple/specialty products: Motion Switch, Rocker, Timer, Nightlight, Fan Speed Controller are NOT smart products. They have NO app connectivity, NO WiFi, NO integrations, NO scheduling. If these products have issues, classify as Hardware symptoms only -- never App, never Integration Link, never Cloud/Remote Access.
PARTNERS: Safe Haven/ADT, D.R. Horton
CUSTOMER SENTIMENT — classify based ONLY on the customer's actual words and tone, not on the severity of the issue:
- Positive: customer says thank you, gives compliments, expresses satisfaction
- Neutral: customer describes what happened without emotional language. THIS IS THE DEFAULT. Most tickets are Neutral. Reporting a broken product, describing a failure, or listing symptoms is Neutral unless the customer explicitly expresses emotion. "My lights don't work" = Neutral. "All my lights stopped working after the update" = Neutral. "I can't connect" = Neutral.
- Frustrated: customer explicitly expresses frustration, uses words like "again," "still," "keeps happening," "I've tried everything," "this is ridiculous," or describes repeated failed attempts with exasperation. The emotion must be in their words, not inferred from the problem.
- Angry/Escalation Risk: customer threatens action (BBB, social media, legal), demands to speak to a manager, uses hostile or aggressive language

For each analysis, identify exactly 3 top themes ranked by ticket count (theme 1 = most tickets). For each theme:
1. Name it using the taxonomy vocabulary above
2. List the ticket IDs that relate to this theme
3. Write 1 sentence explaining the pattern
4. For each ticket in the theme, classify the customer's sentiment individually

You MUST respond with valid JSON only, no other text. Use this exact format:
[
  {"theme": "Taxonomy-based Theme Name", "ticket_ids": [12345, 12346], "ticket_sentiments": {"12345": "Frustrated", "12346": "Neutral"}, "summary": "One sentence."},
  {"theme": "Taxonomy-based Theme Name", "ticket_ids": [12348], "ticket_sentiments": {"12348": "Neutral"}, "summary": "One sentence."},
  {"theme": "Taxonomy-based Theme Name", "ticket_ids": [12350], "ticket_sentiments": {"12350": "Neutral"}, "summary": "One sentence."}
]`;

  const userMessage = `DATE: ${nextDate}
VOLUME: ${tickets.length} total tickets (${emailTickets.length} email/web/messaging, ${callTickets.length} phone calls)

EMAIL/WEB/MESSAGING TICKETS:
${emailSummaries.slice(0, 50).join("\n---\n")}

PHONE CALL TOPICS:
${callIntentSummary || "(no call intent data)"}
Phone call ticket IDs: ${callTicketIds.join(", ")}

Identify the top 3 themes from this day's activity.`;

  try {
    const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method: "post",
      headers: {
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "anthropic-beta": "prompt-caching-2024-07-31",
        "Content-Type": "application/json",
      },
      payload: JSON.stringify({
        model: "claude-haiku-4-5-20251001",
        max_tokens: 1000,
        system: [{ type: "text", text: systemPrompt, cache_control: { type: "ephemeral" } }],
        messages: [{ role: "user", content: userMessage }],
      }),
      muteHttpExceptions: true,
    });

    if (response.getResponseCode() === 200) {
      const respData = JSON.parse(response.getContentText());
      if (respData.content && respData.content[0] && respData.content[0].text) {
        const raw = respData.content[0].text.trim();
        let themes = [];
        try {
          themes = JSON.parse(raw);
        } catch (parseErr) {
          const jsonMatch = raw.match(/\[[\s\S]*\]/);
          if (jsonMatch) themes = JSON.parse(jsonMatch[0]);
        }

        if (themes.length > 0) {
          themes.sort((a, b) => (b.ticket_ids || []).length - (a.ticket_ids || []).length);
          logDailyThemes(nextDate, themes);
          Logger.log("Backfilled themes for " + nextDate + ": " + themes.length + " themes");
        }
      }
    } else {
      Logger.log("Theme backfill API error for " + nextDate + ": " + response.getResponseCode() + " — " + response.getContentText().substring(0, 500));
    }
  } catch (e) {
    Logger.log("Theme backfill failed for " + nextDate + ": " + e.toString());
  }

  props.setProperty("THEME_BACKFILL_NEXT", nextDatePlusOne);
}

function advanceDateStr(dateStr) {
  const d = new Date(dateStr + "T12:00:00");
  d.setDate(d.getDate() + 1);
  return d.toISOString().split("T")[0];
}

function setupThemeBackfillTrigger() {
  cleanupThemeBackfillTrigger();
  ScriptApp.newTrigger("backfillThemesOneDay")
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log("Theme backfill trigger created — runs every 5 minutes. Set THEME_BACKFILL_START and THEME_BACKFILL_END.");
}

function cleanupThemeBackfillTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "backfillThemesOneDay") {
      ScriptApp.deleteTrigger(t);
      Logger.log("Removed theme backfill trigger");
    }
  });
}

// ─── TICKET TAXONOMY SPOT-CHECK ───
// Manually-triggered function that classifies solved tickets from a date range
// using the full 16-field taxonomy. Writes results to a "Ticket Taxonomy" sheet.
// Set Script Properties: CLASSIFY_START_DATE, CLASSIFY_END_DATE
// Run classifyOneTicket() manually or use setupClassifyTrigger() for batch.

const TAXONOMY_SYSTEM_PROMPT = `You are a support ticket classifier for Deako, a smart lighting company. Classify the ticket using these 16 fields. Use ONLY the exact values listed. Do not invent new values.

LIFECYCLE STAGE: Pre-Purchase > (Compatibility Check | Product Selection / Sizing | Pricing / Availability | Where to Buy), Purchase / Fulfillment > (Order Placed | Shipping / Tracking | Delivery Issue), Install > (Wiring / Electrical | Physical Fit / Mounting | App Setup / Pairing | Network Configuration), Post-Install > (Daily Use | Configuration / Settings | Expansion / Adding Devices | Moving / Relocating), Ownership > (Maintenance | End of Life / Disposal), Unknown / Not Inferable
NOTE on Daily Use vs Configuration / Settings: Daily Use = something broke or stopped working during normal operation. Configuration / Settings = customer is actively trying to change or set up how their product works. "App update broke my switches" = Daily Use (customer didn't initiate a config change). "Schedule is firing when disabled" = Daily Use (malfunction, not configuration). "I'm trying to set up a schedule" or "How do I change my WiFi" = Configuration / Settings.

REQUESTER ROLE: Homeowner, Integrator > (Residential | Commercial | Multi-Dwelling Unit), Electrician, Builder / Developer > (Custom Home | Production | Commercial), Retailer / Dealer > (Online Retailer | Local Dealer | Distributor), Internal / Employee, Other

PRODUCT PRIMARY: Smart Switch, Smart Switch Gen 2, Smart Switch Multiway, Single Pole Smart Dimmer, Master Smart Dimmer, Remote Smart Dimmer, Smart Plug, Single Pole Rocker, 3-Way Rocker, Multiway Rocker, Simple Dimmer, Simple Dimmer (Square), Motion Switch, Fan Speed Controller, Astronomical Timer, Nightlight, Ventilation Timer, Backplates (Wired | Quick Wire | Universal), Faceplates (Standard | Medallion | Beswitched), Simple Outlet, USB Outlet, GFCI Outlet, Outlet Covers, Deako App > (iOS | Android | Web / Cloud | Cloud / Backend), No Specific Product, Smart Product - Unspecified, Simple Product - Unspecified, Unknown / Not Mentioned, Not Deako Product

EXTERNAL SYSTEM: Voice Assistant > (Amazon Alexa | Google Assistant | Apple Siri / HomeKit), Smart Home Hub > (Josh.ai | Control4 | Savant | Crestron | Other Hub), Smart Home Platform > (SmartThings | Home Assistant | Hubitat | IFTTT | Other Platform), Security System > (Alarm.com | Ring | Other Security), Network Environment > (WiFi Router / AP | Mesh Network | ISP / Provider), Electrical Environment > (Load Type / Bulb | Circuit / Panel | Transformer), Protocol > (Matter | Thread | Z-Wave | Zigbee), Other External System, or null

REQUEST TYPE: Technical Issue, How-To / Education, Warranty Claim, Return / Exchange, Order Cancellation, Order Inquiry, Account Update, Compatibility Question, Purchase Inquiry, Feature Request, Scheduling / Dispatch, Missed Call / Voicemail, Not Applicable, Other
NOTE: Technical Issue = customer reports their Deako product is malfunctioning. Compatibility Question = customer asks whether Deako works with something, OR a third party claims Deako causes interference/incompatibility. If the customer is not reporting a Deako product failure, do not classify as Technical Issue and do not force a Symptom.

SYMPTOM (only when Technical Issue): Connectivity > WiFi > (Cannot Connect | Intermittent Drops | Slow Response | Won't Reconnect After Change), Connectivity > Bluetooth > (Cannot Pair | Lost Connection | Range Issues), Connectivity > Integration Link > (Device Not Discovered | Commands Unresponsive | Shows Offline | State Out of Sync), Connectivity > Cloud / Remote Access > (Commands Failed | State Out of Sync | Cloud Unreachable), Hardware > Electrical > (No Power | Flickering | Unexpected Noise | Breaker Tripping | Overheating / Thermal Shutoff | Dimming Issues), Hardware > Physical > (Doesn't Fit | Button Stuck / Unresponsive | Cosmetic Damage | Dead on Arrival), Hardware > LED / Light Output > (Won't Dim Properly | Wrong Color Temp | Partial Illumination), App > Crash / Freeze > (Won't Launch | Crashes During Use | Crashes After Update), App > (UI / Display Issue | Firmware / App Update), App > Feature Not Working > (Scheduling | Scenes / Groups | Notifications | Sharing / Multi-User), or null

ROOT CAUSE STATUS: Not Applicable, Hypothesized, Unknown After Troubleshooting
ROOT CAUSE HYPOTHESIS: User Error > Configuration > (Incorrect Credentials | Wrong Setting Selected | Network Misconfiguration), User Error > Installation > (Wiring Issue | Incompatible Load Type | Missing Neutral Wire), User Error > Misuse / Misunderstanding, Software Bug > (App Bug | Firmware Bug | Cloud / Backend Bug), Hardware Defect > (Manufacturing Defect | Component Failure | Wear / Degradation), Environment > (Electrical | Network | Physical), Design / Limitation > (Known Limitation | Feature Not Supported | Compatibility Gap), Third-Party Issue > (Integration Partner | ISP / Network Provider | Carrier / Logistics), or null

ACTION TAKEN (primary + additional array): Troubleshooting > (Power Cycle | Factory Reset | Delete / Re-add Device | Device Linking | Reconfiguration | Reinstall / Rewire Hardware | Reinstall App | Firmware Update | App Update | Permissions / Settings Check | Network Fix), Replacement Initiated > (RMA Created Warranty | RMA Created Goodwill | DOA Replacement), Education Provided > (How-To Guidance | Compatibility Info | Feature Explanation | Referred to Docs / Video), Escalated > (Engineering | Product Team | Sales / Account Mgmt | Integration Partner), Status Communication > (Outage Notice | Known Issue Acknowledgment | Maintenance Window), Redirected to Partner > (Safe Haven / ADT | Other Partner), Scheduled Service Visit, Return / Refund Processed > (Full Refund | Partial Refund | Exchange), Outbound Follow-Up, Information Collected, Pending Customer Response, No Action

CLOSURE OUTCOME: Resolved - Customer Confirmed, Resolved - Agent Determined, Partially Resolved, Replaced, Refunded, Closed - Systemic Issue Pending, Unresolved, Duplicate, Spam / Not Support, No Content / Abandoned, Customer Self-Resolved, No Customer Response

CHANNEL: Email, Phone, Voicemail, Live Chat, Social Media > (Facebook | Instagram | Other Social), In-App, Web Form

RECORD SCOPE: Deako Product Support, Partner / Installer Handoff, Not Deako Product, Order / Fulfillment, Account / Access, Feature Request, Internal Test, Spam / Not Support, No Content / Abandoned Contact, Duplicate

PARTNER / BUILDER: Safe Haven / ADT, D.R. Horton, Alarm.com, Other Partner / Builder, or null

CUSTOMER SENTIMENT — classify based ONLY on the customer's actual words and tone, not on the severity of the issue:
- Positive: customer says thank you, gives compliments, expresses satisfaction
- Neutral: customer describes what happened without emotional language. THIS IS THE DEFAULT. Most tickets are Neutral. Reporting a broken product, describing a failure, or listing symptoms is Neutral unless the customer explicitly expresses emotion.
- Frustrated: customer explicitly expresses frustration, uses words like "again," "still," "keeps happening," "I've tried everything," "this is ridiculous," or describes repeated failed attempts with exasperation. The emotion must be in their words, not inferred from the problem.
- Angry/Escalation Risk: customer threatens action (BBB, social media, legal), demands to speak to a manager, uses hostile or aggressive language

FEATURE REQUESTED: free text, only when Request Type = Feature Request, otherwise null

Respond with valid JSON only. Use this format:
{"ticket_id":0,"lifecycle_stage":"","requester_role":"","product_primary":"","product_related":[],"external_system":null,"request_type":"","symptom":null,"root_cause_status":"","root_cause_hypothesis":null,"primary_action":"","additional_actions":[],"closure_outcome":"","channel":"","record_scope":"","partner_builder":null,"customer_sentiment":"","feature_requested":null,"confidence_notes":""}`;

function classifyOneTicket() {
  loadThresholds();
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty("ANTHROPIC_API_KEY");
  const token = props.getProperty("ZENDESK_TOKEN");
  if (!apiKey || !token) { Logger.log("Missing ANTHROPIC_API_KEY or ZENDESK_TOKEN"); return; }

  const startDate = props.getProperty("CLASSIFY_START_DATE");
  const endDate = props.getProperty("CLASSIFY_END_DATE");
  if (!startDate || !endDate) {
    Logger.log("CLASSIFY_START_DATE or CLASSIFY_END_DATE not set");
    cleanupClassifyTrigger();
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Ticket Taxonomy");
  if (!sheet) {
    sheet = ss.insertSheet("Ticket Taxonomy");
    sheet.appendRow([
      "Ticket ID", "Subject", "Status", "Created", "Channel",
      "Lifecycle Stage", "Requester Role", "Product Primary", "Product Related",
      "External System", "Request Type", "Symptom", "Root Cause Status",
      "Root Cause Hypothesis", "Primary Action", "Additional Actions",
      "Closure Outcome", "Record Scope", "Partner / Builder",
      "Customer Sentiment", "Feature Requested", "Confidence Notes", "Classified At"
    ]);
    sheet.getRange(1, 1, 1, 23).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  // Get already-classified ticket IDs to skip
  const existingData = sheet.getDataRange().getValues();
  const classifiedIds = new Set();
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][0]) classifiedIds.add(String(existingData[i][0]));
  }

  // Fetch solved tickets in date range
  const subdomain = CONFIG.zendesk.subdomain;
  const authHeader = "Basic " + Utilities.base64Encode(token);
  const zdHeaders = { "Authorization": authHeader, "Content-Type": "application/json" };
  const zdOpts = { method: "get", headers: zdHeaders, muteHttpExceptions: true };

  const solvedFilter = `-tags:aircall -tags:internal__testing -tags:auto_close -assignee:"AI Agent"`;
  const query = `type:ticket solved>=${startDate} solved<=${endDate} ${solvedFilter}`;

  // Find the next unclassified ticket
  let page = 1;
  let found = false;
  while (!found && page <= 10) {
    const url = `https://${subdomain}.zendesk.com/api/v2/search.json?query=${encodeURIComponent(query)}&per_page=25&page=${page}&sort_by=created_at&sort_order=asc`;
    const resp = UrlFetchApp.fetch(url, zdOpts);
    if (resp.getResponseCode() !== 200) break;
    const data = JSON.parse(resp.getContentText());
    const results = data.results || [];
    if (results.length === 0) break;

    for (const ticket of results) {
      if (classifiedIds.has(String(ticket.id))) continue;

      // Found an unclassified ticket — fetch full comments
      Logger.log("Classifying ticket #" + ticket.id + ": " + ticket.subject);
      let comments = "";
      try {
        const commUrl = `https://${subdomain}.zendesk.com/api/v2/tickets/${ticket.id}/comments.json?per_page=50`;
        const commResp = UrlFetchApp.fetch(commUrl, zdOpts);
        if (commResp.getResponseCode() === 200) {
          const commData = JSON.parse(commResp.getContentText());
          comments = (commData.comments || []).map(c => {
            const author = c.author_id ? `[Author ${c.author_id}]` : "[Unknown]";
            const pub = c.public ? "PUBLIC" : "INTERNAL";
            const body = (c.plain_body || c.body || "").substring(0, 1500);
            return `${author} (${pub}):\n${body}`;
          }).join("\n---\n");
        }
      } catch (e) {
        Logger.log("Failed to fetch comments for #" + ticket.id + ": " + e);
      }

      const tags = (ticket.tags || []).join(", ");
      const channel = ticket.via && ticket.via.channel ? ticket.via.channel : "unknown";
      const desc = (ticket.description || "").substring(0, 2000);

      const userMessage = `TICKET #${ticket.id}
Subject: ${ticket.subject || "(no subject)"}
Status: ${ticket.status}
Channel: ${channel}
Tags: ${tags}
Created: ${ticket.created_at}
Updated: ${ticket.updated_at}

DESCRIPTION:
${desc}

ALL COMMENTS:
${comments || "(no comments)"}

Classify this ticket.`;

      try {
        const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
          method: "post",
          headers: {
            "x-api-key": apiKey,
            "anthropic-version": "2023-06-01",
            "anthropic-beta": "prompt-caching-2024-07-31",
            "Content-Type": "application/json",
          },
          payload: JSON.stringify({
            model: "claude-haiku-4-5-20251001",
            max_tokens: 1200,
            system: [{ type: "text", text: TAXONOMY_SYSTEM_PROMPT, cache_control: { type: "ephemeral" } }],
            messages: [{ role: "user", content: userMessage }],
          }),
          muteHttpExceptions: true,
        });

        if (response.getResponseCode() === 200) {
          const respData = JSON.parse(response.getContentText());
          if (respData.content && respData.content[0] && respData.content[0].text) {
            const raw = respData.content[0].text.trim();
            let result = {};
            try {
              result = JSON.parse(raw);
            } catch (parseErr) {
              const jsonMatch = raw.match(/\{[\s\S]*\}/);
              if (jsonMatch) result = JSON.parse(jsonMatch[0]);
              else { Logger.log("Could not parse classification for #" + ticket.id); }
            }

            sheet.appendRow([
              ticket.id,
              ticket.subject || "",
              ticket.status || "",
              ticket.created_at || "",
              channel,
              result.lifecycle_stage || "",
              result.requester_role || "",
              result.product_primary || "",
              (result.product_related || []).join(", "),
              result.external_system || "",
              result.request_type || "",
              result.symptom || "",
              result.root_cause_status || "",
              result.root_cause_hypothesis || "",
              result.primary_action || "",
              (result.additional_actions || []).join(", "),
              result.closure_outcome || "",
              result.record_scope || "",
              result.partner_builder || "",
              result.customer_sentiment || "",
              result.feature_requested || "",
              result.confidence_notes || "",
              new Date().toISOString(),
            ]);

            Logger.log("Classified #" + ticket.id + " -> " + result.request_type + " / " + result.customer_sentiment);
          }
        } else {
          Logger.log("Classification API error for #" + ticket.id + ": " + response.getResponseCode());
        }
      } catch (e) {
        Logger.log("Classification failed for #" + ticket.id + ": " + e.toString());
      }

      found = true;
      break;
    }

    if (!found) page++;
  }

  if (!found) {
    Logger.log("All tickets in range " + startDate + " to " + endDate + " have been classified");
    cleanupClassifyTrigger();
  }
}

function setupClassifyTrigger() {
  cleanupClassifyTrigger();
  ScriptApp.newTrigger("classifyOneTicket")
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log("Classify trigger created — runs every 5 minutes. Set CLASSIFY_START_DATE and CLASSIFY_END_DATE.");
}

function cleanupClassifyTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "classifyOneTicket") {
      ScriptApp.deleteTrigger(t);
      Logger.log("Removed classify trigger");
    }
  });
}

// ─── OVERNIGHT THEME TICKET CLASSIFIER ───
// Classifies all tickets referenced in today's Daily Themes Log using the full 16-field taxonomy.
// Designed to run overnight so results are ready for review next business morning.
// Run setupOvernightClassifyTrigger() once — it creates a 2am daily trigger.
// The 2am trigger starts a 5-min recurring trigger that processes one ticket at a time,
// then cleans itself up when all theme tickets are classified.

function startOvernightClassify() {
  // Called by the 2am daily trigger. Reads today's theme ticket IDs and kicks off classification.
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Daily Themes Log");
  if (!logSheet) { Logger.log("No Daily Themes Log sheet — nothing to classify"); return; }

  // Collect all ticket IDs from today's themes
  const data = logSheet.getDataRange().getValues();
  const ticketIds = new Set();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === todayStr && data[i][4]) {
      String(data[i][4]).split(",").forEach(id => {
        const trimmed = id.trim();
        if (trimmed && !isNaN(trimmed)) ticketIds.add(trimmed);
      });
    }
  }

  if (ticketIds.size === 0) {
    Logger.log("No theme ticket IDs for " + todayStr + " — nothing to classify");
    return;
  }

  // Also check yesterday in case recap ran late and themes are dated yesterday
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const yesterdayStr = Utilities.formatDate(yesterday, tz, "yyyy-MM-dd");
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === yesterdayStr && data[i][4]) {
      String(data[i][4]).split(",").forEach(id => {
        const trimmed = id.trim();
        if (trimmed && !isNaN(trimmed)) ticketIds.add(trimmed);
      });
    }
  }

  // Store the queue in Script Properties
  const props = PropertiesService.getScriptProperties();
  props.setProperty("OVERNIGHT_CLASSIFY_IDS", Array.from(ticketIds).join(","));
  Logger.log("Overnight classify queue: " + ticketIds.size + " tickets for " + todayStr);

  // Start the recurring trigger to process one at a time
  cleanupOvernightProcessTrigger();
  ScriptApp.newTrigger("classifyNextThemeTicket")
    .timeBased()
    .everyMinutes(5)
    .create();
  Logger.log("Overnight classify trigger started");
}

function classifyNextThemeTicket() {
  loadThresholds();
  const props = PropertiesService.getScriptProperties();
  const idsStr = props.getProperty("OVERNIGHT_CLASSIFY_IDS") || "";
  const ids = idsStr.split(",").map(s => s.trim()).filter(s => s);

  if (ids.length === 0) {
    Logger.log("Overnight classify complete — no more tickets in queue");
    props.deleteProperty("OVERNIGHT_CLASSIFY_IDS");
    cleanupOvernightProcessTrigger();
    return;
  }

  const apiKey = props.getProperty("ANTHROPIC_API_KEY");
  const token = props.getProperty("ZENDESK_TOKEN");
  if (!apiKey || !token) { Logger.log("Missing API keys"); return; }

  // Check if already classified
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Ticket Taxonomy");
  if (!sheet) {
    sheet = ss.insertSheet("Ticket Taxonomy");
    sheet.appendRow([
      "Ticket ID", "Subject", "Status", "Created", "Channel",
      "Lifecycle Stage", "Requester Role", "Product Primary", "Product Related",
      "External System", "Request Type", "Symptom", "Root Cause Status",
      "Root Cause Hypothesis", "Primary Action", "Additional Actions",
      "Closure Outcome", "Record Scope", "Partner / Builder",
      "Customer Sentiment", "Feature Requested", "Confidence Notes", "Classified At"
    ]);
    sheet.getRange(1, 1, 1, 23).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  const existingData = sheet.getDataRange().getValues();
  const classifiedIds = new Set();
  for (let i = 1; i < existingData.length; i++) {
    if (existingData[i][0]) classifiedIds.add(String(existingData[i][0]));
  }

  // Find the next unclassified ticket
  let processedId = null;
  for (const ticketId of ids) {
    if (classifiedIds.has(ticketId)) {
      processedId = ticketId;
      Logger.log("Ticket #" + ticketId + " already classified — skipping");
      break;
    }

    // Fetch the ticket
    const subdomain = CONFIG.zendesk.subdomain;
    const authHeader = "Basic " + Utilities.base64Encode(token);
    const zdOpts = { method: "get", headers: { "Authorization": authHeader, "Content-Type": "application/json" }, muteHttpExceptions: true };

    try {
      const tickResp = UrlFetchApp.fetch(`https://${subdomain}.zendesk.com/api/v2/tickets/${ticketId}.json`, zdOpts);
      if (tickResp.getResponseCode() !== 200) {
        Logger.log("Could not fetch ticket #" + ticketId + ": " + tickResp.getResponseCode());
        processedId = ticketId;
        break;
      }
      const ticket = JSON.parse(tickResp.getContentText()).ticket;

      // Fetch comments
      let comments = "";
      try {
        const commResp = UrlFetchApp.fetch(`https://${subdomain}.zendesk.com/api/v2/tickets/${ticketId}/comments.json?per_page=50`, zdOpts);
        if (commResp.getResponseCode() === 200) {
          const commData = JSON.parse(commResp.getContentText());
          comments = (commData.comments || []).map(c => {
            const author = c.author_id ? `[Author ${c.author_id}]` : "[Unknown]";
            const pub = c.public ? "PUBLIC" : "INTERNAL";
            const body = (c.plain_body || c.body || "").substring(0, 1500);
            return `${author} (${pub}):\n${body}`;
          }).join("\n---\n");
        }
      } catch (e) { Logger.log("Comments fetch failed for #" + ticketId); }

      const tags = (ticket.tags || []).join(", ");
      const channel = ticket.via && ticket.via.channel ? ticket.via.channel : "unknown";
      const desc = (ticket.description || "").substring(0, 2000);

      const userMessage = `TICKET #${ticketId}
Subject: ${ticket.subject || "(no subject)"}
Status: ${ticket.status}
Channel: ${channel}
Tags: ${tags}
Created: ${ticket.created_at}
Updated: ${ticket.updated_at}

DESCRIPTION:
${desc}

ALL COMMENTS:
${comments || "(no comments)"}

Classify this ticket.`;

      const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
        method: "post",
        headers: {
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01",
          "anthropic-beta": "prompt-caching-2024-07-31",
          "Content-Type": "application/json",
        },
        payload: JSON.stringify({
          model: "claude-haiku-4-5-20251001",
          max_tokens: 1200,
          system: [{ type: "text", text: TAXONOMY_SYSTEM_PROMPT, cache_control: { type: "ephemeral" } }],
          messages: [{ role: "user", content: userMessage }],
        }),
        muteHttpExceptions: true,
      });

      if (response.getResponseCode() === 200) {
        const respData = JSON.parse(response.getContentText());
        if (respData.content && respData.content[0] && respData.content[0].text) {
          const raw = respData.content[0].text.trim();
          let result = {};
          try {
            result = JSON.parse(raw);
          } catch (parseErr) {
            const jsonMatch = raw.match(/\{[\s\S]*\}/);
            if (jsonMatch) result = JSON.parse(jsonMatch[0]);
          }

          sheet.appendRow([
            ticket.id,
            ticket.subject || "",
            ticket.status || "",
            ticket.created_at || "",
            channel,
            result.lifecycle_stage || "",
            result.requester_role || "",
            result.product_primary || "",
            (result.product_related || []).join(", "),
            result.external_system || "",
            result.request_type || "",
            result.symptom || "",
            result.root_cause_status || "",
            result.root_cause_hypothesis || "",
            result.primary_action || "",
            (result.additional_actions || []).join(", "),
            result.closure_outcome || "",
            result.record_scope || "",
            result.partner_builder || "",
            result.customer_sentiment || "",
            result.feature_requested || "",
            result.confidence_notes || "",
            new Date().toISOString(),
          ]);
          Logger.log("Overnight classified #" + ticketId + " -> " + result.request_type);
        }
      } else {
        Logger.log("API error classifying #" + ticketId + ": " + response.getResponseCode());
      }
    } catch (e) {
      Logger.log("Classification failed for #" + ticketId + ": " + e.toString());
    }

    processedId = ticketId;
    break;
  }

  // Remove processed ID from queue
  if (processedId) {
    const remaining = ids.filter(id => id !== processedId);
    if (remaining.length > 0) {
      props.setProperty("OVERNIGHT_CLASSIFY_IDS", remaining.join(","));
      Logger.log(remaining.length + " tickets remaining in overnight queue");
    } else {
      props.deleteProperty("OVERNIGHT_CLASSIFY_IDS");
      cleanupOvernightProcessTrigger();
      Logger.log("Overnight classify complete — all tickets processed");
    }
  }
}

function setupOvernightClassifyTrigger() {
  // Remove any existing overnight triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "startOvernightClassify") {
      ScriptApp.deleteTrigger(t);
    }
  });
  // Daily trigger at 2am to kick off overnight classification
  ScriptApp.newTrigger("startOvernightClassify")
    .timeBased()
    .atHour(2)
    .everyDays(1)
    .create();
  Logger.log("Overnight classify trigger created — runs daily at 2am");
}

function cleanupOvernightProcessTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "classifyNextThemeTicket") {
      ScriptApp.deleteTrigger(t);
      Logger.log("Removed overnight process trigger");
    }
  });
}

// Log daily themes with ticket IDs to a dedicated sheet tab
function logDailyThemes(dateStr, themes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Daily Themes Log");
  if (!sheet) {
    sheet = ss.insertSheet("Daily Themes Log");
    sheet.appendRow(["Date", "Theme #", "Theme Name", "Ticket Count", "Ticket IDs", "Summary", "Sentiment"]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  themes.forEach((t, i) => {
    const ticketIds = (t.ticket_ids || []).join(", ");
    // Summarize per-ticket sentiments for the log
    const sentiments = t.ticket_sentiments || {};
    const sentSummary = Object.entries(sentiments).map(([id, s]) => id + ":" + s).join(", ") || (t.sentiment || "");
    sheet.appendRow([
      dateStr,
      i + 1,
      t.theme || "",
      (t.ticket_ids || []).length,
      ticketIds,
      t.summary || "",
      sentSummary,
    ]);
  });
  Logger.log("Logged " + themes.length + " themes to Daily Themes Log for " + dateStr);
}

// ─── AI RECAP TREND ANALYSIS ───
// Reads from the Daily Themes Log sheet for a date range and sends to Claude
// to identify recurring patterns and trends. Used by weekly and monthly emails.
function analyzeThemeTrends(startDate, endDate, periodLabel) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty("ANTHROPIC_API_KEY");
  if (!apiKey) return null;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Daily Themes Log");
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return null;

  // Filter rows to the date range
  const tz = Session.getScriptTimeZone();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const rawDate = data[i][0]; // Date column -- may be Date object or string
    if (!rawDate) continue;
    const rowDate = (rawDate instanceof Date)
      ? Utilities.formatDate(rawDate, tz, "yyyy-MM-dd")
      : String(rawDate);
    if (rowDate >= startDate && rowDate <= endDate) {
      rows.push({
        date: rowDate,
        rank: data[i][1],
        theme: data[i][2],
        ticketCount: data[i][3],
        ticketIds: data[i][4],
        summary: data[i][5],
        sentiment: data[i][6] || "",
      });
    }
  }

  if (rows.length === 0) return null;

  // Count unique working days (exclude weekends AND holidays)
  const holidays = getDeakoHolidays();
  const uniqueDates = [...new Set(rows.map(r => r.date))];
  const workingDays = uniqueDates.filter(d => {
    const dayObj = new Date(d + "T12:00:00");
    const dow = dayObj.getDay();
    if (dow === 0 || dow === 6) return false; // weekend
    if (holidays.has(d)) return false; // holiday
    return true;
  });
  const totalWorkingDays = workingDays.length;

  // Filter rows to working days only for analysis
  const workingDaySet = new Set(workingDays);
  const workRows = rows.filter(r => workingDaySet.has(r.date));

  const themeSummaries = workRows.map(r =>
    `${r.date} | #${r.rank} | ${r.theme} | ${r.ticketCount} tickets | ${r.sentiment} | ${r.summary}`
  ).join("\n");

  const systemPrompt = `You are analyzing support theme trends for Deako, a smart lighting company. You are reviewing the daily AI Recap theme data for a ${periodLabel} period (${totalWorkingDays} working days) to identify recurring patterns, emerging issues, and notable shifts.

Use Deako's taxonomy vocabulary for consistency. Weekly/monthly trends should use PARENT-LEVEL categories since they aggregate across multiple days with varying specific symptoms. Keep trend names SHORT -- 3-5 words max. Good: "Connectivity > WiFi", "App Issues", "Hardware > Electrical". Bad: listing all sub-symptoms in the title. IMPORTANT: For app-related trends, always use "App Issues" in the trend name. Even if the source data says "Software/App", rename to "App Issues" in your output. Put the specific sub-symptom breakdown in the summary sentence, not the title. In the summary sentence, mention the most affected product type(s) when the data shows a clear pattern (e.g., "primarily affecting Smart Dimmers and Smart Switches").

For frequency, use the total of ${totalWorkingDays} working days as the denominator (e.g., "appeared 3 of ${totalWorkingDays} working days").

Integrate sentiment naturally into the summary sentence itself -- do NOT output sentiment as a separate field. Examples: "WiFi failures persist with customers reporting factual descriptions of connectivity issues." (neutral) or "Integration failures are generating increasing customer frustration, with several expressing exasperation at repeated failures." (frustrated). Always end the summary with a brief tone statement. For neutral: 'Customer tone is matter-of-fact' or 'Tone remains neutral throughout.' For frustrated: 'generating customer frustration' or 'with several customers expressing exasperation.' For mixed: 'Tone is mostly neutral with isolated frustration.' Every theme must include a tone indication.

You MUST respond with valid JSON only, no other text. Use this exact format:
[
  {"trend": "Trend Name", "frequency": "appeared X of ${totalWorkingDays} working days", "summary": "One sentence on the pattern with sentiment woven in naturally."},
  {"trend": "Trend Name", "frequency": "appeared X of ${totalWorkingDays} working days", "summary": "One sentence on the pattern with sentiment woven in naturally."},
  {"trend": "Trend Name", "frequency": "appeared X of ${totalWorkingDays} working days", "summary": "One sentence on the pattern with sentiment woven in naturally."}
]`;

  const userMessage = `PERIOD: ${periodLabel}
DAILY THEME DATA (${workRows.length} theme entries across ${totalWorkingDays} working days):
${themeSummaries}

Identify the top 3 recurring trends or patterns from this period's daily themes. Rank by frequency and impact. When mentioning ticket counts, say "tickets created" not just "tickets" so the reader knows these are new inbound contacts. Note if any theme is new/emerging vs persistent.`;

  try {
    const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method: "post",
      headers: {
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "anthropic-beta": "prompt-caching-2024-07-31",
        "Content-Type": "application/json",
      },
      payload: JSON.stringify({
        model: "claude-haiku-4-5-20251001",
        max_tokens: 800,
        system: [{ type: "text", text: systemPrompt, cache_control: { type: "ephemeral" } }],
        messages: [{ role: "user", content: userMessage }],
      }),
      muteHttpExceptions: true,
    });

    if (response.getResponseCode() === 200) {
      const respData = JSON.parse(response.getContentText());
      if (respData.content && respData.content[0] && respData.content[0].text) {
        const raw = respData.content[0].text.trim();
        let trends = [];
        try {
          trends = JSON.parse(raw);
        } catch (parseErr) {
          const jsonMatch = raw.match(/\[[\s\S]*\]/);
          if (jsonMatch) {
            trends = JSON.parse(jsonMatch[0]);
          } else {
            Logger.log("Could not parse trend analysis as JSON");
            return null;
          }
        }

        // Limit to top 3
        trends = trends.slice(0, 3);

        // Build HTML
        let trendsHtml = "";
        trends.forEach((t, i) => {
          const escapedTrend = (t.trend || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
          const escapedSummary = (t.summary || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
          const frequency = (t.frequency || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
          trendsHtml += `${i + 1}. <strong>${escapedTrend}</strong><br>`;
          trendsHtml += `${frequency} | ${escapedSummary}<br>`;
          if (i < trends.length - 1) trendsHtml += `<br>`;
        });

        return trendsHtml;
      }
    } else {
      Logger.log("Theme trend API error: " + response.getResponseCode());
    }
  } catch (e) {
    Logger.log("Theme trend analysis failed: " + e.toString());
  }
  return null;
}

// Parse DEAKO_HOLIDAYS Script Property into a Set of "YYYY-MM-DD" strings
// Includes multi-day ranges (e.g. "12/28/2026-12/31/2026")
function getDeakoHolidays() {
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty("DEAKO_HOLIDAYS") || "";
  const holidays = new Set();
  if (!raw.trim()) return holidays;

  raw.split(",").forEach(entry => {
    entry = entry.trim();
    if (!entry) return;

    if (entry.includes("-") && entry.split("-").length >= 2) {
      // Could be a date range like "12/28/2026-12/31/2026" or single date "2026-07-03"
      // Detect range: if both sides parse as dates with month/day/year format
      const parts = entry.split("-");
      // Try MM/DD/YYYY-MM/DD/YYYY range format
      if (parts.length >= 4) {
        // Likely "MM/DD/YYYY-MM/DD/YYYY"
        const startStr = parts.slice(0, 3).join("-");
        const endStr = parts.slice(3).join("-");
        const s = new Date(parts[0] + "/" + parts[1] + "/" + parts[2]);
        const e = new Date(parts[3] + "/" + parts[4] + "/" + parts[5]);
        if (!isNaN(s) && !isNaN(e)) {
          for (let d = new Date(s); d <= e; d.setDate(d.getDate() + 1)) {
            holidays.add(Utilities.formatDate(new Date(d), "America/Los_Angeles", "yyyy-MM-dd"));
          }
          return;
        }
      }
      // Try YYYY-MM-DD single date
      const singleDate = new Date(entry);
      if (!isNaN(singleDate)) {
        holidays.add(Utilities.formatDate(singleDate, "America/Los_Angeles", "yyyy-MM-dd"));
        return;
      }
    }

    // Single date: MM/DD/YYYY or YYYY-MM-DD
    const d = new Date(entry);
    if (!isNaN(d)) {
      holidays.add(Utilities.formatDate(d, "America/Los_Angeles", "yyyy-MM-dd"));
    }
  });
  return holidays;
}

// Check if a given date is a non-working day (weekend or Deako holiday)
function isNonWorkingDay(date, holidays) {
  const tz = CONFIG.businessHours.timezone;
  const pacStr = date.toLocaleString("en-US", { timeZone: tz });
  const pac = new Date(pacStr);
  const dow = pac.getDay();  // 0=Sun, 6=Sat

  // Weekend check
  if (!CONFIG.businessHours.workDays.includes(dow)) return true;

  // Holiday check
  if (!holidays) holidays = getDeakoHolidays();
  const dateKey = Utilities.formatDate(date, tz, "yyyy-MM-dd");
  return holidays.has(dateKey);
}

// Calculate business minutes between two dates (6a-5p Mon-Fri PST, excluding holidays)
function calcBusinessMinutes(start, end) {
  const bh = CONFIG.businessHours;
  const minPerDay = (bh.endHour - bh.startHour) * 60; // 660 min for 6a-5p
  const holidays = getDeakoHolidays();

  let totalMin = 0;
  let cursor = new Date(start);

  // Cap at 60 days to prevent infinite loops on ancient tickets
  const maxIterations = 60;
  let iterations = 0;

  while (cursor < end && iterations < maxIterations) {
    // Convert to Pacific time
    const pacStr = cursor.toLocaleString("en-US", { timeZone: bh.timezone });
    const pac = new Date(pacStr);
    const dow = pac.getDay();
    const hour = pac.getHours();
    const min = pac.getMinutes();

    // Skip holidays (treat like weekends)
    const dateKey = Utilities.formatDate(cursor, bh.timezone, "yyyy-MM-dd");
    if (bh.workDays.includes(dow) && !holidays.has(dateKey)) {
      // It's a workday
      const startMin = bh.startHour * 60;
      const endMin = bh.endHour * 60;
      const curMin = hour * 60 + min;

      if (curMin < startMin) {
        // Before business hours — skip to start of business
        cursor = new Date(cursor.getTime() + (startMin - curMin) * 60000);
        continue;
      } else if (curMin >= endMin) {
        // After business hours — skip to next day start
        cursor = new Date(cursor.getTime() + (24 * 60 - curMin + startMin) * 60000);
        iterations++;
        continue;
      } else {
        // During business hours
        const remainToday = Math.min(endMin - curMin, (end - cursor) / 60000);
        totalMin += Math.max(0, remainToday);
        cursor = new Date(cursor.getTime() + remainToday * 60000);
        if (cursor >= end) break;
        // Jump to next day's business start
        const pacEnd = new Date(pac);
        pacEnd.setHours(bh.endHour, 0, 0, 0);
        cursor = new Date(cursor.getTime() + (24 * 60 - endMin + startMin) * 60000);
        iterations++;
        continue;
      }
    } else {
      // Weekend — skip to next day
      cursor = new Date(cursor.getTime() + 24 * 60 * 60000);
      iterations++;
      continue;
    }
  }

  return Math.round(totalMin);
}

// --- AIRCALL API ---
function fetchAircallStatus() {
  const props = PropertiesService.getScriptProperties();
  const apiId = props.getProperty("AIRCALL_API_ID");
  const apiToken = props.getProperty("AIRCALL_API_TOKEN");
  if (!apiId || !apiToken) throw new Error("Aircall API credentials not set in Script Properties");

  const baseUrl = CONFIG.aircall.baseUrl;
  const auth = "Basic " + Utilities.base64Encode(apiId + ":" + apiToken);

  // Today's date range (local timezone)
  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const fromTs = Math.floor(startOfDay.getTime() / 1000);
  const toTs = Math.floor(now.getTime() / 1000);

  // Fetch today's calls with pagination
  let allCalls = [];
  let page = 1;
  let hasMore = true;

  while (hasMore && page <= 10) {
    const url = `${baseUrl}/calls?from=${fromTs}&to=${toTs}&per_page=50&page=${page}&order=desc`;
    const options = {
      method: "get",
      headers: { "Authorization": auth, "Content-Type": "application/json" },
      muteHttpExceptions: true,
    };

    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    if (code !== 200) {
      throw new Error(`Aircall API returned ${code}: ${response.getContentText().substring(0, 200)}`);
    }

    const data = JSON.parse(response.getContentText());
    allCalls = allCalls.concat(data.calls || []);

    hasMore = data.meta && data.meta.next_page_link;
    page++;
  }

  // Filter to CS support lines only, then inbound only
  const supportNumbers = CONFIG.aircall.supportNumbers;
  const supportCalls = allCalls.filter(c => {
    if (!c.number) return false;
    const digits = (c.number.digits || "").replace(/[\s\-\(\)]/g, "");
    return supportNumbers.some(sn => digits.includes(sn.replace(/[\s\-\(\)]/g, "")) || sn.replace(/[\s\-\(\)]/g, "").includes(digits));
  });
  const inboundCalls = supportCalls.filter(c => c.direction === "inbound");

  // Filter to business hours only (6am-5pm Mon-Fri Pacific, excluding Deako holidays)
  const bh = CONFIG.businessHours;
  const holidays = getDeakoHolidays();
  const bizHourCalls = inboundCalls.filter(c => {
    if (!c.started_at) return false;
    // Convert Unix timestamp to Pacific time
    const callDate = new Date(c.started_at * 1000);
    const pacificStr = callDate.toLocaleString("en-US", { timeZone: bh.timezone });
    const pacificDate = new Date(pacificStr);
    const dayOfWeek = pacificDate.getDay(); // 0=Sun, 1=Mon ... 6=Sat
    const hour = pacificDate.getHours();
    const dateKey = Utilities.formatDate(callDate, bh.timezone, "yyyy-MM-dd");
    return bh.workDays.includes(dayOfWeek) && !holidays.has(dateKey) && hour >= bh.startHour && hour < bh.endHour;
  });

  // Also track after-hours calls separately for context
  const afterHoursCalls = inboundCalls.length - bizHourCalls.length;

  // ─── CALL CLASSIFICATION ───
  // Based on actual Aircall API behavior (confirmed from debug data):
  //   - ALL calls have status "done" — status field is useless for classification
  //   - missed_call_reason is the key field:
  //       empty + answered_at + user  → team answered
  //       "agents_did_not_answer"     → team missed, forwarded to answer service
  //       "no_available_agent"        → no agents online, forwarded to answer service
  //       "short_abandoned"           → caller hung up in <6s, not team's fault
  //       "out_of_opening_hours"      → already filtered by biz hours
  //   - Outbound calls TO the answer service number are the forwarding mechanism
  //     and must be excluded from all counts

  // Helper: match a call's user to a CONFIG agent
  function matchAgent(call) {
    const user = call.user;
    if (!user) return null;
    const agentName = (user.name || `${user.first_name || ""} ${user.last_name || ""}`.trim());
    return CONFIG.agents.find(a => {
      const parts = a.toLowerCase().split(" ");
      const callParts = agentName.toLowerCase().split(" ");
      return parts[0] === callParts[0] || (parts[1] && callParts[1] && parts[1] === callParts[1]);
    });
  }

  // ─── FORWARDED COUNT: count outbound calls TO the SAS number ───
  // Aircall doesn't reliably set missed_call_reason on the inbound leg.
  // Instead, it creates a separate outbound call to the SAS number when forwarding.
  // So we count those outbound-to-SAS calls as the true forwarded count.
  const answerSvcNum = (CONFIG.aircall.answeringServiceNumber || "").replace(/[\s\-\(\)]/g, "");
  const forwardedToSAS = supportCalls.filter(call => {
    if (call.direction !== "outbound") return false;
    const rawDigits = (call.raw_digits || "").replace(/[\s\-\(\)]/g, "");
    if (!answerSvcNum || !rawDigits.includes(answerSvcNum.replace("+1", ""))) return false;
    // Apply business hours filter to forwarded calls too
    if (!call.started_at) return false;
    const callDate = new Date(call.started_at * 1000);
    const pacificStr = callDate.toLocaleString("en-US", { timeZone: bh.timezone });
    const pacificDate = new Date(pacificStr);
    const dayOfWeek = pacificDate.getDay();
    const hour = pacificDate.getHours();
    return bh.workDays.includes(dayOfWeek) && hour >= bh.startHour && hour < bh.endHour;
  });

  // Inbound calls only (for team answered + short abandoned classification)
  const customerCalls = bizHourCalls.filter(call => call.direction === "inbound");

  // Categorize inbound calls
  const teamAnswered = [];      // team agent picked up
  const shortAbandoned = [];    // caller hung up too fast to count

  customerCalls.forEach(call => {
    const reason = call.missed_call_reason || "";

    if (reason === "short_abandoned") {
      shortAbandoned.push(call);
    } else if (call.answered_at && call.user) {
      const matched = matchAgent(call);
      if (matched) {
        teamAnswered.push({ call, agent: matched });
      }
      // Calls answered by non-CONFIG users (e.g. SAS) are no longer double-counted here;
      // they are already captured by forwardedToSAS above.
    }
    // Calls with missed_call_reason (agents_did_not_answer, etc.) are also already
    // captured by the outbound-to-SAS count, so we don't double-count them.
  });

  // ─── Determine reason for each SAS-forwarded call ───
  // CONFIRMED BEHAVIOR (Analytics+ data, 2026-05-06):
  //   When NO agents are available on support lines with SAS forwarding,
  //   Aircall does NOT create an inbound call record at all — the entire
  //   customer interaction is logged as a single outbound call to SAS.
  //   The `no_available_agent` missed_call_reason only appears on lines
  //   WITHOUT SAS forwarding (sales lines, personal lines).
  //
  //   Therefore: outbound-to-SAS with no inbound match = "No agents available"
  //   (verified 12/12 on 5/6 — every SAS forward had zero agents available).
  //
  //   On the rare occasion an inbound record exists for a SAS-forwarded call
  //   (e.g. agent was briefly available then went unavailable mid-queue),
  //   we use its missed_call_reason for a more specific reason.
  const allInboundToday = allCalls.filter(c => c.direction === "inbound");

  const forwarded = forwardedToSAS.map(sasCall => {
    const sasStart = sasCall.started_at || 0;

    // Look for a matching inbound call on a support line within 10 min before the SAS outbound
    // Match by raw_digits (customer number) if available, otherwise by time proximity
    let bestMatch = null;
    let bestTimeDiff = Infinity;
    allInboundToday.forEach(inb => {
      // Only consider calls that weren't answered by team agents
      if (inb.answered_at && inb.user) {
        const matched = matchAgent(inb);
        if (matched) return; // team-answered, not related to this SAS forward
      }
      const diff = sasStart - (inb.started_at || 0);
      if (diff >= 0 && diff <= 600 && diff < bestTimeDiff) {
        // Only match if on a support line
        if (!inb.number) return;
        const digits = (inb.number.digits || "").replace(/[\s\-\(\)]/g, "");
        const isSupport = supportNumbers.some(sn =>
          digits.includes(sn.replace(/[\s\-\(\)]/g, "")) ||
          sn.replace(/[\s\-\(\)]/g, "").includes(digits));
        if (!isSupport) return;
        bestTimeDiff = diff;
        bestMatch = inb;
      }
    });

    // Determine the reason — default to "No agents available" (confirmed proxy)
    let reason = "No agents available";
    if (bestMatch) {
      const r = bestMatch.missed_call_reason || "";
      if (r === "agents_did_not_answer" || r === "agents_did_not_pick_up") {
        reason = "Agents didn't answer";
      } else if (r === "no_available_agent" || r === "no_agent_available") {
        reason = "No agents available";
      } else if (r === "out_of_opening_hours") {
        reason = "Outside business hours";
      } else if (r) {
        reason = r;
      }
      // If inbound exists but has no missed reason, still default to no agents available
    }
    return {
      call: sasCall,
      inboundCall: bestMatch,
      reason,
    };
  });

  // Team Answer Rate: short_abandoned excluded from denominator
  const rateEligible = teamAnswered.length + forwarded.length;
  const teamAnswerRate = rateEligible > 0
    ? (teamAnswered.length / rateEligible * 100) : 100;

  // Avg wait time & duration (team-answered calls only)
  const waitTimes = teamAnswered.map(t => t.call.waiting_duration || 0).filter(w => w > 0);
  const avgWaitTime = waitTimes.length > 0 ? waitTimes.reduce((a, b) => a + b, 0) / waitTimes.length : 0;

  const durations = teamAnswered.map(t => t.call.duration || 0).filter(d => d > 0);
  const avgDuration = durations.length > 0 ? durations.reduce((a, b) => a + b, 0) / durations.length : 0;

  // ─── Per-agent breakdown ───
  const agentStats = {};
  CONFIG.agents.forEach(name => agentStats[name] = { answered: 0, outbound: 0, outboundConnected: 0, outboundShort: 0, outboundLong: 0, inboundTalkTime: 0, outboundTalkTime: 0 });

  teamAnswered.forEach(t => {
    agentStats[t.agent].answered++;
    agentStats[t.agent].inboundTalkTime += (t.call.duration || 0);
  });

  // ─── Outbound calls per agent ───
  // Search ALL calls (not just support lines) so agents using direct lines are counted.
  // Exclude calls TO the answering service number (those are forwarding, not real outbound).
  // Only count calls by CONFIG agents to avoid inflating totals with sales/other teams.
  const outboundCalls = allCalls.filter(c => {
    if (c.direction !== "outbound") return false;
    // Exclude outbound to the answering service number
    const rawDigits = (c.raw_digits || "").replace(/[\s\-\(\)]/g, "");
    if (answerSvcNum && rawDigits.includes(answerSvcNum.replace("+1", ""))) return false;
    // Only count calls by recognized CS agents
    const matched = matchAgent(c);
    if (!matched) return false;
    return true;
  });

  const OUTBOUND_SHORT_THRESHOLD = 90; // seconds — under this likely voicemail, over likely conversation

  outboundCalls.forEach(call => {
    const matched = matchAgent(call);
    if (matched && agentStats[matched]) {
      agentStats[matched].outbound++;
      // Connected = has duration > 0 and answered_at
      if (call.duration > 0 && call.answered_at) {
        agentStats[matched].outboundConnected++;
        agentStats[matched].outboundTalkTime += (call.duration || 0);
        if (call.duration < OUTBOUND_SHORT_THRESHOLD) {
          agentStats[matched].outboundShort++;   // likely voicemail
        } else {
          agentStats[matched].outboundLong++;    // likely conversation
        }
      }
    }
  });

  const totalOutbound = outboundCalls.length;
  const totalOutboundConnected = outboundCalls.filter(c => c.duration > 0 && c.answered_at).length;

  // ─── Forwarded call breakdown by reason ───
  const missedSummary = {};
  forwarded.forEach(f => {
    missedSummary[f.reason] = (missedSummary[f.reason] || 0) + 1;
  });

  // Build missed call detail rows
  // Aircall API doesn't expose customer number on outbound-to-SAS calls
  // (confirmed: contact=null, raw_digits=SAS number, no from field in API).
  // The CSV export has the customer number but it's not available via API.
  // We show the line name from the SAS call's number object instead.
  const missedCallDetails = forwarded.map(f => {
    const inb = f.inboundCall;
    const hasInbound = inb && inb.raw_digits;

    // Customer display
    const callerNumber = hasInbound ? inb.raw_digits : "";
    const contactName = hasInbound && inb.contact
      ? `${inb.contact.first_name || ""} ${inb.contact.last_name || ""}`.trim()
      : "";

    // Line name from the SAS outbound call's number object
    const lineName = f.call.number ? (f.call.number.name || "") : "";

    // Time: prefer inbound call time, fall back to SAS outbound time
    const ts = (hasInbound && inb.started_at) ? inb.started_at : f.call.started_at;
    const callTime = ts
      ? Utilities.formatDate(new Date(ts * 1000), Session.getScriptTimeZone(), "h:mm a")
      : "-";

    // Aircall call ID for easy lookup when customer info unavailable
    const callId = f.call.id || "";

    return { callerNumber, contactName, callTime, reason: f.reason, lineName, callId };
  });

  return {
    totalInbound: teamAnswered.length + forwarded.length + shortAbandoned.length,
    teamAnswered: teamAnswered.length,
    forwarded: forwarded.length,
    shortAbandoned: shortAbandoned.length,
    teamAnswerRate,
    totalOutbound,
    totalOutboundConnected,
    avgWaitTime, avgDuration,
    agentStats, calls: allCalls,
    missedSummary,
    missedCallDetails,
    afterHoursCalls: afterHoursCalls,
  };
}

// --- NICEREPLY API (CSAT survey responses) ---
// Docs: https://cdn.nicereply.com/s/api/latest/reference/responses/list
// Auth: Basic (email:token)
// Endpoint: GET https://api.nicereply.com/responses  (NO /v1/ prefix!)
// Date filter: created_after (ISO 8601)
// Response shape: { data: [{ id, answers: [{ question_type, scale: {value}, open_ended: {value} }], ... }], pagination: {...} }
function fetchNicereplyCSAT() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("NICEREPLY_TOKEN");
  if (!token) {
    Logger.log("NICEREPLY_TOKEN not set — skipping CSAT fetch");
    return { score: null, total: 0, satisfied: 0, responses: [] };
  }

  // Auth: Basic HTTP with email:token
  // NICEREPLY_TOKEN must be stored as "email:api_token"
  const auth = "Basic " + Utilities.base64Encode(token);
  const headers = { "Authorization": auth };
  const fetchOpts = { method: "get", headers: headers, muteHttpExceptions: true };

  // Fetch responses from the last 24 hours
  const now = new Date();
  const since = new Date(now.getTime() - 24 * 60 * 60 * 1000);
  // Nicereply requires ISO 8601 WITHOUT milliseconds: 2026-05-04T10:27:00Z
  const sinceISO = since.toISOString().replace(/\.\d{3}Z$/, "Z");

  let allResponses = [];
  let page = 1;
  let hasMore = true;

  while (hasMore && page <= 5) {
    const url = `https://api.nicereply.com/responses?created_after=${encodeURIComponent(sinceISO)}&per_page=50&page=${page}`;
    const resp = UrlFetchApp.fetch(url, fetchOpts);
    const code = resp.getResponseCode();

    if (code === 401) {
      Logger.log("Nicereply auth failed (401) — check NICEREPLY_TOKEN format (must be email:api_token)");
      return { score: null, total: 0, satisfied: 0, responses: [], error: "Auth failed" };
    }
    if (code !== 200) {
      Logger.log("Nicereply API returned " + code + ": " + resp.getContentText().substring(0, 300));
      break;
    }

    const body = JSON.parse(resp.getContentText());
    const responses = body.data || [];
    if (responses.length === 0) break;

    allResponses = allResponses.concat(responses);

    // Pagination: check if there's a next page
    const pagination = body.pagination || {};
    hasMore = pagination.total_pages ? page < pagination.total_pages : responses.length >= 50;
    page++;
  }

  return processNicereplyResponses(allResponses, since);
}

function processNicereplyResponses(responses, since) {
  // Filter to last 24h (belt-and-suspenders) and extract useful fields
  const recent = responses.filter(r => {
    const created = new Date(r.created_at || "");
    return created >= since;
  }).map(r => {
    const created = new Date(r.created_at || "");
    const timeStr = Utilities.formatDate(created, Session.getScriptTimeZone(), "h:mm a");
    const dateStr = Utilities.formatDate(created, Session.getScriptTimeZone(), "MMM d");

    // Nicereply surveys contain multiple SCALE answers (CSAT, CES, NPS).
    // Each has a stable question_id. We care about the CSAT question:
    //   86bae330-... = "Overall support experience" (1-5 scale)
    //   94322eb4-... = "Rate the agent" (1-5 scale)
    //   3ef73361-... = CES "Easy to handle" (1-7 scale)
    //   6c0dc99b-... = NPS "Recommend Deako" (0-10 scale)
    const CSAT_QUESTION_ID = "86bae330-e8bc-4fa3-9af9-91eb2459d348";
    const answers = r.answers || [];

    // Target the CSAT question by ID; fall back to first SCALE answer
    const csatAnswer = answers.find(a => a.question_id === CSAT_QUESTION_ID)
      || answers.find(a => a.question_type === "SCALE");
    const score = csatAnswer && csatAnswer.scale ? csatAnswer.scale.value : 0;
    const maxScore = 5; // Deako's CSAT survey uses a 1-5 scale

    // Extract open-ended comment if present
    const openAnswer = answers.find(a => a.question_type === "OPEN_ENDED");
    const comment = openAnswer
      ? (openAnswer.open_ended ? openAnswer.open_ended.value : (openAnswer.scale ? openAnswer.scale.value : ""))
      : "";

    // Ticket reference and customer
    const ticketId = r.ticket_id || "";
    const customerId = r.customer_id || "";
    // "from" is often null in the Nicereply API — will resolve via Zendesk ticket below
    const email = r.from || "";

    return {
      score,
      maxScore,
      email,
      ticketId,
      comment,
      timeStr,
      dateStr,
      created,
      // Satisfied: 4+ on 5-point CSAT scale
      satisfied: score >= 4,
    };
  });

  // Sort newest first
  recent.sort((a, b) => b.created - a.created);

  // Resolve unknown customer names via Zendesk ticket requester
  const needsLookup = recent.filter(r => !r.email && r.ticketId);
  if (needsLookup.length > 0) {
    try {
      const props = PropertiesService.getScriptProperties();
      const zdToken = props.getProperty("ZENDESK_TOKEN");
      if (zdToken) {
        const subdomain = CONFIG.zendesk.subdomain;
        const zdAuth = "Basic " + Utilities.base64Encode(zdToken);
        const zdHeaders = { "Authorization": zdAuth, "Content-Type": "application/json" };
        const zdOpts = { method: "get", headers: zdHeaders, muteHttpExceptions: true };

        // Batch fetch tickets via show_many
        const ticketIds = needsLookup.map(r => r.ticketId).join(",");
        const url = `https://${subdomain}.zendesk.com/api/v2/tickets/show_many.json?ids=${ticketIds}`;
        const resp = UrlFetchApp.fetch(url, zdOpts);
        if (resp.getResponseCode() === 200) {
          const data = JSON.parse(resp.getContentText());
          // Build requester_id → ticket_id map
          const requesterIds = new Set();
          const ticketRequesterMap = {};
          (data.tickets || []).forEach(t => {
            ticketRequesterMap[t.id] = t.requester_id;
            if (t.requester_id) requesterIds.add(t.requester_id);
          });

          // Batch fetch users
          if (requesterIds.size > 0) {
            const userUrl = `https://${subdomain}.zendesk.com/api/v2/users/show_many.json?ids=${[...requesterIds].join(",")}`;
            const userResp = UrlFetchApp.fetch(userUrl, zdOpts);
            if (userResp.getResponseCode() === 200) {
              const userData = JSON.parse(userResp.getContentText());
              const userNames = {};
              (userData.users || []).forEach(u => { userNames[u.id] = u.name || u.email || ""; });

              // Apply to responses
              needsLookup.forEach(r => {
                const reqId = ticketRequesterMap[r.ticketId];
                if (reqId && userNames[reqId]) {
                  r.email = userNames[reqId];
                }
              });
            }
          }
        }
      }
    } catch (e) {
      Logger.log("CSAT customer lookup failed: " + e.toString());
    }
    // Fill remaining unknowns
    recent.forEach(r => { if (!r.email) r.email = "Unknown"; });
  }

  const total = recent.length;
  const satisfied = recent.filter(r => r.satisfied).length;
  const score = total > 0 ? Math.round((satisfied / total) * 100) : null;

  return {
    score,      // CSAT percentage (0-100) or null if no data
    total,
    satisfied,
    responses: recent,
  };
}

// =============================================================
// DASHBOARD LAYOUT — Calm, minimal, on-brand instrument panel
// =============================================================
function writeDashboard(ss, zendesk, aircall, csat, postCall, sms, meta, qi) {
  // Build on a hidden staging sheet, then swap — eliminates the 5-min blink
  const staging = getOrCreateSheet(ss, "_Staging");
  staging.showSheet(); // ensure it exists and is accessible
  _writeDashboardContent(ss, staging, zendesk, aircall, csat, postCall, sms, meta, qi);

  // Swap staging content → Dashboard in one batch
  const dash = getOrCreateSheet(ss, "Dashboard");
  const lastRow = staging.getLastRow() || 1;
  const lastCol = Math.max(staging.getLastColumn(), 26); // 26 = col Z (social handoff)

  // Preserve column Z (social handoff/spam flags) before clearing — operator-entered data
  const handoffMap = {};  // customerName → handoff value
  const spamNames = [];   // names marked "Spam" to auto-add to Hidden IG Senders
  const dashLastRow = dash.getLastRow();
  if (dashLastRow > 1) {
    const zCol = dash.getRange(1, 26, dashLastRow, 1).getValues();   // col Z
    const rCol = dash.getRange(1, 18, dashLastRow, 1).getValues();   // col R (customer names)
    for (let i = 0; i < dashLastRow; i++) {
      const val = String(zCol[i][0] || "").trim();
      const name = String(rCol[i][0] || "").trim();
      if (val && val !== "Owner" && name) {
        if (val === "Spam") {
          spamNames.push(name);
        } else {
          handoffMap[name] = val;
        }
      }
    }
  }

  // Process spam flags — add to Hidden IG Senders sheet so webhooks filter them too.
  // Once added, they won't appear on the next refresh (filtered by getHiddenIGSenders).
  if (spamNames.length > 0) {
    const hiddenSheet = getOrCreateSheet(ss, "Hidden IG Senders");
    if (hiddenSheet.getLastRow() === 0) {
      hiddenSheet.appendRow(["Sender ID", "Username", "Hidden At", "Hidden By"]);
      hiddenSheet.getRange("1:1").setFontWeight("bold");
    }
    // Look up sender IDs from IG DM Log by matching customer name
    const igSheet = ss.getSheetByName("IG DM Log");
    const igNameToId = {};
    if (igSheet && igSheet.getLastRow() > 1) {
      const igData = igSheet.getRange(2, 2, igSheet.getLastRow() - 1, 2).getValues(); // cols B (ID), C (Name)
      igData.forEach(row => {
        const id = String(row[0] || "");
        const name = String(row[1] || "");
        if (id && name) igNameToId[name] = id;
      });
    }
    // Read existing hidden names to avoid duplicates
    const existingHidden = getHiddenIGSenders(ss);
    spamNames.forEach(name => {
      const senderId = igNameToId[name] || "";
      if (!existingHidden.has(name) && !existingHidden.has(senderId)) {
        hiddenSheet.appendRow([senderId, name, new Date(), "Dashboard (Spam)"]);
      }
    });
  }

  // Clear dashboard and paste all content + formatting from staging
  dash.clear();
  dash.clearFormats();
  if (lastRow > 0 && lastCol > 0) {
    const source = staging.getRange(1, 1, lastRow, lastCol);
    source.copyTo(dash.getRange(1, 1, lastRow, lastCol));
  }

  // Match column widths and row heights
  for (let c = 1; c <= lastCol; c++) {
    dash.setColumnWidth(c, staging.getColumnWidth(c));
  }
  for (let r = 1; r <= lastRow; r++) {
    dash.setRowHeight(r, staging.getRowHeight(r));
  }

  // Restore column Z handoff flags — match by customer name in column R
  // Also add data validation dropdown to social conversation rows
  const socialHandoffRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["CS Team", "Other Team", "Resolved", "Spam"], true)
    .setAllowInvalid(false).build();
  const newLastRow = dash.getLastRow();
  if (newLastRow > 1) {
    for (let r = 1; r <= newLastRow; r++) {
      const cellR = dash.getRange(r, 18).getValue(); // col R
      const cellZ = dash.getRange(r, 26);
      const zVal = String(cellZ.getValue() || "").trim();
      // Look for rows that have the social row marker (written by _writeDashboardContent)
      if (zVal === "_social_dm_row" || zVal === "_social_comment_row") {
        const custName = String(cellR || "").trim();
        const existing = handoffMap[custName] || "";
        const dimStatuses = ["Other Team", "Resolved"];
        cellZ.setValue(existing || "CS Team")
          .setDataValidation(socialHandoffRule)
          .setFontSize(8).setFontColor("#1D1D1D")
          .setBackground(
            existing === "Other Team" ? "#E8E8E8"
            : existing === "Resolved" ? "#D9EAD3"  // light green -- conversation complete
            : dash.getRange(r, 18).getBackground()
          );
        // Dim the entire row if handed off or resolved (no longer needs attention)
        if (dimStatuses.includes(existing)) {
          dash.getRange(r, 18, 1, 8).setFontColor("#AAAAAA");
          // Also dim the excerpt row below
          if (r < newLastRow) {
            dash.getRange(r + 1, 18, 1, 8).setFontColor("#AAAAAA");
          }
        }
      } else if (zVal === "_social_header") {
        cellZ.setValue("Owner").setFontWeight("bold").setFontSize(9)
          .setFontColor("#888888").setBackground(dash.getRange(r, 18).getBackground());
      } else {
        cellZ.setValue(""); // clear marker from non-social rows
      }
    }
  }

  dash.setColumnWidth(26, 75); // Z — Owner column

  // Preserve dashboard appearance
  dash.setTabColor(BRAND.airBlueDark);
  dash.setHiddenGridlines(true);

  // Hide staging
  staging.hideSheet();
}

function _writeDashboardContent(ss, dash, zendesk, aircall, csat, postCall, sms, meta, qi) {
  dash.clear();
  dash.clearFormats();

  const now = new Date();
  const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), "EEE MMM d, h:mm a");
  const slaMinutes = CONFIG.thresholds.oldestUnanswered.green * 60;
  const slaHours = CONFIG.thresholds.oldestUnanswered.green;

  // --- Brand-aligned palette ---
  const bg       = BRAND.white;          // #FAFAFA
  const cardBg   = "#FFFFFF";
  const altRow   = "#F5F5F3";            // subtle zebra stripe for table rows
  const divider  = BRAND.beigeLight;     // #E1DFDD
  const navy     = BRAND.airBlueDark;    // #1B3747
  const darkText = BRAND.black;          // #1D1D1D
  const gray     = BRAND.ashGray;        // #9AA19B
  const green    = BRAND.mossGreen;      // #889578
  const amber    = BRAND.terracotta;     // #BA866A
  const amberLt  = BRAND.terracottaLight;// #DEAC90
  const risk     = "#A85353";            // softened operational red (brand-adjacent)
  const riskLt   = BRAND.roseQuartzLight;// #D6BDC8 — subtle risk tint

  dash.getRange("A:Z").setFontFamily("Inter").setFontColor(darkText).setBackground(bg);
  dash.setTabColor(navy);
  dash.setHiddenGridlines(true);

  // ─── Status logic ───
  const emailStatus = zendesk.totalBreached > 10 ? "At Risk"
    : zendesk.totalBreached > 5 ? "Watch" : "Healthy";
  const emailColor = emailStatus === "At Risk" ? risk
    : emailStatus === "Watch" ? amber : green;

  const phoneGreen = CONFIG.thresholds.phoneAnswerRate.green;
  const phoneYellow = CONFIG.thresholds.phoneAnswerRate.yellow;
  const phoneStatus = aircall.teamAnswerRate >= phoneGreen ? "Healthy"
    : aircall.teamAnswerRate >= phoneYellow ? "Watch" : "At Risk";
  const phoneColor = phoneStatus === "At Risk" ? risk
    : phoneStatus === "Watch" ? amber : green;

  // Social status — based on oldest unread DM response time
  // Read existing handoff flags from the live Dashboard (column Z) so conversations
  // assigned to another team or marked resolved are excluded from the SLA calculation.
  const excludeFromSLA = new Set(["Other Team", "Resolved"]);
  const handedOffNames = new Set();
  const liveDash = ss.getSheetByName("Dashboard");
  if (liveDash && liveDash.getLastRow() > 1) {
    const zVals = liveDash.getRange(1, 26, liveDash.getLastRow(), 1).getValues();
    const rVals = liveDash.getRange(1, 18, liveDash.getLastRow(), 1).getValues();
    for (let i = 0; i < zVals.length; i++) {
      const status = String(zVals[i][0] || "").trim();
      if (excludeFromSLA.has(status) && String(rVals[i][0] || "").trim()) {
        handedOffNames.add(String(rVals[i][0]).trim());
      }
    }
  }
  const socialOldestWaitMin = computeSocialOldestWait(meta, handedOffNames);
  const socialGreen = CONFIG.thresholds.socialResponseTime.green;
  const socialYellow = CONFIG.thresholds.socialResponseTime.yellow;
  const socialStatus = socialOldestWaitMin <= socialGreen ? "Healthy"
    : socialOldestWaitMin <= socialYellow ? "Watch" : "At Risk";
  const socialColor = socialStatus === "At Risk" ? risk
    : socialStatus === "Watch" ? amber : green;

  // Derived values
  const oldestWaitMin = zendesk.longest10.length > 0 ? zendesk.longest10[0].waitBizMin : 0;
  const oldestWaitStr = formatBizMinutes(oldestWaitMin);
  const rateEligible = aircall.teamAnswered + aircall.forwarded;
  const fwdCount = aircall.forwarded;

  // ═══════════════════════════════════════════════
  // ROW 1: Compact title bar with Deako wordmark
  // ═══════════════════════════════════════════════
  // Logo cell — white "deako" wordmark on Air Blue Dark (per brand guidelines)
  dash.getRange("A1:B1").merge()
    .setValue("deako®")
    .setBackground(navy).setFontColor("#FFFFFF")
    .setFontSize(13).setFontWeight("bold")
    .setVerticalAlignment("middle");
  // Title
  dash.getRange("C1:P1").merge()
    .setValue("CS Command Center")
    .setBackground(navy).setFontColor(BRAND.airBlueLight)
    .setFontSize(13).setVerticalAlignment("middle");
  // Timestamp right-aligned
  dash.getRange("R1:Z1").merge()
    .setValue(`Status as of ${timestamp}`)
    .setBackground(navy).setFontColor(BRAND.airBlueMedium)
    .setFontSize(9)
    .setHorizontalAlignment("right").setVerticalAlignment("middle");
  dash.setRowHeight(1, 36);

  // ═══════════════════════════════════════════════
  // ROW 2: Channel status strip (3 channels)
  // ═══════════════════════════════════════════════
  dash.getRange("A2:G2").merge()
    .setValue(`Email: ${emailStatus}`)
    .setBackground(bg).setFontColor(emailColor)
    .setFontSize(14).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  dash.getRange("H2").setBackground(bg);
  dash.getRange("I2:O2").merge()
    .setValue(`Phone: ${phoneStatus}`)
    .setBackground(bg).setFontColor(phoneColor)
    .setFontSize(14).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  dash.getRange("P2").setBackground(bg);
  dash.getRange("R2:Y2").merge()
    .setValue(`Social: ${socialStatus}`)
    .setBackground(bg).setFontColor(socialColor)
    .setFontSize(14).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  dash.setRowHeight(2, 34);

  // ═══════════════════════════════════════════════
  // ROW 3: ALERT BANNER — all-hands-on-deck when unassigned > 100
  // ═══════════════════════════════════════════════
  const UNASSIGNED_ALERT_THRESHOLD = 100;
  const showAlert = zendesk.unassigned > UNASSIGNED_ALERT_THRESHOLD;
  if (showAlert) {
    const alertRed = "#B91C1C";      // deep red background
    const alertRedLight = "#FEE2E2"; // light red for accent
    dash.getRange("A3:Y3").merge()
      .setValue(`⚠  ALL HANDS ON DECK  —  ${zendesk.unassigned} UNASSIGNED TICKETS  ⚠`)
      .setBackground(alertRed).setFontColor("#FFFFFF")
      .setFontSize(18).setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    dash.setRowHeight(3, 56);
    dash.getRange("A4:Y4").merge()
      .setValue("Unassigned ticket count has exceeded " + UNASSIGNED_ALERT_THRESHOLD + ". All available agents should begin triaging unassigned tickets immediately.")
      .setBackground(alertRedLight).setFontColor(alertRed)
      .setFontSize(11).setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    dash.setRowHeight(4, 32);
    // Row 5: divider after alert
    dash.getRange("A5:Y5").setBackground(divider);
    dash.setRowHeight(5, 2);
  } else {
    // Row 3: normal thin divider
    dash.getRange("A3:Y3").setBackground(divider);
    dash.setRowHeight(3, 2);
  }

  // Dynamic row offset — everything below shifts down by 2 when alert is showing
  const alertOffset = showAlert ? 2 : 0;

  // ═══════════════════════════════════════════════
  // KPI panels (numbers on top, labels below)
  // ═══════════════════════════════════════════════

  // Column H & Q: spacers between columns
  dash.setColumnWidth(8, 20);   // H spacer
  dash.setColumnWidth(17, 20);  // Q spacer

  // --- ROW 4: Big numbers ---
  const k1 = 4 + alertOffset;
  dash.setRowHeight(k1, 52);

  // Email: Waiting 12h+ | No Reply 12h+ | Oldest Wait
  const emailNumColor = emailStatus === "At Risk" ? risk : (emailStatus === "Watch" ? amber : darkText);
  dash.getRange(`A${k1}:B${k1}`).merge().setBackground(cardBg)
    .setValue(zendesk.totalBreached)
    .setFontSize(28).setFontWeight("bold").setFontColor(emailNumColor)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");
  // No Reply 12h+ — never received a first response
  const noReplyColor = zendesk.noReplyBreached > 0 ? risk : gray;
  dash.getRange(`C${k1}:D${k1}`).merge().setBackground(cardBg)
    .setValue(zendesk.noReplyBreached)
    .setFontSize(18).setFontWeight("bold").setFontColor(noReplyColor)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");
  const oldestColor = oldestWaitMin > slaMinutes ? emailNumColor : darkText;
  dash.getRange(`E${k1}:G${k1}`).merge().setBackground(cardBg)
    .setValue(oldestWaitStr)
    .setFontSize(18).setFontWeight("bold").setFontColor(oldestColor)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");

  // Phone primary number
  const phoneNumColor = phoneStatus === "At Risk" ? risk : (phoneStatus === "Watch" ? amber : darkText);
  dash.getRange(`I${k1}:J${k1}`).merge().setBackground(cardBg)
    .setValue(`${aircall.teamAnswerRate.toFixed(0)}%`)
    .setFontSize(28).setFontWeight("bold").setFontColor(phoneNumColor)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");
  // Answered fraction
  dash.getRange(`K${k1}:L${k1}`).merge().setBackground(cardBg)
    .setValue(`${aircall.teamAnswered} / ${rateEligible}`)
    .setFontSize(18).setFontWeight("bold").setFontColor(darkText)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");
  // Outbound calls — raw total
  dash.getRange(`M${k1}:P${k1}`).merge().setBackground(cardBg)
    .setValue(aircall.totalOutbound).setNumberFormat("0")
    .setFontSize(18).setFontWeight("bold").setFontColor(darkText)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");

  // Social KPIs — row 1: big numbers
  // Exclude conversations handed off to another team from unread count
  const metaConversations = (meta && meta.recentConversations) || [];
  const metaComments = (meta && meta.recentComments) || [];
  const metaUnread = metaConversations.filter(c => c.unread > 0 && !handedOffNames.has(c.customerName)).length;
  dash.getRange(`R${k1}:S${k1}`).merge().setBackground(cardBg)
    .setValue(metaUnread).setNumberFormat("0")
    .setFontSize(28).setFontWeight("bold").setFontColor(socialColor)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");
  dash.getRange(`T${k1}:U${k1}`).merge().setBackground(cardBg)
    .setValue(Math.round(socialOldestWaitMin / 60)).setNumberFormat("0")
    .setFontSize(18).setFontWeight("bold").setFontColor(darkText)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");
  dash.getRange(`V${k1}:Y${k1}`).merge().setBackground(cardBg)
    .setValue(metaComments.length).setNumberFormat("0")
    .setFontSize(18).setFontWeight("bold").setFontColor(darkText)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");

  // --- ROW 5: Labels under primary numbers ---
  const k2 = 5 + alertOffset;
  const label = navy;  // dark blue for all KPI labels — readable on white
  dash.setRowHeight(k2, 16);

  dash.getRange(`A${k2}:B${k2}`).merge().setBackground(cardBg)
    .setValue(`Total Waiting ${slaHours}+ Biz Hrs`)
    .setFontSize(8).setFontColor(label).setVerticalAlignment("top");
  dash.getRange(`C${k2}:D${k2}`).merge().setBackground(cardBg)
    .setValue(`Waiting ${slaHours}h+ 1st Reply`)
    .setFontSize(8).setFontColor(label).setVerticalAlignment("top");
  dash.getRange(`E${k2}:G${k2}`).merge().setBackground(cardBg)
    .setValue("Oldest Wait")
    .setFontSize(8).setFontColor(label).setVerticalAlignment("top");

  dash.getRange(`I${k2}:J${k2}`).merge().setBackground(cardBg)
    .setValue("Answer Rate")
    .setFontSize(8).setFontColor(label).setVerticalAlignment("top");
  dash.getRange(`K${k2}:L${k2}`).merge().setBackground(cardBg)
    .setValue("Answered")
    .setFontSize(8).setFontColor(label).setVerticalAlignment("top");
  dash.getRange(`M${k2}:P${k2}`).merge().setBackground(cardBg)
    .setValue("Outbound")
    .setFontSize(8).setFontColor(label).setVerticalAlignment("top");

  // Social labels — row 2
  dash.getRange(`R${k2}:S${k2}`).merge().setBackground(cardBg)
    .setValue("Unread DMs")
    .setFontSize(8).setFontColor(label).setVerticalAlignment("top");
  dash.getRange(`T${k2}:U${k2}`).merge().setBackground(cardBg)
    .setValue("Oldest DM Wait (h)")
    .setFontSize(8).setFontColor(label).setVerticalAlignment("top");
  dash.getRange(`V${k2}:Y${k2}`).merge().setBackground(cardBg)
    .setValue("Comments & Mentions (24h)")
    .setFontSize(8).setFontColor(label).setVerticalAlignment("top");

  // --- ROW 6: Queue counts (secondary numbers) ---
  const k3 = 6 + alertOffset;
  dash.setRowHeight(k3, 36);

  // Email queue counts: Open (with SAS sub) | On Hold | Unassigned
  dash.getRange(`A${k3}:B${k3}`).merge().setBackground(cardBg)
    .setValue(zendesk.openQueueCount).setNumberFormat("0")
    .setFontSize(16).setFontWeight("bold").setFontColor(darkText)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");
  dash.getRange(`C${k3}:D${k3}`).merge().setBackground(cardBg)
    .setValue(zendesk.onHoldQueueCount).setNumberFormat("0")
    .setFontSize(16).setFontWeight("bold").setFontColor(navy)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");
  dash.getRange(`E${k3}:F${k3}`).merge().setBackground(cardBg)
    .setValue(zendesk.unassigned).setNumberFormat("0")
    .setFontSize(16).setFontWeight("bold").setFontColor(navy)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");
  dash.getRange(`G${k3}`).setBackground(cardBg);

  // Phone secondary: Sent to Answer Service
  dash.getRange(`I${k3}:J${k3}`).merge().setBackground(cardBg)
    .setValue(fwdCount).setNumberFormat("0")
    .setFontSize(16).setFontWeight("bold").setFontColor(fwdCount > 0 ? amber : navy)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");
  dash.getRange(`K${k3}:P${k3}`).merge().setBackground(cardBg);

  // Social row 3: Total conversations
  const totalConversations = metaConversations.length > 0 ? metaConversations.length : 0;
  dash.getRange(`R${k3}:S${k3}`).merge().setBackground(cardBg)
    .setValue(totalConversations).setNumberFormat("0")
    .setFontSize(16).setFontWeight("bold").setFontColor(darkText)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");
  dash.getRange(`T${k3}:Y${k3}`).merge().setBackground(cardBg);

  // --- ROW 7: Labels under queue counts ---
  const k4 = 7 + alertOffset;
  dash.setRowHeight(k4, 16);

  // Open label shows SAS sub-count and voicemail count
  const sasColor = zendesk.sasTickets > 0 ? amber : navy;
  const vmCount = zendesk.openVoicemails || 0;
  const vmColor = vmCount > 0 ? amber : navy;
  const sasPart = ` · ${zendesk.sasTickets} New SAS`;
  const vmPart = vmCount > 0 ? ` · ${vmCount} VM` : "";
  const sasLabel = `Open${sasPart}${vmPart}`;
  const sasPartStart = 4;  // after "Open"
  const sasPartEnd = sasPartStart + sasPart.length;
  const vmPartStart = sasPartEnd;
  const vmPartEnd = vmPartStart + vmPart.length;

  const rtBuilder = SpreadsheetApp.newRichTextValue()
    .setText(sasLabel)
    .setTextStyle(0, 4, SpreadsheetApp.newTextStyle().setFontSize(8).setForegroundColor(navy).build())
    .setTextStyle(sasPartStart, sasPartEnd, SpreadsheetApp.newTextStyle().setFontSize(8).setForegroundColor(sasColor).build());
  if (vmPart) {
    rtBuilder.setTextStyle(vmPartStart, vmPartEnd, SpreadsheetApp.newTextStyle().setFontSize(8).setForegroundColor(vmColor).build());
  }
  dash.getRange(`A${k4}:B${k4}`).merge().setBackground(cardBg)
    .setRichTextValue(rtBuilder.build())
    .setVerticalAlignment("top");
  dash.getRange(`C${k4}:D${k4}`).merge().setBackground(cardBg)
    .setValue("On Hold")
    .setFontSize(8).setFontColor(navy).setVerticalAlignment("top");
  dash.getRange(`E${k4}:F${k4}`).merge().setBackground(cardBg)
    .setValue("Unassigned")
    .setFontSize(8).setFontColor(navy).setVerticalAlignment("top");
  dash.getRange(`G${k4}`).setBackground(cardBg);

  dash.getRange(`I${k4}:J${k4}`).merge().setBackground(cardBg)
    .setValue("Sent to Answer Service")
    .setFontSize(8).setFontColor(navy).setVerticalAlignment("top");
  dash.getRange(`K${k4}:P${k4}`).merge().setBackground(cardBg);

  dash.getRange(`R${k4}:S${k4}`).merge().setBackground(cardBg)
    .setValue("Active Conversations (24h)")
    .setFontSize(8).setFontColor(navy).setVerticalAlignment("top");
  dash.getRange(`T${k4}:Y${k4}`).merge().setBackground(cardBg);

  // --- ROW 8: Status accent bar (thin colored line under KPI) ---
  const k5 = 8 + alertOffset;
  dash.setRowHeight(k5, 4);
  const emailAccent = emailStatus === "At Risk" ? riskLt
    : emailStatus === "Watch" ? amberLt : BRAND.mossGreenLight;
  const phoneAccent = phoneStatus === "At Risk" ? riskLt
    : phoneStatus === "Watch" ? amberLt : BRAND.mossGreenLight;
  const socialAccent = socialStatus === "At Risk" ? riskLt
    : socialStatus === "Watch" ? amberLt : BRAND.mossGreenLight;
  dash.getRange(`A${k5}:G${k5}`).setBackground(emailAccent);
  dash.getRange(`H${k5}`).setBackground(bg);
  dash.getRange(`I${k5}:P${k5}`).setBackground(phoneAccent);
  dash.getRange(`R${k5}:Y${k5}`).setBackground(socialAccent);

  // Card borders around KPI panels for visual grouping
  const kpiBorder = { style: SpreadsheetApp.BorderStyle.SOLID, color: divider };
  dash.getRange(`A${k1}:G${k4}`).setBorder(true, true, true, true, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange(`I${k1}:P${k4}`).setBorder(true, true, true, true, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange(`R${k1}:Y${k4}`).setBorder(true, true, true, true, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);

  // Spacer row after KPIs
  const spacerRow = k5 + 1;
  dash.setRowHeight(spacerRow, 10);
  dash.getRange(`A${spacerRow}:Y${spacerRow}`).setBackground(bg);

  // ═══════════════════════════════════════════════
  // QUEUE INTELLIGENCE STRIP (full width A:Z, between KPIs and detail tables)
  // ═══════════════════════════════════════════════
  let qiOffset = 0;
  if (qi) {
    const qiHeaderRow = spacerRow + 1;
    qiOffset = 3;  // header + data row + spacer

    // Header bar
    dash.getRange(`A${qiHeaderRow}:Z${qiHeaderRow}`).merge()
      .setValue("Queue Intelligence")
      .setBackground(navy).setFontColor("#FFFFFF")
      .setFontSize(10).setFontWeight("bold")
      .setVerticalAlignment("middle");
    dash.setRowHeight(qiHeaderRow, 22);

    // Data row
    const qiDataRow = qiHeaderRow + 1;
    dash.setRowHeight(qiDataRow, 28);

    // Daily throughput metrics
    const netColor = qi.netToday > 0 ? risk : (qi.netToday < 0 ? green : darkText);
    const netStr = qi.netToday > 0 ? `+${qi.netToday}` : `${qi.netToday}`;

    // Cell 1: Created Today
    dash.getRange(`A${qiDataRow}:D${qiDataRow}`).merge().setBackground(cardBg)
      .setValue(`Created Today: ${qi.createdToday}`)
      .setFontSize(11).setFontWeight("bold").setFontColor(darkText)
      .setVerticalAlignment("middle");

    // Cell 2: Solved Today
    dash.getRange(`E${qiDataRow}:G${qiDataRow}`).merge().setBackground(cardBg)
      .setValue(`Solved Today: ${qi.solvedToday}`)
      .setFontSize(11).setFontWeight("bold").setFontColor(green)
      .setVerticalAlignment("middle");

    // Cell 3: Net (created - solved)
    dash.getRange(`H${qiDataRow}:I${qiDataRow}`).merge().setBackground(cardBg)
      .setValue(`Net: ${netStr}`)
      .setFontSize(11).setFontWeight("bold").setFontColor(netColor)
      .setVerticalAlignment("middle");

    // Cell 4: Est clear time
    const clearStr = qi.estClearHours !== null
      ? (qi.estClearHours < 1 ? "Est Clear: <1h" : `Est Clear: ${qi.estClearHours.toFixed(1)}h`)
      : "Est Clear: --";
    dash.getRange(`J${qiDataRow}:L${qiDataRow}`).merge().setBackground(cardBg)
      .setValue(clearStr)
      .setFontSize(11).setFontColor(darkText)
      .setVerticalAlignment("middle");

    // Cell 5: Spike indicator or normal status (daily grain)
    if (qi.isSpike) {
      const spikeMsg = `SPIKE: ${qi.createdToday} today (avg ${qi.avgDailyCreated}/day)`
        + (qi.spikeAnalysis ? ` -- ${qi.spikeAnalysis}` : "");
      dash.getRange(`M${qiDataRow}:Z${qiDataRow}`).merge()
        .setValue(spikeMsg)
        .setBackground(riskLt).setFontColor(risk)
        .setFontSize(9).setFontWeight("bold")
        .setVerticalAlignment("middle");
    } else {
      const statusMsg = `Volume normal (${qi.createdToday} today vs avg ${qi.avgDailyCreated}/day)`;
      dash.getRange(`M${qiDataRow}:Z${qiDataRow}`).merge().setBackground(cardBg)
        .setValue(statusMsg)
        .setFontSize(9).setFontColor(gray)
        .setVerticalAlignment("middle");
    }

    // Border around the QI strip
    dash.getRange(`A${qiDataRow}:Z${qiDataRow}`).setBorder(
      true, true, true, true, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);

    // Spacer after QI
    const qiSpacer = qiDataRow + 1;
    dash.setRowHeight(qiSpacer, 8);
    dash.getRange(`A${qiSpacer}:Z${qiSpacer}`).setBackground(bg);
  }

  // Gap column backgrounds for full height
  const maxBodyRow = 2000; // will update after we know final row
  dash.getRange(`H1:H${maxBodyRow}`).setBackground(bg);
  dash.getRange(`Q1:Q${maxBodyRow}`).setBackground(bg);

  // ═══════════════════════════════════════════════
  // EMAIL TABLES (Columns A-G)
  // ═══════════════════════════════════════════════

  // ─── Email Queue by Owner ───
  const eqRow = 10 + alertOffset + qiOffset;
  dash.getRange(`A${eqRow}:G${eqRow}`).merge()
    .setValue("Email Queue by Owner")
    .setBackground(navy).setFontColor("#FFFFFF")
    .setFontSize(11).setFontWeight("bold")
    .setVerticalAlignment("middle");
  dash.setRowHeight(eqRow, 26);

  // Visual grouping tints (matching Outbound pattern)
  const slaTint = BRAND.terracottaLight;   // warm tint for SLA/risk columns
  const solvedTint = BRAND.mossGreenLight; // green tint for productivity columns
  const waitTint = BRAND.roseQuartzLight;  // rose tint for wait time

  const ethRow = eqRow + 1;
  dash.setRowHeight(ethRow, 20);
  dash.getRange(`A${ethRow}:B${ethRow}`).merge().setValue("Owner")
    .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
  dash.getRange(`C${ethRow}`).setValue("New+Open")
    .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
  // 12+ Hrs — warm tint group
  dash.getRange(`D${ethRow}`).setValue(`${slaHours}+ Hrs`)
    .setFontWeight("bold").setFontSize(9).setFontColor(BRAND.beigeDark).setBackground(slaTint).setHorizontalAlignment("right");
  dash.getRange(`D${ethRow}`).setBorder(true, true, false, true, false, false, amber, SpreadsheetApp.BorderStyle.SOLID);
  // Solved Today — green tint group
  dash.getRange(`E${ethRow}`).setValue("Solved Today")
    .setFontWeight("bold").setFontSize(9).setFontColor(BRAND.beigeDark).setBackground(solvedTint).setHorizontalAlignment("right");
  dash.getRange(`E${ethRow}`).setBorder(true, true, false, true, false, false, green, SpreadsheetApp.BorderStyle.SOLID);
  // Oldest Wait
  dash.getRange(`F${ethRow}:G${ethRow}`).merge().setValue("Oldest Wait")
    .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
  dash.getRange(`A${ethRow}:G${ethRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);

  let aRow = ethRow + 1;
  const agentsToShow = [...CONFIG.agents];
  const otherAc = zendesk.agentCounts["Other"] || {};
  if ((otherAc.assigned || 0) > 0 || (otherAc.handledToday || 0) > 0) {
    agentsToShow.push("Other");
  }
  const otherLabel = zendesk.otherAgentNames && zendesk.otherAgentNames.length > 0
    ? "Other (" + zendesk.otherAgentNames.join(", ") + ")" : "Other";
  agentsToShow.push("_unassigned_");

  agentsToShow.forEach(agent => {
    const isUnassigned = agent === "_unassigned_";
    const isOther = agent === "Other";
    const displayName = isUnassigned ? "Unassigned" : (isOther ? otherLabel : agent);
    const ac = zendesk.agentCounts[agent] || { assigned: 0, pastSla: 0, longestWaitMin: 0, handledToday: 0 };
    const count = isUnassigned ? zendesk.unassigned : ac.assigned;
    const overSla = isUnassigned ? 0 : ac.pastSla;
    const handled = isUnassigned ? 0 : (ac.handledToday || 0);
    const longestWait = isUnassigned ? 0 : ac.longestWaitMin;
    const longestWaitStr = longestWait > 0 ? formatBizMinutes(longestWait) : "-";

    dash.setRowHeight(aRow, 22);
    dash.getRange(`A${aRow}:B${aRow}`).merge().setValue(displayName)
      .setBackground(cardBg).setFontSize(10);
    dash.getRange(`C${aRow}`).setValue(count)
      .setHorizontalAlignment("right").setBackground(cardBg).setFontSize(10);

    // 12+ Biz Hrs — tinted column with conditional emphasis
    dash.getRange(`D${aRow}`).setValue(overSla).setHorizontalAlignment("right").setFontSize(10)
      .setBackground(slaTint);
    dash.getRange(`D${aRow}`).setBorder(false, true, false, true, false, false, amber, SpreadsheetApp.BorderStyle.SOLID);
    if (overSla > 5) {
      dash.getRange(`D${aRow}`).setFontColor(risk).setFontWeight("bold");
    } else if (overSla > 0) {
      dash.getRange(`D${aRow}`).setFontColor(amber).setFontWeight("bold");
    } else {
      dash.getRange(`D${aRow}`).setFontColor(BRAND.beigeDark);
    }

    // Solved today — tinted column showing productivity
    dash.getRange(`E${aRow}`).setValue(handled).setNumberFormat("0")
      .setHorizontalAlignment("right").setFontSize(10).setBackground(solvedTint);
    dash.getRange(`E${aRow}`).setBorder(false, true, false, true, false, false, green, SpreadsheetApp.BorderStyle.SOLID);
    if (handled > 0) {
      dash.getRange(`E${aRow}`).setFontColor(green).setFontWeight("bold");
    } else {
      dash.getRange(`E${aRow}`).setFontColor(gray);
    }

    // Oldest wait — tint worst only
    dash.getRange(`F${aRow}:G${aRow}`).merge().setValue(longestWaitStr)
      .setHorizontalAlignment("right").setFontSize(10);
    if (longestWait > slaMinutes * 3) {
      dash.getRange(`F${aRow}:G${aRow}`).setBackground(riskLt).setFontColor(risk);
    } else if (longestWait > slaMinutes) {
      dash.getRange(`F${aRow}:G${aRow}`).setBackground(amberLt).setFontColor(amber);
    } else {
      dash.getRange(`F${aRow}:G${aRow}`).setBackground(cardBg);
    }

    dash.getRange(`A${aRow}:G${aRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    aRow++;
  });

  // ─── Oldest Waiting Tickets ───
  const ticketRow = aRow + 1;
  dash.getRange(`A${ticketRow}:G${ticketRow}`).merge()
    .setValue("Oldest Waiting Tickets")
    .setBackground(navy).setFontColor("#FFFFFF")
    .setFontSize(11).setFontWeight("bold").setVerticalAlignment("middle");
  dash.setRowHeight(ticketRow, 26);

  const tthRow = ticketRow + 1;
  dash.setRowHeight(tthRow, 20);
  dash.getRange(`A${tthRow}`).setValue("#").setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
  dash.getRange(`B${tthRow}:D${tthRow}`).merge().setValue("Subject").setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
  dash.getRange(`E${tthRow}`).setValue("Owner").setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
  dash.getRange(`F${tthRow}`).setValue("Wait").setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
  dash.getRange(`G${tthRow}`).setValue("Status").setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
  dash.getRange(`A${tthRow}:G${tthRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);

  let tRow = tthRow + 1;
  zendesk.longest10.forEach((ticket, idx) => {
    dash.setRowHeight(tRow, 22);
    const subj = ticket.subject.length > 42 ? ticket.subject.substring(0, 42) + "..." : ticket.subject;
    const waitStr = formatBizMinutes(ticket.waitBizMin);
    const rowBgT = idx % 2 === 1 ? altRow : cardBg; // zebra stripe

    const ticketUrl = `https://${CONFIG.zendesk.subdomain}.zendesk.com/agent/tickets/${ticket.id}`;
    dash.getRange(`A${tRow}`).setFormula(`=HYPERLINK("${ticketUrl}","${ticket.id}")`)
      .setFontSize(9).setFontColor("#1155CC")
      .setHorizontalAlignment("right").setBackground(rowBgT);
    dash.getRange(`B${tRow}:D${tRow}`).merge().setValue(subj)
      .setFontSize(9).setBackground(rowBgT);
    dash.getRange(`E${tRow}`).setValue(ticket.assignee.split(" ")[0])
      .setFontSize(9).setBackground(rowBgT);

    // Wait — top 3 get risk tint, rest get amber if past SLA
    dash.getRange(`F${tRow}`).setValue(waitStr).setFontSize(9).setHorizontalAlignment("right");
    if (idx < 3 && ticket.pastSla) {
      dash.getRange(`F${tRow}`).setBackground(riskLt).setFontColor(risk);
    } else if (ticket.pastSla) {
      dash.getRange(`F${tRow}`).setBackground(amberLt).setFontColor(amber);
    } else {
      dash.getRange(`F${tRow}`).setBackground(rowBgT).setFontColor(green);
    }

    dash.getRange(`G${tRow}`).setValue(ticket.status).setFontSize(9)
      .setFontColor(gray).setBackground(rowBgT);
    dash.getRange(`A${tRow}:G${tRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    tRow++;
  });

  // ─── Flagged Tickets (High Priority / Tags) ───
  const flagged = zendesk.flaggedTickets || [];
  if (flagged.length > 0) {
    tRow++; // spacer
    dash.getRange(`A${tRow}:G${tRow}`).merge()
      .setValue("Flagged Tickets")
      .setBackground(risk).setFontColor("#FFFFFF")
      .setFontSize(11).setFontWeight("bold").setVerticalAlignment("middle");
    dash.setRowHeight(tRow, 26);
    tRow++;

    // Column headers
    dash.setRowHeight(tRow, 20);
    dash.getRange(`A${tRow}`).setValue("#")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
    dash.getRange(`B${tRow}:C${tRow}`).merge().setValue("Subject")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`D${tRow}`).setValue("Priority")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`E${tRow}`).setValue("Owner")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`F${tRow}`).setValue("Wait")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
    dash.getRange(`G${tRow}`).setValue("Tags")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`A${tRow}:G${tRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    tRow++;

    flagged.forEach(ticket => {
      dash.setRowHeight(tRow, 22);
      const subj = ticket.subject.length > 32 ? ticket.subject.substring(0, 32) + "..." : ticket.subject;
      const waitStr = formatBizMinutes(ticket.waitBizMin);
      const isUrgent = ticket.priority === "urgent" || ticket.priority === "high";
      const rowBg = isUrgent ? riskLt : cardBg;

      const fTicketUrl = `https://${CONFIG.zendesk.subdomain}.zendesk.com/agent/tickets/${ticket.id}`;
      dash.getRange(`A${tRow}`).setFormula(`=HYPERLINK("${fTicketUrl}","${ticket.id}")`)
        .setFontSize(9).setFontColor("#1155CC")
        .setHorizontalAlignment("right").setBackground(rowBg);
      dash.getRange(`B${tRow}:C${tRow}`).merge().setValue(subj)
        .setFontSize(9).setBackground(rowBg);
      dash.getRange(`D${tRow}`).setValue(ticket.priority || "normal")
        .setFontSize(9).setFontColor(isUrgent ? risk : amber).setFontWeight("bold").setBackground(rowBg);
      dash.getRange(`E${tRow}`).setValue(ticket.assignee.split(" ")[0])
        .setFontSize(9).setBackground(rowBg);
      dash.getRange(`F${tRow}`).setValue(waitStr).setFontSize(9)
        .setHorizontalAlignment("right").setBackground(rowBg);
      // Show relevant tags (first 2 for space)
      const relevantTags = (ticket.tags || []).filter(t =>
        ["builder_warranty", "warranty", "escalated", "vip", "urgent"].includes(t)
      ).slice(0, 2).join(", ");
      dash.getRange(`G${tRow}`).setValue(relevantTags || "-")
        .setFontSize(8).setFontColor(amber).setBackground(rowBg);
      dash.getRange(`A${tRow}:G${tRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
      tRow++;
    });
  }

  // ─── CSAT Responses — Nicereply (last 24h) ───
  tRow++; // spacer
  const csatResponses = (csat && csat.responses) || [];
  dash.getRange(`A${tRow}:G${tRow}`).merge()
    .setValue("Email CSAT Survey")
    .setBackground(navy).setFontColor("#FFFFFF")
    .setFontSize(11).setFontWeight("bold").setVerticalAlignment("middle");
  dash.setRowHeight(tRow, 26);
  tRow++;

  if (csatResponses.length === 0) {
    dash.setRowHeight(tRow, 22);
    dash.getRange(`A${tRow}:G${tRow}`).merge()
      .setValue("No surveys submitted in the last 24 hours")
      .setFontSize(9).setFontColor(gray).setFontStyle("italic").setBackground(cardBg);
    dash.getRange(`A${tRow}:G${tRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    tRow++;
  } else {
    // CSAT summary line
    const csatScoreStr = csat.score !== null ? csat.score + "%" : "—";
    const csatSumColor = csat.score >= 90 ? green : (csat.score >= 80 ? amber : risk);
    dash.setRowHeight(tRow, 20);
    dash.getRange(`A${tRow}:B${tRow}`).merge()
      .setValue(csatScoreStr + " CSAT")
      .setFontSize(9).setFontWeight("bold").setFontColor(csat.score !== null ? csatSumColor : gray).setBackground(bg);
    dash.getRange(`C${tRow}:G${tRow}`).merge()
      .setValue(`${csat.satisfied} of ${csat.total} satisfied`)
      .setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`A${tRow}:G${tRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    tRow++;

    // Column headers
    dash.setRowHeight(tRow, 20);
    dash.getRange(`A${tRow}`).setValue("Score")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`B${tRow}:D${tRow}`).merge().setValue("Customer")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`E${tRow}`).setValue("#")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
    dash.getRange(`F${tRow}:G${tRow}`).merge().setValue("Time")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
    dash.getRange(`A${tRow}:G${tRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    tRow++;

    csatResponses.forEach((r, csIdx) => {
      dash.setRowHeight(tRow, 22);
      const csRowBg = csIdx % 2 === 1 ? altRow : cardBg;

      // Score — color-coded
      const scoreStr = r.score + "/" + r.maxScore;
      const scoreColor = r.satisfied ? green : risk;
      const scoreBg = r.satisfied ? csRowBg : riskLt;
      dash.getRange(`A${tRow}`).setValue(scoreStr)
        .setFontSize(9).setFontWeight("bold").setFontColor(scoreColor).setBackground(scoreBg);

      // Customer email
      const emailDisplay = r.email.length > 30 ? r.email.substring(0, 30) + "..." : r.email;
      dash.getRange(`B${tRow}:D${tRow}`).merge().setValue(emailDisplay)
        .setFontSize(9).setBackground(csRowBg);

      // Ticket ID — hyperlinked to Zendesk
      if (r.ticketId) {
        const csatTicketUrl = `https://${CONFIG.zendesk.subdomain}.zendesk.com/agent/tickets/${r.ticketId}`;
        dash.getRange(`E${tRow}`).setFormula(`=HYPERLINK("${csatTicketUrl}","${r.ticketId}")`)
          .setFontSize(9).setFontColor("#1155CC").setBackground(csRowBg).setHorizontalAlignment("right");
      } else {
        dash.getRange(`E${tRow}`).setValue("")
          .setFontSize(9).setFontColor(gray).setBackground(csRowBg).setHorizontalAlignment("right");
      }

      // Time
      dash.getRange(`F${tRow}:G${tRow}`).merge().setValue(r.dateStr + " " + r.timeStr)
        .setFontSize(9).setBackground(csRowBg).setHorizontalAlignment("right");

      dash.getRange(`A${tRow}:G${tRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
      tRow++;
    });
  }

  // ═══════════════════════════════════════════════
  // PHONE TABLES (Columns I-O)
  // ═══════════════════════════════════════════════

  // ─── Phone Activity by Agent ───
  const paHeaderRow = 10 + alertOffset + qiOffset;
  dash.getRange(`I${paHeaderRow}:P${paHeaderRow}`).merge()
    .setValue("Phone Activity by Agent")
    .setBackground(navy).setFontColor("#FFFFFF")
    .setFontSize(11).setFontWeight("bold").setVerticalAlignment("middle");
  dash.setRowHeight(paHeaderRow, 26);

  // Two-row header with visual Outbound grouping:
  // Row 1: Agent | In | ┌────── Outbound ──────┐
  // Row 2:              | Dialed | No Ans | <90s | 90s+
  const outboundBg = BRAND.airBlueLight;  // subtle blue tint to group outbound columns

  const pahRow1 = paHeaderRow + 1;
  dash.setRowHeight(pahRow1, 18);
  dash.getRange(`I${pahRow1}`).setValue("Agent")
    .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
  const inboundBg = BRAND.mossGreenLight;  // green tint for inbound answered
  dash.getRange(`J${pahRow1}`).setValue("In")
    .setFontWeight("bold").setFontSize(9).setFontColor(BRAND.beigeDark).setBackground(inboundBg).setHorizontalAlignment("right");
  dash.getRange(`J${pahRow1}`).setBorder(true, true, true, true, false, false, green, SpreadsheetApp.BorderStyle.SOLID);
  // Outbound group header — spans K-N with tinted background
  dash.getRange(`K${pahRow1}:N${pahRow1}`).merge().setValue("Outbound")
    .setFontWeight("bold").setFontSize(9).setFontColor(navy).setBackground(outboundBg)
    .setHorizontalAlignment("center");
  dash.getRange(`K${pahRow1}:N${pahRow1}`).setBorder(true, true, false, true, false, false, navy, SpreadsheetApp.BorderStyle.SOLID);
  // Talk time group header — spans O-P with terracotta tint (Deako brand)
  const talkBg = BRAND.terracottaLight;    // #DEAC90 warm terracotta tint
  const talkColor = BRAND.beigeDark;       // #523823 dark brown — cohesive with Deako palette
  dash.getRange(`O${pahRow1}:P${pahRow1}`).merge().setValue("Talk Time")
    .setFontWeight("bold").setFontSize(9).setFontColor(talkColor).setBackground(talkBg)
    .setHorizontalAlignment("center");
  dash.getRange(`O${pahRow1}:P${pahRow1}`).setBorder(true, true, false, true, false, false, talkColor, SpreadsheetApp.BorderStyle.SOLID);

  const pahRow2 = pahRow1 + 1;
  dash.setRowHeight(pahRow2, 14);
  dash.getRange(`I${pahRow2}:J${pahRow2}`).setBackground(bg);
  dash.getRange(`K${pahRow2}`).setValue("Dialed")
    .setFontSize(8).setFontColor(navy).setBackground(outboundBg).setHorizontalAlignment("right");
  dash.getRange(`L${pahRow2}`).setValue("No Ans")
    .setFontSize(8).setFontColor(navy).setBackground(outboundBg).setHorizontalAlignment("right");
  dash.getRange(`M${pahRow2}`).setValue("<90s")
    .setFontSize(8).setFontColor(navy).setBackground(outboundBg).setHorizontalAlignment("right");
  dash.getRange(`N${pahRow2}`).setValue("90s+")
    .setFontSize(8).setFontColor(navy).setBackground(outboundBg).setHorizontalAlignment("right");
  dash.getRange(`O${pahRow2}`).setValue("In")
    .setFontSize(8).setFontColor(talkColor).setBackground(talkBg).setHorizontalAlignment("right");
  dash.getRange(`P${pahRow2}`).setValue("Out")
    .setFontSize(8).setFontColor(talkColor).setBackground(talkBg).setHorizontalAlignment("right");
  dash.getRange(`K${pahRow2}:N${pahRow2}`).setBorder(false, true, true, true, false, false, navy, SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange(`O${pahRow2}:P${pahRow2}`).setBorder(false, true, true, true, false, false, talkColor, SpreadsheetApp.BorderStyle.SOLID);

  let paRow = pahRow2 + 1;
  CONFIG.agents.forEach(agent => {
    const stats = aircall.agentStats[agent] || { answered: 0, outbound: 0, outboundConnected: 0, outboundShort: 0, outboundLong: 0, inboundTalkTime: 0, outboundTalkTime: 0 };
    const noAnswer = stats.outbound - stats.outboundConnected;
    dash.setRowHeight(paRow, 22);
    dash.getRange(`I${paRow}`)
      .setValue(agent).setBackground(cardBg).setFontSize(10);
    dash.getRange(`J${paRow}`)
      .setValue(stats.answered).setNumberFormat("0").setHorizontalAlignment("right").setBackground(inboundBg).setFontSize(10);
    dash.getRange(`J${paRow}`).setBorder(false, true, false, true, false, false, green, SpreadsheetApp.BorderStyle.SOLID);
    // Outbound columns — tinted background for visual grouping
    dash.getRange(`K${paRow}`)
      .setValue(stats.outbound).setNumberFormat("0").setHorizontalAlignment("right").setBackground(outboundBg).setFontSize(10);
    dash.getRange(`L${paRow}`)
      .setValue(noAnswer).setNumberFormat("0").setHorizontalAlignment("right").setBackground(outboundBg).setFontSize(10)
      .setFontColor(gray);
    // <90s — likely voicemail
    dash.getRange(`M${paRow}`)
      .setValue(stats.outboundShort).setNumberFormat("0").setHorizontalAlignment("right").setBackground(outboundBg).setFontSize(10)
      .setFontColor(gray);
    // 90s+ — likely conversation
    dash.getRange(`N${paRow}`)
      .setValue(stats.outboundLong).setNumberFormat("0").setHorizontalAlignment("right").setBackground(outboundBg).setFontSize(10)
      .setFontColor(stats.outboundLong > 0 ? green : gray);
    // Talk time columns — In (O) and Out (P)
    const inTalk = stats.inboundTalkTime || 0;
    const outTalk = stats.outboundTalkTime || 0;
    dash.getRange(`O${paRow}`)
      .setValue(inTalk > 0 ? formatTalkTime(inTalk) : "-")
      .setFontSize(9).setHorizontalAlignment("right").setBackground(talkBg)
      .setFontColor(inTalk > 0 ? talkColor : gray);
    dash.getRange(`P${paRow}`)
      .setValue(outTalk > 0 ? formatTalkTime(outTalk) : "-")
      .setFontSize(9).setHorizontalAlignment("right").setBackground(talkBg)
      .setFontColor(outTalk > 0 ? talkColor : gray);
    // Borders
    dash.getRange(`K${paRow}:N${paRow}`).setBorder(false, true, false, true, false, false, navy, SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(`O${paRow}:P${paRow}`).setBorder(false, true, false, true, false, false, talkColor, SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(`I${paRow}:P${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    paRow++;
  });

  // Notes row (after-hours + SMS note)
  const phoneNotes = [];
  if (aircall.afterHoursCalls > 0) {
    phoneNotes.push(`${aircall.afterHoursCalls} call(s) outside biz hrs excluded`);
  }
  // SMS tracking now available via Aircall webhook
  dash.getRange(`I${paRow}:P${paRow}`).merge()
    .setValue(phoneNotes.join("  ·  "))
    .setFontColor(gray).setFontSize(8).setFontStyle("italic").setBackground(bg);
  paRow++;

  // ─── Missed Calls (detail table) ───
  paRow++; // spacer
  const missedDetails = aircall.missedCallDetails || [];
  dash.getRange(`I${paRow}:P${paRow}`).merge()
    .setValue("Missed Calls")
    .setBackground(navy).setFontColor("#FFFFFF")
    .setFontSize(11).setFontWeight("bold").setVerticalAlignment("middle");
  dash.setRowHeight(paRow, 26);
  paRow++;

  if (missedDetails.length === 0) {
    // Quiet empty state
    dash.setRowHeight(paRow, 22);
    dash.getRange(`I${paRow}:P${paRow}`).merge()
      .setValue("No missed calls today")
      .setFontSize(9).setFontColor(gray).setFontStyle("italic").setBackground(cardBg);
    dash.getRange(`I${paRow}:P${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    paRow++;
  } else {
    // Column headers: Line | Customer | Time | Reason
    dash.setRowHeight(paRow, 20);
    dash.getRange(`I${paRow}`).setValue("Line")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`J${paRow}:K${paRow}`).merge().setValue("Customer")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`L${paRow}`).setValue("Time")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
    dash.getRange(`M${paRow}:P${paRow}`).merge().setValue("Reason")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`I${paRow}:P${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    paRow++;

    missedDetails.forEach(detail => {
      dash.setRowHeight(paRow, 22);

      // Abbreviate line name for compact display
      const lineAbbrev = (detail.lineName || "")
        .replace("nonpro support", "nonpro")
        .replace("pro support", "pro")
        .replace("distributor support", "distrib")
        || "—";
      dash.getRange(`I${paRow}`)
        .setValue(lineAbbrev).setFontSize(9).setBackground(cardBg).setFontColor(navy);

      // Build customer display: name + number, just number, or Aircall link hint
      let customerDisplay;
      if (detail.callerNumber && detail.contactName) {
        customerDisplay = `${detail.contactName}  ${detail.callerNumber}`;
      } else if (detail.callerNumber) {
        customerDisplay = detail.callerNumber;
      } else {
        // No customer info via API — show call ID so agent can look it up in Aircall
        customerDisplay = detail.callId
          ? `Check Aircall #${detail.callId}`
          : "Check Aircall";
      }
      const customerColor = detail.callerNumber ? "#000000" : gray;
      dash.getRange(`J${paRow}:K${paRow}`).merge()
        .setValue(customerDisplay).setFontSize(9).setBackground(cardBg)
        .setFontColor(customerColor).setFontStyle(detail.callerNumber ? "normal" : "italic");
      dash.getRange(`L${paRow}`)
        .setValue(detail.callTime).setFontSize(9).setBackground(cardBg).setHorizontalAlignment("right");
      dash.getRange(`M${paRow}:P${paRow}`).merge()
        .setValue(detail.reason).setFontSize(9).setFontColor(amber).setBackground(cardBg);
      dash.getRange(`I${paRow}:P${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
      paRow++;
    });
  }

  // ─── Phone CSAT Survey — PostCall (last 24h) ───
  paRow++; // spacer
  const pcResponses = (postCall && postCall.responses) || [];
  dash.getRange(`I${paRow}:P${paRow}`).merge()
    .setValue("Phone CSAT Survey")
    .setBackground(navy).setFontColor("#FFFFFF")
    .setFontSize(11).setFontWeight("bold").setVerticalAlignment("middle");
  dash.setRowHeight(paRow, 26);
  paRow++;

  if (pcResponses.length === 0) {
    dash.setRowHeight(paRow, 22);
    dash.getRange(`I${paRow}:P${paRow}`).merge()
      .setValue("No surveys submitted in the last 24 hours")
      .setFontSize(9).setFontColor(gray).setFontStyle("italic").setBackground(cardBg);
    dash.getRange(`I${paRow}:P${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    paRow++;
  } else {
    // Summary line
    const pcScoreStr = postCall.score !== null ? postCall.score + "%" : "—";
    const pcSumColor = postCall.score >= 90 ? green : (postCall.score >= 80 ? amber : risk);
    dash.setRowHeight(paRow, 20);
    dash.getRange(`I${paRow}:J${paRow}`).merge()
      .setValue(pcScoreStr + " CSAT")
      .setFontSize(9).setFontWeight("bold").setFontColor(postCall.score !== null ? pcSumColor : gray).setBackground(bg);
    dash.getRange(`K${paRow}:P${paRow}`).merge()
      .setValue(`${postCall.satisfied} of ${postCall.total} satisfied`)
      .setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`I${paRow}:P${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    paRow++;

    // Column headers
    dash.setRowHeight(paRow, 20);
    dash.getRange(`I${paRow}`).setValue("Score")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`J${paRow}:K${paRow}`).merge().setValue("Customer")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`L${paRow}`).setValue("Agent")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`M${paRow}:P${paRow}`).merge().setValue("Time")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
    dash.getRange(`I${paRow}:P${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    paRow++;

    pcResponses.forEach((r, pcIdx) => {
      dash.setRowHeight(paRow, 22);
      const scoreStr = r.score + "/" + r.maxScore;
      const scoreColor = r.satisfied ? green : risk;
      const scoreBg = r.satisfied ? (pcIdx % 2 === 1 ? altRow : cardBg) : riskLt;
      const pcRowBg = pcIdx % 2 === 1 ? altRow : cardBg;
      dash.getRange(`I${paRow}`).setValue(scoreStr)
        .setFontSize(9).setFontWeight("bold").setFontColor(scoreColor).setBackground(scoreBg);
      dash.getRange(`J${paRow}:K${paRow}`).merge().setValue(r.phone)
        .setFontSize(9).setBackground(pcRowBg);
      dash.getRange(`L${paRow}`).setValue(r.agent ? r.agent.split(" ")[0] : "")
        .setFontSize(9).setBackground(pcRowBg);
      dash.getRange(`M${paRow}:P${paRow}`).merge().setValue(r.dateStr + " " + r.timeStr)
        .setFontSize(9).setBackground(pcRowBg).setHorizontalAlignment("right");
      dash.getRange(`I${paRow}:P${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
      paRow++;
    });
  }

  // ─── Text Messages ───
  const smsData = sms || { totalToday: 0, inbound: 0, outbound: 0, agentStats: {}, messages: [] };
  paRow++; // spacer
  dash.getRange(`I${paRow}:P${paRow}`).merge()
    .setValue("Text Messages")
    .setBackground(navy).setFontColor("#FFFFFF")
    .setFontSize(11).setFontWeight("bold").setVerticalAlignment("middle");
  dash.setRowHeight(paRow, 26);
  paRow++;

  if (smsData.totalToday === 0) {
    dash.setRowHeight(paRow, 22);
    dash.getRange(`I${paRow}:P${paRow}`).merge()
      .setValue("No SMS activity today")
      .setFontSize(9).setFontColor(gray).setFontStyle("italic").setBackground(cardBg);
    dash.getRange(`I${paRow}:P${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    paRow++;
  } else {
    // Summary line: In X · Out X · Total X
    const smsSumColor = smsData.totalToday > 0 ? darkText : gray;
    dash.setRowHeight(paRow, 20);
    dash.getRange(`I${paRow}:P${paRow}`).merge()
      .setRichTextValue(
        SpreadsheetApp.newRichTextValue()
          .setText(`In: ${smsData.inbound}  ·  Out: ${smsData.outbound}  ·  Total: ${smsData.totalToday}`)
          .setTextStyle(SpreadsheetApp.newTextStyle().setFontSize(9).setForegroundColor(darkText).build())
          .build()
      ).setBackground(bg);
    dash.getRange(`I${paRow}:P${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    paRow++;

    // Per-agent SMS counts (only show agents with activity)
    const smsAgentsWithActivity = CONFIG.agents.filter(a =>
      smsData.agentStats[a] && (smsData.agentStats[a].sent > 0 || smsData.agentStats[a].received > 0)
    );

    if (smsAgentsWithActivity.length > 0) {
      // Header
      dash.setRowHeight(paRow, 20);
      dash.getRange(`I${paRow}:J${paRow}`).merge().setValue("Agent")
        .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
      dash.getRange(`K${paRow}:L${paRow}`).merge().setValue("Sent")
        .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
      dash.getRange(`M${paRow}:P${paRow}`).merge().setValue("Received")
        .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
      dash.getRange(`I${paRow}:P${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
      paRow++;

      smsAgentsWithActivity.forEach(agent => {
        const s = smsData.agentStats[agent];
        dash.setRowHeight(paRow, 22);
        dash.getRange(`I${paRow}:J${paRow}`).merge().setValue(agent)
          .setFontSize(9).setBackground(cardBg);
        dash.getRange(`K${paRow}:L${paRow}`).merge().setValue(s.sent)
          .setFontSize(9).setBackground(cardBg).setHorizontalAlignment("right");
        dash.getRange(`M${paRow}:P${paRow}`).merge().setValue(s.received)
          .setFontSize(9).setBackground(cardBg).setHorizontalAlignment("right");
        dash.getRange(`I${paRow}:P${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
        paRow++;
      });
    }

    // Recent messages (last 10)
    const recentSMS = smsData.messages.slice(0, 10);
    if (recentSMS.length > 0) {
      // Section sub-header
      dash.setRowHeight(paRow, 20);
      dash.getRange(`I${paRow}:P${paRow}`).merge().setValue("Recent Messages")
        .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
      dash.getRange(`I${paRow}:P${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
      paRow++;

      recentSMS.forEach(m => {
        const isOut = m.direction === "outbound";
        const dirIcon = isOut ? "→ Out" : "← In";
        const dirColor = isOut ? BRAND.airBlueDark : green;

        // Row 1: direction, agent/contact info, time
        // Strip leading apostrophe from phone (stored that way in SMS Log to prevent formula interpretation)
        const safePhone = m.phone ? m.phone.replace(/^'/, "") : "";
        let description = "";
        if (isOut) {
          const agentShort = m.agent ? m.agent.split(" ")[0] : "?";
          const contactStr = m.contact || safePhone || "Unknown";
          description = `${agentShort} → ${contactStr}`;
          if (m.lineName) description += ` via ${m.lineName}`;
        } else {
          const contactStr = m.contact || safePhone || "Unknown";
          description = `${contactStr}`;
          if (m.lineName) description += ` → ${m.lineName}`;
        }

        dash.setRowHeight(paRow, 20);
        dash.getRange(`I${paRow}`).setValue(dirIcon)
          .setFontSize(9).setFontColor(dirColor).setFontWeight("bold").setBackground(cardBg);
        dash.getRange(`J${paRow}:M${paRow}`).merge().setNumberFormat("@").setValue(description)
          .setFontSize(9).setBackground(cardBg);
        dash.getRange(`N${paRow}`).setValue(m.timeStr)
          .setFontSize(9).setBackground(cardBg).setHorizontalAlignment("right");
        paRow++;

        // Row 2: message body (truncated to fit, lighter color)
        if (m.body) {
          const truncBody = m.body.length > 120 ? m.body.substring(0, 117) + "..." : m.body;
          dash.setRowHeight(paRow, 18);
          dash.getRange(`I${paRow}:P${paRow}`).merge().setValue(truncBody)
            .setFontSize(8).setFontColor(gray).setFontStyle("italic").setBackground(cardBg)
            .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        } else {
          dash.setRowHeight(paRow, 4);
          dash.getRange(`I${paRow}:P${paRow}`).merge().setBackground(cardBg);
        }
        dash.getRange(`I${paRow}:P${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
        paRow++;
      });
    }
  }


  // ═══════════════════════════════════════════════
  // SOCIAL TABLES (Columns R-Y)
  // ═══════════════════════════════════════════════
  let sRow = 10 + alertOffset + qiOffset;  // Social row counter (same starting row as phone/email)

  // ─── Social (Meta Business Suite) ───
  sRow++; // spacer to align with other sections
  dash.getRange(`R${sRow}:Z${sRow}`).merge()
    .setRichTextValue(
      SpreadsheetApp.newRichTextValue()
        .setText("Social (Meta Business Suite)")
        .setLinkUrl("https://business.facebook.com/latest/inbox/all")
        .build()
    )
    .setBackground(navy).setFontColor("#FFFFFF")
    .setFontSize(11).setFontWeight("bold").setVerticalAlignment("middle");
  dash.setRowHeight(sRow, 26);
  sRow++;

  // Summary line — DMs + comments totals
  const metaStatusColor = metaUnread > 0 ? amber : green;
  const metaSummary = metaUnread > 0
    ? `${metaUnread} unread DM${metaUnread !== 1 ? "s" : ""}`
    : "DMs: all caught up";
  const commentSummary = metaComments.length > 0
    ? `${metaComments.length} comment${metaComments.length !== 1 ? "s" : ""}/mention${metaComments.length !== 1 ? "s" : ""} (24h)`
    : "No new comments (24h)";
  dash.setRowHeight(sRow, 20);
  dash.getRange(`R${sRow}:T${sRow}`).merge()
    .setValue(metaSummary)
    .setFontSize(9).setFontWeight("bold").setFontColor(metaStatusColor).setBackground(bg);
  dash.getRange(`U${sRow}:Z${sRow}`).merge()
    .setValue(commentSummary)
    .setFontSize(9).setFontColor(metaComments.length > 0 ? amber : gray).setBackground(bg);
  dash.getRange(`R${sRow}:Z${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
  sRow++;

  // Token expiry warning
  if (meta && meta.tokenWarning) {
    dash.setRowHeight(sRow, 20);
    dash.getRange(`R${sRow}:Z${sRow}`).merge()
      .setValue(meta.tokenWarning)
      .setFontSize(9).setFontWeight("bold").setFontColor(risk).setBackground(riskLt);
    dash.getRange(`R${sRow}:Z${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    sRow++;
  }

  // ── DMs sub-header ──
  dash.setRowHeight(sRow, 20);
  dash.getRange(`R${sRow}:Z${sRow}`).merge()
    .setValue("Direct Messages (Last 24h / Unread)")
    .setFontWeight("bold").setFontSize(9).setFontColor(darkText).setBackground(bg);
  dash.getRange(`R${sRow}:Z${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
  sRow++;

  if (metaConversations.length === 0) {
    dash.setRowHeight(sRow, 22);
    dash.getRange(`R${sRow}:Z${sRow}`).merge()
      .setValue("No recent conversations")
      .setFontSize(9).setFontColor(gray).setFontStyle("italic").setBackground(cardBg);
    dash.getRange(`R${sRow}:Z${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    sRow++;
  } else {
    // Column headers
    dash.setRowHeight(sRow, 20);
    dash.getRange(`R${sRow}`).setValue("Customer")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`S${sRow}`).setValue("Via")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`T${sRow}:V${sRow}`).merge().setValue("Last Message")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`W${sRow}`).setValue("From")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`X${sRow}:Y${sRow}`).merge().setValue("Time")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
    dash.getRange(`Z${sRow}`).setValue("_social_header"); // marker for handoff column header
    dash.getRange(`R${sRow}:Z${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    sRow++;

    // Show up to 10 conversations
    metaConversations.slice(0, 10).forEach((convo, cIdx) => {
      // Skip orange highlight for conversations handed off or resolved
      const isHandedOff = handedOffNames && handedOffNames.has(convo.customerName);
      const rowBg = (convo.unread > 0 && !isHandedOff) ? amberLt : (cIdx % 2 === 1 ? altRow : cardBg);

      // Row 1: Customer name, platform badge, from, time
      dash.setRowHeight(sRow, 20);
      dash.getRange(`R${sRow}`)
        .setValue(convo.customerName)
        .setFontSize(9).setFontWeight("bold").setFontColor(darkText).setBackground(rowBg);
      // Platform badge
      const platformShort = convo.platform === "Instagram" ? "IG" : "FB";
      const platformColor = convo.platform === "Instagram" ? "#C13584" : "#1877F2";
      dash.getRange(`S${sRow}`).setValue(platformShort)
        .setFontSize(8).setFontWeight("bold").setFontColor(platformColor).setBackground(rowBg);
      dash.getRange(`T${sRow}:V${sRow}`).merge()
        .setBackground(rowBg);
      dash.getRange(`W${sRow}`).setValue(convo.lastMessageFrom)
        .setFontSize(8).setFontColor(gray).setBackground(rowBg);

      // Format time
      let timeDisplay = "";
      if (convo.time) {
        try {
          const d = new Date(convo.time);
          timeDisplay = Utilities.formatDate(d, Session.getScriptTimeZone(), "MMM d h:mm a");
        } catch (e) { timeDisplay = ""; }
      }
      dash.getRange(`X${sRow}:Y${sRow}`).merge().setValue(timeDisplay)
        .setFontSize(8).setFontColor(gray).setBackground(rowBg).setHorizontalAlignment("right");
      dash.getRange(`Z${sRow}`).setValue("_social_dm_row"); // marker for handoff dropdown
      sRow++;

      // Row 2: message excerpt
      dash.setRowHeight(sRow, 18);
      const excerpt = convo.lastMessage || "";
      dash.getRange(`R${sRow}:Z${sRow}`).merge().setValue(excerpt)
        .setFontSize(8).setFontColor(gray).setFontStyle("italic").setBackground(rowBg)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      dash.getRange(`R${sRow}:Z${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
      sRow++;
    });
  }

  // ── Comments & Mentions sub-header (last 24h) ──
  sRow++; // spacer
  dash.setRowHeight(sRow, 20);
  dash.getRange(`R${sRow}:Z${sRow}`).merge()
    .setValue("Comments & Mentions (Last 24h)")
    .setFontWeight("bold").setFontSize(9).setFontColor(darkText).setBackground(bg);
  dash.getRange(`R${sRow}:Z${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
  sRow++;

  if (metaComments.length === 0) {
    dash.setRowHeight(sRow, 22);
    dash.getRange(`R${sRow}:Z${sRow}`).merge()
      .setValue("No new comments or mentions")
      .setFontSize(9).setFontColor(gray).setFontStyle("italic").setBackground(cardBg);
    dash.getRange(`R${sRow}:Z${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    sRow++;
  } else {
    // Column headers
    dash.setRowHeight(sRow, 20);
    dash.getRange(`R${sRow}`).setValue("Author")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`S${sRow}`).setValue("Source")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`T${sRow}:V${sRow}`).merge().setValue("Comment")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`W${sRow}`).setValue("On")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`X${sRow}:Y${sRow}`).merge().setValue("Time")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
    dash.getRange(`Z${sRow}`).setValue("_social_header"); // marker for handoff column header
    dash.getRange(`R${sRow}:Z${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    sRow++;

    metaComments.slice(0, 8).forEach((c, cmIdx) => {
      const isMention = c.type === "mention";
      const rowBg = isMention ? amberLt : (cmIdx % 2 === 1 ? altRow : cardBg);

      dash.setRowHeight(sRow, 20);
      // Author (hyperlinked to source)
      const safeAuthor = c.author.replace(/"/g, '""');
      dash.getRange(`R${sRow}`)
        .setFormula(`=HYPERLINK("${c.url}","${safeAuthor}")`)
        .setFontSize(9).setFontColor("#1155CC").setBackground(rowBg);
      // Platform + type badge
      const platformShort = c.platform === "Instagram" ? "IG" : "FB";
      const platformColor = c.platform === "Instagram" ? "#C13584" : "#1877F2";
      const label = isMention ? platformShort + " tag" : platformShort;
      dash.getRange(`S${sRow}`).setValue(label)
        .setFontSize(8).setFontWeight("bold").setFontColor(platformColor).setBackground(rowBg);
      dash.getRange(`T${sRow}:V${sRow}`).merge()
        .setBackground(rowBg);
      // Post snippet in "On" column
      dash.getRange(`W${sRow}`).setValue(c.postSnippet || "")
        .setFontSize(8).setFontColor(gray).setBackground(rowBg)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

      // Format time
      let timeDisplay = "";
      if (c.time) {
        try {
          const d = new Date(c.time);
          timeDisplay = Utilities.formatDate(d, Session.getScriptTimeZone(), "MMM d h:mm a");
        } catch (e) { timeDisplay = ""; }
      }
      dash.getRange(`X${sRow}:Y${sRow}`).merge().setValue(timeDisplay)
        .setFontSize(8).setFontColor(gray).setBackground(rowBg).setHorizontalAlignment("right");
      dash.getRange(`Z${sRow}`).setValue("_social_comment_row"); // marker for handoff dropdown
      sRow++;

      // Row 2: comment text
      dash.setRowHeight(sRow, 18);
      dash.getRange(`R${sRow}:Z${sRow}`).merge().setValue(c.text || "")
        .setFontSize(8).setFontColor(gray).setFontStyle("italic").setBackground(rowBg)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      dash.getRange(`R${sRow}:Z${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
      sRow++;
    });
  }


  // ─── COLUMN WIDTHS ───
  dash.setColumnWidth(1, 60);   // A
  dash.setColumnWidth(2, 110);  // B
  dash.setColumnWidth(3, 80);   // C
  dash.setColumnWidth(4, 80);   // D
  dash.setColumnWidth(5, 80);   // E
  dash.setColumnWidth(6, 70);   // F
  dash.setColumnWidth(7, 75);   // G
  dash.setColumnWidth(8, 20);   // H spacer
  dash.setColumnWidth(9, 100);  // I — Agent name / Customer
  dash.setColumnWidth(10, 38);  // J — In / Via badge
  dash.setColumnWidth(11, 48);  // K — Dialed
  dash.setColumnWidth(12, 48);  // L — No Ans
  dash.setColumnWidth(13, 42);  // M — <90s
  dash.setColumnWidth(14, 42);  // N — 90s+
  dash.setColumnWidth(15, 62);  // O — In Talk
  dash.setColumnWidth(16, 62);  // P — Out Talk
  dash.setColumnWidth(17, 20);  // Q spacer
  dash.setColumnWidth(18, 110); // R — Customer/Author
  dash.setColumnWidth(19, 40);  // S — Via/Source badge
  dash.setColumnWidth(20, 70);  // T — Message col 1
  dash.setColumnWidth(21, 70);  // U — Message col 2
  dash.setColumnWidth(22, 60);  // V — Message col 3
  dash.setColumnWidth(23, 60);  // W — From/On
  dash.setColumnWidth(24, 50);  // X — Time col 1
  dash.setColumnWidth(25, 50);  // Y — Time col 2
  dash.setColumnWidth(26, 75);  // Z — Owner (social handoff)

  // Fill remaining
  const lastRow = Math.max(tRow, paRow, sRow) + 2;
  dash.getRange(`A${lastRow}:Z${lastRow + 5}`).setBackground(bg);

  // Footer — version & goals
  dash.getRange(`A${lastRow}:Z${lastRow}`).merge()
    .setValue(`CS Visibility · Command Center v2.5.4  ·  Refreshes every 5 min  ·  Goal: reply within ${slaHours} business hours · answer ${CONFIG.thresholds.phoneAnswerRate.green}%+ inbound calls · Mon-Fri 6a-5p PST`)
    .setFontColor(gray).setFontSize(8).setFontStyle("italic")
    .setHorizontalAlignment("center").setBackground(bg);

  // Footer — legend & logic explanation
  const legendRow = lastRow + 1;
  const legendLines = [
    `Email status: Healthy = 0–5 tickets past SLA, Watch = 6–10 past SLA, At Risk = 11+ past SLA  ·  `
    + `Phone status: Healthy = answer rate ≥ ${CONFIG.thresholds.phoneAnswerRate.green}%, Watch = ${CONFIG.thresholds.phoneAnswerRate.yellow}–${CONFIG.thresholds.phoneAnswerRate.green - 1}%, At Risk = < ${CONFIG.thresholds.phoneAnswerRate.yellow}%  ·  `
    + `Social status: Healthy = oldest DM wait ≤ 2h, Watch = 2–6h, At Risk = > 6h`,
    `Wait times are business hours only (Mon-Fri 6a-5p PST)  ·  Past SLA = waiting > ${CONFIG.thresholds.oldestUnanswered.green} biz hrs without a reply  ·  `
    + `Oldest Waiting table: top 3 past SLA highlighted red, others past SLA amber, within SLA green`,
    `CSAT % = (satisfied ÷ total) × 100  ·  Satisfied = 4+ out of 5  ·  Phone answer rate = answered inbound ÷ total inbound (biz hrs only)  ·  `
    + `Open tickets exclude ${(CONFIG.excludeAgents || []).join(", ")} (not on CS team)`,
    `Social: FB = Facebook Messenger, IG = Instagram DM  ·  Comments & Mentions show last 24h from FB + IG posts  ·  `
    + `Amber highlight = unread DM or @mention  ·  Owner column: set to "Other Team" to exclude from SLA  ·  Meta token expiry warning at 7 days`,
    `API notes: IG DMs powered by Meta webhooks → IG DM Log sheet  ·  `
    + `Missed call customer info shows "Check Aircall #" when contact data is not exposed in the API payload  ·  `
    + `${zendesk.aiAgentCount > 0 ? zendesk.aiAgentCount + " AI agent ticket(s) excluded (auto-close after 4 days idle)" : "AI agent bot tickets excluded from all counts"}`,
    `Queue Intelligence: Spike = today's created > ${qi ? qi.spikeMultiplier + "x" : "2x"} 7-day avg (min 10 tickets)  ·  `
    + `Est Clear = queue size / solve rate per hour  ·  Net = created today - solved today  ·  `
    + `${qi && qi.spikeAnalysis ? "Spike analysis powered by Claude AI" : "Set ANTHROPIC_API_KEY for AI-powered spike analysis"}`,
  ];
  for (let li = 0; li < legendLines.length; li++) {
    dash.getRange(`A${legendRow + li}:Z${legendRow + li}`).merge()
      .setValue(legendLines[li])
      .setFontColor(gray).setFontSize(7).setFontStyle("italic")
      .setHorizontalAlignment("center").setBackground(bg).setWrap(false);
  }
}

// --- SOCIAL OLDEST WAIT HELPER ---
// handedOff is an optional Set of customer names that have been assigned to another team.
// When provided, those conversations are excluded from the SLA wait time calculation.
function computeSocialOldestWait(meta, handedOff) {
  if (!meta || !meta.recentConversations) return 0;
  const now = new Date();
  let oldestMin = 0;
  meta.recentConversations.forEach(c => {
    if (c.unread > 0 && c.time) {
      // Skip conversations handed off to another team
      if (handedOff && handedOff.has(c.customerName)) return;
      const msgTime = new Date(c.time);
      const diffMin = (now - msgTime) / 60000;
      if (diffMin > oldestMin) oldestMin = diffMin;
    }
  });
  return Math.round(oldestMin);
}


// --- WRITE RAW DATA TABS ---
function writeZendeskRaw(ss, data) {
  const sheet = getOrCreateSheet(ss, "Zendesk Raw");
  sheet.clear();

  sheet.getRange("A1").setValue("Last Fetched").setFontWeight("bold");
  sheet.getRange("B1").setValue(new Date());
  sheet.getRange("A2").setValue("Awaiting Response").setFontWeight("bold");
  sheet.getRange("B2").setValue(data.totalOpen);
  sheet.getRange("A3").setValue("Past SLA / Unassigned").setFontWeight("bold");
  sheet.getRange("B3").setValue(`${data.totalBreached} / ${data.unassigned}`);

  const headers = ["ID", "Subject", "Requester", "Assignee", "Status", "Priority", "Created", "Wait (biz)", "Tags"];
  headers.forEach((h, i) => sheet.getRange(5, i + 1).setValue(h).setFontWeight("bold").setBackground(BRAND.beigeLight));

  data.tickets.forEach((ticket, i) => {
    const row = i + 6;
    sheet.getRange(row, 1).setValue(ticket.id);
    sheet.getRange(row, 2).setValue(ticket.subject);
    sheet.getRange(row, 3).setValue(ticket.requester);
    sheet.getRange(row, 4).setValue(ticket.assignee);
    sheet.getRange(row, 5).setValue(ticket.status);
    sheet.getRange(row, 6).setValue(ticket.priority);
    sheet.getRange(row, 7).setValue(ticket.created);
    sheet.getRange(row, 8).setValue(formatBizMinutes(ticket.waitBizMin));
    sheet.getRange(row, 9).setValue((ticket.tags || []).join(", "));
  });
}

function writeAircallRaw(ss, data) {
  const sheet = getOrCreateSheet(ss, "Aircall Raw");
  sheet.clear();

  sheet.getRange("A1").setValue("Last Fetched").setFontWeight("bold");
  sheet.getRange("B1").setValue(new Date());
  sheet.getRange("A2").setValue("Total Inbound").setFontWeight("bold");
  sheet.getRange("B2").setValue(data.totalInbound);
  sheet.getRange("A3").setValue("Team Answer Rate").setFontWeight("bold");
  sheet.getRange("B3").setValue(data.teamAnswerRate.toFixed(1) + "%");
  sheet.getRange("A4").setValue("Team / Fwd to SAS / Short Hangup").setFontWeight("bold");
  sheet.getRange("B4").setValue(`${data.teamAnswered} / ${data.forwarded} / ${data.shortAbandoned}`);

  const headers = ["ID", "Direction", "Status", "From", "To", "Agent", "Duration (s)", "Wait (s)", "Started At"];
  headers.forEach((h, i) => sheet.getRange(5, i + 1).setValue(h).setFontWeight("bold").setBackground(BRAND.beigeLight));

  data.calls.forEach((call, i) => {
    const row = i + 6;
    const agent = call.user
      ? (call.user.name || `${call.user.first_name || ""} ${call.user.last_name || ""}`.trim())
      : "None";
    sheet.getRange(row, 1).setValue(call.id);
    sheet.getRange(row, 2).setValue(call.direction);
    sheet.getRange(row, 3).setValue(call.status || "");
    sheet.getRange(row, 4).setValue(call.raw_digits || "");
    sheet.getRange(row, 5).setValue(call.number ? call.number.digits : "");
    sheet.getRange(row, 6).setValue(agent);
    sheet.getRange(row, 7).setValue(call.duration || 0);
    sheet.getRange(row, 8).setValue(call.waiting_duration || 0);
    sheet.getRange(row, 9).setValue(call.started_at ? new Date(call.started_at * 1000) : "");
  });
}

// --- THRESHOLDS TAB ---
function setupThresholdsTab(ss) {
  const sheet = getOrCreateSheet(ss, "Thresholds");
  sheet.clear();
  sheet.setTabColor(BRAND.terracotta);

  sheet.getRange("A1").setValue("CS Command Center — Thresholds")
    .setFontSize(14).setFontWeight("bold").setFontFamily("Inter").setFontColor(BRAND.airBlueDark);
  sheet.getRange("A2").setValue("Edit the CONFIG object in Code.gs to change these values")
    .setFontColor(BRAND.textSecondary).setFontFamily("Inter");

  const headers = ["Metric", "Green If", "Yellow If", "Red If", "Unit"];
  headers.forEach((h, i) => sheet.getRange(4, i + 1).setValue(h).setFontWeight("bold")
    .setBackground(BRAND.beigeLight).setFontFamily("Inter"));

  const rows = [
    ["Oldest Unanswered Email", "< 12", "12 — 24", "> 24", "hours (12h SLA)"],
    ["Open Backlog Count", "< 30", "30 — 50", "> 50", "tickets"],
    ["Phone Answer Rate", "> 90%", "75 — 90%", "< 75%", "% (set your target)"],
    ["Median First Response Time", "< 12", "12 — 24", "> 24", "hours (12h SLA)"],
    ["Avg Caller Wait Time", "< 30", "30 — 60", "> 60", "seconds"],
  ];

  rows.forEach((r, i) => {
    r.forEach((val, j) => {
      const cell = sheet.getRange(5 + i, j + 1).setFontFamily("Inter");
      cell.setValue(val);
      if (j === 1) cell.setFontColor(BRAND.statusGreen).setFontWeight("bold");
      if (j === 2) cell.setFontColor(BRAND.statusYellow).setFontWeight("bold");
      if (j === 3) cell.setFontColor(BRAND.statusRed).setFontWeight("bold");
    });
  });
}

// --- AGENT MAP TAB ---
function setupAgentMapTab(ss) {
  const sheet = getOrCreateSheet(ss, "Agent Map");
  sheet.clear();
  sheet.setTabColor(BRAND.mossGreen);

  sheet.getRange("A1").setValue("Agent Name Mapping")
    .setFontSize(14).setFontWeight("bold").setFontFamily("Inter").setFontColor(BRAND.airBlueDark);
  sheet.getRange("A2").setValue("Update CONFIG.agents in Code.gs to add/remove agents")
    .setFontColor(BRAND.textSecondary).setFontFamily("Inter");

  const headers = ["Display Name", "Zendesk Name", "Aircall Name", "Role"];
  headers.forEach((h, i) => sheet.getRange(4, i + 1).setValue(h).setFontWeight("bold")
    .setBackground(BRAND.beigeLight).setFontFamily("Inter"));

  const agents = [
    ["Agent A", "Agent A", "Agent A", "Anchor"],
    ["Agent B", "Agent B", "Agent B", "Full-Time (Mar 1)"],
    ["Agent C", "Agent C", "Agent C", "Ramping"],
    ["Manager", "Manager", "Manager", "Manager"],
  ];

  agents.forEach((a, i) => {
    a.forEach((val, j) => sheet.getRange(5 + i, j + 1).setValue(val).setFontFamily("Inter"));
  });
}

// --- RUN LOG ---
function logRun(sheet, startTime, status, error) {
  const lastRow = Math.max(sheet.getLastRow(), 1);

  if (lastRow <= 1) {
    sheet.getRange("A1").setValue("Timestamp").setFontWeight("bold").setFontFamily("Inter");
    sheet.getRange("B1").setValue("Status").setFontWeight("bold").setFontFamily("Inter");
    sheet.getRange("C1").setValue("Duration (s)").setFontWeight("bold").setFontFamily("Inter");
    sheet.getRange("D1").setValue("Error").setFontWeight("bold").setFontFamily("Inter");
    sheet.setTabColor(BRAND.ashGray);
  }

  const row = lastRow + 1;
  const duration = ((new Date() - startTime) / 1000).toFixed(1);

  sheet.getRange(`A${row}`).setValue(startTime);
  sheet.getRange(`B${row}`).setValue(status)
    .setFontColor(status === "ERROR" ? BRAND.statusRed : BRAND.statusGreen)
    .setFontWeight(status === "ERROR" ? "bold" : "normal");
  sheet.getRange(`C${row}`).setValue(duration);
  sheet.getRange(`D${row}`).setValue(error);

  if (row > 502) sheet.deleteRow(2);
}

// --- SETUP FUNCTIONS ---

function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "refreshDashboard") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger("refreshDashboard").timeBased().everyMinutes(5).create();
  Logger.log("5-minute trigger created for refreshDashboard");
}

function initializeSheet() {
  loadThresholds();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  getOrCreateSheet(ss, "Dashboard");
  getOrCreateSheet(ss, "Zendesk Raw");
  getOrCreateSheet(ss, "Aircall Raw");
  setupThresholdsTab(ss);
  setupAgentMapTab(ss);
  getOrCreateSheet(ss, "PostCall Log");
  getOrCreateSheet(ss, "SMS Log");
  getOrCreateSheet(ss, "Run Log");

  // Hide raw data tabs
  const zRaw = ss.getSheetByName("Zendesk Raw");
  const aRaw = ss.getSheetByName("Aircall Raw");
  const pcLog = ss.getSheetByName("PostCall Log");
  const smsLog = ss.getSheetByName("SMS Log");
  if (zRaw) zRaw.hideSheet();
  if (aRaw) aRaw.hideSheet();
  if (pcLog) pcLog.hideSheet();
  if (smsLog) smsLog.hideSheet();

  // Dashboard first
  const dash = ss.getSheetByName("Dashboard");
  if (dash) { ss.setActiveSheet(dash); ss.moveActiveSheet(1); }

  // Remove default Sheet1
  const sheet1 = ss.getSheetByName("Sheet1");
  if (sheet1 && ss.getSheets().length > 1) ss.deleteSheet(sheet1);

  Logger.log("Sheet initialized — running first refresh...");
  refreshDashboard();
}

// --- UNIFIED WEBHOOK RECEIVER ---
// Handles both PostCall (survey) and Aircall (SMS) webhooks via the same web app URL.
// Deploy: Deploy > New deployment > Web app > Execute as "me", access "anyone"
// Paste the /exec URL into both PostCall and Aircall webhook settings.

function doPost(e) {
  try {
    const ss = SpreadsheetApp.openById("1db-1Zlny6ryoAYc4CCkjXPtgXyGCGBEgWBfg5xnq7rU");
    const payload = JSON.parse(e.postData.contents);

    // Route based on payload shape:
    //   Meta webhooks use { object: "instagram", entry: [...] }
    //   Aircall webhooks use { event: "message.received", data: {...} }
    if (payload.object === "instagram" || payload.object === "page") {
      return handleInstagramDM(ss, payload);
    }

    const event = payload.event || "";
    if (event.startsWith("message.") || event.startsWith("group_message.")) {
      return handleAircallSMS(ss, payload);
    } else {
      return handlePostCallWebhook(ss, payload);
    }
  } catch (err) {
    Logger.log("Webhook error: " + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// --- INSTAGRAM DM WEBHOOK HANDLER ---
// Meta sends Instagram webhooks in TWO possible formats:
//   Format A (changes): { object: "instagram", entry: [{ id, time, changes: [{ field: "messages", value: { sender, recipient, timestamp, message } }] }] }
//   Format B (messaging): { object: "instagram", entry: [{ id, time, messaging: [{ sender, recipient, timestamp, message }] }] }
// We handle both.
function handleInstagramDM(ss, payload) {
  const timestamp = new Date();

  // Log raw payload to debug sheet
  const debugSheet = getOrCreateSheet(ss, "IG DM Debug");
  const debugRow = Math.min(debugSheet.getLastRow() + 1, 50);
  debugSheet.getRange(debugRow, 1).setValue(timestamp);
  debugSheet.getRange(debugRow, 2).setValue(JSON.stringify(payload).substring(0, 50000));

  const igSheet = getOrCreateSheet(ss, "IG DM Log");

  // Ensure headers exist
  if (igSheet.getLastRow() === 0) {
    igSheet.appendRow(["Timestamp", "Sender ID", "Sender Name", "Recipient ID", "Message", "Message ID", "Is Echo", "Direction"]);
    igSheet.getRange("1:1").setFontWeight("bold").setBackground(BRAND.beigeLight);
  }

  // Our IG-scoped user ID — messages FROM this ID are outbound (agent replies)
  const props = PropertiesService.getScriptProperties();
  const ourIgId = props.getProperty("META_IG_USER_ID") || "";

  const entries = payload.entry || [];
  let rowsAdded = 0;

  // Load hidden senders (pass ss since we're in webhook context, not UI)
  const hiddenSenders = getHiddenIGSenders(ss);

  // Spam filter — messages matching these patterns are logged to debug but skipped from IG DM Log.
  // Case-insensitive partial match. Add new patterns as spam evolves.
  const SPAM_PATTERNS = [
    "followers instantly",
    "skyrocket your social",
    "buy followers",
    "10k-100k",
    "shoutout for shoutout",
    "get verified now",
  ];

  function isSpam(text) {
    if (!text) return false;
    const lower = text.toLowerCase();
    return SPAM_PATTERNS.some(p => lower.includes(p));
  }

  // Helper: process a single message event object (same shape from both formats)
  function processMessage(evt) {
    const senderId = (evt.sender && evt.sender.id) || "";
    const recipientId = (evt.recipient && evt.recipient.id) || "";
    const message = evt.message || {};
    const messageId = message.mid || "";
    const messageText = message.text || "";
    const isEcho = message.is_echo || false;

    // Skip spam (inbound only — never filter our own outbound replies)
    if (!isEcho && isSpam(messageText)) {
      Logger.log("IG DM spam filtered: " + messageText.substring(0, 80));
      return;
    }

    // Skip messages from hidden senders (inbound only)
    if (!isEcho && hiddenSenders.has(senderId)) {
      Logger.log("IG DM hidden sender skipped: " + senderId);
      return;
    }

    // Also detect outbound by checking if sender matches our IG user ID
    const isOutbound = isEcho || (ourIgId && senderId === ourIgId);
    const direction = isOutbound ? "outbound" : "inbound";

    const senderName = isOutbound ? "Deako" : senderId;

    // Deduplicate by message ID
    if (messageId && igSheet.getLastRow() > 1) {
      const existingMids = igSheet.getRange(2, 6, igSheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < existingMids.length; i++) {
        if (String(existingMids[i][0]) === String(messageId)) {
          return; // skip duplicate
        }
      }
    }

    // Try to resolve sender name from IG profile (only for inbound / customer messages)
    let resolvedName = senderName;
    if (!isOutbound && senderId) {
      const igToken = props.getProperty("META_IG_TOKEN");
      if (igToken) {
        try {
          const profileUrl = `https://graph.instagram.com/v25.0/${senderId}?fields=name,username&access_token=${igToken}`;
          const resp = UrlFetchApp.fetch(profileUrl, { muteHttpExceptions: true });
          if (resp.getResponseCode() === 200) {
            const profile = JSON.parse(resp.getContentText());
            resolvedName = profile.username || profile.name || senderId;
          }
        } catch (e) {
          Logger.log("IG profile lookup failed for " + senderId + ": " + e.toString());
        }
      }
    }

    igSheet.appendRow([
      timestamp, senderId, resolvedName, recipientId,
      messageText, messageId, isOutbound, direction
    ]);
    rowsAdded++;
  }

  entries.forEach(entry => {
    // Format A: entry.changes[] (Instagram webhook standard format)
    const changes = entry.changes || [];
    changes.forEach(change => {
      if (change.field === "messages" && change.value) {
        processMessage(change.value);
      }
    });

    // Format B: entry.messaging[] (Messenger-style format, may also be used)
    // Only process events that contain a message — skip read receipts, reactions, etc.
    const messagingEvents = entry.messaging || [];
    messagingEvents.forEach(evt => {
      if (evt.message) {
        processMessage(evt);
      } else {
        Logger.log("IG webhook skipped non-message event: " + (evt.read ? "read_receipt" : evt.reaction ? "reaction" : "other"));
      }
    });
  });

  // Keep log manageable — trim to last 500 rows
  const totalRows = igSheet.getLastRow();
  if (totalRows > 501) {
    igSheet.deleteRows(2, totalRows - 501);
  }

  Logger.log("IG DM webhook processed: " + rowsAdded + " messages logged");
  return ContentService.createTextOutput(JSON.stringify({ status: "ok", messages_logged: rowsAdded }))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- AIRCALL SMS WEBHOOK HANDLER ---
function handleAircallSMS(ss, payload) {
  const event = payload.event || "unknown";
  const d = payload.data || {};
  const timestamp = new Date();

  // Log raw payload to debug sheet
  const debugSheet = getOrCreateSheet(ss, "SMS Debug");
  const debugRow = Math.min(debugSheet.getLastRow() + 1, 50);
  debugSheet.getRange(debugRow, 1).setValue(timestamp);
  debugSheet.getRange(debugRow, 2).setValue(JSON.stringify(payload).substring(0, 50000));

  // Extract SMS fields from real Aircall webhook payload
  // Events: message.sent, message.received, message.status_updated
  const messageId = d.id || "";
  const direction = event === "message.received" ? "inbound" : "outbound";
  const body = d.body || "";
  const status = d.status || "";

  // Contact info: data.contact.first_name / last_name
  const contactFirst = (d.contact && d.contact.first_name) || "";
  const contactLast = (d.contact && d.contact.last_name) || "";
  const contactName = (contactFirst + " " + contactLast).trim() || "";
  const contactPhoneRaw = d.external_number || "";
  const contactPhone = contactPhoneRaw ? "'" + contactPhoneRaw : "";  // apostrophe prefix prevents Sheets formula interpretation
  const contactEmail = (d.contact && d.contact.emails
    && d.contact.emails[0] && d.contact.emails[0].value) || "";

  // Agent / user info (present on message.sent, absent on message.received)
  const agentName = (d.user && d.user.name) || "";
  const agentEmail = (d.user && d.user.email) || "";

  // Aircall line info: data.number.name / digits
  const lineName = (d.number && d.number.name) || "";
  const lineNumberRaw = (d.number && d.number.digits) || "";
  const lineNumber = lineNumberRaw ? "'" + lineNumberRaw : "";  // apostrophe prevents Sheets formula interpretation

  const smsSheet = getOrCreateSheet(ss, "SMS Log");

  // Ensure headers exist
  if (smsSheet.getLastRow() === 0) {
    smsSheet.appendRow(["Timestamp", "Event", "Direction", "Agent", "Contact Name", "Contact Phone", "Contact Email", "Message", "Status", "Aircall Line", "Line Number", "Message ID"]);
    smsSheet.getRange("1:1").setFontWeight("bold").setBackground(BRAND.beigeLight);
  }

  // Deduplicate — Aircall sometimes sends the same webhook twice.
  // Check if this message ID + event combo already exists in the log.
  if (messageId && smsSheet.getLastRow() > 1) {
    const existingIds = smsSheet.getRange(2, 12, smsSheet.getLastRow() - 1, 1).getValues(); // column 12 = Message ID
    const existingEvents = smsSheet.getRange(2, 2, smsSheet.getLastRow() - 1, 1).getValues(); // column 2 = Event
    for (let i = 0; i < existingIds.length; i++) {
      if (String(existingIds[i][0]) === String(messageId) && String(existingEvents[i][0]) === event) {
        // Duplicate — skip logging, still return ok
        return ContentService.createTextOutput(JSON.stringify({ status: "ok", deduplicated: true }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
  }

  // Append the SMS event
  smsSheet.appendRow([timestamp, event, direction, agentName, contactName, contactPhone, contactEmail, body, status, lineName, lineNumber, messageId]);

  // Keep log manageable — trim to last 1000 rows
  const totalRows = smsSheet.getLastRow();
  if (totalRows > 1001) {
    smsSheet.deleteRows(2, totalRows - 1001);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- POSTCALL WEBHOOK HANDLER ---
function handlePostCallWebhook(ss, payload) {
  const sheet = getOrCreateSheet(ss, "PostCall Log");

  // Log raw payload to debug sheet
  const debugSheet = getOrCreateSheet(ss, "PostCall Debug");
  const debugRow = Math.min(debugSheet.getLastRow() + 1, 50);
  debugSheet.getRange(debugRow, 1).setValue(new Date());
  debugSheet.getRange(debugRow, 2).setValue(JSON.stringify(payload).substring(0, 50000));

  const event = payload.event || "unknown";
  const d = payload.data || {};
  const timestamp = new Date();

  // Extract fields from PostCall payload
  const agentName = (d.call && d.call.agent && d.call.agent.name) || "";
  const customerName = (d.contact && d.contact.name) || "";
  const customerPhone = (d.contact && d.contact.phone_numbers
    && d.contact.phone_numbers[0] && d.contact.phone_numbers[0].number) || "";
  const callId = (d.call && d.call.external_id) || "";
  const callDuration = (d.call && d.call.duration) || "";
  const surveyUrl = d.url || "";
  const answeredAt = d.answered_at || "";

  // Parse answers array (survey-completed events)
  const CSAT_MAP = { "great": 5, "good": 4, "okay": 3, "bad": 2, "terrible": 1 };
  let csatScore = "";
  let npsScore = "";
  let comment = "";
  const answers = d.answers || [];
  answers.forEach(a => {
    if (a.question_type === "csat-5-emoji" && a.answer) {
      csatScore = CSAT_MAP[a.answer.toLowerCase()] || a.answer;
    } else if (a.question_type === "nps" && a.answer) {
      npsScore = Number(a.answer) || a.answer;
    } else if (a.question_type === "longtext" && a.answer) {
      comment = a.answer;
    }
  });

  // Ensure headers exist
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Event", "CSAT (1-5)", "NPS (0-10)", "Agent", "Customer", "Phone", "Call ID", "Duration (s)", "Comment", "Answered At", "Survey URL"]);
    sheet.getRange("1:1").setFontWeight("bold").setBackground(BRAND.beigeLight);
  }

  sheet.appendRow([timestamp, event, csatScore, npsScore, agentName, customerName, customerPhone, callId, callDuration, comment, answeredAt, surveyUrl]);

  // Keep log manageable — trim to last 500 rows
  const totalRows = sheet.getLastRow();
  if (totalRows > 501) {
    sheet.deleteRows(2, totalRows - 501);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "ok" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// Handle GET — serves two purposes:
// 1. PostCall may ping the URL to verify
// 2. Meta webhook verification: GET with hub.mode=subscribe, hub.verify_token, hub.challenge
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  // Meta webhook verification challenge
  if (params["hub.mode"] === "subscribe" && params["hub.verify_token"]) {
    const props = PropertiesService.getScriptProperties();
    const expectedToken = props.getProperty("META_WEBHOOK_VERIFY_TOKEN") || "";
    if (params["hub.verify_token"] === expectedToken) {
      // Return the challenge value as plain text (Meta requires this exact format)
      return ContentService.createTextOutput(params["hub.challenge"]);
    } else {
      Logger.log("Meta webhook verify failed — token mismatch");
      return ContentService.createTextOutput("Forbidden").setMimeType(ContentService.MimeType.TEXT);
    }
  }

  // Default response for PostCall or other pings
  return ContentService.createTextOutput(JSON.stringify({ status: "ok", service: "CS Command Center Webhook" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- REPROCESS POSTCALL DEBUG DATA ---
// Run this once to re-parse raw debug payloads into the PostCall Log with correct field mapping.
// Safe to run multiple times — it clears PostCall Log first.
// --- DEBUG: ZENDESK AUDIT ---
// Run this to dump all Zendesk metrics to a "Zendesk Audit" sheet.
// Shows: search query counts, view ticket list, per-agent breakdown, SAS detection.
// Compare against Zendesk UI to verify dashboard accuracy.
function debugZendesk() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("ZENDESK_TOKEN");
  if (!token) throw new Error("ZENDESK_TOKEN not set");

  const subdomain = CONFIG.zendesk.subdomain;
  const viewId = CONFIG.zendesk.viewId;
  const authHeader = "Basic " + Utilities.base64Encode(token);
  const headers = { "Authorization": authHeader, "Content-Type": "application/json" };
  const fetchOpts = { method: "get", headers: headers, muteHttpExceptions: true };

  function searchCount(query) {
    try {
      const url = `https://${subdomain}.zendesk.com/api/v2/search/count.json?query=${encodeURIComponent(query)}`;
      const resp = UrlFetchApp.fetch(url, fetchOpts);
      if (resp.getResponseCode() === 200) return JSON.parse(resp.getContentText()).count || 0;
    } catch (e) { Logger.log("Search failed: " + query); }
    return 0;
  }

  let sheet = ss.getSheetByName("Zendesk Audit");
  if (!sheet) sheet = ss.insertSheet("Zendesk Audit");
  sheet.clear();

  let row = 1;

  // ─── PART 1: Search query counts (what powers the dashboard KPIs) ───
  sheet.getRange(row, 1, 1, 3).setValues([["SEARCH QUERY AUDIT", "Query", "Count"]]).setFontWeight("bold");
  row++;

  const queries = [
    ["Open tickets", "type:ticket status:open"],
    ["New tickets", "type:ticket status:new"],
    ["Open + New", "type:ticket status<solved status>pending"],
    ["All unsolved (new/open/pending)", "type:ticket status<solved"],
    ["On Hold", "type:ticket status:hold"],
    ["Unassigned (all)", "type:ticket status<solved assignee:none"],
    ["SAS by tag (new)", "type:ticket status:new tags:sas_flex"],
    ["SAS by subject (new)", 'type:ticket status:new subject:"You have a new call from SAS Flex"'],
    ["SAS by requester (new)", "type:ticket status:new requester:notifications@sasdesk.com"],
    ["SAS by tag (all unsolved)", "type:ticket status<solved tags:sas_flex"],
    ["SAS by requester (all unsolved)", "type:ticket status<solved requester:notifications@sasdesk.com"],
  ];

  // Add per-agent solved today
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  CONFIG.agents.forEach(agent => {
    queries.push([`Solved today: ${agent}`, `type:ticket solved>=${todayStr} assignee:"${agent}"`]);
  });
  queries.push(["Solved today: TOTAL", `type:ticket solved>=${todayStr}`]);

  queries.forEach(q => {
    const count = searchCount(q[1]);
    sheet.getRange(row, 1).setValue(q[0]);
    sheet.getRange(row, 2).setValue(q[1]).setFontColor("#666666");
    sheet.getRange(row, 3).setValue(count).setFontWeight("bold");
    row++;
  });

  row++;

  // ─── PART 2: View tickets (what the dashboard iterates over) ───
  sheet.getRange(row, 1, 1, 2).setValues([["VIEW TICKETS (ID: " + viewId + ")", ""]]).setFontWeight("bold");
  row++;

  const viewUrl = `https://${subdomain}.zendesk.com/api/v2/views/${viewId}/tickets.json?per_page=100&include=users`;
  const viewResp = UrlFetchApp.fetch(viewUrl, fetchOpts);
  const viewData = JSON.parse(viewResp.getContentText());
  const tickets = viewData.tickets || [];

  const userMap = {};
  const userEmailMap = {};
  if (viewData.users) {
    viewData.users.forEach(u => {
      userMap[u.id] = u.name || u.email || "Unknown";
      userEmailMap[u.id] = u.email || "";
    });
  }

  const viewHeaders = ["ID", "Subject", "Requester", "Req Email", "Assignee", "Status", "Priority", "Created", "Updated", "Tags", "SAS?"];
  sheet.getRange(row, 1, 1, viewHeaders.length).setValues([viewHeaders]).setFontWeight("bold").setBackground(BRAND.beigeLight);
  row++;

  tickets.forEach(t => {
    const reqName = t.requester_id ? (userMap[t.requester_id] || "?") : "?";
    const reqEmail = t.requester_id ? (userEmailMap[t.requester_id] || "") : "";
    const assignee = t.assignee_id ? (userMap[t.assignee_id] || "?") : "Unassigned";
    const tags = (t.tags || []).join(", ");
    const isSAS = (t.subject && t.subject.toLowerCase().includes("you have a new call from sas flex"))
      || (reqEmail.toLowerCase() === "notifications@sasdesk.com")
      || (t.tags && t.tags.includes("sas_flex"));

    sheet.getRange(row, 1).setValue(t.id);
    sheet.getRange(row, 2).setValue(t.subject || "");
    sheet.getRange(row, 3).setValue(reqName);
    sheet.getRange(row, 4).setValue(reqEmail);
    sheet.getRange(row, 5).setValue(assignee);
    sheet.getRange(row, 6).setValue(t.status);
    sheet.getRange(row, 7).setValue(t.priority || "normal");
    sheet.getRange(row, 8).setValue(t.created_at);
    sheet.getRange(row, 9).setValue(t.updated_at);
    sheet.getRange(row, 10).setValue(tags);
    sheet.getRange(row, 11).setValue(isSAS ? "YES" : "");
    if (isSAS) sheet.getRange(row, 11).setFontColor("#BA866A").setFontWeight("bold");
    row++;
  });

  row++;

  // ─── PART 3: Summary comparison ───
  sheet.getRange(row, 1, 1, 2).setValues([["SUMMARY", ""]]).setFontWeight("bold");
  row++;

  const excludeLower = (CONFIG.excludeAgents || []).map(n => n.toLowerCase());
  const filtered = tickets.filter(t => {
    const assignee = t.assignee_id ? (userMap[t.assignee_id] || "") : "";
    if (!assignee || assignee === "Unassigned") return true;
    return !excludeLower.some(ex => assignee.toLowerCase().includes(ex));
  });

  const viewSAS = filtered.filter(t =>
    t.status === "new" && (
      (t.subject && t.subject.toLowerCase().includes("you have a new call from sas flex"))
      || ((t.requester_id ? (userEmailMap[t.requester_id] || "") : "").toLowerCase() === "notifications@sasdesk.com")
      || (t.tags && t.tags.includes("sas_flex"))
    )
  ).length;

  const summaryRows = [
    ["Tickets in view (raw)", tickets.length],
    ["Tickets in view (after exclude filter)", filtered.length],
    ["View: status=new", filtered.filter(t => t.status === "new").length],
    ["View: status=open", filtered.filter(t => t.status === "open").length],
    ["View: unassigned", filtered.filter(t => !t.assignee_id).length],
    ["View: SAS (new only)", viewSAS],
    ["", ""],
    ["Per-agent assigned (from view):", ""],
  ];

  CONFIG.agents.forEach(agent => {
    const count = filtered.filter(t => {
      const assignee = t.assignee_id ? (userMap[t.assignee_id] || "") : "";
      const parts = agent.toLowerCase().split(/\s+/);
      const aLower = assignee.toLowerCase();
      return parts.some(p => p.length > 1 && aLower.split(/\s+/).some(ap => ap === p));
    }).length;
    summaryRows.push(["  " + agent, count]);
  });
  summaryRows.push(["  Unassigned", filtered.filter(t => !t.assignee_id).length]);

  summaryRows.forEach(r => {
    sheet.getRange(row, 1).setValue(r[0]);
    sheet.getRange(row, 2).setValue(r[1]).setFontWeight("bold");
    row++;
  });

  SpreadsheetApp.flush();
  Logger.log("Zendesk Audit complete: " + tickets.length + " tickets in view, " + queries.length + " search queries run.");
}

// --- DEBUG: FORWARDED CALL CLASSIFICATION ---
// Run this to dump raw Nicereply API responses to a "Nicereply Debug" sheet.
// Shows the full answers array structure so we can see exact field names and scales.
function debugNicereply() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty("NICEREPLY_TOKEN");
  if (!token) throw new Error("NICEREPLY_TOKEN not set");

  const auth = "Basic " + Utilities.base64Encode(token);
  const headers = { "Authorization": auth };
  const fetchOpts = { method: "get", headers: headers, muteHttpExceptions: true };

  // Fetch last 7 days of responses for a good sample
  const since = new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
  const sinceISO = since.toISOString().replace(/\.\d{3}Z$/, "Z");
  const url = `https://api.nicereply.com/responses?created_after=${encodeURIComponent(sinceISO)}&per_page=50&page=1`;
  const resp = UrlFetchApp.fetch(url, fetchOpts);

  const sheet = ss.getSheetByName("Nicereply Debug") || ss.insertSheet("Nicereply Debug");
  sheet.clear();
  let row = 1;

  // API response metadata
  sheet.getRange(row, 1).setValue("URL").setFontWeight("bold");
  sheet.getRange(row, 2, 1, 4).merge().setValue(url);
  row++;
  sheet.getRange(row, 1).setValue("Status").setFontWeight("bold");
  sheet.getRange(row, 2).setValue(resp.getResponseCode());
  row += 2;

  if (resp.getResponseCode() !== 200) {
    sheet.getRange(row, 1).setValue("Error body:");
    sheet.getRange(row + 1, 1, 1, 6).merge().setValue(resp.getContentText().substring(0, 1000));
    return;
  }

  const body = JSON.parse(resp.getContentText());
  const responses = body.data || [];

  // Pagination info
  sheet.getRange(row, 1).setValue("Pagination").setFontWeight("bold");
  sheet.getRange(row, 2, 1, 4).merge().setValue(JSON.stringify(body.pagination || {}));
  row++;
  sheet.getRange(row, 1).setValue("Responses count").setFontWeight("bold");
  sheet.getRange(row, 2).setValue(responses.length);
  row += 2;

  // Raw JSON of first response (full structure)
  if (responses.length > 0) {
    sheet.getRange(row, 1).setValue("RAW FIRST RESPONSE (full JSON)").setFontWeight("bold");
    row++;
    sheet.getRange(row, 1, 1, 8).merge().setValue(JSON.stringify(responses[0], null, 2)).setWrap(true);
    row += 2;
  }

  // Table header
  sheet.getRange(row, 1).setValue("#").setFontWeight("bold");
  sheet.getRange(row, 2).setValue("created_at").setFontWeight("bold");
  sheet.getRange(row, 3).setValue("from").setFontWeight("bold");
  sheet.getRange(row, 4).setValue("ticket_id").setFontWeight("bold");
  sheet.getRange(row, 5).setValue("answers (raw JSON)").setFontWeight("bold");
  sheet.getRange(row, 6).setValue("top-level keys").setFontWeight("bold");
  row++;

  // Each response — dump answers array as raw JSON so we can see the exact structure
  responses.forEach((r, idx) => {
    sheet.getRange(row, 1).setValue(idx + 1);
    sheet.getRange(row, 2).setValue(r.created_at || "");
    sheet.getRange(row, 3).setValue(r.from || r.email || "");
    sheet.getRange(row, 4).setValue(r.ticket_id || "");
    sheet.getRange(row, 5).setValue(JSON.stringify(r.answers || [])).setWrap(true);
    sheet.getRange(row, 6).setValue(Object.keys(r).join(", "));
    row++;
  });

  // Auto-size
  sheet.autoResizeColumns(1, 6);
  SpreadsheetApp.flush();
  Logger.log("Nicereply debug complete — " + responses.length + " responses dumped");
}

// Run this to dump all of today's inbound support calls with their classification
// to a "Call Debug" sheet. Compare against what you see in Aircall to find mismatches.
function debugForwardedCalls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const apiId = props.getProperty("AIRCALL_API_ID");
  const apiToken = props.getProperty("AIRCALL_API_TOKEN");
  const baseUrl = CONFIG.aircall.baseUrl;
  const auth = "Basic " + Utilities.base64Encode(apiId + ":" + apiToken);

  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const fromTs = Math.floor(startOfDay.getTime() / 1000);
  const toTs = Math.floor(now.getTime() / 1000);

  // Fetch all calls today
  let allCalls = [];
  let page = 1;
  let hasMore = true;
  while (hasMore && page <= 10) {
    const url = `${baseUrl}/calls?from=${fromTs}&to=${toTs}&per_page=50&page=${page}&order=desc`;
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "Authorization": auth, "Content-Type": "application/json" },
      muteHttpExceptions: true,
    });
    if (response.getResponseCode() !== 200) break;
    const data = JSON.parse(response.getContentText());
    allCalls = allCalls.concat(data.calls || []);
    hasMore = data.meta && data.meta.next_page_link;
    page++;
  }

  // Filter to support lines + inbound
  const supportNumbers = CONFIG.aircall.supportNumbers;
  const supportCalls = allCalls.filter(c => {
    if (!c.number) return false;
    const digits = (c.number.digits || "").replace(/[\s\-\(\)]/g, "");
    return supportNumbers.some(sn => digits.includes(sn.replace(/[\s\-\(\)]/g, "")) || sn.replace(/[\s\-\(\)]/g, "").includes(digits));
  });
  const inboundCalls = supportCalls.filter(c => c.direction === "inbound");

  // Classify each call
  function matchAgent(call) {
    const user = call.user;
    if (!user) return null;
    const agentName = (user.name || `${user.first_name || ""} ${user.last_name || ""}`.trim());
    return CONFIG.agents.find(a => {
      const parts = a.toLowerCase().split(" ");
      const callParts = agentName.toLowerCase().split(" ");
      return parts[0] === callParts[0] || (parts[1] && callParts[1] && parts[1] === callParts[1]);
    });
  }

  const rows = inboundCalls.map(call => {
    const reason = call.missed_call_reason || "";
    const userName = call.user ? (call.user.name || `${call.user.first_name || ""} ${call.user.last_name || ""}`) : "NONE";
    const matched = matchAgent(call);
    const lineName = call.number ? call.number.name : "";

    let classification = "???";
    if (reason === "short_abandoned") {
      classification = "SHORT_ABANDONED";
    } else if (reason === "agents_did_not_answer" || reason === "agents_did_not_pick_up"
            || reason === "no_available_agent" || reason === "no_agent_available") {
      classification = "FORWARDED";
    } else if (call.answered_at && call.user) {
      classification = matched ? `TEAM_ANSWERED (${matched})` : `FORWARDED (unknown_agent: ${userName})`;
    } else if (reason) {
      classification = `FORWARDED (${reason})`;
    } else {
      classification = "SKIPPED (no reason, no answered_at)";
    }

    const callTime = call.started_at ? new Date(call.started_at * 1000) : "";
    const callerNum = call.raw_digits || "";
    const contactName = call.contact ? `${call.contact.first_name || ""} ${call.contact.last_name || ""}`.trim() : "";

    return [
      callTime,
      call.id,
      lineName,
      callerNum,
      contactName,
      call.direction,
      call.status,
      reason || "(empty)",
      call.answered_at ? "YES" : "NO",
      userName,
      matched || "(no match)",
      classification,
      call.duration || 0,
      call.waiting_duration || 0,
    ];
  });

  // Write to Call Debug sheet
  let sheet = ss.getSheetByName("Call Debug");
  if (!sheet) {
    sheet = ss.insertSheet("Call Debug");
  }
  sheet.clear();

  const headers = ["Time", "Call ID", "Line", "Caller #", "Contact", "Direction", "Status",
                   "missed_call_reason", "answered_at?", "User Name", "Agent Match", "Classification",
                   "Duration (s)", "Wait (s)"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  // Summary at bottom
  const gap = rows.length + 3;
  const teamCount = rows.filter(r => r[11].startsWith("TEAM_ANSWERED")).length;
  const fwdCount = rows.filter(r => r[11].startsWith("FORWARDED")).length;
  const shortCount = rows.filter(r => r[11].startsWith("SHORT_ABANDONED")).length;
  const skipCount = rows.filter(r => r[11].startsWith("SKIPPED")).length;
  const unknownCount = rows.filter(r => r[11] === "???").length;

  sheet.getRange(gap, 1, 6, 2).setValues([
    ["SUMMARY", ""],
    ["Team Answered", teamCount],
    ["Forwarded (SAS)", fwdCount],
    ["Short Abandoned", shortCount],
    ["Skipped", skipCount],
    ["Unknown (???)", unknownCount],
  ]);

  // ─── PART 2: Outbound calls TO the SAS number (the actual forwarding mechanism) ───
  const answerSvcNum = (CONFIG.aircall.answeringServiceNumber || "").replace(/[\s\-\(\)]/g, "");
  const outboundToSAS = allCalls.filter(c => {
    if (c.direction !== "outbound") return false;
    const rawDigits = (c.raw_digits || "").replace(/[\s\-\(\)]/g, "");
    return answerSvcNum && (rawDigits.includes(answerSvcNum.replace("+1", "")) || answerSvcNum.includes(rawDigits));
  });

  const sasGap = gap + 8;
  sheet.getRange(sasGap, 1, 1, 2).setValues([
    ["OUTBOUND CALLS TO SAS NUMBER", `Count: ${outboundToSAS.length}`],
  ]).setFontWeight("bold");

  if (outboundToSAS.length > 0) {
    const sasHeaders = ["Time", "Call ID", "Line", "raw_digits", "Contact Name", "Contact Phone",
                        "participants", "transferred_from", "comments", "tags", "All Top-Level Keys"];
    sheet.getRange(sasGap + 1, 1, 1, sasHeaders.length).setValues([sasHeaders]).setFontWeight("bold");

    const sasRows = outboundToSAS.map(call => {
      const callTime = call.started_at ? new Date(call.started_at * 1000) : "";
      const lineName = call.number ? call.number.name : "";
      const contactName = call.contact
        ? `${call.contact.first_name || ""} ${call.contact.last_name || ""}`.trim()
        : "NO CONTACT";
      const contactPhone = call.contact && call.contact.phone_numbers
        ? call.contact.phone_numbers.map(p => p.value).join(", ")
        : "";
      const participants = JSON.stringify(call.participants || call.teams || []).substring(0, 200);
      const transferred = call.transferred_from || call.transfer_from || call.transferred_to || "";
      const comments = (call.comments || []).map(c => c.body || "").join("; ").substring(0, 200);
      const tags = (call.tags || []).map(t => t.name || t).join(", ");
      const allKeys = Object.keys(call).join(", ");
      return [
        callTime, call.id, lineName, call.raw_digits || "",
        contactName, contactPhone,
        participants, JSON.stringify(transferred),
        comments, tags, allKeys,
      ];
    });
    sheet.getRange(sasGap + 2, 1, sasRows.length, sasHeaders.length).setValues(sasRows);
  }

  // ─── PART 3: ALL other outbound calls (to see full picture) ───
  const allOutbound = allCalls.filter(c => c.direction === "outbound");
  const outGap = sasGap + (outboundToSAS.length > 0 ? outboundToSAS.length + 4 : 3);
  sheet.getRange(outGap, 1, 1, 2).setValues([
    ["ALL OUTBOUND CALLS TODAY", `Count: ${allOutbound.length}`],
  ]).setFontWeight("bold");

  if (allOutbound.length > 0) {
    const outHeaders = ["Time", "Call ID", "Line", "Dialed #", "Status", "missed_call_reason",
                        "answered_at?", "User Name", "Duration (s)"];
    sheet.getRange(outGap + 1, 1, 1, outHeaders.length).setValues([outHeaders]).setFontWeight("bold");

    const outRows = allOutbound.map(call => {
      const callTime = call.started_at ? new Date(call.started_at * 1000) : "";
      const userName = call.user ? (call.user.name || `${call.user.first_name || ""} ${call.user.last_name || ""}`) : "NONE";
      const lineName = call.number ? call.number.name : "";
      return [
        callTime, call.id, lineName, call.raw_digits || "",
        call.status, call.missed_call_reason || "(empty)",
        call.answered_at ? "YES" : "NO", userName,
        call.duration || 0,
      ];
    });
    sheet.getRange(outGap + 2, 1, outRows.length, outHeaders.length).setValues(outRows);
  }

  // ─── PART 4: Fetch first SAS call individually to see ALL available fields ───
  if (outboundToSAS.length > 0) {
    const testCallId = outboundToSAS[0].id;
    const detailGap = outGap + (allOutbound.length > 0 ? allOutbound.length + 4 : 3);
    try {
      const detailUrl = `${baseUrl}/calls/${testCallId}`;
      const detailResp = UrlFetchApp.fetch(detailUrl, {
        method: "get",
        headers: { "Authorization": auth, "Content-Type": "application/json" },
        muteHttpExceptions: true,
      });
      const detailData = JSON.parse(detailResp.getContentText());
      const call = detailData.call || detailData;

      sheet.getRange(detailGap, 1, 1, 3).setValues([
        ["SINGLE CALL DETAIL (ID: " + testCallId + ")", "Response Code: " + detailResp.getResponseCode(), ""],
      ]).setFontWeight("bold");

      // Dump all fields as key=value pairs
      let dumpRow = detailGap + 1;
      const dumpFields = (obj, prefix) => {
        Object.keys(obj).forEach(key => {
          const val = obj[key];
          if (val !== null && typeof val === "object" && !Array.isArray(val)) {
            dumpFields(val, prefix + key + ".");
          } else {
            const display = Array.isArray(val)
              ? JSON.stringify(val).substring(0, 500)
              : String(val).substring(0, 500);
            sheet.getRange(dumpRow, 1).setValue(prefix + key);
            sheet.getRange(dumpRow, 2, 1, 3).merge().setValue(display);
            dumpRow++;
          }
        });
      };
      dumpFields(call, "");
    } catch (err) {
      sheet.getRange(detailGap, 1).setValue("Error fetching call detail: " + err.message);
    }
  }

  SpreadsheetApp.flush();
  Logger.log(`Call Debug: ${rows.length} inbound. Team=${teamCount}, Fwd=${fwdCount}, Short=${shortCount}. Outbound to SAS=${outboundToSAS.length}. Total outbound=${allOutbound.length}`);
}

// --- REPROCESS SMS DEBUG DATA ---
// Run once to re-parse raw SMS debug payloads into SMS Log with correct field mapping.
function reprocessSMSDebug() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const debugSheet = ss.getSheetByName("SMS Debug");
  const logSheet = getOrCreateSheet(ss, "SMS Log");

  if (!debugSheet || debugSheet.getLastRow() < 1) {
    Logger.log("No SMS Debug data to reprocess.");
    return;
  }

  logSheet.clear();
  logSheet.appendRow(["Timestamp", "Event", "Direction", "Agent", "Contact Name", "Contact Phone", "Contact Email", "Message", "Status", "Aircall Line", "Line Number", "Message ID"]);
  logSheet.getRange("1:1").setFontWeight("bold").setBackground(BRAND.beigeLight);

  const rows = debugSheet.getRange(1, 1, debugSheet.getLastRow(), 2).getValues();
  let count = 0;
  const seen = new Set(); // deduplicate by messageId + event

  rows.forEach(row => {
    try {
      const ts = row[0];
      const payload = JSON.parse(row[1]);
      const event = payload.event || "unknown";
      const d = payload.data || {};
      const messageId = d.id || "";

      // Skip duplicates
      const dedupeKey = messageId + "|" + event;
      if (messageId && seen.has(dedupeKey)) return;
      if (messageId) seen.add(dedupeKey);

      const direction = event === "message.received" ? "inbound" : "outbound";
      const contactFirst = (d.contact && d.contact.first_name) || "";
      const contactLast = (d.contact && d.contact.last_name) || "";
      const contactName = (contactFirst + " " + contactLast).trim();
      const contactPhone = d.external_number || "";
      const contactEmail = (d.contact && d.contact.emails
        && d.contact.emails[0] && d.contact.emails[0].value) || "";
      const agentName = (d.user && d.user.name) || "";
      const body = d.body || "";
      const status = d.status || "";
      const lineName = (d.number && d.number.name) || "";
      const lineNumber = (d.number && d.number.digits) || "";

      logSheet.appendRow([ts, event, direction, agentName, contactName, contactPhone, contactEmail, body, status, lineName, lineNumber, messageId]);
      count++;
    } catch (err) {
      Logger.log("Skipped SMS debug row: " + err.toString());
    }
  });

  Logger.log("Reprocessed " + count + " SMS debug entries into SMS Log (deduped).");
}

function reprocessPostCallDebug() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const debugSheet = ss.getSheetByName("PostCall Debug");
  const logSheet = getOrCreateSheet(ss, "PostCall Log");

  if (!debugSheet || debugSheet.getLastRow() < 1) {
    Logger.log("No PostCall Debug data to reprocess.");
    return;
  }

  // Clear existing log
  logSheet.clear();

  // Add headers
  logSheet.appendRow(["Timestamp", "Event", "CSAT (1-5)", "NPS (0-10)", "Agent", "Customer", "Phone", "Call ID", "Duration (s)", "Comment", "Answered At", "Survey URL"]);
  logSheet.getRange("1:1").setFontWeight("bold").setBackground(BRAND.beigeLight);

  const CSAT_MAP = { "great": 5, "good": 4, "okay": 3, "bad": 2, "terrible": 1 };
  const rows = debugSheet.getRange(1, 1, debugSheet.getLastRow(), 2).getValues();
  let count = 0;

  rows.forEach(row => {
    try {
      const ts = row[0];
      const payload = JSON.parse(row[1]);
      const event = payload.event || "unknown";
      const d = payload.data || {};

      const agentName = (d.call && d.call.agent && d.call.agent.name) || "";
      const customerName = (d.contact && d.contact.name) || "";
      const customerPhone = (d.contact && d.contact.phone_numbers
        && d.contact.phone_numbers[0] && d.contact.phone_numbers[0].number) || "";
      const callId = (d.call && d.call.external_id) || "";
      const callDuration = (d.call && d.call.duration) || "";
      const surveyUrl = d.url || "";
      const answeredAt = d.answered_at || "";

      let csatScore = "";
      let npsScore = "";
      let comment = "";
      const answers = d.answers || [];
      answers.forEach(a => {
        if (a.question_type === "csat-5-emoji" && a.answer) {
          csatScore = CSAT_MAP[a.answer.toLowerCase()] || a.answer;
        } else if (a.question_type === "nps" && a.answer) {
          npsScore = Number(a.answer) || a.answer;
        } else if (a.question_type === "longtext" && a.answer) {
          comment = a.answer;
        }
      });

      logSheet.appendRow([ts, event, csatScore, npsScore, agentName, customerName, customerPhone, callId, callDuration, comment, answeredAt, surveyUrl]);
      count++;
    } catch (err) {
      Logger.log("Skipped debug row: " + err.toString());
    }
  });

  Logger.log("Reprocessed " + count + " PostCall debug entries into PostCall Log.");
}

// --- TEST: Trigger alert banner ---
// Temporarily overrides unassigned count to test the "all hands on deck" alert banner.
// Run this from the script editor, then run refreshDashboard to see the normal view.
function testAlertBanner() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const zendesk = fetchZendeskStatus();
  const aircall = fetchAircallStatus();
  const csat = fetchNicereplyCSAT();
  const postCall = readPostCallCSAT();
  const smsData = readSMSActivity();

  // Override unassigned to trigger the alert
  zendesk.unassigned = 142;

  Logger.log("Testing alert banner with " + zendesk.unassigned + " unassigned tickets...");
  writeDashboard(ss, zendesk, aircall, csat, postCall, smsData);
  Logger.log("Alert banner test complete. Run refreshDashboard() to restore normal view.");
}

// --- READ POSTCALL DATA FOR DASHBOARD ---
function readPostCallCSAT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("PostCall Log");
  if (!sheet || sheet.getLastRow() <= 1) {
    return { score: null, total: 0, satisfied: 0, responses: [] };
  }

  const now = new Date();
  const since = new Date(now.getTime() - 24 * 60 * 60 * 1000);

  // Read all data rows (skip header)
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 12).getValues();

  // Filter to last 24 hours, only survey-completed events with CSAT scores
  // Columns: [0] Timestamp, [1] Event, [2] CSAT (1-5), [3] NPS (0-10), [4] Agent,
  //          [5] Customer, [6] Phone, [7] Call ID, [8] Duration, [9] Comment, [10] Answered At, [11] Survey URL
  const recent = [];
  data.forEach(row => {
    const ts = new Date(row[0]);
    if (ts < since) return;

    const event = String(row[1] || "");
    // Only count survey-completed events
    if (event !== "survey-completed") return;

    // Skip excluded agents (not on CS team)
    const agentName = String(row[4] || "");
    const excluded = CONFIG.excludePostCallAgents.some(ex => agentName.toLowerCase().includes(ex.toLowerCase()));
    if (excluded) return;

    const csatScore = Number(row[2]) || 0;
    const npsScore = Number(row[3]) || "";
    if (!csatScore) return;  // skip if no CSAT score

    const timeStr = Utilities.formatDate(ts, Session.getScriptTimeZone(), "h:mm a");
    const dateStr = Utilities.formatDate(ts, Session.getScriptTimeZone(), "MMM d");

    recent.push({
      score: csatScore,
      maxScore: 5,
      nps: npsScore,
      customer: String(row[5] || "Unknown"),
      phone: String(row[6] || "Unknown"),
      agent: String(row[4] || ""),
      callId: String(row[7] || ""),
      comment: String(row[9] || ""),
      timeStr,
      dateStr,
      created: ts,
      satisfied: csatScore >= 4,  // 4 or 5 out of 5 = satisfied
    });
  });

  // Sort newest first
  recent.sort((a, b) => b.created - a.created);

  const total = recent.length;
  const satisfied = recent.filter(r => r.satisfied).length;
  const csatPct = total > 0 ? Math.round((satisfied / total) * 100) : null;

  // Also compute average NPS if available
  const npsResponses = recent.filter(r => r.nps !== "");
  const avgNps = npsResponses.length > 0
    ? Math.round(npsResponses.reduce((sum, r) => sum + Number(r.nps), 0) / npsResponses.length)
    : null;

  return { score: csatPct, total, satisfied, avgNps, responses: recent };
}

// --- FETCH META BUSINESS SUITE (Facebook Messenger + Instagram DMs) ---
// Requires Script Properties: META_PAGE_TOKEN, META_PAGE_ID
// Optional: META_IG_ID (Instagram Business Account ID — for IG-specific data)
function fetchMetaStatus() {
  const props = PropertiesService.getScriptProperties();
  const pageToken = props.getProperty("META_PAGE_TOKEN");
  const pageId = props.getProperty("META_PAGE_ID");

  const emptyResult = {
    unreadDMs: 0, totalConversations: 0, recentConversations: [],
    recentComments: [], tokenWarning: null, error: null
  };

  if (!pageToken || !pageId) {
    Logger.log("META_PAGE_TOKEN or META_PAGE_ID not set — skipping Meta fetch");
    return emptyResult;
  }

  const baseUrl = "https://graph.facebook.com/v25.0";
  const fetchOpts = { method: "get", muteHttpExceptions: true };
  const now = new Date();
  const oneDayAgo = new Date(now.getTime() - 24 * 60 * 60 * 1000);

  // Helper to make Graph API calls
  function graphGet(path) {
    const separator = path.includes("?") ? "&" : "?";
    const raw = `${baseUrl}/${path}${separator}access_token=${pageToken}`;
    const url = raw.replace(/\{/g, "%7B").replace(/\}/g, "%7D");
    try {
      const resp = UrlFetchApp.fetch(url, fetchOpts);
      if (resp.getResponseCode() !== 200) {
        const errText = resp.getContentText().substring(0, 300);
        Logger.log("Meta API " + resp.getResponseCode() + ": " + errText);
        return null;
      }
      return JSON.parse(resp.getContentText());
    } catch (e) {
      Logger.log("Meta API fetch error: " + e.toString());
      return null;
    }
  }

  // ── Helper: parse a conversations response into structured items ──
  function parseConversations(data, platform) {
    if (!data || !data.data) return [];
    const items = [];
    (data.data || []).forEach(convo => {
      const unread = convo.unread_count || 0;
      const participants = (convo.participants && convo.participants.data) || [];
      const customer = participants.find(p => p.id !== pageId);
      const customerName = customer ? customer.name : "Unknown";

      const messages = (convo.messages && convo.messages.data) || [];
      const latestMsg = messages[0] || {};
      const msgText = latestMsg.message || "";
      const msgTime = latestMsg.created_time || convo.updated_time || "";
      const isFromPage = latestMsg.from && latestMsg.from.id === pageId;

      const inboxUrl = `https://business.facebook.com/latest/inbox/all?asset_id=${pageId}&thread_id=${convo.id}`;

      items.push({
        id: convo.id,
        customerName,
        unread,
        lastMessage: msgText.length > 80 ? msgText.substring(0, 80) + "..." : msgText,
        lastMessageFrom: isFromPage ? "Deako" : customerName.split(" ")[0],
        time: msgTime,
        inboxUrl,
        platform,
      });
    });
    return items;
  }

  // ─── Step 1: Messenger DMs ───
  const messengerData = graphGet(
    `${pageId}/conversations?fields=id,updated_time,unread_count,participants,messages.limit(1){message,from,created_time}&limit=25`
  );
  const messengerConvos = parseConversations(messengerData, "Messenger");

  // ─── Step 2: Instagram DMs (from webhook log sheet) ───
  // IG DMs come in via Meta webhooks → doPost() → handleInstagramDM() → "IG DM Log" sheet.
  // We read the sheet and group messages into conversations by sender.
  let igConvos = [];
  const igToken = props.getProperty("META_IG_TOKEN");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hiddenIGSenders = getHiddenIGSenders(ss);
  try {
    const igSheet = ss.getSheetByName("IG DM Log");
    if (igSheet && igSheet.getLastRow() > 1) {
      // Columns: Timestamp(1), Sender ID(2), Sender Name(3), Recipient ID(4), Message(5), Message ID(6), Is Echo(7), Direction(8)
      const igData = igSheet.getRange(2, 1, igSheet.getLastRow() - 1, 8).getValues();

      // Group by sender into conversations (keyed by sender ID for inbound, recipient for outbound)
      const convoMap = {};  // senderId → { messages: [...], senderName }
      igData.forEach(row => {
        const ts = row[0] instanceof Date ? row[0] : new Date(row[0]);
        const senderId = String(row[1] || "");
        const senderName = String(row[2] || "Unknown");
        const msgText = String(row[4] || "");
        const isEcho = row[6] === true || row[6] === "TRUE" || row[6] === "true";
        const direction = String(row[7] || "");

        // Key the conversation by the customer's ID (senderId for inbound, recipientId for outbound)
        const customerId = isEcho ? String(row[3] || "") : senderId;
        const customerName = isEcho ? "customer" : senderName;

        if (!customerId) return;

        // Skip hidden senders (filtered from dashboard, not blocked on IG)
        // Checks both sender ID and username since operators may paste either into the sheet
        if (hiddenIGSenders.has(customerId) || hiddenIGSenders.has(customerName)) return;

        if (!convoMap[customerId]) {
          convoMap[customerId] = { customerName: "Unknown", messages: [] };
        }
        // Update customer name if this is an inbound message (customer sent it)
        if (!isEcho && senderName !== senderId) {
          convoMap[customerId].customerName = senderName;
        }
        convoMap[customerId].messages.push({
          text: msgText, time: ts, isEcho, direction
        });
      });

      // Convert to conversation objects matching the same shape as Messenger convos
      const inboxUrl = `https://business.facebook.com/latest/inbox/instagram?asset_id=${pageId}`;
      Object.keys(convoMap).forEach(customerId => {
        const convo = convoMap[customerId];
        // Sort messages newest first
        convo.messages.sort((a, b) => b.time - a.time);
        const latest = convo.messages[0];
        const latestTime = latest.time;

        // Only include if last activity was within 24h
        if (latestTime < oneDayAgo) return;

        // Determine unread: if the latest message is from the customer (not echo), it's "unread"
        const isUnread = !latest.isEcho;
        const customerName = convo.customerName !== "Unknown" ? convo.customerName : ("IG User " + customerId.substring(0, 6));

        const msgText = latest.text || "";
        igConvos.push({
          id: "ig_webhook_" + customerId,
          customerName,
          unread: isUnread ? 1 : 0,
          lastMessage: msgText.length > 80 ? msgText.substring(0, 80) + "..." : msgText,
          lastMessageFrom: latest.isEcho ? "Deako" : customerName.split(" ")[0],
          time: latestTime.toISOString(),
          inboxUrl,
          platform: "Instagram",
        });
      });
      Logger.log("IG DMs from webhook log: " + igConvos.length + " active conversations");
    } else {
      Logger.log("IG DM Log sheet empty or missing — no IG DM data");
    }
  } catch (e) {
    Logger.log("Error reading IG DM Log sheet: " + e.toString());
  }

  // Merge conversations, deduplicate by id, filter to actionable items only:
  //   - Within last 24h (recent activity), OR
  //   - Unread (unread_count > 0 per Meta API — the real "needs action" signal)
  // Old conversations where the customer sent the last message months ago are NOT
  // considered actionable — unread_count from the API is the authoritative indicator.
  const seenIds = new Set();
  const allConversations = [];
  [...messengerConvos, ...igConvos].forEach(c => {
    if (seenIds.has(c.id)) return;
    seenIds.add(c.id);
    const convoTime = c.time ? new Date(c.time) : new Date(0);
    const isRecent = convoTime >= oneDayAgo;
    const isUnread = c.unread > 0;
    if (isRecent || isUnread) {
      allConversations.push(c);
    }
  });
  allConversations.sort((a, b) => {
    if (a.unread !== b.unread) return b.unread - a.unread;
    return new Date(b.time) - new Date(a.time);
  });

  let totalUnread = 0;
  allConversations.forEach(c => { totalUnread += c.unread; });

  // ─── Step 3: Auto-discover Instagram Business Account ID ───
  let igAccountId = props.getProperty("META_IG_ID") || null;
  if (!igAccountId) {
    const pageInfo = graphGet(`${pageId}?fields=instagram_business_account`);
    if (pageInfo && pageInfo.instagram_business_account) {
      igAccountId = pageInfo.instagram_business_account.id;
      Logger.log("Auto-discovered IG account: " + igAccountId);
    }
  }

  // ─── Step 4: Facebook Post Comments (last 24h) ───
  const recentComments = [];
  const fbFeed = graphGet(
    `${pageId}/feed?fields=id,message,permalink_url,comments.limit(10){id,from,message,created_time}&limit=5`
  );
  if (fbFeed && fbFeed.data) {
    fbFeed.data.forEach(post => {
      const comments = (post.comments && post.comments.data) || [];
      comments.forEach(c => {
        const cTime = new Date(c.created_time);
        if (cTime < oneDayAgo) return; // only last 24h
        const fromName = (c.from && c.from.name) || "Unknown";
        // Skip comments from our own page
        if (c.from && c.from.id === pageId) return;
        const postSnippet = post.message
          ? (post.message.length > 40 ? post.message.substring(0, 40) + "..." : post.message)
          : "Post";
        recentComments.push({
          type: "comment",
          platform: "Facebook",
          author: fromName,
          text: c.message.length > 80 ? c.message.substring(0, 80) + "..." : c.message,
          time: c.created_time,
          postSnippet,
          url: post.permalink_url || `https://business.facebook.com/latest/inbox/all?asset_id=${pageId}`,
        });
      });
    });
  }

  // ─── Step 5: Instagram Post Comments (last 24h) ───
  if (igAccountId) {
    const igMedia = graphGet(
      `${igAccountId}/media?fields=id,caption,permalink,timestamp,comments.limit(10){id,text,username,timestamp}&limit=5`
    );
    if (igMedia && igMedia.data) {
      igMedia.data.forEach(media => {
        const comments = (media.comments && media.comments.data) || [];
        comments.forEach(c => {
          const cTime = new Date(c.timestamp);
          if (cTime < oneDayAgo) return;
          const postSnippet = media.caption
            ? (media.caption.length > 40 ? media.caption.substring(0, 40) + "..." : media.caption)
            : "Post";
          recentComments.push({
            type: "comment",
            platform: "Instagram",
            author: c.username || "Unknown",
            text: c.text.length > 80 ? c.text.substring(0, 80) + "..." : c.text,
            time: c.timestamp,
            postSnippet,
            url: media.permalink || `https://business.facebook.com/latest/inbox/all?asset_id=${pageId}`,
          });
        });
      });
    }

    // ─── Step 6: Instagram Mentions / Tags (last 24h) ───
    const mentionData = graphGet(
      `${igAccountId}/tags?fields=id,caption,permalink,timestamp,username&limit=10`
    );
    if (mentionData && mentionData.data) {
      mentionData.data.forEach(m => {
        const mTime = new Date(m.timestamp);
        if (mTime < oneDayAgo) return;
        recentComments.push({
          type: "mention",
          platform: "Instagram",
          author: m.username || "Unknown",
          text: m.caption
            ? (m.caption.length > 80 ? m.caption.substring(0, 80) + "..." : m.caption)
            : "Tagged Deako",
          time: m.timestamp,
          postSnippet: "Mention",
          url: m.permalink || `https://business.facebook.com/latest/inbox/all?asset_id=${pageId}`,
        });
      });
    }
  }

  // Sort comments: newest first
  recentComments.sort((a, b) => new Date(b.time) - new Date(a.time));

  // ─── Step 7: Token expiry check ───
  let tokenWarning = null;
  // Check Page token expiry
  try {
    const debugData = graphGet(`debug_token?input_token=${pageToken}`);
    if (debugData && debugData.data && debugData.data.expires_at) {
      const expiresAt = new Date(debugData.data.expires_at * 1000);
      const daysLeft = Math.floor((expiresAt - new Date()) / (1000 * 60 * 60 * 24));
      if (daysLeft <= 7) {
        tokenWarning = `Meta Page token expires in ${daysLeft} day${daysLeft !== 1 ? "s" : ""}!`;
      }
    }
  } catch (e) {
    Logger.log("Page token debug check failed: " + e.toString());
  }
  // Check IG token expiry (uses graph.instagram.com, separate long-lived token)
  if (igToken && !tokenWarning) {
    try {
      const igRefreshUrl = `https://graph.instagram.com/refresh_access_token?grant_type=ig_refresh_token&access_token=${igToken}`;
      // We don't actually refresh here — just checking via token debug on graph.facebook.com
      const igDebug = graphGet(`debug_token?input_token=${igToken}`);
      if (igDebug && igDebug.data && igDebug.data.expires_at) {
        const expiresAt = new Date(igDebug.data.expires_at * 1000);
        const daysLeft = Math.floor((expiresAt - new Date()) / (1000 * 60 * 60 * 24));
        if (daysLeft <= 7) {
          tokenWarning = `IG token expires in ${daysLeft} day${daysLeft !== 1 ? "s" : ""}! Refresh via: /refresh_access_token`;
        }
      }
    } catch (e) {
      Logger.log("IG token debug check failed: " + e.toString());
    }
  }

  return {
    unreadDMs: totalUnread,
    totalConversations: allConversations.length,
    recentConversations: allConversations.slice(0, 10),
    recentComments: recentComments.slice(0, 8), // top 8 for dashboard
    tokenWarning,
    error: null,
  };
}

// --- READ SMS ACTIVITY FOR DASHBOARD ---
function readSMSActivity() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SMS Log");
  if (!sheet || sheet.getLastRow() <= 1) {
    return { totalToday: 0, inbound: 0, outbound: 0, agentStats: {}, messages: [] };
  }

  const now = new Date();
  const todayStart = new Date(now);
  todayStart.setHours(0, 0, 0, 0);

  // Read all data rows (skip header)
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

  // Columns: [0] Timestamp, [1] Event, [2] Direction, [3] Agent, [4] Contact Name,
  //          [5] Contact Phone, [6] Contact Email, [7] Message, [8] Status, [9] Aircall Line, [10] Line Number
  let inbound = 0;
  let outbound = 0;
  const agentStats = {};
  const messages = [];

  // Initialize agent stats for tracked agents
  CONFIG.agents.forEach(a => { agentStats[a] = { sent: 0, received: 0 }; });

  data.forEach(row => {
    const ts = new Date(row[0]);
    if (ts < todayStart) return;

    // Skip status_updated events — only count sent/received
    const event = String(row[1] || "");
    if (event.includes("status_updated")) return;

    const direction = String(row[2] || "");
    const agent = String(row[3] || "");
    const contact = String(row[4] || "");
    const phone = String(row[5] || "");
    const body = String(row[7] || "");
    const lineName = String(row[9] || "");

    // Skip excluded Aircall lines
    const excludeLines = CONFIG.excludeSMSLines || [];
    if (excludeLines.some(ex => lineName.toLowerCase() === ex.toLowerCase())) return;
    const timeStr = Utilities.formatDate(ts, Session.getScriptTimeZone(), "h:mm a");

    if (direction === "inbound") {
      inbound++;
    } else {
      outbound++;
    }

    // Match to tracked agent
    const matched = CONFIG.agents.find(a => {
      const parts = a.toLowerCase().split(/\s+/);
      const agentLower = agent.toLowerCase();
      return a === agent || parts.some(p => p.length > 1 && agentLower.includes(p));
    });
    if (matched) {
      if (direction === "inbound") {
        agentStats[matched].received++;
      } else {
        agentStats[matched].sent++;
      }
    }

    messages.push({ ts, direction, agent, contact, phone, body, lineName, timeStr });
  });

  // Sort newest first
  messages.sort((a, b) => b.ts - a.ts);

  return {
    totalToday: inbound + outbound,
    inbound,
    outbound,
    agentStats,
    messages: messages.slice(0, 10), // last 10 for display
  };
}

// --- DAILY RECAP EMAIL ---
// Sends a formatted HTML email summarizing the day's CS metrics.
// Recipients configured via RECAP_RECIPIENTS Script Property (comma-separated emails).
// Schedule: run setupDailyRecapTrigger() once to set up the 6pm PST daily trigger.

function sendDailyRecap(recipientOverride, skipSave) {
  // When called by a time-driven trigger, the first arg is an event object -- ignore it
  if (recipientOverride && typeof recipientOverride !== "string") {
    recipientOverride = null;
    skipSave = false;
  }
  loadThresholds();
  const props = PropertiesService.getScriptProperties();
  const recipients = recipientOverride || props.getProperty("RECAP_RECIPIENTS") || "";
  if (!recipients) {
    Logger.log("RECAP_RECIPIENTS not set — skipping daily recap");
    return;
  }

  // Gather all data (same calls the dashboard uses)
  const zendesk = fetchZendeskStatus();
  const aircall = fetchAircallStatus();
  const csat = fetchNicereplyCSAT();
  const postCall = readPostCallCSAT();
  const sms = readSMSActivity();
  const meta = fetchMetaStatus();

  const tz = CONFIG.businessHours.timezone;
  const now = new Date();
  const dateStr = Utilities.formatDate(now, tz, "EEEE, MMMM d, yyyy");

  // ── Check if today is a non-working day (weekend or Deako holiday) ──
  const holidays = getDeakoHolidays();
  const nonWorkingDay = isNonWorkingDay(now, holidays);

  if (nonWorkingDay) {
    sendNonWorkingDaySnapshot(zendesk, meta, now, dateStr, tz, recipients);
    return;
  }

  // ── AI-powered daily theme analysis (runs early, non-blocking) ──
  let dailyThemesHtml = null;
  try {
    const themesResult = analyzeDailyThemes();
    if (themesResult) {
      dailyThemesHtml = typeof themesResult === "string" ? themesResult : themesResult.html;
    }
  } catch (e) {
    Logger.log("Daily themes analysis failed (non-fatal): " + e.toString());
  }

  // ── Load previous day's snapshot for comparison ──
  const prevJson = props.getProperty("RECAP_PREV_SNAPSHOT") || "{}";
  let prev = {};
  try { prev = JSON.parse(prevJson); } catch (e) { prev = {}; }

  // Determine if we have previous data to show the Yesterday column
  const hasPrev = prev.date ? true : false;

  // Fixed column widths for consistent alignment across all tables
  const colLabel = "width:50%;";
  const colToday = "width:25%;text-align:right;";
  const colYest = "width:25%;text-align:right;";

  // Yesterday cell helper — returns a <td> for the yesterday column (or empty string if no prev data)
  function prevTd(prevVal, unit) {
    if (!hasPrev || prevVal === undefined || prevVal === null) return hasPrev ? `<td style="${colYest}color:#AAA;font-size:12px;">-</td>` : "";
    const u = unit || "";
    return `<td style="${colYest}color:#888;font-size:12px;">${prevVal}${u}</td>`;
  }

  // Format the comparison date label (e.g., "Fri 5/22" instead of generic "Yesterday")
  let compLabel = "Yesterday";
  if (hasPrev && prev.date) {
    const parts = prev.date.split("-");
    const prevDate = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    const dayNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    compLabel = dayNames[prevDate.getDay()] + " " + (prevDate.getMonth() + 1) + "/" + prevDate.getDate();
  }

  // Table header row with Today / comparison date columns
  const compHeader = hasPrev
    ? `<tr style="border-bottom:1px solid #E1DFDD;"><td style="${colLabel}"></td><td style="${colToday}font-size:11px;color:#888;padding-bottom:6px;font-weight:bold;">Today</td><td style="${colYest}font-size:11px;color:#888;padding-bottom:6px;font-weight:bold;">${compLabel}</td></tr>`
    : "";

  // ── Health status calculations (must match dashboard logic exactly) ──
  const th = CONFIG.thresholds;

  // Email: >5 breached = At Risk, >0 = Watch, 0 = Healthy
  const emailHealth = zendesk.totalBreached > 10 ? "red"
    : zendesk.totalBreached > 5 ? "yellow" : "green";

  // Phone: >= green threshold = Healthy, >= yellow = Watch, else At Risk
  const phoneHealth = aircall.teamAnswerRate >= th.phoneAnswerRate.green ? "green"
    : aircall.teamAnswerRate >= th.phoneAnswerRate.yellow ? "yellow" : "red";

  // Social: uses computeSocialOldestWait() — same helper the dashboard uses
  const metaConvos = meta.recentConversations || [];
  const unreadConvos = metaConvos.filter(c => c.unread > 0);
  const socialOldestWaitMin = computeSocialOldestWait(meta);
  const socialHealth = socialOldestWaitMin <= th.socialResponseTime.green ? "green"
    : socialOldestWaitMin <= th.socialResponseTime.yellow ? "yellow" : "red";

  const healthColors = { green: "#2E7D32", yellow: "#F57F17", red: "#C62828" };
  const healthLabels = { green: "Healthy", yellow: "Watch", red: "At Risk" };

  function healthBadge(level) {
    return `<span style="background:${healthColors[level]};color:#fff;padding:2px 10px;border-radius:10px;font-size:12px;font-weight:bold;">${healthLabels[level]}</span>`;
  }

  // ── Build HTML email ──
  const navy = "#1B3747";
  const cardBg = "#F8F7F6";
  const borderColor = "#E1DFDD";
  const answerRateRounded = Math.round(aircall.teamAnswerRate || 0);
  const sasCount = (zendesk.sasTickets && zendesk.sasTickets.length) || 0;

  let html = `
  <div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;color:#1D1D1D;">
    <div style="background:${navy};color:#fff;padding:16px 24px;border-radius:8px 8px 0 0;">
      <h1 style="margin:0;font-size:20px;">CS AI Recap - Daily</h1>
      <p style="margin:4px 0 0;font-size:13px;color:#C3D3D7;">${dateStr}</p>
    </div>

    <div style="padding:20px 24px;background:#fff;border:1px solid ${borderColor};border-top:none;">

      <!-- Health Status -->
      <table style="width:100%;margin-bottom:20px;"><tr>
        <td style="text-align:center;padding:8px;">
          <div style="font-size:11px;color:#666;margin-bottom:4px;">EMAIL</div>${healthBadge(emailHealth)}
        </td>
        <td style="text-align:center;padding:8px;">
          <div style="font-size:11px;color:#666;margin-bottom:4px;">PHONE</div>${healthBadge(phoneHealth)}
        </td>
        <td style="text-align:center;padding:8px;">
          <div style="font-size:11px;color:#666;margin-bottom:4px;">SOCIAL</div>${healthBadge(socialHealth)}
        </td>
      </tr></table>`;

  // ── DAILY THEMES (AI-powered) — shown first, right after health badges ──
  if (dailyThemesHtml) {
    html += `
      <div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
        <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">AI Recap - Today's Top Themes</h2>
        <div style="font-size:13px;line-height:1.6;color:#1D1D1D;">${dailyThemesHtml}</div>
        <div style="margin-top:10px;font-size:10px;color:#999;">Analysis powered by Claude AI across all tickets and call intents</div>
      </div>`;
  }

  // ── EMAIL SECTION (single continuous table with per-agent subheadings) ──
  const subH1 = `font-size:12px;color:${navy};font-weight:bold;padding:10px 0 2px;border-top:1px solid #E1DFDD;`;
  const subH2 = `font-size:11px;color:#888;font-weight:bold;padding:6px 0 2px;`;
  const agentRow = `padding:1px 0 1px 12px;font-size:12px;`;
  const prevAgentEmail = (hasPrev && prev.agentEmail) || {};

  html += `
      <div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
        <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Email (Zendesk)</h2>
        <table style="width:100%;font-size:13px;border-collapse:collapse;">
          ${compHeader}
          <tr><td style="${colLabel}padding:4px 0;">Tickets Created</td><td style="${colToday}font-weight:bold;">${zendesk.ticketsCreatedToday || 0}</td>${prevTd(prev.ticketsCreated)}</tr>
          <tr><td style="${colLabel}padding:4px 0;">Open Tickets</td><td style="${colToday}font-weight:bold;">${zendesk.totalOpen}</td>${prevTd(prev.openTickets)}</tr>
          <tr><td style="${colLabel}padding:4px 0;">On Hold</td><td style="${colToday}font-weight:bold;">${zendesk.onHoldQueueCount}</td>${prevTd(prev.onHold)}</tr>
          <tr><td style="${colLabel}padding:4px 0;">Unassigned</td><td style="${colToday}font-weight:bold;">${zendesk.unassigned}</td>${prevTd(prev.unassigned)}</tr>
          <tr><td style="${colLabel}padding:4px 0;">SAS Tickets</td><td style="${colToday}font-weight:bold;">${sasCount}</td>${prevTd(prev.sasTickets)}</tr>
          <tr><td style="${colLabel}padding:4px 0;">Voicemails (Open)</td><td style="${colToday}font-weight:bold;">${zendesk.openVoicemails || 0}</td>${prevTd(prev.openVoicemails)}</tr>
          <tr><td style="${colLabel}padding:4px 0;">Past SLA (${zendesk.slaHours}h)</td><td style="${colToday}font-weight:bold;color:${healthColors[emailHealth]};">${zendesk.totalBreached}</td>${prevTd(prev.pastSla)}</tr>
          <tr><td style="${colLabel}padding:4px 0;">Tickets Solved</td><td style="${colToday}font-weight:bold;">${zendesk.totalHandledToday}</td>${prevTd(prev.solvedTotal)}</tr>`;

  html += `</table>`;

  // Per-agent horizontal table (agent names as rows, metrics as columns)
  if (zendesk.agentCounts && Object.keys(zendesk.agentCounts).length > 0) {
    const agentHeaderStyle = `font-size:11px;color:#888;font-weight:bold;text-align:center;padding:4px 6px;border-bottom:1px solid #E1DFDD;`;
    const agentNameStyle = `font-size:12px;padding:3px 0;font-weight:bold;`;
    const agentCellStyle = `font-size:12px;text-align:center;padding:3px 6px;`;

    html += `
        <div style="margin-top:12px;font-size:12px;color:${navy};font-weight:bold;border-top:1px solid #E1DFDD;padding-top:10px;">Per Agent</div>
        <table style="width:100%;font-size:12px;border-collapse:collapse;margin-top:6px;">
          <tr>
            <td style="${agentHeaderStyle}text-align:left;"></td>
            <td style="${agentHeaderStyle}">Assigned</td>
            <td style="${agentHeaderStyle}">Past SLA</td>
            <td style="${agentHeaderStyle}">Solved</td>
          </tr>`;
    CONFIG.agents.forEach(agent => {
      const ac = zendesk.agentCounts[agent];
      if (ac) {
        const slaColor = (ac.pastSla || 0) > 0 ? "color:#C62828;" : "";
        html += `<tr>
            <td style="${agentNameStyle}">${agent.split(" ")[0]}</td>
            <td style="${agentCellStyle}">${ac.assigned || 0}</td>
            <td style="${agentCellStyle}${slaColor}">${ac.pastSla || 0}</td>
            <td style="${agentCellStyle}">${ac.handledToday || 0}</td>
          </tr>`;
      }
    });
    html += `</table>`;
  }

  html += `</div>`;

  // ── PHONE SECTION ──
  html += `
      <div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
        <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Phone (Aircall)</h2>
        <table style="width:100%;font-size:13px;border-collapse:collapse;">
          ${compHeader}
          <tr><td style="${colLabel}padding:4px 0;">Answer Rate</td><td style="${colToday}font-weight:bold;color:${healthColors[phoneHealth]};">${answerRateRounded}%</td>${prevTd(prev.answerRate, "%")}</tr>
          <tr><td style="${colLabel}padding:4px 0;">Inbound Answered</td><td style="${colToday}font-weight:bold;">${aircall.teamAnswered}</td>${prevTd(prev.inboundCalls)}</tr>
          <tr><td style="${colLabel}padding:4px 0;">Sent to Answer Service</td><td style="${colToday}font-weight:bold;">${aircall.forwarded}</td>${prevTd(prev.forwardedToSAS)}</tr>
          <tr><td style="${colLabel}padding:4px 0;">Outbound Calls</td><td style="${colToday}font-weight:bold;">${aircall.totalOutbound}</td>${prevTd(prev.outboundCalls)}</tr>
          <tr><td style="${colLabel}padding:4px 0;">Avg Call Duration</td><td style="${colToday}font-weight:bold;">${formatSeconds(aircall.avgDuration || 0)}</td>${hasPrev && prev.avgCallDuration ? `<td style="${colYest}color:#888;font-size:12px;">${formatSeconds(prev.avgCallDuration)}</td>` : (hasPrev ? `<td style="${colYest}color:#AAA;font-size:12px;">-</td>` : '')}</tr>`;

  html += `</table>`;

  // Per-agent phone horizontal table
  if (aircall.agentStats) {
    const agentHeaderStyle = `font-size:11px;color:#888;font-weight:bold;text-align:center;padding:4px 6px;border-bottom:1px solid #E1DFDD;`;
    const agentNameStyle = `font-size:12px;padding:3px 0;font-weight:bold;`;
    const agentCellStyle = `font-size:12px;text-align:center;padding:3px 6px;`;

    html += `
        <div style="margin-top:12px;font-size:12px;color:${navy};font-weight:bold;border-top:1px solid #E1DFDD;padding-top:10px;">Per Agent</div>
        <table style="width:100%;font-size:12px;border-collapse:collapse;margin-top:6px;">
          <tr>
            <td style="${agentHeaderStyle}text-align:left;"></td>
            <td style="${agentHeaderStyle}">Inbound</td>
            <td style="${agentHeaderStyle}">Outbound</td>
            <td style="${agentHeaderStyle}">In Talk</td>
            <td style="${agentHeaderStyle}">Out Talk</td>
          </tr>`;
    CONFIG.agents.forEach(agent => {
      const as = aircall.agentStats[agent];
      if (as) {
        html += `<tr>
            <td style="${agentNameStyle}">${agent.split(" ")[0]}</td>
            <td style="${agentCellStyle}">${as.answered || 0}</td>
            <td style="${agentCellStyle}">${as.outbound || 0}</td>
            <td style="${agentCellStyle}">${formatTalkTime(as.inboundTalkTime || 0)}</td>
            <td style="${agentCellStyle}">${formatTalkTime(as.outboundTalkTime || 0)}</td>
          </tr>`;
      }
    });
    html += `</table>`;
  }

  html += `</div>`;

  // ── TEXT MESSAGES SECTION ──
  if (sms.totalToday > 0 || (hasPrev && (prev.smsInbound || prev.smsOutbound))) {
    const smsTotal = (sms.inbound || 0) + (sms.outbound || 0);
    const prevSmsTotal = hasPrev ? ((prev.smsInbound || 0) + (prev.smsOutbound || 0)) : null;
    html += `
      <div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
        <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Text Messages</h2>
        <table style="width:100%;font-size:13px;border-collapse:collapse;">
          ${compHeader}
          <tr><td style="${colLabel}padding:4px 0;">Inbound</td><td style="${colToday}font-weight:bold;">${sms.inbound || 0}</td>${prevTd(prev.smsInbound)}</tr>
          <tr><td style="${colLabel}padding:4px 0;">Outbound</td><td style="${colToday}font-weight:bold;">${sms.outbound || 0}</td>${prevTd(prev.smsOutbound)}</tr>
          <tr><td style="${colLabel}padding:4px 0;">Total</td><td style="${colToday}font-weight:bold;">${smsTotal}</td>${prevTd(prevSmsTotal)}</tr>
        </table>
      </div>`;
  }

  // ── SOCIAL SECTION ──
  const metaUnread = meta.unreadDMs || 0;
  const metaComments = meta.recentComments || [];
  html += `
      <div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
        <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Social (Meta Business Suite)</h2>
        <table style="width:100%;font-size:13px;border-collapse:collapse;">
          ${compHeader}
          <tr><td style="${colLabel}padding:4px 0;">Unread DMs</td><td style="${colToday}font-weight:bold;color:${healthColors[socialHealth]};">${metaUnread}</td>${prevTd(prev.unreadDMs)}</tr>
          <tr><td style="${colLabel}padding:4px 0;">Active Conversations (24h)</td><td style="${colToday}font-weight:bold;">${metaConvos.length}</td>${prevTd(prev.activeConversations)}</tr>
          <tr><td style="${colLabel}padding:4px 0;">Comments & Mentions (24h)</td><td style="${colToday}font-weight:bold;">${metaComments.length}</td>${prevTd(prev.commentsAndMentions)}</tr>
        </table>`;

  // List unread DMs
  if (unreadConvos.length > 0) {
    html += `<div style="margin-top:12px;font-size:12px;color:#666;font-weight:bold;">Unread DMs</div>
        <table style="width:100%;font-size:12px;border-collapse:collapse;margin-top:4px;">`;
    unreadConvos.slice(0, 5).forEach(c => {
      const platform = c.platform === "Instagram" ? "IG" : "FB";
      const msg = c.lastMessage ? (c.lastMessage.length > 60 ? c.lastMessage.substring(0, 57) + "..." : c.lastMessage) : "";
      html += `<tr><td style="padding:2px 0;"><strong>${c.customerName}</strong> <span style="color:${c.platform === 'Instagram' ? '#C13584' : '#1877F2'};font-size:10px;">${platform}</span></td></tr>`;
      if (msg) html += `<tr><td style="padding:0 0 4px;color:#888;font-style:italic;">${msg}</td></tr>`;
    });
    html += `</table>`;
  }
  html += `</div>`;

  // ── CSAT SECTION (summary only) ──
  const emailCsat = csat || {};
  const phoneCsat = postCall || {};
  const totalCsatResponses = (emailCsat.total || 0) + (phoneCsat.total || 0);
  if (totalCsatResponses > 0) {
    html += `
      <div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
        <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Customer Satisfaction (Last 24h)</h2>
        <table style="width:100%;font-size:13px;border-collapse:collapse;">`;
    if (emailCsat.score !== null && emailCsat.score !== undefined) {
      html += `<tr><td style="padding:4px 0;">Email CSAT</td><td style="text-align:right;font-weight:bold;">${emailCsat.score}% (${emailCsat.total} reviews)</td></tr>`;
    }
    if (phoneCsat.score !== null && phoneCsat.score !== undefined) {
      html += `<tr><td style="padding:4px 0;">Phone CSAT</td><td style="text-align:right;font-weight:bold;">${phoneCsat.score}% (${phoneCsat.total} reviews)</td></tr>`;
    }
    html += `</table>
      </div>`;
  }

  // (Daily themes block moved to top of email — see above)

  // ── TOKEN WARNING ──
  if (meta.tokenWarning) {
    html += `<div style="background:#FFF3CD;border:1px solid #FFE69C;border-radius:6px;padding:12px 16px;margin-bottom:16px;font-size:13px;color:#856404;">Warning: ${meta.tokenWarning}</div>`;
  }

  // Footer with health status definitions
  html += `
      <div style="border-top:1px solid ${borderColor};margin-top:20px;padding:16px 0 4px;font-size:11px;color:#999;">
        <div style="margin-bottom:8px;font-weight:bold;color:#888;">Health Status Thresholds</div>
        <div style="margin-bottom:4px;">Email: Healthy = 0-5 tickets past ${zendesk.slaHours}h SLA · Watch = 6-10 past SLA · At Risk = 11+ past SLA</div>
        <div style="margin-bottom:4px;">Phone: Healthy = ${th.phoneAnswerRate.green}%+ answer rate · Watch = ${th.phoneAnswerRate.yellow}-${th.phoneAnswerRate.green - 1}% · At Risk = below ${th.phoneAnswerRate.yellow}%</div>
        <div style="margin-bottom:4px;">Social: Healthy = oldest unread DM under ${th.socialResponseTime.green / 60}h · Watch = ${th.socialResponseTime.green / 60}-${th.socialResponseTime.yellow / 60}h · At Risk = over ${th.socialResponseTime.yellow / 60}h</div>
        <div style="margin-bottom:4px;">The "${compLabel}" column shows the previous working day's end-of-day values for comparison.</div>
      </div>
      <div style="text-align:center;padding:8px 0;font-size:11px;color:#999;">
        CS Visibility · AI Recap v2.5.4 · End of Day Summary · ${dateStr}
      </div>
    </div>
  </div>`;

  // ── SAVE DATA FIRST (before email send) ──
  // This ensures snapshot + metrics log persist even if the email send fails,
  // preventing stale comparison dates (e.g., the May 26 crash that caused Fri 5/22 to show on May 27).
  const snapshot = {
    date: Utilities.formatDate(now, tz, "yyyy-MM-dd"),
    // Email
    openTickets: zendesk.totalOpen,
    onHold: zendesk.onHoldQueueCount,
    unassigned: zendesk.unassigned,
    sasTickets: sasCount,
    pastSla: zendesk.totalBreached,
    openVoicemails: zendesk.openVoicemails || 0,
    ticketsCreated: zendesk.ticketsCreatedToday || 0,
    solvedTotal: zendesk.totalHandledToday || 0,
    agentEmail: {},
    // Phone
    answerRate: answerRateRounded,
    inboundCalls: aircall.teamAnswered || 0,
    forwardedToSAS: aircall.forwarded || 0,
    outboundCalls: aircall.totalOutbound || 0,
    avgCallDuration: Math.round(aircall.avgDuration || 0),
    agentPhone: {},
    // Social
    unreadDMs: metaUnread,
    activeConversations: metaConvos.length,
    commentsAndMentions: metaComments.length,
    // SMS
    smsInbound: sms.inbound || 0,
    smsOutbound: sms.outbound || 0,
    // CSAT
    csatPct: csat.score,
    csatResponses: csat.total || 0,
    phoneCsatPct: postCall.score,
    phoneCsatResponses: postCall.total || 0,
  };

  // ── Per-agent CSAT attribution ──
  // Email CSAT: match Nicereply response ticket_id to Zendesk ticket assignee
  const recapAgentEmailCsat = {};
  CONFIG.agents.forEach(a => { recapAgentEmailCsat[a] = { satisfied: 0, total: 0, tickets: [] }; });
  if (csat && csat.responses) {
    csat.responses.forEach(r => {
      if (!r.ticketId) return;
      // Find which agent owns this ticket from the fetched ticket list
      const ticket = (zendesk.tickets || []).find(t => String(t.id) === String(r.ticketId));
      if (!ticket) return;
      const matched = CONFIG.agents.find(ca => {
        const parts = ca.toLowerCase().split(/\s+/);
        const tLower = (ticket.assignee || "").toLowerCase();
        return ca === ticket.assignee || parts.some(p => p.length > 1 && tLower.split(/\s+/).some(ap => ap === p));
      });
      if (matched) {
        recapAgentEmailCsat[matched].total++;
        if (r.satisfied) recapAgentEmailCsat[matched].satisfied++;
        recapAgentEmailCsat[matched].tickets.push({ ticketId: r.ticketId, score: r.score, satisfied: r.satisfied });
      }
    });
  }

  // Phone CSAT: match by agent name in PostCall responses
  const recapAgentPhoneCsat = {};
  CONFIG.agents.forEach(a => { recapAgentPhoneCsat[a] = { satisfied: 0, total: 0 }; });
  if (postCall && postCall.responses) {
    postCall.responses.forEach(r => {
      const matched = CONFIG.agents.find(a => {
        const parts = a.toLowerCase().split(/\s+/);
        const rParts = (r.agent || "").toLowerCase().split(/\s+/);
        return parts[0] === rParts[0] || (parts[1] && rParts[1] && parts[1] === rParts[1]);
      });
      if (matched) {
        recapAgentPhoneCsat[matched].total++;
        if (r.satisfied) recapAgentPhoneCsat[matched].satisfied++;
      }
    });
  }

  // Per-agent snapshots
  CONFIG.agents.forEach(agent => {
    const ac = zendesk.agentCounts ? zendesk.agentCounts[agent] : null;
    snapshot.agentEmail[agent] = {
      assigned: ac ? (ac.assigned || 0) : 0,
      pastSla: ac ? (ac.pastSla || 0) : 0,
      solved: ac ? (ac.handledToday || 0) : 0,
    };
    const as = aircall.agentStats ? aircall.agentStats[agent] : null;
    snapshot.agentPhone[agent] = {
      inbound: as ? (as.answered || 0) : 0,
      outbound: as ? (as.outbound || 0) : 0,
      inTalk: as ? (as.inboundTalkTime || 0) : 0,
      outTalk: as ? (as.outboundTalkTime || 0) : 0,
    };
  });
  if (!skipSave) {
    props.setProperty("RECAP_PREV_SNAPSHOT", JSON.stringify(snapshot));
    Logger.log("Saved daily snapshot for comparison: " + JSON.stringify(snapshot));
  } else {
    Logger.log("Test mode — skipped saving snapshot and metrics log");
  }

  // ── Append to Daily Metrics Log sheet (powers weekly/monthly summaries) ──
  if (!skipSave) {
  logDailyMetrics({
    date: Utilities.formatDate(now, tz, "yyyy-MM-dd"),
    dayOfWeek: Utilities.formatDate(now, tz, "EEEE"),
    openTickets: zendesk.totalOpen,
    unassigned: zendesk.unassigned,
    onHold: zendesk.onHoldQueueCount,
    pastSla: zendesk.totalBreached,
    sasTickets: sasCount,
    openVoicemails: zendesk.openVoicemails || 0,
    aiAgentTickets: zendesk.aiAgentCount || 0,
    ticketsCreated: zendesk.ticketsCreatedToday || "",
    solvedTotal: zendesk.totalHandledToday || 0,
    agentSolved: CONFIG.agents.map(a => (zendesk.agentCounts[a] || {}).handledToday || 0),
    answerRate: answerRateRounded,
    inboundCalls: aircall.teamAnswered || 0,
    forwardedToSAS: aircall.forwarded || 0,
    outboundCalls: aircall.totalOutbound || 0,
    avgWaitTime: Math.round(aircall.avgWaitTime || 0),
    avgCallDuration: Math.round(aircall.avgDuration || 0),
    agentInbound: CONFIG.agents.map(a => (aircall.agentStats[a] || {}).answered || 0),
    agentOutbound: CONFIG.agents.map(a => (aircall.agentStats[a] || {}).outbound || 0),
    agentInTalk: CONFIG.agents.map(a => (aircall.agentStats[a] || {}).inboundTalkTime || 0),
    agentOutTalk: CONFIG.agents.map(a => (aircall.agentStats[a] || {}).outboundTalkTime || 0),
    csatPct: csat.score,
    csatResponses: csat.total || 0,
    csatSatisfied: csat.satisfied || 0,
    phoneCsatPct: postCall.score,
    phoneCsatResponses: postCall.total || 0,
    // Per-agent CSAT
    agentEmailCsatSat: CONFIG.agents.map(a => recapAgentEmailCsat[a].satisfied),
    agentEmailCsatTot: CONFIG.agents.map(a => recapAgentEmailCsat[a].total),
    agentPhoneCsatSat: CONFIG.agents.map(a => recapAgentPhoneCsat[a].satisfied),
    agentPhoneCsatTot: CONFIG.agents.map(a => recapAgentPhoneCsat[a].total),
    unreadDMs: metaUnread,
    smsInbound: sms.inbound || 0,
    smsOutbound: sms.outbound || 0,
  });
  } // end if (!skipSave)

  // Update Health Trends tab (historical chart of health statuses)
  try { updateHealthTrends(); } catch (e) {
    Logger.log("Health Trends update failed (non-fatal): " + e.toString());
  }

  // ── SEND EMAIL (after data is safely persisted) ──
  const subject = `CS AI Recap - Daily - ${Utilities.formatDate(now, tz, "MMM d")} - Email: ${healthLabels[emailHealth]} | Phone: ${healthLabels[phoneHealth]} | Social: ${healthLabels[socialHealth]}`;

  GmailApp.sendEmail(recipients, subject, "View this email with HTML enabled.", {
    htmlBody: html,
    name: "CS Visibility",
  });

  Logger.log("End of day summary sent to: " + recipients);
}

// ─── DAILY METRICS LOG ───
// Appends one row per working day to a "Daily Metrics Log" sheet.
// Powers weekly/monthly summary emails and Hex historical dashboard.
// Column order must match METRICS_LOG_HEADERS exactly.

const METRICS_LOG_HEADERS = [
  "Date", "Day", "Open Tickets", "Unassigned", "On Hold", "Past SLA",
  "SAS Tickets", "Open Voicemails", "AI Agent Tickets",
  "Tickets Created", "Solved Total",
  // Per-agent solved columns are dynamically named from CONFIG.agents
  // e.g. "Solved: AgentA", "Solved: AgentB", "Solved: AgentC"
  ...CONFIG.agents.map(a => "Solved: " + a.split(" ")[0]),
  "Answer Rate %", "Inbound Calls", "Forwarded to SAS", "Outbound Calls",
  "Avg Wait (sec)", "Avg Duration (sec)",
  // Per-agent inbound/outbound
  ...CONFIG.agents.map(a => "In: " + a.split(" ")[0]),
  ...CONFIG.agents.map(a => "Out: " + a.split(" ")[0]),
  // Per-agent talk time (seconds)
  ...CONFIG.agents.map(a => "In Talk: " + a.split(" ")[0]),
  ...CONFIG.agents.map(a => "Out Talk: " + a.split(" ")[0]),
  "Email CSAT %", "Email CSAT Responses", "Email CSAT Satisfied",
  "Phone CSAT %", "Phone CSAT Responses",
  // Per-agent CSAT (satisfied / total pairs)
  ...CONFIG.agents.map(a => "Email CSAT Sat: " + a.split(" ")[0]),
  ...CONFIG.agents.map(a => "Email CSAT Tot: " + a.split(" ")[0]),
  ...CONFIG.agents.map(a => "Phone CSAT Sat: " + a.split(" ")[0]),
  ...CONFIG.agents.map(a => "Phone CSAT Tot: " + a.split(" ")[0]),
  "Unread DMs", "SMS Inbound", "SMS Outbound",
];

function getOrCreateMetricsLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Daily Metrics Log");
  if (!sheet) {
    sheet = ss.insertSheet("Daily Metrics Log");
    // Write header row
    sheet.getRange(1, 1, 1, METRICS_LOG_HEADERS.length).setValues([METRICS_LOG_HEADERS]);
    sheet.getRange(1, 1, 1, METRICS_LOG_HEADERS.length)
      .setFontWeight("bold").setFontSize(9).setBackground("#E1DFDD");
    sheet.setFrozenRows(1);
    // Set date column format
    sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd");
    Logger.log("Created 'Daily Metrics Log' sheet with " + METRICS_LOG_HEADERS.length + " columns");
  }
  return sheet;
}

function logDailyMetrics(m) {
  const sheet = getOrCreateMetricsLog();

  // Prevent duplicate rows for the same date
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const lastDate = sheet.getRange(lastRow, 1).getValue();
    if (lastDate) {
      const lastDateStr = (lastDate instanceof Date)
        ? Utilities.formatDate(lastDate, CONFIG.businessHours.timezone, "yyyy-MM-dd")
        : String(lastDate);
      if (lastDateStr === m.date) {
        Logger.log("Daily Metrics Log already has entry for " + m.date + " — skipping");
        return;
      }
    }
  }

  // Build row in exact column order matching METRICS_LOG_HEADERS
  const row = [
    m.date,
    m.dayOfWeek,
    m.openTickets,
    m.unassigned,
    m.onHold,
    m.pastSla,
    m.sasTickets,
    m.openVoicemails,
    m.aiAgentTickets,
    m.ticketsCreated !== undefined ? m.ticketsCreated : "",
    m.solvedTotal,
    ...m.agentSolved,          // per-agent solved (matches CONFIG.agents order)
    m.answerRate,
    m.inboundCalls,
    m.forwardedToSAS,
    m.outboundCalls,
    m.avgWaitTime,
    m.avgCallDuration,
    ...m.agentInbound,         // per-agent inbound calls
    ...m.agentOutbound,        // per-agent outbound calls
    ...m.agentInTalk,          // per-agent inbound talk time (seconds)
    ...m.agentOutTalk,         // per-agent outbound talk time (seconds)
    m.csatPct !== null && m.csatPct !== undefined ? m.csatPct : "",
    m.csatResponses,
    m.csatSatisfied,
    m.phoneCsatPct !== null && m.phoneCsatPct !== undefined ? m.phoneCsatPct : "",
    m.phoneCsatResponses,
    // Per-agent CSAT
    ...(m.agentEmailCsatSat || CONFIG.agents.map(() => 0)),
    ...(m.agentEmailCsatTot || CONFIG.agents.map(() => 0)),
    ...(m.agentPhoneCsatSat || CONFIG.agents.map(() => 0)),
    ...(m.agentPhoneCsatTot || CONFIG.agents.map(() => 0)),
    m.unreadDMs,
    m.smsInbound,
    m.smsOutbound,
  ];

  sheet.appendRow(row);
  Logger.log("Logged daily metrics for " + m.date + " (" + row.length + " columns)");
}

// Read metrics log data for a date range (inclusive).
// Returns array of objects keyed by header name.
function readMetricsLog(startDate, endDate) {
  const sheet = getOrCreateMetricsLog();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];  // header only

  const data = sheet.getRange(1, 1, lastRow, METRICS_LOG_HEADERS.length).getValues();
  const headers = data[0];
  const tz = CONFIG.businessHours.timezone;
  const rows = [];

  for (let i = 1; i < data.length; i++) {
    const dateVal = data[i][0];
    if (!dateVal) continue;
    const dateStr = (dateVal instanceof Date)
      ? Utilities.formatDate(dateVal, tz, "yyyy-MM-dd")
      : String(dateVal);
    if (dateStr >= startDate && dateStr <= endDate) {
      const obj = {};
      headers.forEach((h, j) => { obj[h] = data[i][j]; });
      obj._dateStr = dateStr;
      rows.push(obj);
    }
  }
  return rows;
}

// ─── WEEKLY & MONTHLY SUMMARY HELPERS ───

// Compute aggregate metrics from an array of daily log rows.
// Returns object with averages and totals.
function aggregateMetrics(rows) {
  if (rows.length === 0) return null;
  const n = rows.length;

  function avg(key) {
    const vals = rows.map(r => Number(r[key]) || 0);
    return vals.reduce((a, b) => a + b, 0) / n;
  }
  function sum(key) {
    return rows.map(r => Number(r[key]) || 0).reduce((a, b) => a + b, 0);
  }
  function avgNonEmpty(key) {
    const vals = rows.map(r => r[key]).filter(v => v !== "" && v !== null && v !== undefined).map(Number);
    return vals.length > 0 ? vals.reduce((a, b) => a + b, 0) / vals.length : null;
  }
  function countNonEmpty(key) {
    return rows.map(r => r[key]).filter(v => v !== "" && v !== null && v !== undefined).length;
  }

  return {
    days: n,
    // Queue-state metrics use avgNonEmpty because backfilled rows have blank values
    // (these are live snapshots that can't be reconstructed historically)
    queueStateDays: countNonEmpty("Open Tickets"), // how many days have live queue data
    avgOpenTickets: avgNonEmpty("Open Tickets") !== null ? Math.round(avgNonEmpty("Open Tickets") * 10) / 10 : null,
    avgUnassigned: avgNonEmpty("Unassigned") !== null ? Math.round(avgNonEmpty("Unassigned") * 10) / 10 : null,
    avgOnHold: avgNonEmpty("On Hold") !== null ? Math.round(avgNonEmpty("On Hold") * 10) / 10 : null,
    avgPastSla: avgNonEmpty("Past SLA") !== null ? Math.round(avgNonEmpty("Past SLA") * 10) / 10 : null,
    avgVoicemails: avgNonEmpty("Open Voicemails") !== null ? Math.round(avgNonEmpty("Open Voicemails") * 10) / 10 : null,
    totalCreated: sum("Tickets Created"),
    totalSolved: sum("Solved Total"),
    agentSolved: CONFIG.agents.map(a => ({
      name: a,
      total: sum("Solved: " + a.split(" ")[0]),
    })),
    avgAnswerRate: Math.round(avg("Answer Rate %")),
    totalInbound: sum("Inbound Calls"),
    totalForwarded: sum("Forwarded to SAS"),
    totalOutbound: sum("Outbound Calls"),
    avgWaitTime: Math.round(avg("Avg Wait (sec)")),
    agentInbound: CONFIG.agents.map(a => ({
      name: a,
      total: sum("In: " + a.split(" ")[0]),
    })),
    agentOutbound: CONFIG.agents.map(a => ({
      name: a,
      total: sum("Out: " + a.split(" ")[0]),
    })),
    agentInTalk: CONFIG.agents.map(a => ({
      name: a,
      total: sum("In Talk: " + a.split(" ")[0]),
    })),
    agentOutTalk: CONFIG.agents.map(a => ({
      name: a,
      total: sum("Out Talk: " + a.split(" ")[0]),
    })),
    avgCsatPct: avgNonEmpty("Email CSAT %"),
    totalCsatResponses: sum("Email CSAT Responses"),
    totalCsatSatisfied: sum("Email CSAT Satisfied"),
    avgPhoneCsatPct: avgNonEmpty("Phone CSAT %"),
    totalPhoneCsatResponses: sum("Phone CSAT Responses"),
    avgUnreadDMs: Math.round(avg("Unread DMs") * 10) / 10,
    totalSmsInbound: sum("SMS Inbound"),
    totalSmsOutbound: sum("SMS Outbound"),
  };
}

// Format a comparison cell: "78%" with optional delta "(was 72%)"
function fmtComp(current, prior, unit, opts) {
  const u = unit || "";
  const round = (opts && opts.round !== undefined) ? opts.round : 0;
  const curStr = current !== null && current !== undefined
    ? (round > 0 ? Number(current).toFixed(round) : Math.round(current)) + u
    : "—";
  if (prior === null || prior === undefined) return curStr;
  const priorStr = (round > 0 ? Number(prior).toFixed(round) : Math.round(prior)) + u;
  return `${curStr} <span style="color:#888;font-size:11px;">(was ${priorStr})</span>`;
}

// Get date range for the current work-week (Mon–Fri of this week)
function getThisWeekRange(now, tz) {
  const d = new Date(Utilities.formatDate(now, tz, "yyyy-MM-dd") + "T12:00:00");
  const dow = d.getDay(); // 0=Sun ... 6=Sat
  // Monday of this week
  const mondayOffset = dow === 0 ? -6 : 1 - dow;
  const monday = new Date(d);
  monday.setDate(d.getDate() + mondayOffset);
  const friday = new Date(monday);
  friday.setDate(monday.getDate() + 4);
  return {
    start: Utilities.formatDate(monday, tz, "yyyy-MM-dd"),
    end: Utilities.formatDate(friday, tz, "yyyy-MM-dd"),
  };
}

// Get date range for the prior work-week
function getLastWeekRange(now, tz) {
  const d = new Date(Utilities.formatDate(now, tz, "yyyy-MM-dd") + "T12:00:00");
  const dow = d.getDay();
  const mondayOffset = dow === 0 ? -6 : 1 - dow;
  const thisMonday = new Date(d);
  thisMonday.setDate(d.getDate() + mondayOffset);
  const lastFriday = new Date(thisMonday);
  lastFriday.setDate(thisMonday.getDate() - 3);
  const lastMonday = new Date(lastFriday);
  lastMonday.setDate(lastFriday.getDate() - 4);
  return {
    start: Utilities.formatDate(lastMonday, tz, "yyyy-MM-dd"),
    end: Utilities.formatDate(lastFriday, tz, "yyyy-MM-dd"),
  };
}

// Get date range for a full calendar month (1st–last day)
function getMonthRange(year, month, tz) {
  const start = `${year}-${String(month).padStart(2, "0")}-01`;
  const lastDay = new Date(year, month, 0).getDate(); // day 0 of next month = last day of this month
  const end = `${year}-${String(month).padStart(2, "0")}-${String(lastDay).padStart(2, "0")}`;
  return { start, end };
}

// Check if today is the last working day of the week (usually Friday, Thursday if Fri is holiday)
function isLastWorkingDayOfWeek(now, holidays) {
  const tz = CONFIG.businessHours.timezone;
  const dow = new Date(Utilities.formatDate(now, tz, "yyyy-MM-dd") + "T12:00:00").getDay();
  // Must be a working day itself
  if (!CONFIG.businessHours.workDays.includes(dow)) return false;
  // Check if tomorrow through Sunday are all non-working days
  for (let i = 1; i <= (7 - dow); i++) {
    const next = new Date(now);
    next.setDate(now.getDate() + i);
    const nextDow = new Date(Utilities.formatDate(next, tz, "yyyy-MM-dd") + "T12:00:00").getDay();
    if (nextDow === 0 || nextDow === 6) continue; // weekend
    const nextKey = Utilities.formatDate(next, tz, "yyyy-MM-dd");
    if (holidays.has(nextKey)) continue; // holiday
    return false; // found a working day before end of week
  }
  return true;
}

// Check if today is the last business day of the month
function isLastBusinessDayOfMonth(now, holidays) {
  const tz = CONFIG.businessHours.timezone;
  const todayStr = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  const todayDate = new Date(todayStr + "T12:00:00");
  const thisMonth = todayDate.getMonth();
  // Must be a working day
  const dow = todayDate.getDay();
  if (!CONFIG.businessHours.workDays.includes(dow)) return false;
  if (holidays.has(todayStr)) return false;
  // Check remaining days of the month
  for (let d = 1; d <= 10; d++) {
    const next = new Date(todayDate);
    next.setDate(todayDate.getDate() + d);
    if (next.getMonth() !== thisMonth) break; // past end of month
    const nextDow = next.getDay();
    if (nextDow === 0 || nextDow === 6) continue;
    const nextKey = Utilities.formatDate(next, tz, "yyyy-MM-dd");
    if (holidays.has(nextKey)) continue;
    return false; // found a later business day in this month
  }
  return true;
}

// ─── HEALTH TREND DATA FOR EMAILS ───
// Reads Health Badge Log + Daily Metrics Log for a date range and returns
// an array of { date, dateLabel, email, phone, social } objects (status strings).
// Merges both sources — badge log is authoritative, metrics fill gaps.
function readHealthTrends(startDate, endDate, tz) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  function toKey(d) {
    if (!d) return null;
    if (d instanceof Date) return Utilities.formatDate(d, tz, "yyyy-MM-dd");
    const s = String(d).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const parsed = new Date(s);
    return !isNaN(parsed.getTime()) ? Utilities.formatDate(parsed, tz, "yyyy-MM-dd") : null;
  }

  // Badge log (authoritative)
  const badgeByDate = {};
  const badgeSheet = ss.getSheetByName("Health Badge Log");
  if (badgeSheet && badgeSheet.getLastRow() > 1) {
    const bData = badgeSheet.getRange(1, 1, badgeSheet.getLastRow(), 4).getValues();
    for (let i = 1; i < bData.length; i++) {
      const key = toKey(bData[i][0]);
      if (!key || key < startDate || key > endDate) continue;
      badgeByDate[key] = { email: String(bData[i][1] || ""), phone: String(bData[i][2] || ""), social: String(bData[i][3] || "") };
    }
  }

  // Metrics log (derived fallback)
  const th = CONFIG.thresholds;
  const metricsByDate = {};
  const logSheet = ss.getSheetByName("Daily Metrics Log");
  if (logSheet && logSheet.getLastRow() > 1) {
    const data = logSheet.getRange(1, 1, logSheet.getLastRow(), logSheet.getLastColumn()).getValues();
    const headers = data[0];
    const dateIdx = headers.indexOf("Date");
    const pastSlaIdx = headers.indexOf("Past SLA");
    const answerRateIdx = headers.indexOf("Answer Rate %");
    const unreadDMsIdx = headers.indexOf("Unread DMs");
    for (let i = 1; i < data.length; i++) {
      const key = toKey(data[i][dateIdx]);
      if (!key || key < startDate || key > endDate) continue;
      const pastSla = pastSlaIdx >= 0 ? data[i][pastSlaIdx] : "";
      const answerRate = answerRateIdx >= 0 ? data[i][answerRateIdx] : "";
      const unreadDMs = unreadDMsIdx >= 0 ? data[i][unreadDMsIdx] : "";
      metricsByDate[key] = {
        email: pastSla !== "" && !isNaN(pastSla) ? (pastSla > 10 ? "At Risk" : pastSla > 5 ? "Watch" : "Healthy") : "",
        phone: answerRate !== "" && !isNaN(answerRate) ? (answerRate >= th.phoneAnswerRate.green ? "Healthy" : answerRate >= th.phoneAnswerRate.yellow ? "Watch" : "At Risk") : "",
        social: unreadDMs !== "" && !isNaN(unreadDMs) ? (unreadDMs === 0 ? "Healthy" : unreadDMs <= 2 ? "Watch" : "At Risk") : "",
      };
    }
  }

  // Merge, filter to business days only, and sort
  const allDates = new Set([...Object.keys(badgeByDate), ...Object.keys(metricsByDate)]);
  const days = [];
  allDates.forEach(key => {
    const b = badgeByDate[key] || {};
    const m = metricsByDate[key] || {};
    const parts = key.split("-");
    const dt = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    // Skip weekends — showing Fri status for Sat/Sun overweights that day's health
    const dow = dt.getDay(); // 0=Sun, 6=Sat
    if (dow === 0 || dow === 6) return;
    const dayNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    days.push({
      date: key,
      dateLabel: dayNames[dow] + " " + parseInt(parts[1]) + "/" + parseInt(parts[2]),
      email: b.email || m.email || "",
      phone: b.phone || m.phone || "",
      social: b.social || m.social || "",
    });
  });
  days.sort((a, b) => a.date.localeCompare(b.date));
  return days;
}

// Build HTML inline mini-charts for health trend data.
// Returns an HTML string with a bar chart showing daily status for each channel.
function buildHealthTrendHtml(trendDays, navy, borderColor) {
  if (!trendDays || trendDays.length === 0) return "";
  const cardBg = "#F8F7F6";
  const statusColor = { "Healthy": "#2E7D32", "Watch": "#F57F17", "At Risk": "#C62828", "Holiday": "#D5D5D5" };
  const barHeight = { "Healthy": "100%", "Watch": "66%", "At Risk": "33%", "Holiday": "100%" };

  function miniChart(channel, label) {
    let barsHtml = "";
    trendDays.forEach(d => {
      const status = d[channel] || "";
      const color = statusColor[status] || "#E0E0E0";
      const height = barHeight[status] || "10%";
      const tooltip = d.dateLabel + ": " + (status || "No data");
      barsHtml += `<td style="vertical-align:bottom;padding:0 2px;width:${Math.floor(100 / trendDays.length)}%;" title="${tooltip}">
        <div style="background:${color};height:${height};min-height:4px;border-radius:2px 2px 0 0;"></div>
      </td>`;
    });

    // Day labels row
    let labelsHtml = "";
    trendDays.forEach(d => {
      const isHoliday = (d.email === "Holiday" || d.phone === "Holiday" || d.social === "Holiday");
      const parts = d.dateLabel.split(" ");
      const dayName = parts[0] + (isHoliday ? "*" : "");
      const dateNum = parts[1] || "";
      const label = dateNum ? `${dayName}<br><span style="font-size:8px;color:#BBB;">${dateNum}</span>` : dayName;
      labelsHtml += `<td style="text-align:center;font-size:9px;color:${isHoliday ? '#AAA' : '#999'};padding-top:2px;">${label}</td>`;
    });

    return `<div style="margin-bottom:12px;">
      <div style="font-size:11px;font-weight:bold;color:${navy};margin-bottom:6px;">${label}</div>
      <table style="width:100%;border-collapse:collapse;height:40px;"><tr>${barsHtml}</tr></table>
      <table style="width:100%;border-collapse:collapse;"><tr>${labelsHtml}</tr></table>
    </div>`;
  }

  // Legend
  // Check if any day is a holiday to conditionally show the legend item
  const hasHoliday = trendDays.some(d => d.email === "Holiday" || d.phone === "Holiday" || d.social === "Holiday");
  const hasNoData = trendDays.some(d => {
    return (d.email === "" || d.phone === "" || d.social === "") && d.email !== "Holiday";
  });
  const legend = `<div style="text-align:center;margin-top:8px;font-size:10px;">
    <span style="display:inline-block;background:#2E7D32;color:#fff;padding:1px 8px;border-radius:8px;margin:0 4px;">Healthy</span>
    <span style="display:inline-block;background:#F57F17;color:#fff;padding:1px 8px;border-radius:8px;margin:0 4px;">Watch</span>
    <span style="display:inline-block;background:#C62828;color:#fff;padding:1px 8px;border-radius:8px;margin:0 4px;">At Risk</span>
    ${hasHoliday ? '<span style="display:inline-block;background:#D5D5D5;color:#666;padding:1px 8px;border-radius:8px;margin:0 4px;">* Holiday</span>' : ''}
    ${hasNoData ? '<span style="display:inline-block;background:#E0E0E0;color:#666;padding:1px 8px;border-radius:8px;margin:0 4px;">No data</span>' : ''}
  </div>`;

  return `<div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
    <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Health Trends</h2>
    ${miniChart("email", "Email")}
    ${miniChart("phone", "Phone")}
    ${miniChart("social", "Social")}
    ${legend}
  </div>`;
}

// ─── POSITIVE HIGHLIGHTS ───
// Pulls positive CSAT comments for a date range from Nicereply + PostCall Log.
// Returns HTML with: 1 general experience quote at top, up to 1 per agent.
function buildPositiveHighlights(startDate, endDate, navy, borderColor) {
  const props = PropertiesService.getScriptProperties();
  const allComments = [];

  // 1. Nicereply email CSAT comments
  try {
    const nrToken = props.getProperty("NICEREPLY_TOKEN");
    if (nrToken) {
      const nrAuth = "Basic " + Utilities.base64Encode(nrToken);
      const sinceISO = startDate + "T00:00:00Z";
      const untilISO = endDate + "T23:59:59Z";
      let nrPage = 1;
      let nrMore = true;
      while (nrMore && nrPage <= 5) {
        const url = `https://api.nicereply.com/responses?created_after=${encodeURIComponent(sinceISO)}&created_before=${encodeURIComponent(untilISO)}&per_page=50&page=${nrPage}`;
        const resp = UrlFetchApp.fetch(url, { method: "get", headers: { "Authorization": nrAuth }, muteHttpExceptions: true });
        if (resp.getResponseCode() !== 200) break;
        const body = JSON.parse(resp.getContentText());
        const responses = body.data || [];
        responses.forEach(r => {
          const answers = r.answers || [];
          const CSAT_QID = "86bae330-e8bc-4fa3-9af9-91eb2459d348";
          const csatAnswer = answers.find(a => a.question_id === CSAT_QID) || answers.find(a => a.question_type === "SCALE");
          const score = csatAnswer && csatAnswer.scale ? csatAnswer.scale.value : 0;
          const openAnswer = answers.find(a => a.question_type === "OPEN_ENDED");
          const comment = openAnswer ? (openAnswer.open_ended ? openAnswer.open_ended.value : "") : "";
          if (score >= 4 && comment && comment.trim().length > 10) {
            // Resolve agent from ticket
            let agentName = "";
            if (r.ticket_id) {
              try {
                const zdToken = props.getProperty("ZENDESK_TOKEN");
                const subdomain = CONFIG.zendesk.subdomain;
                const zdAuth = "Basic " + Utilities.base64Encode(zdToken);
                const tResp = UrlFetchApp.fetch(`https://${subdomain}.zendesk.com/api/v2/tickets/${r.ticket_id}.json`, {
                  method: "get", headers: { "Authorization": zdAuth, "Content-Type": "application/json" }, muteHttpExceptions: true
                });
                if (tResp.getResponseCode() === 200) {
                  const ticket = JSON.parse(tResp.getContentText()).ticket;
                  const assigneeId = ticket.assignee_id;
                  if (assigneeId) {
                    const uResp = UrlFetchApp.fetch(`https://${subdomain}.zendesk.com/api/v2/users/${assigneeId}.json`, {
                      method: "get", headers: { "Authorization": zdAuth, "Content-Type": "application/json" }, muteHttpExceptions: true
                    });
                    if (uResp.getResponseCode() === 200) {
                      agentName = JSON.parse(uResp.getContentText()).user.name || "";
                    }
                  }
                }
              } catch (e) { /* skip agent lookup */ }
            }
            allComments.push({ comment: comment.trim(), agent: agentName, source: "email", score });
          }
        });
        const pag = body.pagination || {};
        nrMore = pag.total_pages ? nrPage < pag.total_pages : responses.length >= 50;
        nrPage++;
      }
    }
  } catch (e) {
    Logger.log("Nicereply comment fetch failed: " + e.toString());
  }

  // 2. PostCall Log comments (from sheet if available)
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pcSheet = ss.getSheetByName("PostCall Log");
    if (pcSheet && pcSheet.getLastRow() > 1) {
      const pcData = pcSheet.getRange(2, 1, pcSheet.getLastRow() - 1, 12).getValues();
      pcData.forEach(row => {
        const ts = new Date(row[0]);
        const tsStr = Utilities.formatDate(ts, Session.getScriptTimeZone(), "yyyy-MM-dd");
        if (tsStr < startDate || tsStr > endDate) return;
        const event = String(row[1] || "");
        if (event !== "survey-completed") return;
        const csat = Number(row[2]) || 0;
        const comment = String(row[9] || "").trim();
        const agentName = String(row[4] || "");
        if (csat >= 4 && comment && comment.length > 10) {
          allComments.push({ comment, agent: agentName, source: "phone", score: csat });
        }
      });
    }
  } catch (e) {
    Logger.log("PostCall comment fetch failed: " + e.toString());
  }

  if (allComments.length === 0) return "";

  // Use LLM to pick the best quotes
  const apiKey = props.getProperty("ANTHROPIC_API_KEY");
  if (!apiKey) {
    Logger.log("No API key for positive highlights LLM call");
    return "";
  }

  const agentNames = CONFIG.agents.map(a => a.split(" ")[0]);
  const commentList = allComments.map((c, i) =>
    `[${i}] Agent: ${c.agent || "unknown"} | Score: ${c.score}/5 | Source: ${c.source} | "${c.comment}"`
  ).join("\n");

  try {
    const llmResp = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method: "post",
      headers: {
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "Content-Type": "application/json",
      },
      payload: JSON.stringify({
        model: "claude-haiku-4-5-20251001",
        max_tokens: 500,
        messages: [{ role: "user", content: `Pick the best positive customer quotes from this week's CSAT feedback for a support team weekly email.

AGENTS: ${agentNames.join(", ")}

ALL POSITIVE COMMENTS THIS WEEK:
${commentList}

Pick:
1. ONE best overall experience quote (warm, specific, representative of great service -- not about a specific agent)
2. Up to ONE quote per agent that specifically praises them (skip agents with no good quotes)

Return valid JSON only:
{"general": {"index": 0, "quote": "the comment"}, "agents": {"AgentA": {"index": 1, "quote": "the comment"}, "AgentB": {"index": 2, "quote": "the comment"}}}

If no good general quote exists, set general to null. If an agent has no praiseworthy comment, omit them from agents. Prefer specific, descriptive praise over generic "great service" comments.` }],
      }),
      muteHttpExceptions: true,
    });

    if (llmResp.getResponseCode() !== 200) {
      Logger.log("Positive highlights LLM error: " + llmResp.getResponseCode());
      return "";
    }

    const llmData = JSON.parse(llmResp.getContentText());
    const raw = llmData.content[0].text.trim();
    let picks = {};
    try {
      picks = JSON.parse(raw);
    } catch (e) {
      const jsonMatch = raw.match(/\{[\s\S]*\}/);
      if (jsonMatch) picks = JSON.parse(jsonMatch[0]);
      else return "";
    }

    let html = `<div style="background:#F0F7F0;border:1px solid #C8E6C9;border-radius:6px;padding:16px;margin-bottom:16px;">
      <h2 style="margin:0 0 10px;font-size:15px;color:#2E7D32;">Customer Kudos</h2>`;

    // General highlight
    if (picks.general && picks.general.quote) {
      const escaped = picks.general.quote.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
      html += `<div style="font-size:13px;color:#333;font-style:italic;border-left:3px solid #2E7D32;padding-left:10px;margin-bottom:12px;">"${escaped}"</div>`;
    }

    // Per-agent highlights
    if (picks.agents) {
      const agentEntries = Object.entries(picks.agents);
      if (agentEntries.length > 0) {
        agentEntries.forEach(([name, c]) => {
          if (c && c.quote) {
            const escaped = c.quote.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
            html += `<div style="font-size:12px;color:#555;margin-bottom:6px;"><strong>${name}:</strong> <span style="font-style:italic;">"${escaped}"</span></div>`;
          }
        });
      }
    }

    html += `</div>`;

    // Only return if we have at least one quote
    if ((!picks.general || !picks.general.quote) && (!picks.agents || Object.keys(picks.agents).length === 0)) return "";
    return html;

  } catch (e) {
    Logger.log("Positive highlights LLM call failed: " + e.toString());
    return "";
  }
}

// ─── WEEKLY SUMMARY EMAIL ───
// Triggered daily at 6:15pm — only sends on the last working day of the week.
function checkAndSendWeeklySummary() {
  loadThresholds();
  const tz = CONFIG.businessHours.timezone;
  const now = new Date();
  const holidays = getDeakoHolidays();

  if (!isLastWorkingDayOfWeek(now, holidays)) {
    Logger.log("Not the last working day of the week — skipping weekly summary");
    return;
  }

  sendWeeklySummary(now, tz);
}

function sendWeeklySummary(now, tz, recipientOverride) {
  const props = PropertiesService.getScriptProperties();
  const recipients = recipientOverride || props.getProperty("RECAP_RECIPIENTS") || "";
  if (!recipients) { Logger.log("RECAP_RECIPIENTS not set — skipping weekly summary"); return; }

  const thisWeek = getThisWeekRange(now, tz);
  const lastWeek = getLastWeekRange(now, tz);

  const thisData = readMetricsLog(thisWeek.start, thisWeek.end);
  const lastData = readMetricsLog(lastWeek.start, lastWeek.end);

  const curr = aggregateMetrics(thisData);
  const prev = aggregateMetrics(lastData);

  if (!curr || curr.days === 0) {
    Logger.log("No metrics data for this week — skipping weekly summary");
    return;
  }

  const navy = "#1B3747";
  const cardBg = "#F8F7F6";
  const borderColor = "#E1DFDD";
  const weekLabel = `${formatDateShort(thisWeek.start, tz)} – ${formatDateShort(thisWeek.end, tz)}`;
  const prevWeekLabel = prev ? `${formatDateShort(lastWeek.start, tz)} – ${formatDateShort(lastWeek.end, tz)}` : "";
  const noData = '<span style="color:#AAA;">no prior data</span>';

  const colLabel = "width:50%;padding:4px 0;";
  const colThis = "width:25%;text-align:right;padding:4px 0;";
  const colPrev = "width:25%;text-align:right;padding:4px 0;color:#888;font-size:12px;";

  function prevCell(currVal, prevVal, unit, opts) {
    if (!prev) return `<td style="${colPrev}">${noData}</td>`;
    return `<td style="${colPrev}">${fmtComp(prevVal, null, unit, opts)}</td>`;
  }

  // ── Compute health badge summary for subject line ──
  let healthSummary = "";
  try {
    const trendDays = readHealthTrends(thisWeek.start, thisWeek.end, tz);
    if (trendDays && trendDays.length > 0) {
      // Find the last non-holiday day with actual status data
      const realDays = trendDays.filter(d => d.email && d.email !== "Holiday" && d.email !== "");
      if (realDays.length > 0) {
        const lastDay = realDays[realDays.length - 1];
        healthSummary = ` — Email: ${lastDay.email} | Phone: ${lastDay.phone} | Social: ${lastDay.social}`;
      }
    }
  } catch (e) { /* non-fatal */ }

  let html = `
  <div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;color:#1D1D1D;">
    <div style="background:${navy};color:#fff;padding:16px 24px;border-radius:8px 8px 0 0;">
      <h1 style="margin:0;font-size:20px;">CS AI Recap - Weekly</h1>
      <p style="margin:4px 0 0;font-size:13px;color:#C3D3D7;">${weekLabel} (${curr.days} working days)</p>
    </div>
    <div style="padding:20px 24px;">`;

  // ── 0. POSITIVE HIGHLIGHT ──
  try {
    const positiveHtml = buildPositiveHighlights(thisWeek.start, thisWeek.end, navy, borderColor);
    if (positiveHtml) html += positiveHtml;
  } catch (e) {
    Logger.log("Positive highlights failed (non-fatal): " + e.toString());
  }

  // ── 1. HEALTH TRENDS (visual first) ──
  try {
    const trendDays = readHealthTrends(thisWeek.start, thisWeek.end, tz);
    // Insert markers for non-working days (holidays/weekends) so the chart shows the full week
    const holidays = getDeakoHolidays();
    const dayNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    const fullWeekDays = [];
    const existingDates = new Set((trendDays || []).map(d => d.date));
    // Walk Mon-Fri of the week
    const weekStart = new Date(thisWeek.start + "T12:00:00");
    for (let i = 0; i < 5; i++) {
      const d = new Date(weekStart);
      d.setDate(weekStart.getDate() + i);
      const dateStr = Utilities.formatDate(d, tz, "yyyy-MM-dd");
      const dow = d.getDay();
      const dayLabel = dayNames[dow];
      if (existingDates.has(dateStr)) {
        // Use the real data
        const existing = trendDays.find(td => td.date === dateStr);
        if (existing) fullWeekDays.push(existing);
      } else if (holidays.has(dateStr)) {
        // Holiday -- show as labeled gap
        const mm = d.getMonth() + 1;
        const dd = d.getDate();
        fullWeekDays.push({ date: dateStr, dateLabel: dayLabel + " " + mm + "/" + dd + " (Holiday)", email: "Holiday", phone: "Holiday", social: "Holiday" });
      } else if (dow === 0 || dow === 6) {
        // Weekend -- skip
      } else {
        // Missing working day
        fullWeekDays.push({ date: dateStr, dateLabel: dayLabel, email: "", phone: "", social: "" });
      }
    }
    html += buildHealthTrendHtml(fullWeekDays.length > 0 ? fullWeekDays : trendDays, navy, borderColor);
  } catch (e) {
    Logger.log("Health trends in weekly email failed (non-fatal): " + e.toString());
  }

  // ── 2. AI RECAP — WEEKLY THEME TRENDS ──
  try {
    const trendHtml = analyzeThemeTrends(thisWeek.start, thisWeek.end, "Week of " + weekLabel);
    if (trendHtml) {
      html += `<div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
        <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">AI Recap - Weekly Theme Trends</h2>
        <div style="font-size:13px;line-height:1.6;">${trendHtml}</div>
        <div style="margin-top:10px;font-size:10px;color:#999;">Analysis powered by Claude AI across all tickets and call intents</div>
      </div>`;
    }
  } catch (e) {
    Logger.log("Weekly theme trend analysis failed (non-fatal): " + e.toString());
  }

  // Table header row
  const headerRow = `<tr style="border-bottom:1px solid ${borderColor};">
    <td style="${colLabel}"></td>
    <td style="${colThis}font-size:11px;color:#888;font-weight:bold;">This Week</td>
    <td style="${colPrev}font-size:11px;font-weight:bold;">${prev ? "Last Week" : ""}</td></tr>`;

  // ── 3. EMAIL / ZENDESK SECTION ──
  html += `<div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
    <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Email (Zendesk)</h2>
    <table style="width:100%;font-size:13px;border-collapse:collapse;">
      ${headerRow}
      <tr><td style="${colLabel}">Tickets Created</td><td style="${colThis}font-weight:bold;">${curr.totalCreated}</td>${prevCell(curr.totalCreated, prev ? prev.totalCreated : null, "")}</tr>
      <tr><td style="${colLabel}">Tickets Solved</td><td style="${colThis}font-weight:bold;">${curr.totalSolved}</td>${prevCell(curr.totalSolved, prev ? prev.totalSolved : null, "")}</tr>
      <tr><td style="${colLabel}">Avg Open Tickets</td><td style="${colThis}font-weight:bold;">${curr.avgOpenTickets !== null ? curr.avgOpenTickets : '<span style="color:#AAA;font-weight:normal;">no data</span>'}</td>${prevCell(curr.avgOpenTickets, prev && prev.queueStateDays >= Math.ceil(prev.days / 2) ? prev.avgOpenTickets : null, "")}</tr>
      <tr><td style="${colLabel}">Avg Past SLA</td><td style="${colThis}font-weight:bold;${curr.avgPastSla !== null && curr.avgPastSla > 0 ? 'color:#C62828;' : ''}">${curr.avgPastSla !== null ? curr.avgPastSla : '<span style="color:#AAA;font-weight:normal;">no data</span>'}</td>${prevCell(curr.avgPastSla, prev && prev.queueStateDays >= Math.ceil(prev.days / 2) ? prev.avgPastSla : null, "")}</tr>
      <tr><td style="${colLabel}">Avg Unassigned</td><td style="${colThis}font-weight:bold;">${curr.avgUnassigned !== null ? curr.avgUnassigned : '<span style="color:#AAA;font-weight:normal;">no data</span>'}</td>${prevCell(curr.avgUnassigned, prev && prev.queueStateDays >= Math.ceil(prev.days / 2) ? prev.avgUnassigned : null, "")}</tr>`;

  html += `</table>`;

  // Per-agent email table (horizontal: agent names as rows, Solved as column)
  const wAgentHdrStyle = `font-size:11px;color:#888;font-weight:bold;text-align:center;padding:4px 6px;border-bottom:1px solid #E1DFDD;`;
  const wAgentNameStyle = `font-size:12px;font-weight:bold;color:${navy};padding:4px 6px;`;
  const wAgentCellStyle = `font-size:12px;text-align:center;padding:4px 6px;`;
  const wAgentPrevStyle = `font-size:11px;text-align:center;padding:4px 6px;color:#888;`;

  html += `<div style="margin-top:12px;font-size:12px;color:${navy};font-weight:bold;border-top:1px solid #E1DFDD;padding-top:10px;">Per Agent</div>
    <table style="width:100%;font-size:12px;border-collapse:collapse;margin-top:6px;">
      <tr>
        <td style="${wAgentHdrStyle}text-align:left;"></td>
        <td style="${wAgentHdrStyle}">Solved</td>
        ${prev ? `<td style="${wAgentHdrStyle}">Prev</td>` : ""}
        ${prev ? `<td style="${wAgentHdrStyle}">Change</td>` : ""}
      </tr>`;
  curr.agentSolved.forEach((a, i) => {
    const prevVal = prev ? prev.agentSolved[i].total : null;
    const delta = prevVal !== null ? a.total - prevVal : null;
    const deltaStr = delta !== null ? (delta >= 0 ? `<span style="color:#2E7D32;">+${delta}</span>` : `<span style="color:#C62828;">${delta}</span>`) : "-";
    const prevStr = prevVal !== null ? String(prevVal) : "-";
    html += `<tr>
        <td style="${wAgentNameStyle}">${a.name.split(" ")[0]}</td>
        <td style="${wAgentCellStyle}">${a.total}</td>
        ${prev ? `<td style="${wAgentPrevStyle}">${prevStr}</td>` : ""}
        ${prev ? `<td style="${wAgentPrevStyle}">${deltaStr}</td>` : ""}
      </tr>`;
  });
  html += `</table></div>`;

  // ── PHONE SECTION ──
  html += `<div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
    <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Phone (Aircall)</h2>
    <table style="width:100%;font-size:13px;border-collapse:collapse;">
      ${headerRow}
      <tr><td style="${colLabel}">Avg Answer Rate</td><td style="${colThis}font-weight:bold;color:${curr.avgAnswerRate >= 75 ? '#2E7D32' : '#C62828'};">${curr.avgAnswerRate}%</td>${prevCell(curr.avgAnswerRate, prev ? prev.avgAnswerRate : null, "%")}</tr>
      <tr><td style="${colLabel}">Inbound Answered</td><td style="${colThis}font-weight:bold;">${curr.totalInbound}</td>${prevCell(curr.totalInbound, prev ? prev.totalInbound : null, "")}</tr>
      <tr><td style="${colLabel}">Sent to Answer Service</td><td style="${colThis}font-weight:bold;">${curr.totalForwarded}</td>${prevCell(curr.totalForwarded, prev ? prev.totalForwarded : null, "")}</tr>
      <tr><td style="${colLabel}">Total Outbound</td><td style="${colThis}font-weight:bold;">${curr.totalOutbound}</td>${prevCell(curr.totalOutbound, prev ? prev.totalOutbound : null, "")}</tr>
`;

  html += `</table>`;

  // Per-agent phone table (horizontal: agent names as rows, metrics as columns)
  html += `<div style="margin-top:12px;font-size:12px;color:${navy};font-weight:bold;border-top:1px solid #E1DFDD;padding-top:10px;">Per Agent</div>
    <table style="width:100%;font-size:12px;border-collapse:collapse;margin-top:6px;">
      <tr>
        <td style="${wAgentHdrStyle}text-align:left;"></td>
        <td style="${wAgentHdrStyle}">Inbound</td>
        <td style="${wAgentHdrStyle}">Outbound</td>
        <td style="${wAgentHdrStyle}">In Talk</td>
        <td style="${wAgentHdrStyle}">Out Talk</td>
      </tr>`;
  curr.agentInbound.forEach((a, i) => {
    const inTotal = a.total;
    const outTotal = curr.agentOutbound[i].total;
    const inTalk = curr.agentInTalk ? curr.agentInTalk[i].total : 0;
    const outTalk = curr.agentOutTalk ? curr.agentOutTalk[i].total : 0;
    html += `<tr>
        <td style="${wAgentNameStyle}">${a.name.split(" ")[0]}</td>
        <td style="${wAgentCellStyle}">${inTotal}</td>
        <td style="${wAgentCellStyle}">${outTotal}</td>
        <td style="${wAgentCellStyle}">${formatTalkTime(inTalk)}</td>
        <td style="${wAgentCellStyle}">${formatTalkTime(outTalk)}</td>
      </tr>`;
  });
  html += `</table></div>`;

  // ── CSAT SECTION ──
  if (curr.totalCsatResponses > 0 || curr.totalPhoneCsatResponses > 0) {
    html += `<div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
      <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Customer Satisfaction</h2>
      <table style="width:100%;font-size:13px;border-collapse:collapse;">
        ${headerRow}`;
    if (curr.avgCsatPct !== null && !isNaN(curr.avgCsatPct)) {
      const prevCsatDisplay = prev && prev.avgCsatPct !== null && !isNaN(prev.avgCsatPct) ? Math.round(prev.avgCsatPct) + "% (" + prev.totalCsatResponses + ")" : noData;
      html += `<tr><td style="${colLabel}">Email CSAT</td><td style="${colThis}font-weight:bold;">${Math.round(curr.avgCsatPct)}% (${curr.totalCsatResponses} reviews)</td><td style="${colPrev}">${prevCsatDisplay}</td></tr>`;
    }
    if (curr.avgPhoneCsatPct !== null && !isNaN(curr.avgPhoneCsatPct)) {
      const prevPCsatDisplay = prev && prev.avgPhoneCsatPct !== null && !isNaN(prev.avgPhoneCsatPct) ? Math.round(prev.avgPhoneCsatPct) + "% (" + prev.totalPhoneCsatResponses + ")" : noData;
      html += `<tr><td style="${colLabel}">Phone CSAT</td><td style="${colThis}font-weight:bold;">${Math.round(curr.avgPhoneCsatPct)}% (${curr.totalPhoneCsatResponses} reviews)</td><td style="${colPrev}">${prevPCsatDisplay}</td></tr>`;
    }
    html += `</table></div>`;
  }

  // ── SOCIAL SECTION ──
  html += `<div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
    <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Social (Meta Business Suite)</h2>
    <table style="width:100%;font-size:13px;border-collapse:collapse;">
      ${headerRow}
      <tr><td style="${colLabel}">Avg Unread DMs at EOD</td><td style="${colThis}font-weight:bold;">${curr.avgUnreadDMs}</td>${prevCell(curr.avgUnreadDMs, prev ? prev.avgUnreadDMs : null, "")}</tr>
    </table></div>`;

  // ── TEXT MESSAGES SECTION ──
  html += `<div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
    <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Text Messages</h2>
    <table style="width:100%;font-size:13px;border-collapse:collapse;">
      ${headerRow}
      <tr><td style="${colLabel}">Total In/Out</td><td style="${colThis}font-weight:bold;">${curr.totalSmsInbound} / ${curr.totalSmsOutbound}</td>${prev ? `<td style="${colPrev}">${prev.totalSmsInbound} / ${prev.totalSmsOutbound}</td>` : `<td style="${colPrev}">${noData}</td>`}</tr>
    </table></div>`;

  // Footer with health status definitions
  const th = CONFIG.thresholds;
  const slaHours = th.oldestUnanswered.green;
  html += `
      <div style="border-top:1px solid ${borderColor};margin-top:20px;padding:16px 0 4px;font-size:11px;color:#999;">
        <div style="margin-bottom:8px;font-weight:bold;color:#888;">Health Status Thresholds</div>
        <div style="margin-bottom:4px;">Email: Healthy = 0-5 tickets past ${slaHours}h SLA · Watch = 6-10 past SLA · At Risk = 11+ past SLA</div>
        <div style="margin-bottom:4px;">Phone: Healthy = ${th.phoneAnswerRate.green}%+ answer rate · Watch = ${th.phoneAnswerRate.yellow}-${th.phoneAnswerRate.green - 1}% · At Risk = below ${th.phoneAnswerRate.yellow}%</div>
        <div style="margin-bottom:4px;">Social: Healthy = oldest unread DM under ${th.socialResponseTime.green / 60}h · Watch = ${th.socialResponseTime.green / 60}-${th.socialResponseTime.yellow / 60}h · At Risk = over ${th.socialResponseTime.yellow / 60}h</div>
      </div>`;
  html += `<div style="text-align:center;font-size:11px;color:#999;padding-top:8px;">
      CS Visibility · AI Recap v2.5.4 · Weekly Summary · ${weekLabel}
      ${prev ? '<br>Comparison: ' + prevWeekLabel : '<br>No prior week data available for comparison'}
    </div>
    </div>
  </div>`;

  const subject = `CS AI Recap - Weekly - ${weekLabel}`;

  GmailApp.sendEmail(recipients, subject, "View this email with HTML enabled.", {
    htmlBody: html,
    name: "CS Visibility",
  });

  Logger.log("Weekly summary sent to: " + recipients);
}

// ─── MONTHLY SUMMARY EMAIL ───
// Triggered daily at 6:30pm — only sends on the last business day of the month.
function checkAndSendMonthlySummary() {
  loadThresholds();
  const tz = CONFIG.businessHours.timezone;
  const now = new Date();
  const holidays = getDeakoHolidays();

  if (!isLastBusinessDayOfMonth(now, holidays)) {
    Logger.log("Not the last business day of the month — skipping monthly summary");
    return;
  }

  sendMonthlySummary(now, tz);
}

function sendMonthlySummary(now, tz, recipientOverride) {
  const props = PropertiesService.getScriptProperties();
  const recipients = recipientOverride || props.getProperty("RECAP_RECIPIENTS") || "";
  if (!recipients) { Logger.log("RECAP_RECIPIENTS not set — skipping monthly summary"); return; }

  const todayDate = new Date(Utilities.formatDate(now, tz, "yyyy-MM-dd") + "T12:00:00");
  const thisMonth = todayDate.getMonth() + 1; // 1-indexed
  const thisYear = todayDate.getFullYear();
  const lastMonth = thisMonth === 1 ? 12 : thisMonth - 1;
  const lastYear = thisMonth === 1 ? thisYear - 1 : thisYear;

  const thisRange = getMonthRange(thisYear, thisMonth, tz);
  const lastRange = getMonthRange(lastYear, lastMonth, tz);

  const thisData = readMetricsLog(thisRange.start, thisRange.end);
  const lastData = readMetricsLog(lastRange.start, lastRange.end);

  const curr = aggregateMetrics(thisData);
  const prev = aggregateMetrics(lastData);

  if (!curr || curr.days === 0) {
    Logger.log("No metrics data for this month — skipping monthly summary");
    return;
  }

  const navy = "#1B3747";
  const cardBg = "#F8F7F6";
  const borderColor = "#E1DFDD";
  const monthNames = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const monthLabel = monthNames[thisMonth - 1] + " " + thisYear;
  const prevMonthLabel = prev ? monthNames[lastMonth - 1] + " " + lastYear : "";
  const noData = '<span style="color:#AAA;">no prior data</span>';

  const monthStart = thisYear + "-" + String(thisMonth).padStart(2, "0") + "-01";
  const monthEndDay = new Date(thisYear, thisMonth, 0).getDate();
  const monthEnd = thisYear + "-" + String(thisMonth).padStart(2, "0") + "-" + String(monthEndDay).padStart(2, "0");

  const colLabel = "width:50%;padding:4px 0;";
  const colThis = "width:25%;text-align:right;padding:4px 0;";
  const colPrev = "width:25%;text-align:right;padding:4px 0;color:#888;font-size:12px;";

  function prevCell(currVal, prevVal, unit) {
    if (!prev) return `<td style="${colPrev}">${noData}</td>`;
    return `<td style="${colPrev}">${prevVal !== null && prevVal !== undefined ? (typeof prevVal === 'number' ? Math.round(prevVal * 10) / 10 : prevVal) + (unit || "") : "—"}</td>`;
  }

  // Table header row
  const headerRow = `<tr style="border-bottom:1px solid ${borderColor};">
    <td style="${colLabel}"></td>
    <td style="${colThis}font-size:11px;color:#888;font-weight:bold;">${monthLabel}</td>
    <td style="${colPrev}font-size:11px;font-weight:bold;">${prevMonthLabel}</td></tr>`;

  let html = `
  <div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;color:#1D1D1D;">
    <div style="background:${navy};color:#fff;padding:16px 24px;border-radius:8px 8px 0 0;">
      <h1 style="margin:0;font-size:20px;">CS AI Recap - Monthly</h1>
      <p style="margin:4px 0 0;font-size:13px;color:#C3D3D7;">${monthLabel} (${curr.days} working days)</p>
    </div>
    <div style="padding:20px 24px;">`;

  // ── HEALTH TRENDS MINI-CHARTS ──
  try {
    const trendDays = readHealthTrends(thisRange.start, thisRange.end, tz);
    html += buildHealthTrendHtml(trendDays, navy, borderColor);
  } catch (e) {
    Logger.log("Health trends in monthly email failed (non-fatal): " + e.toString());
  }

  // ── AI RECAP — MONTHLY TREND ANALYSIS ──
  try {
    const mStart = thisYear + "-" + String(thisMonth).padStart(2, "0") + "-01";
    const mEndDay = new Date(thisYear, thisMonth, 0).getDate();
    const mEnd = thisYear + "-" + String(thisMonth).padStart(2, "0") + "-" + String(mEndDay).padStart(2, "0");
    const trendHtml = analyzeThemeTrends(mStart, mEnd, monthLabel);
    if (trendHtml) {
      html += `<div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
        <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">AI Recap - Monthly Theme Trends</h2>
        <div style="font-size:13px;line-height:1.6;">${trendHtml}</div>
        <div style="margin-top:10px;font-size:10px;color:#999;">Analysis powered by Claude AI across all tickets and call intents</div>
      </div>`;
    }
  } catch (e) {
    Logger.log("Monthly theme trend analysis failed (non-fatal): " + e.toString());
  }

  // ── FEATURE REQUESTS ──
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const taxSheet = ss.getSheetByName("Ticket Taxonomy");
    if (taxSheet && taxSheet.getLastRow() > 1) {
      const taxData = taxSheet.getDataRange().getValues();
      const taxHeaders = taxData[0];
      const reqTypeCol = taxHeaders.indexOf("Request Type");
      const createdCol = taxHeaders.indexOf("Created");
      const subjectCol = taxHeaders.indexOf("Subject");
      const ticketIdCol = taxHeaders.indexOf("Ticket ID");
      const featureCol = taxHeaders.indexOf("Feature Requested");

      if (reqTypeCol >= 0) {
        const frTz = Session.getScriptTimeZone();
        const featureRequests = [];
        for (let i = 1; i < taxData.length; i++) {
          if (String(taxData[i][reqTypeCol]).trim() === "Feature Request") {
            const created = taxData[i][createdCol];
            const createdStr = created instanceof Date
              ? Utilities.formatDate(created, frTz, "yyyy-MM-dd")
              : String(created || "").substring(0, 10);
            // Check if within this month
            if (createdStr >= monthStart && createdStr <= monthEnd) {
              featureRequests.push({
                id: taxData[i][ticketIdCol],
                subject: taxData[i][subjectCol] || "",
                feature: taxData[i][featureCol] || "",
              });
            }
          }
        }

        if (featureRequests.length > 0) {
          const subdomain = CONFIG.zendesk.subdomain;
          html += `<div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
            <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Feature & Product Requests</h2>
            <div style="font-size:13px;color:#555;margin-bottom:8px;">${featureRequests.length} feature request${featureRequests.length > 1 ? 's' : ''} this month:</div>`;
          featureRequests.forEach(fr => {
            const escaped = (fr.feature || fr.subject || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
            const link = fr.id ? `<a href="https://${subdomain}.zendesk.com/agent/tickets/${fr.id}" style="color:#7597A0;text-decoration:none;">#${fr.id}</a>` : "";
            html += `<div style="margin-bottom:6px;padding-left:12px;border-left:2px solid #CCC6C0;">
              <span style="font-size:12px;">${escaped}</span> ${link ? `<span style="font-size:11px;color:#999;">${link}</span>` : ""}
            </div>`;
          });
          html += `</div>`;
        }
      }
    }
  } catch (e) {
    Logger.log("Feature requests section failed (non-fatal): " + e.toString());
  }

  // ── EMAIL / ZENDESK ──
  html += `<div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
    <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Email (Zendesk)</h2>
    <table style="width:100%;font-size:13px;border-collapse:collapse;">
      ${headerRow}
      <tr><td style="${colLabel}">Tickets Created</td><td style="${colThis}font-weight:bold;">${curr.totalCreated}</td>${prevCell(curr.totalCreated, prev ? prev.totalCreated : null)}</tr>
      <tr><td style="${colLabel}">Tickets Solved</td><td style="${colThis}font-weight:bold;">${curr.totalSolved}</td>${prevCell(curr.totalSolved, prev ? prev.totalSolved : null)}</tr>
      <tr><td style="${colLabel}">Avg Daily Open Tickets</td><td style="${colThis}font-weight:bold;">${curr.avgOpenTickets !== null ? curr.avgOpenTickets : '<span style="color:#AAA;font-weight:normal;">no data</span>'}</td>${prevCell(curr.avgOpenTickets, prev ? prev.avgOpenTickets : null)}</tr>
      <tr><td style="${colLabel}">Avg Daily Past SLA</td><td style="${colThis}font-weight:bold;${curr.avgPastSla !== null && curr.avgPastSla > 0 ? 'color:#C62828;' : ''}">${curr.avgPastSla !== null ? curr.avgPastSla : '<span style="color:#AAA;font-weight:normal;">no data</span>'}</td>${prevCell(curr.avgPastSla, prev ? prev.avgPastSla : null)}</tr>
      <tr><td style="${colLabel}">Avg Daily Unassigned</td><td style="${colThis}font-weight:bold;">${curr.avgUnassigned !== null ? curr.avgUnassigned : '<span style="color:#AAA;font-weight:normal;">no data</span>'}</td>${prevCell(curr.avgUnassigned, prev ? prev.avgUnassigned : null)}</tr>
      <tr><td style="${colLabel}">Avg Solved/Day</td><td style="${colThis}font-weight:bold;">${(curr.totalSolved / curr.days).toFixed(1)}</td>${prev ? `<td style="${colPrev}">${(prev.totalSolved / prev.days).toFixed(1)}</td>` : `<td style="${colPrev}">${noData}</td>`}</tr>`;

  html += `</table>`;

  // Per-agent email table (horizontal layout)
  const mAHdr = `font-size:11px;color:#888;font-weight:bold;text-align:center;padding:4px 6px;border-bottom:1px solid #E1DFDD;`;
  const mAName = `font-size:12px;font-weight:bold;color:${navy};padding:4px 6px;`;
  const mACell = `font-size:12px;text-align:center;padding:4px 6px;`;
  const mAPrev = `font-size:11px;text-align:center;padding:4px 6px;color:#888;`;

  const shortMonthNames = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const prevMonthShort = prev ? shortMonthNames[lastMonth - 1] : "";

  html += `<div style="margin-top:12px;font-size:12px;color:${navy};font-weight:bold;border-top:1px solid #E1DFDD;padding-top:10px;">Per Agent</div>
    <table style="width:100%;font-size:12px;border-collapse:collapse;margin-top:6px;">
      <tr>
        <td style="${mAHdr}text-align:left;"></td>
        <td style="${mAHdr}">Solved</td>
        ${prev ? `<td style="${mAHdr}">${prevMonthShort}</td>` : ""}
        ${prev ? `<td style="${mAHdr}">Change</td>` : ""}
      </tr>`;
  curr.agentSolved.forEach((a, i) => {
    const prevVal = prev ? prev.agentSolved[i].total : null;
    const delta = prevVal !== null ? a.total - prevVal : null;
    const deltaStr = delta !== null ? (delta >= 0 ? `<span style="color:#2E7D32;">+${delta}</span>` : `<span style="color:#C62828;">${delta}</span>`) : "-";
    const prevStr = prevVal !== null ? String(prevVal) : "-";
    html += `<tr>
        <td style="${mAName}">${a.name.split(" ")[0]}</td>
        <td style="${mACell}">${a.total}</td>
        ${prev ? `<td style="${mAPrev}">${prevStr}</td>` : ""}
        ${prev ? `<td style="${mAPrev}">${deltaStr}</td>` : ""}
      </tr>`;
  });
  html += `</table></div>`;

  // ── PHONE ──
  html += `<div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
    <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Phone (Aircall)</h2>
    <table style="width:100%;font-size:13px;border-collapse:collapse;">
      ${headerRow}
      <tr><td style="${colLabel}">Avg Answer Rate</td><td style="${colThis}font-weight:bold;color:${curr.avgAnswerRate >= 75 ? '#2E7D32' : '#C62828'};">${curr.avgAnswerRate}%</td>${prevCell(curr.avgAnswerRate, prev ? prev.avgAnswerRate : null, "%")}</tr>
      <tr><td style="${colLabel}">Inbound Answered</td><td style="${colThis}font-weight:bold;">${curr.totalInbound}</td>${prevCell(curr.totalInbound, prev ? prev.totalInbound : null)}</tr>
      <tr><td style="${colLabel}">Sent to Answer Service</td><td style="${colThis}font-weight:bold;">${curr.totalForwarded}</td>${prevCell(curr.totalForwarded, prev ? prev.totalForwarded : null)}</tr>
      <tr><td style="${colLabel}">Total Outbound</td><td style="${colThis}font-weight:bold;">${curr.totalOutbound}</td>${prevCell(curr.totalOutbound, prev ? prev.totalOutbound : null)}</tr>`;

  html += `</table>`;

  // Per-agent phone table (horizontal layout)
  html += `<div style="margin-top:12px;font-size:12px;color:${navy};font-weight:bold;border-top:1px solid #E1DFDD;padding-top:10px;">Per Agent</div>
    <table style="width:100%;font-size:12px;border-collapse:collapse;margin-top:6px;">
      <tr>
        <td style="${mAHdr}text-align:left;"></td>
        <td style="${mAHdr}">Inbound</td>
        <td style="${mAHdr}">Outbound</td>
        <td style="${mAHdr}">In Talk</td>
        <td style="${mAHdr}">Out Talk</td>
      </tr>`;
  curr.agentInbound.forEach((a, i) => {
    const inTotal = a.total;
    const outTotal = curr.agentOutbound[i].total;
    const inTalk = curr.agentInTalk ? curr.agentInTalk[i].total : 0;
    const outTalk = curr.agentOutTalk ? curr.agentOutTalk[i].total : 0;
    html += `<tr>
        <td style="${mAName}">${a.name.split(" ")[0]}</td>
        <td style="${mACell}">${inTotal}</td>
        <td style="${mACell}">${outTotal}</td>
        <td style="${mACell}">${formatTalkTime(inTalk)}</td>
        <td style="${mACell}">${formatTalkTime(outTalk)}</td>
      </tr>`;
  });
  html += `</table></div>`;

  // ── CSAT ──
  if (curr.totalCsatResponses > 0 || curr.totalPhoneCsatResponses > 0) {
    html += `<div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
      <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Customer Satisfaction</h2>
      <table style="width:100%;font-size:13px;border-collapse:collapse;">
        ${headerRow}`;
    if (curr.avgCsatPct !== null) {
      html += `<tr><td style="${colLabel}">Email CSAT</td><td style="${colThis}font-weight:bold;">${Math.round(curr.avgCsatPct)}% (${curr.totalCsatResponses} reviews)</td>${prev && prev.avgCsatPct !== null ? `<td style="${colPrev}">${Math.round(prev.avgCsatPct)}% (${prev.totalCsatResponses})</td>` : `<td style="${colPrev}">${noData}</td>`}</tr>`;
    }
    if (curr.avgPhoneCsatPct !== null) {
      html += `<tr><td style="${colLabel}">Phone CSAT</td><td style="${colThis}font-weight:bold;">${Math.round(curr.avgPhoneCsatPct)}% (${curr.totalPhoneCsatResponses} reviews)</td>${prev && prev.avgPhoneCsatPct !== null ? `<td style="${colPrev}">${Math.round(prev.avgPhoneCsatPct)}% (${prev.totalPhoneCsatResponses})</td>` : `<td style="${colPrev}">${noData}</td>`}</tr>`;
    }
    html += `</table></div>`;
  }

  // ── SOCIAL ──
  html += `<div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
    <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Social (Meta Business Suite)</h2>
    <table style="width:100%;font-size:13px;border-collapse:collapse;">
      ${headerRow}
      <tr><td style="${colLabel}">Avg Unread DMs at EOD</td><td style="${colThis}font-weight:bold;">${curr.avgUnreadDMs}</td>${prevCell(curr.avgUnreadDMs, prev ? prev.avgUnreadDMs : null)}</tr>
    </table></div>`;

  // ── TEXT MESSAGES ──
  html += `<div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
    <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Text Messages</h2>
    <table style="width:100%;font-size:13px;border-collapse:collapse;">
      ${headerRow}
      <tr><td style="${colLabel}">Total In/Out</td><td style="${colThis}font-weight:bold;">${curr.totalSmsInbound} / ${curr.totalSmsOutbound}</td>${prev ? `<td style="${colPrev}">${prev.totalSmsInbound} / ${prev.totalSmsOutbound}</td>` : `<td style="${colPrev}">${noData}</td>`}</tr>
    </table></div>`;

  // Footer with health status definitions
  const th = CONFIG.thresholds;
  const slaHours = th.oldestUnanswered.green;
  html += `
      <div style="border-top:1px solid ${borderColor};margin-top:20px;padding:16px 0 4px;font-size:11px;color:#999;">
        <div style="margin-bottom:8px;font-weight:bold;color:#888;">Health Status Thresholds</div>
        <div style="margin-bottom:4px;">Email: Healthy = 0-5 tickets past ${slaHours}h SLA · Watch = 6-10 past SLA · At Risk = 11+ past SLA</div>
        <div style="margin-bottom:4px;">Phone: Healthy = ${th.phoneAnswerRate.green}%+ answer rate · Watch = ${th.phoneAnswerRate.yellow}-${th.phoneAnswerRate.green - 1}% · At Risk = below ${th.phoneAnswerRate.yellow}%</div>
        <div style="margin-bottom:4px;">Social: Healthy = oldest unread DM under ${th.socialResponseTime.green / 60}h · Watch = ${th.socialResponseTime.green / 60}-${th.socialResponseTime.yellow / 60}h · At Risk = over ${th.socialResponseTime.yellow / 60}h</div>
      </div>`;
  html += `<div style="text-align:center;font-size:11px;color:#999;padding-top:8px;">
      CS Visibility · AI Recap v2.5.4 · Monthly Summary · ${monthLabel}
      ${prev ? '<br>Comparison: ' + prevMonthLabel + ' (' + prev.days + ' working days)' : '<br>No prior month data available for comparison'}
    </div>
    </div>
  </div>`;

  const subject = `CS AI Recap - Monthly - ${monthLabel}`;

  GmailApp.sendEmail(recipients, subject, "View this email with HTML enabled.", {
    htmlBody: html,
    name: "CS Visibility",
  });

  Logger.log("Monthly summary sent to: " + recipients);
}

// Helper: format "2026-05-19" as "May 19"
function formatDateShort(dateStr, tz) {
  const d = new Date(dateStr + "T12:00:00");
  return Utilities.formatDate(d, tz, "MMM d");
}

// ─── TEST FUNCTIONS FOR WEEKLY / MONTHLY ───
function testWeeklySummary() {
  loadThresholds();
  const tz = CONFIG.businessHours.timezone;
  sendWeeklySummary(new Date(), tz);
}

function testMonthlySummary() {
  loadThresholds();
  const tz = CONFIG.businessHours.timezone;
  sendMonthlySummary(new Date(), tz);
}

// ═══════════════════════════════════════════════════════════
// AGENT PERFORMANCE DASHBOARDS
// Each CS agent gets a separate spreadsheet showing their
// today's metrics alongside the team median for benchmarking.
//
// SETUP:
// 1. Create a blank Google Sheet for each agent
// 2. Share each sheet with the agent (Viewer) and the script's
//    service account or your email (Editor)
// 3. Add Script Property AGENT_SHEETS as JSON:
//    {"Agent A":"SPREADSHEET_ID","Agent B":"SPREADSHEET_ID","Agent C":"SPREADSHEET_ID"}
// ═══════════════════════════════════════════════════════════

function updateAgentDashboards(zendesk, aircall, csat, postCall) {
  const props = PropertiesService.getScriptProperties();
  const sheetsJson = props.getProperty("AGENT_SHEETS");
  if (!sheetsJson) return;

  let agentSheets;
  try { agentSheets = JSON.parse(sheetsJson); } catch (e) {
    Logger.log("AGENT_SHEETS parse error: " + e.toString());
    return;
  }

  const agents = CONFIG.agents;
  const tz = CONFIG.businessHours.timezone;
  const now = new Date();
  const dateStr = Utilities.formatDate(now, tz, "EEEE, MMMM d, yyyy");
  const timeStr = Utilities.formatDate(now, tz, "h:mm a z");

  // ─── Gather per-agent metrics ───
  const agentMetrics = {};
  agents.forEach(agent => {
    const ac = (zendesk.agentCounts || {})[agent] || {};
    const as = (aircall.agentStats || {})[agent] || {};
    agentMetrics[agent] = {
      assigned:   ac.assigned || 0,
      pastSla:    ac.pastSla || 0,
      solved:     ac.handledToday || 0,
      inbound:    as.answered || 0,
      outbound:   as.outbound || 0,
      inTalk:     as.inboundTalkTime || 0,
      outTalk:    as.outboundTalkTime || 0,
    };
  });

  // ─── Per-agent oldest waiting tickets ───
  const agentOldest = {};
  agents.forEach(agent => {
    const myTickets = (zendesk.tickets || []).filter(t => {
      if (t.assignee === "Unassigned") return false;
      const matched = agents.find(ca => {
        const parts = ca.toLowerCase().split(/\s+/);
        const tLower = t.assignee.toLowerCase();
        return ca === t.assignee || parts.some(p => p.length > 1 && tLower.split(/\s+/).some(ap => ap === p));
      });
      return matched === agent;
    });
    agentOldest[agent] = myTickets
      .filter(t => t.waitBizMin > 0)
      .sort((a, b) => b.waitBizMin - a.waitBizMin)
      .slice(0, 5);
  });

  // ─── Per-agent flagged tickets (past SLA + high priority) ───
  const agentFlagged = {};
  agents.forEach(agent => {
    const pastSla = (zendesk.tickets || []).filter(t => {
      if (!t.pastSla) return false;
      const matched = agents.find(ca => {
        const parts = ca.toLowerCase().split(/\s+/);
        const tLower = t.assignee.toLowerCase();
        return ca === t.assignee || parts.some(p => p.length > 1 && tLower.split(/\s+/).some(ap => ap === p));
      });
      return matched === agent;
    });
    const highPri = (zendesk.flaggedTickets || []).filter(t => {
      const matched = agents.find(ca => {
        const parts = ca.toLowerCase().split(/\s+/);
        const tLower = (t.assignee || "").toLowerCase();
        return ca === t.assignee || parts.some(p => p.length > 1 && tLower.split(/\s+/).some(ap => ap === p));
      });
      return matched === agent;
    });
    // Deduplicate by ticket id
    const seen = new Set();
    const combined = [];
    pastSla.forEach(t => { if (!seen.has(t.id)) { seen.add(t.id); combined.push({ ...t, flag: "Past SLA" }); }});
    highPri.forEach(t => { if (!seen.has(t.id)) { seen.add(t.id); combined.push({ ...t, flag: t.priority === "urgent" ? "Urgent" : "High Priority" }); }});
    agentFlagged[agent] = combined;
  });

  // ─── Per-agent CSAT ───
  // Email CSAT: match by ticketId → zendesk ticket assignee
  const agentEmailCsat = {};
  agents.forEach(a => agentEmailCsat[a] = { total: 0, satisfied: 0 });
  if (csat && csat.responses) {
    csat.responses.forEach(r => {
      if (!r.ticketId) return;
      // Find which agent owns this ticket
      const ticket = (zendesk.tickets || []).find(t => String(t.id) === String(r.ticketId));
      if (!ticket) return;
      const matched = agents.find(ca => {
        const parts = ca.toLowerCase().split(/\s+/);
        const tLower = (ticket.assignee || "").toLowerCase();
        return ca === ticket.assignee || parts.some(p => p.length > 1 && tLower.split(/\s+/).some(ap => ap === p));
      });
      if (matched) {
        agentEmailCsat[matched].total++;
        if (r.satisfied) agentEmailCsat[matched].satisfied++;
      }
    });
  }

  // Phone CSAT: match by agent name
  const agentPhoneCsat = {};
  agents.forEach(a => agentPhoneCsat[a] = { total: 0, satisfied: 0 });
  if (postCall && postCall.responses) {
    postCall.responses.forEach(r => {
      const matched = agents.find(a => {
        const parts = a.toLowerCase().split(/\s+/);
        const rParts = (r.agent || "").toLowerCase().split(/\s+/);
        return parts[0] === rParts[0] || (parts[1] && rParts[1] && parts[1] === rParts[1]);
      });
      if (matched) {
        agentPhoneCsat[matched].total++;
        if (r.satisfied) agentPhoneCsat[matched].satisfied++;
      }
    });
  }

  // ─── Compute team averages ───
  function teamAvgFn(arr) {
    if (arr.length === 0) return 0;
    return Math.round(arr.reduce((a, b) => a + b, 0) / arr.length * 10) / 10;
  }

  const vals = agents.map(a => agentMetrics[a]);
  const teamAvg = {
    assigned: teamAvgFn(vals.map(v => v.assigned)),
    pastSla:  teamAvgFn(vals.map(v => v.pastSla)),
    solved:   teamAvgFn(vals.map(v => v.solved)),
    inbound:  teamAvgFn(vals.map(v => v.inbound)),
    outbound: teamAvgFn(vals.map(v => v.outbound)),
    inTalk:   teamAvgFn(vals.map(v => v.inTalk)),
    outTalk:  teamAvgFn(vals.map(v => v.outTalk)),
  };

  // ─── Brand colors ───
  const navy = BRAND.airBlueDark;
  const bg = BRAND.white;
  const headerBg = navy;
  const labelFg = BRAND.black;
  const dimFg = BRAND.airBlueMedium;
  const green = "#2E7D32";
  const red = "#C62828";
  const amber = "#F57F17";
  const gray = BRAND.beigeMedium;
  const borderColor = BRAND.beigeLight;
  const sectionBg = BRAND.airBlueLight;
  // ─── Layout constants ───
  // Left column: A-E (cols 1-5) = agent's own metrics
  // Gap: F (col 6) = spacer
  // Right column: G-K (cols 7-11) = team performance
  const leftCols = 5;   // A-E
  const gapCol = 6;     // F
  const rightStart = 7; // G
  const rightCols = 5;  // G-K
  const totalCols = 11; // A-K

  const zdUrl = `https://${CONFIG.zendesk.subdomain}.zendesk.com/agent/tickets/`;

  // ─── Phone team data (shared across all agents) ───
  const answerRate = aircall.teamAnswerRate || 0;
  const answered = aircall.teamAnswered || 0;
  const forwarded = aircall.forwarded || 0;
  const totalInbound = answered + forwarded;
  const totalOutbound = aircall.totalOutbound || 0;
  const missedDetails = aircall.missedCallDetails || [];

  // ─── "Last Week" data from Daily Metrics Log (prior Mon-Fri) ───
  const today = new Date(Utilities.formatDate(now, tz, "yyyy-MM-dd") + "T12:00:00");
  const todayDow = today.getDay(); // 0=Sun ... 6=Sat

  // Find prior Monday: go back to this Monday, then subtract 7 days
  const thisMon = new Date(today);
  thisMon.setDate(today.getDate() - ((todayDow + 6) % 7)); // this week's Monday
  const lastMon = new Date(thisMon);
  lastMon.setDate(thisMon.getDate() - 7);
  const lastFri = new Date(lastMon);
  lastFri.setDate(lastMon.getDate() + 4);

  const lwStart = Utilities.formatDate(lastMon, tz, "yyyy-MM-dd");
  const lwEnd = Utilities.formatDate(lastFri, tz, "yyyy-MM-dd");
  const lwLabel = Utilities.formatDate(lastMon, tz, "M/d") + " - " + Utilities.formatDate(lastFri, tz, "M/d");

  let lwRows = [];
  try { lwRows = readMetricsLog(lwStart, lwEnd); } catch (e) {
    Logger.log("Last week metrics read failed: " + e);
  }

  // Aggregate per-agent last-week data
  const lwAgentData = {};
  agents.forEach((agent, idx) => {
    const firstName = agent.split(" ")[0];
    const solved = lwRows.reduce((s, r) => s + (Number(r["Solved: " + firstName]) || 0), 0);
    const inbound = lwRows.reduce((s, r) => s + (Number(r["In: " + firstName]) || 0), 0);
    const inTalk = lwRows.reduce((s, r) => s + (Number(r["In Talk: " + firstName]) || 0), 0);
    const outbound = lwRows.reduce((s, r) => s + (Number(r["Out: " + firstName]) || 0), 0);
    const outTalk = lwRows.reduce((s, r) => s + (Number(r["Out Talk: " + firstName]) || 0), 0);
    const emailCsatSat = lwRows.reduce((s, r) => s + (Number(r["Email CSAT Sat: " + firstName]) || 0), 0);
    const emailCsatTot = lwRows.reduce((s, r) => s + (Number(r["Email CSAT Tot: " + firstName]) || 0), 0);
    const phoneCsatSat = lwRows.reduce((s, r) => s + (Number(r["Phone CSAT Sat: " + firstName]) || 0), 0);
    const phoneCsatTot = lwRows.reduce((s, r) => s + (Number(r["Phone CSAT Tot: " + firstName]) || 0), 0);
    lwAgentData[agent] = { solved, inbound, inTalk, outbound, outTalk, emailCsatSat, emailCsatTot, phoneCsatSat, phoneCsatTot };
  });

  // Team averages for last week
  function lwTeamAvg(field) {
    const vals = agents.map(a => lwAgentData[a][field]);
    if (vals.length === 0) return 0;
    return Math.round(vals.reduce((a, b) => a + b, 0) / vals.length * 10) / 10;
  }
  const lwTeamAvgs = {
    solved: lwTeamAvg("solved"),
    inbound: lwTeamAvg("inbound"),
    inTalk: lwTeamAvg("inTalk"),
    outbound: lwTeamAvg("outbound"),
    outTalk: lwTeamAvg("outTalk"),
  };
  // Team CSAT averages (aggregate all, not per-agent avg)
  const lwTeamEmailCsatSat = agents.reduce((s, a) => s + lwAgentData[a].emailCsatSat, 0);
  const lwTeamEmailCsatTot = agents.reduce((s, a) => s + lwAgentData[a].emailCsatTot, 0);
  const lwTeamPhoneCsatSat = agents.reduce((s, a) => s + lwAgentData[a].phoneCsatSat, 0);
  const lwTeamPhoneCsatTot = agents.reduce((s, a) => s + lwAgentData[a].phoneCsatTot, 0);

  const hasLastWeekData = lwRows.length > 0;

  // ─── Write each agent's sheet ───
  agents.forEach(agent => {
    const sheetId = agentSheets[agent];
    if (!sheetId) return;

    let ss;
    try { ss = SpreadsheetApp.openById(sheetId); } catch (e) {
      Logger.log(`Cannot open sheet for ${agent}: ${e.toString()}`);
      return;
    }

    const me = agentMetrics[agent];
    const firstName = agent.split(" ")[0];
    const sheetName = firstName + " - Agent Hub";

    // Clean up old sheet names from previous versions (e.g. "AgentA - Agent Command Center")
    const oldSheetName = firstName + " - Agent Command Center";
    const oldSheet = ss.getSheetByName(oldSheetName);
    if (oldSheet) {
      try { ss.deleteSheet(oldSheet); } catch (e) {
        Logger.log(`Could not delete old sheet "${oldSheetName}": ${e.toString()}`);
      }
    }

    // Build on a hidden staging sheet, then swap — eliminates blink on refresh
    const staging = getOrCreateSheet(ss, "_Staging");
    staging.showSheet();
    staging.clear();
    staging.clearFormats();
    try { staging.showRows(1, staging.getMaxRows()); } catch (e) { /* ignore */ }
    try { staging.showColumns(1, staging.getMaxColumns()); } catch (e) { /* ignore */ }
    const sheet = staging; // all writes go to staging

    // Column widths: Left = A(label) B(you) C(avg) D+E(tables); F(gap); Right = G(label) H(val) I+J+K(tables)
    sheet.setColumnWidth(1, 200);  // A - label
    sheet.setColumnWidth(2, 110);  // B - you
    sheet.setColumnWidth(3, 110);  // C - team avg
    sheet.setColumnWidth(4, 110);  // D - table col
    sheet.setColumnWidth(5, 140);  // E - table col
    sheet.setColumnWidth(gapCol, 20); // F - gap
    sheet.setColumnWidth(7, 180);  // G - label
    sheet.setColumnWidth(8, 120);  // H - value
    sheet.setColumnWidth(9, 120);  // I - table col
    sheet.setColumnWidth(10, 120); // J - table col
    sheet.setColumnWidth(11, 140); // K - table col
    const maxCol = sheet.getMaxColumns();
    if (maxCol > totalCols) {
      try { sheet.hideColumns(totalCols + 1, maxCol - totalCols); } catch (e) { /* ignore */ }
    }

    let row = 1;

    // ─── HEADER (spans full width) ───
    sheet.setRowHeight(row, 44);
    sheet.getRange(row, 1, 1, totalCols).merge()
      .setValue(`${firstName} - Agent Hub`)
      .setBackground(headerBg).setFontColor("#FFFFFF")
      .setFontSize(16).setFontWeight("bold").setFontFamily("Arial")
      .setVerticalAlignment("middle").setHorizontalAlignment("center");
    row++;

    sheet.setRowHeight(row, 24);
    sheet.getRange(row, 1, 1, totalCols).merge()
      .setValue(`${dateStr}  ·  Last updated ${timeStr}`)
      .setBackground(headerBg).setFontColor(BRAND.airBlueLight)
      .setFontSize(9).setFontStyle("italic").setFontFamily("Arial")
      .setVerticalAlignment("middle").setHorizontalAlignment("center");
    row++;

    // Spacer
    sheet.setRowHeight(row, 8);
    sheet.getRange(row, 1, 1, totalCols).setBackground(bg);
    row++;

    const contentStartRow = row; // both columns start here

    // ─── Helper: left-column metric row (3-col: label, you, team avg) ───
    function metricRow(label, myVal, avgVal, fmt, opts) {
      const f = fmt || "num";
      const highlight = opts && opts.highlight;
      sheet.setRowHeight(row, 26);
      sheet.getRange(row, 1).setValue(label)
        .setFontColor(labelFg).setFontSize(11).setFontFamily("Arial").setBackground(bg);
      sheet.getRange(row, 4, 1, 2).setBackground(bg).setValue("");

      const myCell = sheet.getRange(row, 2);
      if (f === "talk") { myCell.setValue(formatTalkTime(myVal)); }
      else if (f === "pct") { myCell.setValue(myVal !== null ? myVal + "%" : "—"); }
      else { myCell.setValue(myVal).setNumberFormat("0"); }
      myCell.setHorizontalAlignment("right").setFontWeight("bold")
        .setFontSize(12).setFontFamily("Arial").setBackground(bg);

      if (highlight === "higher") {
        myCell.setFontColor(myVal >= avgVal ? green : red);
      } else if (highlight === "lower") {
        myCell.setFontColor(myVal <= avgVal ? green : red);
      } else { myCell.setFontColor(labelFg); }

      const avgCell = sheet.getRange(row, 3);
      if (f === "talk") { avgCell.setValue(formatTalkTime(avgVal)); }
      else if (f === "pct") { avgCell.setValue(avgVal !== null ? avgVal + "%" : "—"); }
      else { avgCell.setValue(avgVal).setNumberFormat("0"); }
      avgCell.setHorizontalAlignment("right").setFontColor(dimFg)
        .setFontSize(11).setFontFamily("Arial").setBackground(bg);
      row++;
    }

    // ─── Helper: left-column section header ───
    function sectionHeader(title) {
      sheet.setRowHeight(row, 32);
      sheet.getRange(row, 1, 1, leftCols).merge()
        .setValue(title)
        .setBackground(sectionBg).setFontColor(navy)
        .setFontWeight("bold").setFontSize(12).setFontFamily("Arial")
        .setVerticalAlignment("middle");
      sheet.getRange(row, 1, 1, leftCols).setBorder(true, false, true, false, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);
      row++;
    }

    // ─── Helper: left-column col headers (You / Team Avg) ───
    function colHeaders() {
      sheet.setRowHeight(row, 24);
      sheet.getRange(row, 1).setValue("").setBackground(bg);
      sheet.getRange(row, 2).setValue("You")
        .setBackground(bg).setFontColor(navy).setFontWeight("bold")
        .setFontSize(10).setHorizontalAlignment("right").setFontFamily("Arial");
      sheet.getRange(row, 3).setValue("Team Avg")
        .setBackground(bg).setFontColor(dimFg).setFontWeight("bold")
        .setFontSize(10).setHorizontalAlignment("right").setFontFamily("Arial");
      sheet.getRange(row, 4, 1, 2).setBackground(bg).setValue("");
      sheet.getRange(row, 1, 1, leftCols).setBorder(false, false, true, false, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);
      row++;
    }

    // ════════════════════════════════════════
    // LEFT COLUMN: EMAIL (TICKETS) — agent-specific
    // ════════════════════════════════════════
    sectionHeader("Email (Zendesk)");
    colHeaders();
    metricRow("New+Open Assigned Tickets", me.assigned, teamAvg.assigned, "num");
    metricRow("Past SLA", me.pastSla, teamAvg.pastSla, "num", { highlight: "lower" });
    metricRow("Solved Today", me.solved, teamAvg.solved, "num", { highlight: "higher" });

    // Email CSAT
    const ec = agentEmailCsat[agent];
    const ecPct = ec.total > 0 ? Math.round(ec.satisfied / ec.total * 100) : null;
    const ecAllTotal = csat ? csat.total : 0;
    const ecAllSat = csat ? csat.satisfied : 0;
    const ecTeamPct = ecAllTotal > 0 ? Math.round(ecAllSat / ecAllTotal * 100) : null;
    metricRow("Email CSAT (24h)", ecPct, ecTeamPct, "pct", { highlight: "higher" });

    // Oldest waiting tickets for this agent
    const oldest = agentOldest[agent];
    if (oldest.length > 0) {
      sheet.setRowHeight(row, 6);
      sheet.getRange(row, 1, 1, leftCols).setBackground(bg);
      row++;
      sheet.setRowHeight(row, 22);
      sheet.getRange(row, 1, 1, leftCols).merge()
        .setValue("Your Oldest Waiting Tickets")
        .setFontColor(navy).setFontWeight("bold").setFontSize(10).setFontFamily("Arial").setBackground(bg);
      row++;
      // Header
      sheet.setRowHeight(row, 20);
      ["#", "Subject", "", "Wait", "Status"].forEach((h, ci) => {
        sheet.getRange(row, ci + 1).setValue(h)
          .setFontColor(dimFg).setFontSize(9).setFontWeight("bold").setFontFamily("Arial").setBackground(bg);
      });
      sheet.getRange(row, 1, 1, leftCols).setBorder(false, false, true, false, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);
      row++;
      oldest.forEach(t => {
        sheet.setRowHeight(row, 20);
        const subj = (t.subject || "").length > 30 ? t.subject.substring(0, 27) + "..." : (t.subject || "");
        const waitHrs = Math.round(t.waitBizMin / 60 * 10) / 10;
        // Ticket ID as clickable hyperlink
        sheet.getRange(row, 1).setFormula(`=HYPERLINK("${zdUrl}${t.id}","${t.id}")`)
          .setFontSize(9).setFontColor("#1155CC").setFontFamily("Arial").setBackground(bg);
        sheet.getRange(row, 2, 1, 2).merge().setValue(subj)
          .setFontSize(9).setFontFamily("Arial").setFontColor(labelFg).setBackground(bg);
        sheet.getRange(row, 4).setValue(waitHrs + "h")
          .setFontSize(9).setFontFamily("Arial").setFontColor(t.pastSla ? red : green).setBackground(bg);
        sheet.getRange(row, 5).setValue(t.status || "")
          .setFontSize(9).setFontFamily("Arial").setFontColor(dimFg).setBackground(bg);
        row++;
      });
    }

    // Flagged tickets for this agent
    const flagged = agentFlagged[agent];
    if (flagged.length > 0) {
      sheet.setRowHeight(row, 6);
      sheet.getRange(row, 1, 1, leftCols).setBackground(bg);
      row++;
      sheet.setRowHeight(row, 22);
      sheet.getRange(row, 1, 1, leftCols).merge()
        .setValue("Flagged Tickets")
        .setFontColor(navy).setFontWeight("bold").setFontSize(10).setFontFamily("Arial").setBackground(bg);
      row++;
      sheet.setRowHeight(row, 20);
      ["#", "Subject", "Flag", "Status", "Wait"].forEach((h, ci) => {
        sheet.getRange(row, ci + 1).setValue(h)
          .setFontColor(dimFg).setFontSize(9).setFontWeight("bold").setFontFamily("Arial").setBackground(bg);
      });
      sheet.getRange(row, 1, 1, leftCols).setBorder(false, false, true, false, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);
      row++;
      flagged.slice(0, 8).forEach(t => {
        sheet.setRowHeight(row, 20);
        const subj = (t.subject || "").length > 28 ? t.subject.substring(0, 25) + "..." : (t.subject || "");
        const waitHrs = t.waitBizMin ? Math.round(t.waitBizMin / 60 * 10) / 10 + "h" : "";
        // Ticket ID as clickable hyperlink
        sheet.getRange(row, 1).setFormula(`=HYPERLINK("${zdUrl}${t.id}","${t.id}")`)
          .setFontSize(9).setFontColor("#1155CC").setFontFamily("Arial").setBackground(bg);
        sheet.getRange(row, 2).setValue(subj)
          .setFontSize(9).setFontFamily("Arial").setFontColor(labelFg).setBackground(bg);
        sheet.getRange(row, 3).setValue(t.flag || "")
          .setFontSize(9).setFontFamily("Arial").setFontColor(t.flag === "Urgent" ? red : amber).setBackground(bg);
        sheet.getRange(row, 4).setValue(t.status || "")
          .setFontSize(9).setFontFamily("Arial").setFontColor(dimFg).setBackground(bg);
        sheet.getRange(row, 5).setValue(waitHrs)
          .setFontSize(9).setFontFamily("Arial").setFontColor(dimFg).setBackground(bg);
        row++;
      });
    }

    // Spacer
    sheet.setRowHeight(row, 10);
    sheet.getRange(row, 1, 1, leftCols).setBackground(bg);
    row++;

    // ════════════════════════════════════════
    // LEFT COLUMN: PHONE — agent-specific
    // ════════════════════════════════════════
    sectionHeader("Phone (Aircall) — Your Activity");
    colHeaders();
    metricRow("Inbound Answered", me.inbound, teamAvg.inbound, "num", { highlight: "higher" });
    metricRow("Outbound Calls", me.outbound, teamAvg.outbound, "num", { highlight: "higher" });
    metricRow("Inbound Talk Time", me.inTalk, teamAvg.inTalk, "talk", { highlight: "higher" });
    metricRow("Outbound Talk Time", me.outTalk, teamAvg.outTalk, "talk", { highlight: "higher" });

    // Phone CSAT
    const pc = agentPhoneCsat[agent];
    const pcPct = pc.total > 0 ? Math.round(pc.satisfied / pc.total * 100) : null;
    const pcAllTotal = postCall ? postCall.total : 0;
    const pcAllSat = postCall ? postCall.satisfied : 0;
    const pcTeamPct = pcAllTotal > 0 ? Math.round(pcAllSat / pcAllTotal * 100) : null;
    metricRow("Phone CSAT (24h)", pcPct, pcTeamPct, "pct", { highlight: "higher" });

    // ─── CSAT Ticket Details (email CSAT with hyperlinks) ───
    const myEmailCsatResponses = (csat && csat.responses || []).filter(r => {
      if (!r.ticketId) return false;
      const ticket = (zendesk.tickets || []).find(t => String(t.id) === String(r.ticketId));
      if (!ticket) return false;
      const matched = agents.find(ca => {
        const parts = ca.toLowerCase().split(/\s+/);
        const tLower = (ticket.assignee || "").toLowerCase();
        return ca === ticket.assignee || parts.some(p => p.length > 1 && tLower.split(/\s+/).some(ap => ap === p));
      });
      return matched === agent;
    });
    const myPhoneCsatResponses = (postCall && postCall.responses || []).filter(r => {
      const matched = agents.find(a => {
        const parts = a.toLowerCase().split(/\s+/);
        const rParts = (r.agent || "").toLowerCase().split(/\s+/);
        return parts[0] === rParts[0] || (parts[1] && rParts[1] && parts[1] === rParts[1]);
      });
      return matched === agent;
    });
    const allCsatForAgent = [
      ...myEmailCsatResponses.map(r => ({ type: "Email", ticketId: r.ticketId, score: r.score, maxScore: r.maxScore, customer: r.email || "Unknown", time: r.timeStr })),
      ...myPhoneCsatResponses.map(r => ({ type: "Phone", ticketId: null, score: r.score, maxScore: r.maxScore, customer: r.customer || "Unknown", time: r.timeStr })),
    ];
    if (allCsatForAgent.length > 0) {
      sheet.setRowHeight(row, 6);
      sheet.getRange(row, 1, 1, leftCols).setBackground(bg);
      row++;
      sheet.setRowHeight(row, 22);
      sheet.getRange(row, 1, 1, leftCols).merge()
        .setValue("Your CSAT Reviews (24h)")
        .setFontColor(navy).setFontWeight("bold").setFontSize(10).setFontFamily("Arial").setBackground(bg);
      row++;
      // Header
      sheet.setRowHeight(row, 20);
      ["Type", "Ticket", "Score", "Customer", "Time"].forEach((h, ci) => {
        sheet.getRange(row, ci + 1).setValue(h)
          .setFontColor(dimFg).setFontSize(9).setFontWeight("bold").setFontFamily("Arial").setBackground(bg);
      });
      sheet.getRange(row, 1, 1, leftCols).setBorder(false, false, true, false, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);
      row++;
      allCsatForAgent.slice(0, 10).forEach(c => {
        sheet.setRowHeight(row, 20);
        sheet.getRange(row, 1).setValue(c.type)
          .setFontSize(9).setFontFamily("Arial").setFontColor(labelFg).setBackground(bg);
        if (c.ticketId) {
          sheet.getRange(row, 2).setFormula(`=HYPERLINK("${zdUrl}${c.ticketId}","${c.ticketId}")`)
            .setFontSize(9).setFontColor("#1155CC").setFontFamily("Arial").setBackground(bg);
        } else {
          sheet.getRange(row, 2).setValue("-")
            .setFontSize(9).setFontFamily("Arial").setFontColor(dimFg).setBackground(bg);
        }
        const scoreColor = c.score >= 4 ? green : (c.score >= 3 ? amber : red);
        sheet.getRange(row, 3).setValue(c.score + "/" + c.maxScore)
          .setFontSize(9).setFontFamily("Arial").setFontWeight("bold").setFontColor(scoreColor).setBackground(bg);
        const custShort = (c.customer || "").length > 18 ? c.customer.substring(0, 15) + "..." : c.customer;
        sheet.getRange(row, 4).setValue(custShort)
          .setFontSize(9).setFontFamily("Arial").setFontColor(labelFg).setBackground(bg);
        sheet.getRange(row, 5).setValue(c.time || "")
          .setFontSize(9).setFontFamily("Arial").setFontColor(dimFg).setBackground(bg);
        row++;
      });
    }

    // ─── LAST WEEK SECTION ───
    if (hasLastWeekData) {
      const lw = lwAgentData[agent];
      sheet.setRowHeight(row, 10);
      sheet.getRange(row, 1, 1, leftCols).setBackground(bg);
      row++;

      // Section header
      sheet.getRange(row, 1, 1, leftCols).merge()
        .setValue("Last Week (" + lwLabel + ")")
        .setBackground(sectionBg).setFontColor(navy)
        .setFontWeight("bold").setFontSize(12).setFontFamily("Arial")
        .setVerticalAlignment("middle");
      sheet.getRange(row, 1, 1, leftCols).setBorder(true, false, true, false, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);
      row++;

      // Column headers
      sheet.setRowHeight(row, 24);
      sheet.getRange(row, 1).setValue("").setBackground(bg);
      sheet.getRange(row, 2).setValue("You")
        .setBackground(bg).setFontColor(navy).setFontWeight("bold")
        .setFontSize(10).setHorizontalAlignment("right").setFontFamily("Arial");
      sheet.getRange(row, 3).setValue("Team Avg")
        .setBackground(bg).setFontColor(dimFg).setFontWeight("bold")
        .setFontSize(10).setHorizontalAlignment("right").setFontFamily("Arial");
      sheet.getRange(row, 4, 1, 2).setBackground(bg).setValue("");
      sheet.getRange(row, 1, 1, leftCols).setBorder(false, false, true, false, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);
      row++;

      // Reuse metricRow for last week data
      metricRow("Tickets Solved", lw.solved, lwTeamAvgs.solved, "num", { highlight: "higher" });
      metricRow("Inbound Answered", lw.inbound, lwTeamAvgs.inbound, "num", { highlight: "higher" });
      metricRow("Inbound Talk Time", lw.inTalk, lwTeamAvgs.inTalk, "talk", { highlight: "higher" });
      metricRow("Outbound Calls", lw.outbound, lwTeamAvgs.outbound, "num", { highlight: "higher" });
      metricRow("Outbound Talk Time", lw.outTalk, lwTeamAvgs.outTalk, "talk", { highlight: "higher" });

      // Email CSAT for last week
      const lwEcPct = lw.emailCsatTot > 0 ? Math.round(lw.emailCsatSat / lw.emailCsatTot * 100) : null;
      const lwEcTeamPct = lwTeamEmailCsatTot > 0 ? Math.round(lwTeamEmailCsatSat / lwTeamEmailCsatTot * 100) : null;
      metricRow("Email CSAT", lwEcPct, lwEcTeamPct, "pct", { highlight: "higher" });

      // Phone CSAT for last week
      const lwPcPct = lw.phoneCsatTot > 0 ? Math.round(lw.phoneCsatSat / lw.phoneCsatTot * 100) : null;
      const lwPcTeamPct = lwTeamPhoneCsatTot > 0 ? Math.round(lwTeamPhoneCsatSat / lwTeamPhoneCsatTot * 100) : null;
      metricRow("Phone CSAT", lwPcPct, lwPcTeamPct, "pct", { highlight: "higher" });
    }

    const leftEndRow = row; // track where left column ends

    // ════════════════════════════════════════
    // RIGHT COLUMN: TEAM PERFORMANCE (starts at same row as left content)
    // ════════════════════════════════════════
    let rRow = contentStartRow;

    // Gap column background
    for (let r = contentStartRow; r <= Math.max(leftEndRow + 20, contentStartRow + 30); r++) {
      sheet.getRange(r, gapCol).setBackground(bg).setValue("");
    }

    // ─── Right-column section header ───
    sheet.getRange(rRow, rightStart, 1, rightCols).merge()
      .setValue("Phone (Aircall) — Team Overview")
      .setBackground(sectionBg).setFontColor(navy)
      .setFontWeight("bold").setFontSize(12).setFontFamily("Arial")
      .setVerticalAlignment("middle");
    sheet.getRange(rRow, rightStart, 1, rightCols).setBorder(true, false, true, false, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);
    rRow++;

    // KPI rows
    function rightKpiRow(label, value, color) {
      sheet.getRange(rRow, rightStart).setValue(label)
        .setFontColor(labelFg).setFontSize(11).setFontFamily("Arial").setBackground(bg);
      sheet.getRange(rRow, rightStart + 1).setValue(value)
        .setHorizontalAlignment("right").setFontWeight("bold").setFontSize(12).setFontFamily("Arial")
        .setFontColor(color || labelFg).setBackground(bg);
      sheet.getRange(rRow, rightStart + 2, 1, 3).setBackground(bg).setValue("");
      rRow++;
    }

    rightKpiRow("Answer Rate", Math.round(answerRate) + "%",
      answerRate >= (CONFIG.thresholds.phoneAnswerRate.green || 75) ? green : red);
    rightKpiRow("Answered / Total", `${answered} / ${totalInbound}`);
    rightKpiRow("Sent to Answer Service", forwarded);
    rightKpiRow("Team Outbound", totalOutbound);

    // Missed calls table
    if (missedDetails.length > 0) {
      sheet.getRange(rRow, rightStart, 1, rightCols).setBackground(bg);
      rRow++;
      sheet.getRange(rRow, rightStart, 1, rightCols).merge()
        .setValue("Missed Calls")
        .setFontColor(navy).setFontWeight("bold").setFontSize(10).setFontFamily("Arial").setBackground(bg);
      rRow++;
      // Header
      ["Line", "Customer", "Time", "Reason", ""].forEach((h, ci) => {
        sheet.getRange(rRow, rightStart + ci).setValue(h)
          .setFontColor(dimFg).setFontSize(9).setFontWeight("bold").setFontFamily("Arial").setBackground(bg);
      });
      sheet.getRange(rRow, rightStart, 1, rightCols).setBorder(false, false, true, false, false, false, borderColor, SpreadsheetApp.BorderStyle.SOLID);
      rRow++;
      missedDetails.slice(0, 10).forEach(mc => {
        const lineName = (mc.lineName || "").replace(/ support/i, "").trim() || "support";
        const customer = mc.contactName || (mc.callId ? "Check Aircall #" + mc.callId : mc.callerNumber || "Unknown");
        sheet.getRange(rRow, rightStart).setValue(lineName)
          .setFontSize(9).setFontFamily("Arial").setFontColor(labelFg).setBackground(bg);
        sheet.getRange(rRow, rightStart + 1).setValue(customer)
          .setFontSize(9).setFontFamily("Arial").setFontColor(labelFg).setBackground(bg);
        sheet.getRange(rRow, rightStart + 2).setValue(mc.callTime || "")
          .setFontSize(9).setFontFamily("Arial").setFontColor(dimFg).setBackground(bg);
        sheet.getRange(rRow, rightStart + 3, 1, 2).merge().setValue(mc.reason || "")
          .setFontSize(9).setFontFamily("Arial")
          .setFontColor(mc.reason === "No agents available" ? amber : dimFg).setBackground(bg);
        rRow++;
      });
    }

    // ─── Ensure consistent row heights across both columns ───
    const maxRow = Math.max(row, rRow);

    // Fill remaining right-column rows with bg if left column is longer
    for (let r = rRow; r < maxRow; r++) {
      sheet.getRange(r, rightStart, 1, rightCols).setBackground(bg);
    }

    // ─── FOOTER (spans full width, below both columns) ───
    const footerRow = maxRow + 1;
    sheet.setRowHeight(footerRow, 22);
    sheet.getRange(footerRow, 1, 1, totalCols).merge()
      .setValue(`CS Visibility · Agent Hub v2.5.4  ·  ${dateStr}  ·  Team avg = average across ${agents.length} agents`)
      .setFontColor(dimFg).setFontSize(8).setFontStyle("italic")
      .setHorizontalAlignment("center").setBackground(bg).setFontFamily("Arial");

    sheet.setRowHeight(footerRow + 1, 18);
    sheet.getRange(footerRow + 1, 1, 1, totalCols).merge()
      .setValue("Green = at or above team avg  ·  Red = below team avg")
      .setFontColor(dimFg).setFontSize(7).setFontStyle("italic")
      .setHorizontalAlignment("center").setBackground(bg).setFontFamily("Arial");

    // ─── Swap staging → main sheet in one batch ───
    SpreadsheetApp.flush();
    const dash = getOrCreateSheet(ss, sheetName);
    dash.clear();
    dash.clearFormats();
    try { dash.showRows(1, dash.getMaxRows()); } catch (e) { /* ignore */ }
    try { dash.showColumns(1, dash.getMaxColumns()); } catch (e) { /* ignore */ }

    const lastRow = staging.getLastRow() || 1;
    const lastCol = Math.max(staging.getLastColumn(), totalCols);
    if (lastRow > 0 && lastCol > 0) {
      const source = staging.getRange(1, 1, lastRow, lastCol);
      source.copyTo(dash.getRange(1, 1, lastRow, lastCol));
    }

    // Match column widths and row heights
    for (let c = 1; c <= lastCol; c++) {
      dash.setColumnWidth(c, staging.getColumnWidth(c));
    }
    for (let r = 1; r <= lastRow; r++) {
      dash.setRowHeight(r, staging.getRowHeight(r));
    }

    // Hide extra rows/cols on the main sheet
    const dashMaxRow = dash.getMaxRows();
    if (dashMaxRow > footerRow + 2) {
      try { dash.hideRows(footerRow + 3, dashMaxRow - footerRow - 2); } catch (e) { /* ignore */ }
    }
    const dashMaxCol = dash.getMaxColumns();
    if (dashMaxCol > totalCols) {
      try { dash.hideColumns(totalCols + 1, dashMaxCol - totalCols); } catch (e) { /* ignore */ }
    }

    dash.setHiddenGridlines(true);
    staging.hideSheet();

    // Ensure the main sheet is the first/active sheet
    ss.setActiveSheet(dash);
    ss.moveActiveSheet(1);

    SpreadsheetApp.flush();
    Logger.log(`Updated agent dashboard for ${agent}`);
  });
}

// ─── HEALTH TRENDS TAB ───
// Merges two data sources to build a historical health status chart:
//   1. Daily Metrics Log — derives status from raw metrics (Past SLA, Answer Rate, Unread DMs)
//   2. Health Badge Log — authoritative statuses parsed from email subject lines
// Email-sourced data takes priority since it captures the actual badge shown.
// Called once daily from sendDailyRecap.
function updateHealthTrends() {
  loadThresholds();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const th = CONFIG.thresholds;
  const tz = CONFIG.businessHours.timezone;

  // Helper: normalize any date value to "yyyy-MM-dd" string
  function toDateKey(d) {
    if (!d) return null;
    if (d instanceof Date) return Utilities.formatDate(d, tz, "yyyy-MM-dd");
    // String like "2026-05-11" or "Wed May 11 2026..."
    const s = String(d).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    // Try parsing as date
    const parsed = new Date(s);
    if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, tz, "yyyy-MM-dd");
    return null;
  }

  // ─── Source 1: Health Badge Log (from emails — authoritative) ───
  const badgeByDate = {};
  const badgeSheet = ss.getSheetByName("Health Badge Log");
  if (badgeSheet && badgeSheet.getLastRow() > 1) {
    const bData = badgeSheet.getRange(1, 1, badgeSheet.getLastRow(), 4).getValues();
    for (let i = 1; i < bData.length; i++) {
      const dateKey = toDateKey(bData[i][0]);
      if (!dateKey) continue;
      badgeByDate[dateKey] = {
        email: String(bData[i][1] || "").trim(),
        phone: String(bData[i][2] || "").trim(),
        social: String(bData[i][3] || "").trim(),
      };
    }
    Logger.log("Health Badge Log: " + Object.keys(badgeByDate).length + " days loaded. Sample: " + JSON.stringify(Object.entries(badgeByDate).slice(0, 3)));
  }

  // ─── Source 2: Daily Metrics Log (derive from raw data) ───
  const logSheet = ss.getSheetByName("Daily Metrics Log");
  const metricsByDate = {};
  if (logSheet && logSheet.getLastRow() > 1) {
    const data = logSheet.getRange(1, 1, logSheet.getLastRow(), logSheet.getLastColumn()).getValues();
    const headers = data[0];
    const dateIdx = headers.indexOf("Date");
    const pastSlaIdx = headers.indexOf("Past SLA");
    const answerRateIdx = headers.indexOf("Answer Rate %");
    const unreadDMsIdx = headers.indexOf("Unread DMs");

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dateKey = toDateKey(row[dateIdx]);
      if (!dateKey) continue;

      const pastSla = pastSlaIdx >= 0 ? row[pastSlaIdx] : "";
      const answerRate = answerRateIdx >= 0 ? row[answerRateIdx] : "";
      const unreadDMs = unreadDMsIdx >= 0 ? row[unreadDMsIdx] : "";

      let emailStatus = "";
      if (pastSla !== "" && pastSla !== null && !isNaN(pastSla)) {
        emailStatus = pastSla > 10 ? "At Risk" : pastSla > 5 ? "Watch" : "Healthy";
      }
      let phoneStatus = "";
      if (answerRate !== "" && answerRate !== null && !isNaN(answerRate)) {
        phoneStatus = answerRate >= th.phoneAnswerRate.green ? "Healthy"
          : answerRate >= th.phoneAnswerRate.yellow ? "Watch" : "At Risk";
      }
      let socialStatus = "";
      if (unreadDMs !== "" && unreadDMs !== null && !isNaN(unreadDMs)) {
        socialStatus = unreadDMs === 0 ? "Healthy" : unreadDMs <= 2 ? "Watch" : "At Risk";
      }

      metricsByDate[dateKey] = { email: emailStatus, phone: phoneStatus, social: socialStatus };
    }
    Logger.log("Daily Metrics Log: " + Object.keys(metricsByDate).length + " days loaded");
  }

  // ─── Merge: email badges take priority, metrics fill gaps ───
  const allDates = new Set([...Object.keys(badgeByDate), ...Object.keys(metricsByDate)]);
  const days = [];
  allDates.forEach(dateKey => {
    const badge = badgeByDate[dateKey] || {};
    const metric = metricsByDate[dateKey] || {};
    days.push({
      date: dateKey,
      emailStatus: badge.email || metric.email || "",
      phoneStatus: badge.phone || metric.phone || "",
      socialStatus: badge.social || metric.social || "",
    });
  });
  days.sort((a, b) => a.date.localeCompare(b.date));
  Logger.log("Merged: " + days.length + " days. First 3: " + JSON.stringify(days.slice(0, 3)));

  if (days.length === 0) return;

  // ─── Write directly to Health Trends sheet ───
  // (No staging — charts must reference the sheet they live on)
  const sheet = getOrCreateSheet(ss, "Health Trends");
  sheet.clear();
  sheet.clearFormats();
  try { sheet.showRows(1, sheet.getMaxRows()); } catch (e) { /* ignore */ }
  try { sheet.showColumns(1, sheet.getMaxColumns()); } catch (e) { /* ignore */ }
  sheet.getCharts().forEach(c => sheet.removeChart(c));

  const navy = BRAND.airBlueDark;
  const bg = BRAND.white;
  const dimFg = BRAND.airBlueMedium;

  // Status → numeric value for charting (3=Healthy, 2=Watch, 1=At Risk)
  function statusNum(s) { return s === "Healthy" ? 3 : s === "Watch" ? 2 : s === "At Risk" ? 1 : ""; }

  // ─── Data table (cols A-G) ───
  // Cols A-D: Date, Email, Phone, Social (status values 1-3)
  // Cols E-G: Layered area bands (painted back-to-front to create colored zones)
  //   E = 3.5 (green, fills entire area behind), F = 2.5 (yellow, covers below 2.5), G = 1.5 (red, covers below 1.5)
  const chartHeaders = ["Date", "Email", "Phone", "Social", "Zone1", "Zone2", "Zone3"];
  sheet.getRange(1, 1, 1, 7).setValues([chartHeaders])
    .setFontWeight("bold").setFontSize(9).setBackground(BRAND.beigeLight).setFontFamily("Arial");

  // Convert date strings to actual Date objects so charts recognize the x-axis
  const chartData = days.map(d => {
    const parts = d.date.split("-");
    const dt = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
    // Band values: layered areas (NOT stacked) — drawn back-to-front
    // E=3.5 (green, fills 0→3.5, rendered behind), F=2.5 (yellow, 0→2.5), G=1.5 (red, 0→1.5, on top)
    return [dt, statusNum(d.emailStatus), statusNum(d.phoneStatus), statusNum(d.socialStatus), 3.5, 2.5, 1.5];
  });
  if (chartData.length > 0) {
    sheet.getRange(2, 1, chartData.length, 7).setValues(chartData);
    sheet.getRange(2, 1, chartData.length, 1).setNumberFormat("M/d");
  }

  // Flush data to sheet BEFORE creating charts — charts read committed data
  SpreadsheetApp.flush();

  // ─── Title + subtitle (row 1-2) ───
  const dateRange = days.length > 0
    ? Utilities.formatDate(new Date(days[0].date), tz, "MMM d") + " – " + Utilities.formatDate(new Date(days[days.length - 1].date), tz, "MMM d, yyyy")
    : "";

  // ─── Header bar (matching main dashboard style) ───
  // Wordmark
  sheet.getRange(1, 8, 1, 2).merge()
    .setValue("deako®")
    .setBackground(navy).setFontColor("#FFFFFF")
    .setFontSize(13).setFontWeight("bold")
    .setVerticalAlignment("middle").setFontFamily("Arial");
  // Title
  sheet.getRange(1, 10, 1, 4).merge()
    .setValue("Health Trends")
    .setBackground(navy).setFontColor(BRAND.airBlueLight)
    .setFontSize(13).setVerticalAlignment("middle").setFontFamily("Arial");
  // Date range right-aligned
  sheet.getRange(1, 14, 1, 4).merge()
    .setValue(dateRange ? `${dateRange}  ·  ${days.length} business days` : "")
    .setBackground(navy).setFontColor(BRAND.airBlueMedium)
    .setFontSize(9).setHorizontalAlignment("right").setVerticalAlignment("middle").setFontFamily("Arial");
  sheet.setRowHeight(1, 36);
  // Subtitle
  sheet.getRange(2, 8, 1, 10).merge()
    .setValue(`Source: "CS End of Day Summary" daily emails  ·  Business days only (Mon-Fri)  ·  Weekends excluded`)
    .setFontSize(9).setFontColor(dimFg).setFontFamily("Arial").setBackground(bg);

  // ─── Create charts ───
  const dataRows = chartData.length + 1;
  const lineColor = BRAND.airBlueDark; // Deako navy for all chart lines
  const chartConfigs = [
    { title: "Email", col: 2, topRow: 4 },
    { title: "Phone", col: 3, topRow: 18 },
    { title: "Social", col: 4, topRow: 32 },
  ];

  // Summary legend — positioned to the right of each chart
  // Chart is 650px starting at col H; at ~80px/col that's cols H-P, so legend starts at Q
  const sumCol = 17; // column Q
  // Legend colors = exact chart band colors (#2E7D32, #F57F17, #C62828 at 18% opacity on white)
  const legendBg    = { "Healthy": "#D9E8DA", "Watch": "#FDE8D5", "At Risk": "#F5D8D8" };
  const legendText  = { "Healthy": "#2E7D32", "Watch": "#F57F17", "At Risk": "#C62828" };

  const channels = [
    { name: "EMAIL", data: days.map(d => d.emailStatus), topRow: 4 },
    { name: "PHONE", data: days.map(d => d.phoneStatus), topRow: 18 },
    { name: "SOCIAL", data: days.map(d => d.socialStatus), topRow: 32 },
  ];

  channels.forEach(ch => {
    const valid = ch.data.filter(s => s !== "");
    const counts = { "At Risk": 0, "Watch": 0, "Healthy": 0 };
    valid.forEach(s => { if (counts[s] !== undefined) counts[s]++; });

    let sr = ch.topRow;
    // Channel header + date range
    sheet.getRange(sr, sumCol, 1, 2).merge()
      .setValue(ch.name)
      .setFontWeight("bold").setFontSize(10).setFontColor(navy).setFontFamily("Arial").setBackground(bg);
    sr++;
    sheet.getRange(sr, sumCol, 1, 2).merge()
      .setValue(dateRange)
      .setFontSize(8).setFontColor(dimFg).setFontFamily("Arial").setBackground(bg);
    sr++;

    ["Healthy", "Watch", "At Risk"].forEach(status => {
      sheet.getRange(sr, sumCol, 1, 2).merge()
        .setValue(`${counts[status]}  ${status}`)
        .setFontSize(10).setFontWeight("bold").setFontFamily("Arial")
        .setBackground(legendBg[status]).setFontColor(legendText[status])
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
      sr++;
    });
  });

  // Column widths — hide all data columns A-G (charts still read them)
  for (let c = 1; c <= 7; c++) sheet.setColumnWidth(c, 40);
  sheet.hideColumns(1, 7); // cols A-G hidden
  for (let c = 8; c <= 16; c++) sheet.setColumnWidth(c, 80);
  sheet.setColumnWidth(sumCol, 120);     // legend pill column
  sheet.setColumnWidth(sumCol + 1, 10);  // narrow since merged into sumCol

  // Build combo charts: 3 layered area bands (green behind, yellow, red on top) + line
  // Series order: 0=Green(3.5), 1=Yellow(2.5), 2=Red(1.5), 3=actual line
  // Areas fill from 0 to their value; layering paints the visible bands
  // Band colors match the health badge pill colors used across the app
  chartConfigs.forEach(cfg => {
    const chart = sheet.newChart()
      .setChartType(Charts.ChartType.COMBO)
      .addRange(sheet.getRange(1, 1, dataRows, 1))       // Date (col A)
      .addRange(sheet.getRange(1, 5, dataRows, 1))        // Green band (col E, value=3.5)
      .addRange(sheet.getRange(1, 6, dataRows, 1))        // Yellow band (col F, value=2.5)
      .addRange(sheet.getRange(1, 7, dataRows, 1))        // Red band (col G, value=1.5)
      .addRange(sheet.getRange(1, cfg.col, dataRows, 1))  // Actual status line
      .setOption("useFirstColumnAsDomain", true)
      .setOption("title", cfg.title)
      .setOption("legend.position", "none")
      // Series 0: Healthy zone (green #2E7D32) — area filling 0→3.5, rendered behind
      .setOption("series.0.type", "area")
      .setOption("series.0.color", "#2E7D32")
      .setOption("series.0.areaOpacity", 0.18)
      .setOption("series.0.lineWidth", 0)
      .setOption("series.0.visibleInLegend", false)
      // Series 1: Watch zone (amber #F57F17) — area filling 0→2.5, covers green below 2.5
      .setOption("series.1.type", "area")
      .setOption("series.1.color", "#F57F17")
      .setOption("series.1.areaOpacity", 0.18)
      .setOption("series.1.lineWidth", 0)
      .setOption("series.1.visibleInLegend", false)
      // Series 2: At Risk zone (red #C62828) — area filling 0→1.5, covers yellow below 1.5
      .setOption("series.2.type", "area")
      .setOption("series.2.color", "#C62828")
      .setOption("series.2.areaOpacity", 0.18)
      .setOption("series.2.lineWidth", 0)
      .setOption("series.2.visibleInLegend", false)
      // Series 3: Actual status line — Deako navy blue
      .setOption("series.3.type", "line")
      .setOption("series.3.color", lineColor)
      .setOption("series.3.lineWidth", 2)
      .setOption("series.3.pointSize", 6)
      // Single vAxis for everything
      .setOption("vAxis.minValue", 0.5)
      .setOption("vAxis.maxValue", 3.5)
      .setOption("vAxis.ticks", [1, 2, 3])
      .setOption("vAxis.gridlines.color", "#E0E0E0")
      .setOption("hAxis.textStyle.fontSize", 9)
      .setPosition(cfg.topRow, 8, 0, 0)
      .setOption("width", 650)
      .setOption("height", 220)
      .build();
    sheet.insertChart(chart);
  });

  sheet.setTabColor("#6A1B9A");
  sheet.setHiddenGridlines(true);

  // ─── Footer — thresholds and notes ───
  const footerRow = 46; // below all 3 charts + legends
  const footerCols = 10; // cols H (8) through Q (17)
  const footerLines = [
    `CS Visibility · Command Center v2.5.4  ·  Health Trends  ·  Business days only (Mon-Fri, weekends excluded)`,
    `Email: Healthy = 0-5 past SLA, Watch = 6-10 past SLA, At Risk = 11+ past SLA`,
    `Phone: Healthy = answer rate >= ${th.phoneAnswerRate.green}%, Watch = ${th.phoneAnswerRate.yellow}-${th.phoneAnswerRate.green - 1}%, At Risk = < ${th.phoneAnswerRate.yellow}%`,
    `Social: Healthy = oldest unread DM <= ${th.socialResponseTime.green / 60}h, Watch = ${th.socialResponseTime.green / 60}-${th.socialResponseTime.yellow / 60}h, At Risk = > ${th.socialResponseTime.yellow / 60}h`,
    `Chart bands: green = Healthy zone, yellow = Watch zone, red = At Risk zone  ·  Line = actual daily status  ·  Y-axis: 3 = Healthy, 2 = Watch, 1 = At Risk`,
    `Data sourced from Health Badge Log + Daily Metrics Log sheets  ·  Badge log takes priority when both exist for a date`,
  ];
  for (let i = 0; i < footerLines.length; i++) {
    sheet.getRange(footerRow + i, 8, 1, footerCols).merge()
      .setValue(footerLines[i])
      .setFontSize(8).setFontColor(dimFg).setFontStyle("italic")
      .setHorizontalAlignment("center").setBackground(bg).setFontFamily("Arial").setWrap(false);
  }

  SpreadsheetApp.flush();
  Logger.log("Updated Health Trends tab with " + days.length + " business days of data");
}

// ─── BACKFILL HEALTH BADGES FROM EMAILS (one-time) ───
// Searches Gmail for "CS End of Day Summary" emails, extracts the ACTUAL
// metrics from the HTML body (Past SLA count, Answer Rate %, Unread DMs),
// then derives health badges using the CURRENT threshold logic.
// This ensures historical data reflects any threshold changes.
//
// Usage: Run backfillHealthFromEmails() once from the script editor.
function backfillHealthFromEmails() {
  loadThresholds();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = CONFIG.businessHours.timezone;
  const th = CONFIG.thresholds;

  // Search Gmail for daily recap emails
  const query = 'subject:"CS End of Day Summary"';
  const threads = GmailApp.search(query, 0, 100); // up to 100 most recent
  Logger.log(`Found ${threads.length} email threads matching "${query}"`);

  if (threads.length === 0) return;

  const entries = [];
  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(msg => {
      const date = msg.getDate();
      const body = msg.getBody(); // HTML body

      // Extract Past SLA count from: "Past SLA (Xh)</td><td ...>NUMBER ..."
      // Simplified: just grab the first number after the td opens
      const pastSlaMatch = body.match(/Past SLA[^<]*<\/td>\s*<td[^>]*>\s*(\d+)/i);
      // Extract Answer Rate from: "Answer Rate</td><td ...>NUMBER% ..."
      const answerRateMatch = body.match(/Answer Rate<\/td>\s*<td[^>]*>\s*([\d.]+)%/i);
      // Extract Unread DMs from: "Unread DMs</td><td ...>NUMBER ..."
      const unreadDMsMatch = body.match(/Unread DMs<\/td>\s*<td[^>]*>\s*(\d+)/i);

      const pastSla = pastSlaMatch ? parseInt(pastSlaMatch[1]) : null;
      const answerRate = answerRateMatch ? parseFloat(answerRateMatch[1]) : null;
      const unreadDMs = unreadDMsMatch ? parseInt(unreadDMsMatch[1]) : null;

      // Derive health badges using CURRENT thresholds
      const emailStatus = pastSla !== null
        ? (pastSla > 10 ? "At Risk" : pastSla > 5 ? "Watch" : "Healthy") : "";
      const phoneStatus = answerRate !== null
        ? (answerRate >= th.phoneAnswerRate.green ? "Healthy"
          : answerRate >= th.phoneAnswerRate.yellow ? "Watch" : "At Risk") : "";
      const socialStatus = unreadDMs !== null
        ? (unreadDMs === 0 ? "Healthy" : unreadDMs <= 2 ? "Watch" : "At Risk") : "";

      if (emailStatus || phoneStatus || socialStatus) {
        const dateStr = Utilities.formatDate(date, tz, "yyyy-MM-dd");
        entries.push({
          date: dateStr,
          email: emailStatus,
          phone: phoneStatus,
          social: socialStatus,
          // Raw values for logging
          _pastSla: pastSla,
          _answerRate: answerRate,
          _unreadDMs: unreadDMs,
        });
      }
    });
  });

  // Deduplicate by date (keep first/latest per day)
  const byDate = {};
  entries.forEach(e => {
    if (!byDate[e.date]) byDate[e.date] = e;
  });
  const sorted = Object.values(byDate).sort((a, b) => a.date.localeCompare(b.date));
  Logger.log(`Extracted ${sorted.length} unique days of health data from email bodies`);
  // Log a few samples so we can verify parsing worked
  sorted.slice(0, 5).forEach(e => {
    Logger.log(`  ${e.date}: PastSLA=${e._pastSla} -> ${e.email} | Rate=${e._answerRate}% -> ${e.phone} | DMs=${e._unreadDMs} -> ${e.social}`);
  });

  // Write to Health Badge Log sheet
  const sheet = getOrCreateSheet(ss, "Health Badge Log");
  sheet.clear();
  sheet.clearFormats();

  const headers = ["Date", "Email", "Phone", "Social"];
  sheet.getRange(1, 1, 1, 4).setValues([headers])
    .setFontWeight("bold").setFontSize(9).setBackground(BRAND.beigeLight).setFontFamily("Arial");
  sheet.setFrozenRows(1);
  sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd");

  if (sorted.length > 0) {
    const rows = sorted.map(e => [e.date, e.email, e.phone, e.social]);
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }

  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 80);

  Logger.log("Wrote " + sorted.length + " rows to Health Badge Log (derived from current thresholds)");

  // Now regenerate the Health Trends chart using combined data
  try { updateHealthTrends(); } catch (e) {
    Logger.log("Health Trends update after backfill failed: " + e.toString());
  }
}

// ─── BACKFILL HISTORICAL METRICS ───
// Processes ONE day per execution to stay well under the 6-minute Apps Script limit.
// Pulls "flow" metrics (solved tickets, calls, CSAT) from APIs.
// "Stock" metrics (open count, unassigned, past SLA) are left blank — can't reconstruct point-in-time.
//
// Usage:
//   1. Set Script Property BACKFILL_START_DATE = "2026-04-01" (or however far back you want)
//   2. Run setupBackfillTrigger() from the editor — it creates a trigger every 8 minutes
//   3. Watch the Daily Metrics Log sheet fill up
//   4. When caught up, the trigger auto-deletes itself
function backfillOneDay() {
  loadThresholds();
  const props = PropertiesService.getScriptProperties();
  const tz = CONFIG.businessHours.timezone;
  const startDate = props.getProperty("BACKFILL_START_DATE");
  if (!startDate) {
    Logger.log("BACKFILL_START_DATE not set — nothing to backfill");
    cleanupBackfillTrigger();
    return;
  }

  // Determine the next date to process
  const nextDate = props.getProperty("BACKFILL_NEXT_DATE") || startDate;
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");

  if (nextDate >= todayStr) {
    Logger.log("Backfill complete — caught up to today (" + todayStr + ")");
    props.deleteProperty("BACKFILL_NEXT_DATE");
    props.deleteProperty("BACKFILL_START_DATE");
    cleanupBackfillTrigger();
    return;
  }

  // Check if this is a non-working day — skip it
  const dateObj = new Date(nextDate + "T12:00:00");
  const holidays = getDeakoHolidays();
  if (isNonWorkingDay(dateObj, holidays)) {
    Logger.log("Skipping non-working day: " + nextDate);
    advanceBackfillDate(props, nextDate);
    return;
  }

  // Check if already logged
  const existing = readMetricsLog(nextDate, nextDate);
  if (existing.length > 0) {
    Logger.log("Already have data for " + nextDate + " — skipping");
    advanceBackfillDate(props, nextDate);
    return;
  }

  Logger.log("Backfilling metrics for: " + nextDate);
  const dayOfWeek = Utilities.formatDate(dateObj, tz, "EEEE");

  // ── Zendesk: Solved tickets that day ──
  const token = props.getProperty("ZENDESK_TOKEN");
  const subdomain = CONFIG.zendesk.subdomain;
  const authHeader = "Basic " + Utilities.base64Encode(token);
  const zdHeaders = { "Authorization": authHeader, "Content-Type": "application/json" };
  const zdOpts = { method: "get", headers: zdHeaders, muteHttpExceptions: true };

  function zdCount(query) {
    try {
      const url = `https://${subdomain}.zendesk.com/api/v2/search/count.json?query=${encodeURIComponent(query)}`;
      const resp = UrlFetchApp.fetch(url, zdOpts);
      if (resp.getResponseCode() === 200) return JSON.parse(resp.getContentText()).count || 0;
    } catch (e) { Logger.log("Zendesk search failed: " + e); }
    return 0;
  }

  const nextDatePlusOne = Utilities.formatDate(new Date(dateObj.getTime() + 86400000), tz, "yyyy-MM-dd");
  // Apply same exclusion filters as the daily recap: no aircall, no internal testing, no auto-close, no AI Agent
  const solvedFilter = `-tags:aircall -tags:internal__testing -tags:auto_close -assignee:"AI Agent"`;
  const solvedTotal = zdCount(`type:ticket solved>=${nextDate} solved<${nextDatePlusOne} ${solvedFilter}`);
  const ticketsCreated = zdCount(`type:ticket created>=${nextDate} created<${nextDatePlusOne} ${solvedFilter}`);
  const agentSolved = CONFIG.agents.map(a =>
    zdCount(`type:ticket solved>=${nextDate} solved<${nextDatePlusOne} ${solvedFilter} assignee:"${a}"`)
  );

  // ── Aircall: Calls that day ──
  let inboundCalls = 0;
  let forwardedToSAS = 0;
  let outboundCalls = 0;
  let answeredCalls = 0;
  let avgWaitTime = 0;
  let avgDuration = 0;
  const agentIn = CONFIG.agents.map(() => 0);
  const agentOut = CONFIG.agents.map(() => 0);
  const agentInTalk = CONFIG.agents.map(() => 0);
  const agentOutTalk = CONFIG.agents.map(() => 0);

  try {
    const apiId = props.getProperty("AIRCALL_API_ID");
    const apiToken = props.getProperty("AIRCALL_API_TOKEN");
    if (apiId && apiToken) {
      const baseUrl = CONFIG.aircall.baseUrl;
      const auth = "Basic " + Utilities.base64Encode(apiId + ":" + apiToken);

      // Start/end of the day in local timezone
      const dayStart = new Date(nextDate + "T00:00:00");
      const dayEnd = new Date(nextDate + "T23:59:59");
      // Approximate: use Pacific timezone offset manually
      const fromTs = Math.floor(dayStart.getTime() / 1000) + 7 * 3600; // rough PST offset
      const toTs = Math.floor(dayEnd.getTime() / 1000) + 7 * 3600;

      let allCalls = [];
      let page = 1;
      let hasMore = true;
      while (hasMore && page <= 10) {
        const url = `${baseUrl}/calls?from=${fromTs}&to=${toTs}&per_page=50&page=${page}&order=desc`;
        const resp = UrlFetchApp.fetch(url, { method: "get", headers: { "Authorization": auth }, muteHttpExceptions: true });
        if (resp.getResponseCode() !== 200) break;
        const data = JSON.parse(resp.getContentText());
        allCalls = allCalls.concat(data.calls || []);
        hasMore = data.meta && data.meta.next_page_link;
        page++;
      }

      const supportNumbers = CONFIG.aircall.supportNumbers;
      const answerSvcNum = (CONFIG.aircall.answeringServiceNumber || "").replace(/[\s\-\(\)]/g, "");
      const supportCalls = allCalls.filter(c => {
        if (!c.number) return false;
        const digits = (c.number.digits || "").replace(/[\s\-\(\)]/g, "");
        return supportNumbers.some(sn => digits.includes(sn.replace(/[\s\-\(\)]/g, "")) || sn.replace(/[\s\-\(\)]/g, "").includes(digits));
      });

      // Inbound during business hours
      const bh = CONFIG.businessHours;
      const bizInbound = supportCalls.filter(c => {
        if (c.direction !== "inbound" || !c.started_at) return false;
        const cd = new Date(c.started_at * 1000);
        const ps = cd.toLocaleString("en-US", { timeZone: bh.timezone });
        const pd = new Date(ps);
        const dow = pd.getDay();
        const hr = pd.getHours();
        return bh.workDays.includes(dow) && hr >= bh.startHour && hr < bh.endHour;
      });

      // Team answered
      const teamAns = bizInbound.filter(c => c.answered_at && c.user && (c.missed_call_reason || "") !== "short_abandoned");
      answeredCalls = teamAns.length;
      inboundCalls = answeredCalls;

      // Forwarded to SAS
      forwardedToSAS = supportCalls.filter(c => {
        if (c.direction !== "outbound") return false;
        const rd = (c.raw_digits || "").replace(/[\s\-\(\)]/g, "");
        return answerSvcNum && rd.includes(answerSvcNum.replace("+1", ""));
      }).length;

      // Outbound (all lines, CS agents only, exclude SAS)
      outboundCalls = allCalls.filter(c => {
        if (c.direction !== "outbound") return false;
        const rd = (c.raw_digits || "").replace(/[\s\-\(\)]/g, "");
        if (answerSvcNum && rd.includes(answerSvcNum.replace("+1", ""))) return false;
        // Only count recognized CS agents
        if (!c.user) return false;
        const name = (c.user.name || `${c.user.first_name || ""} ${c.user.last_name || ""}`.trim()).toLowerCase();
        return CONFIG.agents.some(a => {
          const parts = a.toLowerCase().split(" ");
          const cparts = name.split(" ");
          return parts[0] === cparts[0] || (parts[1] && cparts[1] && parts[1] === cparts[1]);
        });
      }).length;

      // Answer rate
      const answerRate = (answeredCalls + forwardedToSAS) > 0
        ? Math.round(answeredCalls / (answeredCalls + forwardedToSAS) * 100) : 100;

      // Avg wait/duration
      const waits = teamAns.map(c => c.waiting_duration || 0).filter(w => w > 0);
      avgWaitTime = waits.length > 0 ? Math.round(waits.reduce((a, b) => a + b, 0) / waits.length) : 0;
      const durs = teamAns.map(c => c.duration || 0).filter(d => d > 0);
      avgDuration = durs.length > 0 ? Math.round(durs.reduce((a, b) => a + b, 0) / durs.length) : 0;

      // Per-agent (counts + talk time)
      teamAns.forEach(c => {
        const user = c.user;
        if (!user) return;
        const name = (user.name || `${user.first_name || ""} ${user.last_name || ""}`.trim());
        CONFIG.agents.forEach((a, idx) => {
          const parts = a.toLowerCase().split(" ");
          const cparts = name.toLowerCase().split(" ");
          if (parts[0] === cparts[0] || (parts[1] && cparts[1] && parts[1] === cparts[1])) {
            agentIn[idx]++;
            agentInTalk[idx] += (c.duration || 0);
          }
        });
      });

      allCalls.filter(c => c.direction === "outbound" && c.user).forEach(c => {
        const rd = (c.raw_digits || "").replace(/[\s\-\(\)]/g, "");
        if (answerSvcNum && rd.includes(answerSvcNum.replace("+1", ""))) return;
        const name = (c.user.name || `${c.user.first_name || ""} ${c.user.last_name || ""}`.trim());
        CONFIG.agents.forEach((a, idx) => {
          const parts = a.toLowerCase().split(" ");
          const cparts = name.toLowerCase().split(" ");
          if (parts[0] === cparts[0] || (parts[1] && cparts[1] && parts[1] === cparts[1])) {
            agentOut[idx]++;
            if (c.duration > 0 && c.answered_at) {
              agentOutTalk[idx] += (c.duration || 0);
            }
          }
        });
      });

      // Store the computed answer rate (overwrite the variable)
      avgWaitTime = avgWaitTime; // already set
      // We need to store answerRate — we'll put it in the row
      inboundCalls = answeredCalls;
      // Store answer rate in a variable we can use below
      var computedAnswerRate = answerRate;
    }
  } catch (e) {
    Logger.log("Aircall backfill error: " + e);
  }

  // ── Nicereply: CSAT for that day (team + per-agent) ──
  let csatPct = "";
  let csatResponses = 0;
  let csatSatisfied = 0;
  const bfAgentEmailCsatSat = CONFIG.agents.map(() => 0);
  const bfAgentEmailCsatTot = CONFIG.agents.map(() => 0);
  const bfAgentPhoneCsatSat = CONFIG.agents.map(() => 0);
  const bfAgentPhoneCsatTot = CONFIG.agents.map(() => 0);
  try {
    const nrToken = props.getProperty("NICEREPLY_TOKEN");
    if (nrToken) {
      const nrAuth = "Basic " + Utilities.base64Encode(nrToken);
      const sinceISO = nextDate + "T00:00:00Z";
      const untilISO = nextDatePlusOne + "T00:00:00Z";
      let allResponses = [];
      let nrPage = 1;
      let nrMore = true;
      while (nrMore && nrPage <= 5) {
        const url = `https://api.nicereply.com/responses?created_after=${encodeURIComponent(sinceISO)}&created_before=${encodeURIComponent(untilISO)}&per_page=50&page=${nrPage}`;
        const resp = UrlFetchApp.fetch(url, { method: "get", headers: { "Authorization": nrAuth }, muteHttpExceptions: true });
        if (resp.getResponseCode() !== 200) break;
        const body = JSON.parse(resp.getContentText());
        allResponses = allResponses.concat(body.data || []);
        const pag = body.pagination || {};
        nrMore = pag.total_pages ? nrPage < pag.total_pages : (body.data || []).length >= 50;
        nrPage++;
      }

      const CSAT_QID = "86bae330-e8bc-4fa3-9af9-91eb2459d348";
      // Parse scores from each response
      const parsed = allResponses.map(r => {
        const answers = r.answers || [];
        const csatAnswer = answers.find(a => a.question_id === CSAT_QID) || answers.find(a => a.question_type === "SCALE");
        const score = csatAnswer && csatAnswer.scale ? csatAnswer.scale.value : 0;
        return { ticketId: r.ticket_id || "", score, satisfied: score >= 4 };
      });

      csatResponses = parsed.length;
      csatSatisfied = parsed.filter(p => p.satisfied).length;
      csatPct = csatResponses > 0 ? Math.round((csatSatisfied / csatResponses) * 100) : "";

      // Per-agent email CSAT: look up ticket assignees in Zendesk
      const ticketIdsForCsat = parsed.map(p => p.ticketId).filter(Boolean);
      if (ticketIdsForCsat.length > 0) {
        const ticketAssignees = {};
        for (let i = 0; i < ticketIdsForCsat.length; i += 100) {
          const batch = ticketIdsForCsat.slice(i, i + 100).join(",");
          const tUrl = `https://${subdomain}.zendesk.com/api/v2/tickets/show_many.json?ids=${batch}`;
          const tResp = UrlFetchApp.fetch(tUrl, zdOpts);
          if (tResp.getResponseCode() === 200) {
            const tData = JSON.parse(tResp.getContentText());
            // Need user names for assignees
            const assigneeIds = new Set();
            (tData.tickets || []).forEach(t => { if (t.assignee_id) assigneeIds.add(t.assignee_id); });
            const userNames = {};
            if (assigneeIds.size > 0) {
              const uUrl = `https://${subdomain}.zendesk.com/api/v2/users/show_many.json?ids=${[...assigneeIds].join(",")}`;
              const uResp = UrlFetchApp.fetch(uUrl, zdOpts);
              if (uResp.getResponseCode() === 200) {
                (JSON.parse(uResp.getContentText()).users || []).forEach(u => {
                  userNames[u.id] = u.name || "";
                });
              }
            }
            (tData.tickets || []).forEach(t => {
              ticketAssignees[t.id] = userNames[t.assignee_id] || "";
            });
          }
        }
        // Attribute to agents
        parsed.forEach(p => {
          if (!p.ticketId) return;
          const assigneeName = ticketAssignees[p.ticketId] || "";
          if (!assigneeName) return;
          const idx = CONFIG.agents.findIndex(ca => {
            const parts = ca.toLowerCase().split(/\s+/);
            const tLower = assigneeName.toLowerCase();
            return ca === assigneeName || parts.some(pt => pt.length > 1 && tLower.split(/\s+/).some(ap => ap === pt));
          });
          if (idx >= 0) {
            bfAgentEmailCsatTot[idx]++;
            if (p.satisfied) bfAgentEmailCsatSat[idx]++;
          }
        });
      }
    }
  } catch (e) {
    Logger.log("Nicereply backfill error: " + e);
  }

  // ── PostCall CSAT for that day (per-agent) ──
  try {
    const pcSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("PostCall Log");
    if (pcSheet && pcSheet.getLastRow() > 1) {
      const pcData = pcSheet.getRange(2, 1, pcSheet.getLastRow() - 1, 12).getValues();
      const dayStart = new Date(nextDate + "T00:00:00");
      const dayEnd = new Date(nextDatePlusOne + "T00:00:00");
      pcData.forEach(row => {
        const ts = new Date(row[0]);
        if (ts < dayStart || ts >= dayEnd) return;
        if (String(row[1] || "") !== "survey-completed") return;
        const agentName = String(row[4] || "");
        const excluded = CONFIG.excludePostCallAgents.some(ex => agentName.toLowerCase().includes(ex.toLowerCase()));
        if (excluded) return;
        const csatScore = Number(row[2]) || 0;
        if (!csatScore) return;
        const idx = CONFIG.agents.findIndex(a => {
          const parts = a.toLowerCase().split(/\s+/);
          const rParts = agentName.toLowerCase().split(/\s+/);
          return parts[0] === rParts[0] || (parts[1] && rParts[1] && parts[1] === rParts[1]);
        });
        if (idx >= 0) {
          bfAgentPhoneCsatTot[idx]++;
          if (csatScore >= 4) bfAgentPhoneCsatSat[idx]++;
        }
      });
    }
  } catch (e) {
    Logger.log("PostCall CSAT backfill error: " + e);
  }

  // ── Log the row ──
  logDailyMetrics({
    date: nextDate,
    dayOfWeek: dayOfWeek,
    openTickets: "",        // can't reconstruct
    unassigned: "",         // can't reconstruct
    onHold: "",             // can't reconstruct
    pastSla: "",            // can't reconstruct
    sasTickets: "",         // can't reconstruct
    openVoicemails: "",     // can't reconstruct
    aiAgentTickets: "",     // can't reconstruct
    ticketsCreated: ticketsCreated,
    solvedTotal: solvedTotal,
    agentSolved: agentSolved,
    agentInbound: agentIn,
    agentOutbound: agentOut,
    agentInTalk: agentInTalk,
    agentOutTalk: agentOutTalk,
    answerRate: typeof computedAnswerRate !== "undefined" ? computedAnswerRate : "",
    inboundCalls: inboundCalls,
    forwardedToSAS: forwardedToSAS,
    outboundCalls: outboundCalls,
    avgWaitTime: avgWaitTime,
    avgCallDuration: avgDuration,
    csatPct: csatPct,
    csatResponses: csatResponses,
    csatSatisfied: csatSatisfied,
    phoneCsatPct: "",       // team phone CSAT not easily backfillable
    phoneCsatResponses: 0,
    agentEmailCsatSat: bfAgentEmailCsatSat,
    agentEmailCsatTot: bfAgentEmailCsatTot,
    agentPhoneCsatSat: bfAgentPhoneCsatSat,
    agentPhoneCsatTot: bfAgentPhoneCsatTot,
    unreadDMs: "",          // can't reconstruct
    smsInbound: "",         // SMS log could be read, but skip for now
    smsOutbound: "",
  });

  advanceBackfillDate(props, nextDate);
  Logger.log("Backfill complete for " + nextDate);
}

function advanceBackfillDate(props, currentDate) {
  const next = new Date(currentDate + "T12:00:00");
  next.setDate(next.getDate() + 1);
  const tz = CONFIG.businessHours.timezone;
  const nextStr = Utilities.formatDate(next, tz, "yyyy-MM-dd");
  props.setProperty("BACKFILL_NEXT_DATE", nextStr);
}

function setupBackfillTrigger() {
  cleanupBackfillTrigger();
  ScriptApp.newTrigger("backfillOneDay")
    .timeBased()
    .everyMinutes(10)
    .create();
  Logger.log("Backfill trigger created — runs every 10 minutes. Set BACKFILL_START_DATE to begin.");
}

function cleanupBackfillTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "backfillOneDay") {
      ScriptApp.deleteTrigger(t);
      Logger.log("Removed backfill trigger");
    }
  });
}

// Slimmed-down email for weekends and Deako holidays — queue state only, no performance metrics
function sendNonWorkingDaySnapshot(zendesk, meta, now, dateStr, tz, recipients) {
  const props = PropertiesService.getScriptProperties();
  const navy = "#1B3747";
  const cardBg = "#F8F7F6";
  const borderColor = "#E1DFDD";
  const sasCount = (zendesk.sasTickets && zendesk.sasTickets.length) || 0;
  const vmCount = zendesk.openVoicemails || 0;
  const metaUnread = meta.unreadDMs || 0;
  const metaConvos = (meta.recentConversations || []);

  // Load last working day snapshot for comparison
  const prevJson = props.getProperty("RECAP_PREV_SNAPSHOT") || "{}";
  let prev = {};
  try { prev = JSON.parse(prevJson); } catch (e) { prev = {}; }
  const prevDate = prev.date || "";
  // Format prev date as "May 15 EOD" for column header
  let prevLabel = "";
  if (prevDate) {
    const pd = new Date(prevDate + "T12:00:00");  // noon to avoid timezone issues
    prevLabel = Utilities.formatDate(pd, tz, "MMM d") + " EOD";
  }

  // Determine if we have previous data for the EOD column
  const hasPrev = prev.date ? true : false;
  function prevTd(prevVal, unit) {
    if (!hasPrev || prevVal === undefined || prevVal === null) return hasPrev ? `<td style="text-align:right;color:#AAA;font-size:12px;">-</td>` : "";
    const u = unit || "";
    return `<td style="text-align:right;color:#888;font-size:12px;">${prevVal}${u}</td>`;
  }
  const nwdCompHeader = hasPrev
    ? `<tr style="border-bottom:1px solid #E1DFDD;"><td style="width:50%;"></td><td style="width:25%;text-align:right;font-size:11px;color:#888;padding-bottom:6px;font-weight:bold;">Now</td><td style="width:25%;text-align:right;font-size:11px;color:#888;padding-bottom:6px;font-weight:bold;">${prevLabel}</td></tr>`
    : "";

  let html = `
  <div style="font-family:Arial,sans-serif;max-width:700px;margin:0 auto;color:#1D1D1D;">
    <div style="background:${navy};color:#fff;padding:16px 24px;border-radius:8px 8px 0 0;">
      <h1 style="margin:0;font-size:20px;">CS AI Recap - Non-Working Day</h1>
      <p style="margin:4px 0 0;font-size:13px;color:#C3D3D7;">${dateStr}</p>
    </div>

    <div style="padding:20px 24px;background:#fff;border:1px solid ${borderColor};border-top:none;">

      <div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
        <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Email Queue (Zendesk)</h2>
        <table style="width:100%;font-size:13px;border-collapse:collapse;">
          ${nwdCompHeader}
          <tr><td style="padding:4px 0;">Open Tickets</td><td style="text-align:right;font-weight:bold;">${zendesk.totalOpen}</td>${prevTd(prev.openTickets)}</tr>
          <tr><td style="padding:4px 0;">On Hold</td><td style="text-align:right;font-weight:bold;">${zendesk.onHoldQueueCount}</td>${prevTd(prev.onHold)}</tr>
          <tr><td style="padding:4px 0;">Unassigned</td><td style="text-align:right;font-weight:bold;">${zendesk.unassigned}</td>${prevTd(prev.unassigned)}</tr>
          <tr><td style="padding:4px 0;">SAS Tickets</td><td style="text-align:right;font-weight:bold;">${sasCount}</td>${prevTd(prev.sasTickets)}</tr>
        </table>
      </div>

      <div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
        <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Phone (Voicemails)</h2>
        <table style="width:100%;font-size:13px;border-collapse:collapse;">
          ${nwdCompHeader}
          <tr><td style="padding:4px 0;">Open Voicemails</td><td style="text-align:right;font-weight:bold;">${vmCount}</td>${prevTd(prev.openVoicemails)}</tr>
        </table>
      </div>

      <div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
        <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">Social (Meta Business Suite)</h2>
        <table style="width:100%;font-size:13px;border-collapse:collapse;">
          ${nwdCompHeader}
          <tr><td style="padding:4px 0;">Unread DMs</td><td style="text-align:right;font-weight:bold;color:${metaUnread > 0 ? '#C62828' : '#2E7D32'};">${metaUnread}</td>${prevTd(prev.unreadDMs)}</tr>
          <tr><td style="padding:4px 0;">Active Conversations (24h)</td><td style="text-align:right;font-weight:bold;">${metaConvos.length}</td>${prevTd(prev.activeConversations)}</tr>
        </table>
      </div>

`;

  // Run AI Recap theme analysis for non-working days too
  let dailyThemesHtml = "";
  try {
    const themesResult = analyzeDailyThemes();
    if (themesResult) {
      dailyThemesHtml = typeof themesResult === "string" ? themesResult : themesResult.html;
    }
  } catch (e) {
    Logger.log("Non-working day theme analysis failed (non-fatal): " + e.toString());
  }

  if (dailyThemesHtml) {
    html += `
      <div style="background:${cardBg};border:1px solid ${borderColor};border-radius:6px;padding:16px;margin-bottom:16px;">
        <h2 style="margin:0 0 12px;font-size:15px;color:${navy};">AI Recap - Top 3 Themes</h2>
        <div style="font-size:13px;line-height:1.6;">${dailyThemesHtml}</div>
        <div style="margin-top:10px;font-size:10px;color:#999;">Analysis powered by Claude AI across all tickets and call intents</div>
      </div>`;
  }

  html += `
      <div style="text-align:center;padding:16px 0 8px;font-size:11px;color:#999;">
        CS Visibility · AI Recap v2.5.4 · Non-Working Day · ${dateStr}
      </div>
    </div>
  </div>`;

  const subject = `CS AI Recap - Non-Working Day - ${Utilities.formatDate(now, tz, "MMM d")} - Open: ${zendesk.totalOpen} · VM: ${vmCount} · Unread DMs: ${metaUnread}`;

  GmailApp.sendEmail(recipients, subject, "View this email with HTML enabled.", {
    htmlBody: html,
    name: "CS Visibility",
  });

  Logger.log("Non-working day snapshot sent to: " + recipients);
}

// ─── CLEANUP & PATCH DAILY METRICS LOG ───
// Run once to sort rows by date and patch in phone CSAT from PostCall survey data.
function cleanupDailyMetricsLog() {
  loadThresholds();
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Daily Metrics Log");
  if (!sheet || sheet.getLastRow() <= 1) { Logger.log("No data to clean up"); return; }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getLastRow();
  const data = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

  // Find column indices
  const dateCol = headers.indexOf("Date");
  const phoneCsatCol = headers.indexOf("Phone CSAT %");
  const phoneCsatRespCol = headers.indexOf("Phone CSAT Responses");

  if (dateCol < 0) { Logger.log("Can't find Date column"); return; }

  // Phone CSAT data from Aircall PostCall survey export (March-June 2026)
  const phoneCsatData = {
    "2026-04-01": { pct: 83, responses: 6 }, "2026-04-02": { pct: 100, responses: 7 },
    "2026-04-03": { pct: 71, responses: 7 }, "2026-04-04": { pct: 100, responses: 1 },
    "2026-04-06": { pct: 100, responses: 8 }, "2026-04-07": { pct: 100, responses: 3 },
    "2026-04-08": { pct: 100, responses: 5 }, "2026-04-09": { pct: 100, responses: 6 },
    "2026-04-10": { pct: 88, responses: 8 }, "2026-04-11": { pct: 100, responses: 1 },
    "2026-04-12": { pct: 100, responses: 1 }, "2026-04-13": { pct: 75, responses: 4 },
    "2026-04-14": { pct: 100, responses: 2 }, "2026-04-15": { pct: 100, responses: 7 },
    "2026-04-17": { pct: 100, responses: 2 }, "2026-04-20": { pct: 100, responses: 1 },
    "2026-04-21": { pct: 100, responses: 10 }, "2026-04-22": { pct: 100, responses: 5 },
    "2026-04-23": { pct: 100, responses: 2 }, "2026-04-24": { pct: 100, responses: 4 },
    "2026-04-27": { pct: 100, responses: 6 }, "2026-04-28": { pct: 75, responses: 4 },
    "2026-04-29": { pct: 100, responses: 3 }, "2026-04-30": { pct: 100, responses: 3 },
    "2026-05-01": { pct: 100, responses: 4 }, "2026-05-04": { pct: 100, responses: 3 },
    "2026-05-05": { pct: 100, responses: 4 }, "2026-05-06": { pct: 100, responses: 3 },
    "2026-05-07": { pct: 100, responses: 7 }, "2026-05-08": { pct: 100, responses: 5 },
    "2026-05-09": { pct: 100, responses: 2 }, "2026-05-11": { pct: 100, responses: 3 },
    "2026-05-12": { pct: 100, responses: 3 }, "2026-05-13": { pct: 100, responses: 6 },
    "2026-05-14": { pct: 100, responses: 8 }, "2026-05-15": { pct: 67, responses: 3 },
    "2026-05-18": { pct: 100, responses: 2 }, "2026-05-19": { pct: 83, responses: 6 },
    "2026-05-20": { pct: 100, responses: 6 }, "2026-05-21": { pct: 100, responses: 2 },
    "2026-05-22": { pct: 100, responses: 4 }, "2026-05-26": { pct: 100, responses: 4 },
    "2026-05-27": { pct: 90, responses: 10 }, "2026-05-28": { pct: 100, responses: 3 },
    "2026-05-29": { pct: 100, responses: 5 }, "2026-05-30": { pct: 100, responses: 1 },
    "2026-06-01": { pct: 100, responses: 3 },
  };

  // SMS data from Aircall messages export (April-May 2026)
  const smsData = {
    "2026-04-01": { inbound: 32, outbound: 23 }, "2026-04-02": { inbound: 30, outbound: 13 },
    "2026-04-03": { inbound: 30, outbound: 18 }, "2026-04-06": { inbound: 66, outbound: 47 },
    "2026-04-07": { inbound: 33, outbound: 24 }, "2026-04-08": { inbound: 50, outbound: 52 },
    "2026-04-09": { inbound: 56, outbound: 62 }, "2026-04-10": { inbound: 49, outbound: 37 },
    "2026-04-13": { inbound: 25, outbound: 17 }, "2026-04-14": { inbound: 23, outbound: 25 },
    "2026-04-15": { inbound: 30, outbound: 23 }, "2026-04-16": { inbound: 17, outbound: 15 },
    "2026-04-17": { inbound: 42, outbound: 40 }, "2026-04-20": { inbound: 23, outbound: 17 },
    "2026-05-19": { inbound: 24, outbound: 53 }, "2026-05-20": { inbound: 46, outbound: 51 },
    "2026-05-21": { inbound: 32, outbound: 31 }, "2026-05-22": { inbound: 27, outbound: 30 },
    "2026-05-26": { inbound: 33, outbound: 52 }, "2026-05-27": { inbound: 42, outbound: 49 },
    "2026-05-28": { inbound: 32, outbound: 30 }, "2026-05-29": { inbound: 35, outbound: 25 },
  };

  const smsInCol = headers.indexOf("SMS Inbound");
  const smsOutCol = headers.indexOf("SMS Outbound");
  const ticketsCreatedCol = headers.indexOf("Tickets Created");

  const tz = Session.getScriptTimeZone();

  // Normalize dates and patch phone CSAT
  data.forEach(row => {
    // Normalize date to string
    const rawDate = row[dateCol];
    if (rawDate instanceof Date) {
      row[dateCol] = Utilities.formatDate(rawDate, tz, "yyyy-MM-dd");
    }
    const dateStr = String(row[dateCol]);

    // Patch phone CSAT if we have data and the cell is empty/zero
    if (phoneCsatCol >= 0 && phoneCsatRespCol >= 0 && phoneCsatData[dateStr]) {
      const existing = row[phoneCsatCol];
      if (existing === "" || existing === 0 || existing === null || existing === undefined) {
        row[phoneCsatCol] = phoneCsatData[dateStr].pct;
        row[phoneCsatRespCol] = phoneCsatData[dateStr].responses;
        Logger.log("Patched phone CSAT for " + dateStr + ": " + phoneCsatData[dateStr].pct + "% (" + phoneCsatData[dateStr].responses + " responses)");
      }
    }

    // Patch Tickets Created by querying Zendesk if empty
    if (ticketsCreatedCol >= 0) {
      const existingCreated = row[ticketsCreatedCol];
      if (existingCreated === "" || existingCreated === null || existingCreated === undefined) {
        try {
          const zdToken = props.getProperty("ZENDESK_TOKEN");
          if (zdToken) {
            const subdomain = CONFIG.zendesk.subdomain;
            const zdAuth = "Basic " + Utilities.base64Encode(zdToken);
            const nextDay = new Date(dateStr + "T12:00:00");
            nextDay.setDate(nextDay.getDate() + 1);
            const nextDayStr = nextDay.toISOString().split("T")[0];
            const solvedFilter = `-tags:aircall -tags:internal__testing -tags:auto_close -assignee:"AI Agent"`;
            const url = `https://${subdomain}.zendesk.com/api/v2/search/count.json?query=${encodeURIComponent("type:ticket created>=" + dateStr + " created<" + nextDayStr + " " + solvedFilter)}`;
            const resp = UrlFetchApp.fetch(url, { method: "get", headers: { "Authorization": zdAuth, "Content-Type": "application/json" }, muteHttpExceptions: true });
            if (resp.getResponseCode() === 200) {
              const count = JSON.parse(resp.getContentText()).count || 0;
              row[ticketsCreatedCol] = count;
              Logger.log("Patched Tickets Created for " + dateStr + ": " + count);
            }
          }
        } catch (e) { Logger.log("Tickets Created patch failed for " + dateStr + ": " + e); }
      }
    }

    // Patch SMS if we have data and the cell is empty/zero
    if (smsInCol >= 0 && smsOutCol >= 0 && smsData[dateStr]) {
      const existingSms = row[smsInCol];
      if (existingSms === "" || existingSms === 0 || existingSms === null || existingSms === undefined) {
        row[smsInCol] = smsData[dateStr].inbound;
        row[smsOutCol] = smsData[dateStr].outbound;
        Logger.log("Patched SMS for " + dateStr + ": " + smsData[dateStr].inbound + " in / " + smsData[dateStr].outbound + " out");
      }
    }
  });

  // Sort by date ascending
  data.sort((a, b) => {
    const da = String(a[dateCol]);
    const db = String(b[dateCol]);
    return da < db ? -1 : da > db ? 1 : 0;
  });

  // Write back — clear existing data rows first, then write sorted data
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);

  Logger.log("Daily Metrics Log cleaned: " + data.length + " rows sorted by date, phone CSAT patched");
}

// ─── CLEANUP DAILY THEMES LOG ───
// Run once to deduplicate theme entries. Keeps the LAST set of 3 themes per date
// (most recent backfill/run), removes earlier duplicates, and sorts by date.
function cleanupDailyThemesLog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Daily Themes Log");
  if (!sheet || sheet.getLastRow() <= 1) { Logger.log("No theme data to clean up"); return; }

  const tz = Session.getScriptTimeZone();
  const lastRow = sheet.getLastRow();
  const numCols = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();

  // Normalize dates and group by date
  const byDate = {};
  data.forEach(row => {
    const rawDate = row[0];
    if (!rawDate) return;
    const dateStr = (rawDate instanceof Date)
      ? Utilities.formatDate(rawDate, tz, "yyyy-MM-dd")
      : String(rawDate);
    row[0] = dateStr;
    if (!byDate[dateStr]) byDate[dateStr] = [];
    byDate[dateStr].push(row);
  });

  // For each date, keep only the last 3 rows (most recent run)
  const cleaned = [];
  Object.keys(byDate).sort().forEach(date => {
    const rows = byDate[date];
    // Take the last 3 entries (most recent backfill overwrites earlier ones)
    const keep = rows.slice(-3);
    cleaned.push(...keep);
  });

  // Clear and rewrite
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, numCols).clearContent();
  }
  if (cleaned.length > 0) {
    sheet.getRange(2, 1, cleaned.length, numCols).setValues(cleaned);
  }

  const dateCount = Object.keys(byDate).length;
  const removedCount = data.length - cleaned.length;
  Logger.log("Daily Themes Log cleaned: " + dateCount + " unique dates, " + cleaned.length + " rows kept, " + removedCount + " duplicates removed");
}

// ─── RERUN WEEKLY SUMMARY ───
// Regenerates a weekly summary email for a specific past week.
// Set WEEKLY_RERUN_DATE in Script Properties to any date within the target week (e.g., "2026-05-27").
// Sends to TEST_EMAIL so it doesn't spam the full recipient list.
function rerunWeeklySummary() {
  loadThresholds();
  const props = PropertiesService.getScriptProperties();
  const targetDate = props.getProperty("WEEKLY_RERUN_DATE");
  if (!targetDate) {
    Logger.log("WEEKLY_RERUN_DATE not set — set it to any date within the target week (e.g., 2026-05-27)");
    return;
  }
  const testEmail = props.getProperty("TEST_EMAIL") || "";
  if (!testEmail) {
    Logger.log("TEST_EMAIL not set — need a recipient for the rerun");
    return;
  }

  const tz = CONFIG.businessHours.timezone;
  // Create a Date object from the target date, treated as if "now" is that Friday
  const d = new Date(targetDate + "T17:00:00");
  const thisWeek = getThisWeekRange(d, tz);
  const lastWeek = getLastWeekRange(d, tz);
  Logger.log("Rerunning weekly summary for week containing " + targetDate);
  Logger.log("This week range: " + thisWeek.start + " to " + thisWeek.end);
  Logger.log("Last week range: " + lastWeek.start + " to " + lastWeek.end);
  const debugData = readMetricsLog(thisWeek.start, thisWeek.end);
  Logger.log("Metrics rows found for this week: " + debugData.length);
  if (debugData.length > 0) Logger.log("First row date: " + debugData[0]._dateStr);
  sendWeeklySummary(d, tz, testEmail);
  Logger.log("Weekly rerun sent to " + testEmail);
}

// ─── TEST EMAIL FUNCTIONS ───
// All test functions send to TEST_EMAIL script property only.
// Set TEST_EMAIL in Script Properties to your email address before running.
// This prevents accidentally modifying RECAP_RECIPIENTS during testing.

function getTestEmail_() {
  const email = PropertiesService.getScriptProperties().getProperty("TEST_EMAIL") || "";
  if (!email) {
    Logger.log("TEST_EMAIL script property not set. Add it in Project Settings > Script Properties.");
    throw new Error("TEST_EMAIL not set - add your email in Script Properties before running test functions.");
  }
  return email;
}

// Test: Daily Recap email with live data
function testDailyRecapEmail() {
  const testEmail = getTestEmail_();
  sendDailyRecap(testEmail, true);  // skipSave = true for test mode
  Logger.log("Test daily recap sent to: " + testEmail);
}

// Test: Weekly Summary email with this week's data
function testWeeklySummaryEmail() {
  const testEmail = getTestEmail_();
  loadThresholds();
  const tz = CONFIG.businessHours.timezone;
  const now = new Date();
  sendWeeklySummary(now, tz, testEmail);
  Logger.log("Test weekly summary sent to: " + testEmail);
}

// Test: Monthly Summary email with this month's data
// Rerun monthly summary for a specific month.
// Set MONTHLY_RERUN_DATE to any date within the target month (e.g., "2026-05-15").
// Sends to TEST_EMAIL.
function rerunMonthlySummary() {
  loadThresholds();
  const props = PropertiesService.getScriptProperties();
  const targetDate = props.getProperty("MONTHLY_RERUN_DATE");
  if (!targetDate) {
    Logger.log("MONTHLY_RERUN_DATE not set — set it to any date within the target month (e.g., 2026-05-15)");
    return;
  }
  const testEmail = props.getProperty("TEST_EMAIL") || "";
  if (!testEmail) {
    Logger.log("TEST_EMAIL not set — need a recipient for the rerun");
    return;
  }
  const tz = CONFIG.businessHours.timezone;
  const d = new Date(targetDate + "T17:00:00");
  Logger.log("Rerunning monthly summary for month containing " + targetDate);
  sendMonthlySummary(d, tz, testEmail);
  Logger.log("Monthly rerun sent to " + testEmail);
}

function testMonthlySummaryEmail() {
  const testEmail = getTestEmail_();
  loadThresholds();
  const tz = CONFIG.businessHours.timezone;
  const now = new Date();
  sendMonthlySummary(now, tz, testEmail);
  Logger.log("Test monthly summary sent to: " + testEmail);
}

// Test: Non-working day snapshot email with live data
function testNonWorkingDayEmail() {
  const testEmail = getTestEmail_();
  loadThresholds();
  const zendesk = fetchZendeskStatus();
  const meta = fetchMetaStatus();
  const tz = CONFIG.businessHours.timezone;
  const now = new Date();
  const dateStr = Utilities.formatDate(now, tz, "EEEE, MMMM d, yyyy");
  sendNonWorkingDaySnapshot(zendesk, meta, now, dateStr, tz, testEmail);
  Logger.log("Test non-working day email sent to: " + testEmail);
}

function setupDailyRecapTrigger() {
  // Remove any existing recap triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === "sendDailyRecap") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Create new trigger: daily at 6pm PST (18:00)
  ScriptApp.newTrigger("sendDailyRecap")
    .timeBased()
    .atHour(18)
    .everyDays(1)
    .inTimezone("America/Los_Angeles")
    .create();

  Logger.log("Daily recap trigger set for 6:00 PM PST");
}

// Set up weekly + monthly summary triggers (run once from Apps Script editor)
function setupWeeklyMonthlyTriggers() {
  // Remove existing weekly/monthly triggers
  ScriptApp.getProjectTriggers().forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === "checkAndSendWeeklySummary" || fn === "checkAndSendMonthlySummary") {
      ScriptApp.deleteTrigger(t);
    }
  });

  // Weekly check: daily at 6:15pm — only sends on last working day of week
  ScriptApp.newTrigger("checkAndSendWeeklySummary")
    .timeBased()
    .atHour(18)
    .nearMinute(15)
    .everyDays(1)
    .inTimezone("America/Los_Angeles")
    .create();

  // Monthly check: daily at 6:30pm — only sends on last business day of month
  ScriptApp.newTrigger("checkAndSendMonthlySummary")
    .timeBased()
    .atHour(18)
    .nearMinute(30)
    .everyDays(1)
    .inTimezone("America/Los_Angeles")
    .create();

  Logger.log("Weekly summary trigger set for ~6:15 PM PST (sends on last working day of week)");
  Logger.log("Monthly summary trigger set for ~6:30 PM PST (sends on last business day of month)");
}

// --- HELPER FUNCTIONS ---

// --- ONE-TIME CLEANUP: Fix #ERROR! phone numbers in SMS Log ---
// Run this once from the Apps Script editor to fix historical phone number cells
// that Sheets interpreted as formulas. Safe to run multiple times.
function fixSMSLogPhoneErrors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SMS Log");
  if (!sheet || sheet.getLastRow() <= 1) {
    Logger.log("SMS Log empty or missing — nothing to fix");
    return;
  }

  const lastRow = sheet.getLastRow();
  // Column F (6) = Contact Phone, Column K (11) = Line Number
  const colsToFix = [6, 11];
  let fixed = 0;

  colsToFix.forEach(col => {
    const range = sheet.getRange(2, col, lastRow - 1, 1);
    // Set entire column to plain text first
    range.setNumberFormat("@");
    const formulas = sheet.getRange(2, col, lastRow - 1, 1).getFormulas();
    const values = sheet.getRange(2, col, lastRow - 1, 1).getValues();

    for (let i = 0; i < formulas.length; i++) {
      const formula = formulas[i][0];
      const value = values[i][0];
      // If cell has a formula (Sheets misinterpreted +number as formula) or shows error
      if (formula) {
        // The formula IS the original value Sheets tried to evaluate (e.g., "+14155551234")
        const originalValue = formula.startsWith("=") ? formula.substring(1) : formula;
        sheet.getRange(i + 2, col).setValue("'" + originalValue);
        fixed++;
      } else if (String(value).includes("#ERROR") || String(value).includes("#REF") || String(value).includes("#VALUE")) {
        // Can't recover the original — mark it
        sheet.getRange(i + 2, col).setValue("(phone error — check Aircall)");
        fixed++;
      }
    }
  });

  Logger.log("Fixed " + fixed + " phone number cells in SMS Log");
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function getHealthLevel(value, greenThreshold, yellowThreshold, direction) {
  if (direction === "lower") {
    if (value < greenThreshold) return "green";
    if (value < yellowThreshold) return "yellow";
    return "red";
  } else {
    if (value > greenThreshold) return "green";
    if (value > yellowThreshold) return "yellow";
    return "red";
  }
}

function formatSeconds(totalSeconds) {
  if (totalSeconds === 0) return "0s";
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = Math.round(totalSeconds % 60);
  if (minutes === 0) return `${seconds}s`;
  return `${minutes}m ${seconds}s`;
}

// Format talk time — hours and minutes only, no seconds.
// Examples: 0 → "0m", 50s → "0m", 15m 50s → "15m", 186m 4s → "3h 6m"
function formatTalkTime(totalSeconds) {
  if (!totalSeconds || totalSeconds <= 0) return "0m";
  const totalMinutes = Math.round(totalSeconds / 60);
  if (totalMinutes < 60) return `${totalMinutes}m`;
  const hours = Math.floor(totalMinutes / 60);
  const mins = totalMinutes % 60;
  return mins > 0 ? `${hours}h ${mins}m` : `${hours}h`;
}

// Format business minutes as business hours (e.g. "2.5h", "125h")
function formatBizMinutes(min) {
  if (min === 0) return "-";
  const hours = min / 60;
  if (hours < 1) return Math.round(min) + "m";
  if (hours < 10) return hours.toFixed(1) + "h";
  return Math.round(hours) + "h";
}


