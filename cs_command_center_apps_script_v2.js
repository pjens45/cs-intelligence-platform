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
  // Aircall lines to exclude from SMS activity tracking
  excludeSMSLines: ["nonpro sales (post close)", "Agent D", "Agent E", "Agent F", "Agent G", "Agent H", "Agent I", "Agent J", "Agent K", "Agent L"],
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

// --- MAIN REFRESH FUNCTION ---
function refreshDashboard() {
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

    writeZendeskRaw(ss, zendeskData);
    writeAircallRaw(ss, aircallData);
    writeDashboard(ss, zendeskData, aircallData, csatData, postCallData, smsData, metaData);

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

  // "Emails handled today" — tickets solved or closed today, per agent
  // Uses Zendesk search: type:ticket status:solved solved>=today assignee:"Name"
  const todayStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const handledToday = {};
  CONFIG.agents.forEach(agent => {
    // Zendesk search supports solved>=date for tickets solved on or after that date
    const query = `type:ticket solved>=${todayStr} assignee:"${agent}"`;
    handledToday[agent] = zendeskSearchCount(query);
  });
  // Also count "Other" and total
  const totalHandledQuery = `type:ticket solved>=${todayStr}`;
  const totalHandledToday = zendeskSearchCount(totalHandledQuery);
  const knownHandled = Object.values(handledToday).reduce((a, b) => a + b, 0);
  handledToday["Other"] = Math.max(0, totalHandledToday - knownHandled);

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

  return {
    totalOpen,
    totalBreached,
    noReplyBreached,
    sasTickets,
    slaHours,
    otherAgentNames: [...otherAgentNames],
    unassigned,
    openQueueCount,
    onHoldQueueCount,
    agentCounts,
    tickets: filtered,
    flaggedTickets,
    totalHandledToday,
    // Top 10 longest-waiting tickets for the detail table
    longest10: filtered.slice(0, 10),
  };
}

// Calculate business minutes between two dates (6a-5p Mon-Fri PST)
function calcBusinessMinutes(start, end) {
  const bh = CONFIG.businessHours;
  const minPerDay = (bh.endHour - bh.startHour) * 60; // 660 min for 6a-5p

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

    if (bh.workDays.includes(dow)) {
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

  // Filter to business hours only (6am-5pm Mon-Fri Pacific)
  const bh = CONFIG.businessHours;
  const bizHourCalls = inboundCalls.filter(c => {
    if (!c.started_at) return false;
    // Convert Unix timestamp to Pacific time
    const callDate = new Date(c.started_at * 1000);
    const pacificStr = callDate.toLocaleString("en-US", { timeZone: bh.timezone });
    const pacificDate = new Date(pacificStr);
    const dayOfWeek = pacificDate.getDay(); // 0=Sun, 1=Mon ... 6=Sat
    const hour = pacificDate.getHours();
    return bh.workDays.includes(dayOfWeek) && hour >= bh.startHour && hour < bh.endHour;
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
  CONFIG.agents.forEach(name => agentStats[name] = { answered: 0, outbound: 0, outboundConnected: 0, outboundShort: 0, outboundLong: 0 });

  teamAnswered.forEach(t => {
    agentStats[t.agent].answered++;
  });

  // ─── Outbound calls per agent ───
  // Filter all today's support calls that are outbound (exclude calls TO the answer service)
  const outboundCalls = supportCalls.filter(c => {
    if (c.direction !== "outbound") return false;
    // Exclude outbound to the answering service number
    const rawDigits = (c.raw_digits || "").replace(/[\s\-\(\)]/g, "");
    if (answerSvcNum && rawDigits.includes(answerSvcNum.replace("+1", ""))) return false;
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
function writeDashboard(ss, zendesk, aircall, csat, postCall, sms, meta) {
  // Build on a hidden staging sheet, then swap — eliminates the 5-min blink
  const staging = getOrCreateSheet(ss, "_Staging");
  staging.showSheet(); // ensure it exists and is accessible
  _writeDashboardContent(ss, staging, zendesk, aircall, csat, postCall, sms, meta);

  // Swap staging content → Dashboard in one batch
  const dash = getOrCreateSheet(ss, "Dashboard");
  const lastRow = staging.getLastRow() || 1;
  const lastCol = Math.max(staging.getLastColumn(), 24);

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

  // Preserve dashboard appearance
  dash.setTabColor(BRAND.airBlueDark);
  dash.setHiddenGridlines(true);

  // Hide staging
  staging.hideSheet();
}

function _writeDashboardContent(ss, dash, zendesk, aircall, csat, postCall, sms, meta) {
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

  dash.getRange("A:X").setFontFamily("Inter").setFontColor(darkText).setBackground(bg);
  dash.setTabColor(navy);
  dash.setHiddenGridlines(true);

  // ─── Status logic ───
  const emailStatus = zendesk.totalBreached > 5 ? "At Risk"
    : zendesk.totalBreached > 0 ? "Watch" : "Healthy";
  const emailColor = emailStatus === "At Risk" ? risk
    : emailStatus === "Watch" ? amber : green;

  const phoneGreen = CONFIG.thresholds.phoneAnswerRate.green;
  const phoneYellow = CONFIG.thresholds.phoneAnswerRate.yellow;
  const phoneStatus = aircall.teamAnswerRate >= phoneGreen ? "Healthy"
    : aircall.teamAnswerRate >= phoneYellow ? "Watch" : "At Risk";
  const phoneColor = phoneStatus === "At Risk" ? risk
    : phoneStatus === "Watch" ? amber : green;

  // Social status — based on oldest unread DM response time
  const socialOldestWaitMin = computeSocialOldestWait(meta);
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
  dash.getRange("Q1:X1").merge()
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
  dash.getRange("Q2:X2").merge()
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
    dash.getRange("A3:X3").merge()
      .setValue(`⚠  ALL HANDS ON DECK  —  ${zendesk.unassigned} UNASSIGNED TICKETS  ⚠`)
      .setBackground(alertRed).setFontColor("#FFFFFF")
      .setFontSize(18).setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    dash.setRowHeight(3, 56);
    dash.getRange("A4:X4").merge()
      .setValue("Unassigned ticket count has exceeded " + UNASSIGNED_ALERT_THRESHOLD + ". All available agents should begin triaging unassigned tickets immediately.")
      .setBackground(alertRedLight).setFontColor(alertRed)
      .setFontSize(11).setFontWeight("bold")
      .setHorizontalAlignment("center").setVerticalAlignment("middle");
    dash.setRowHeight(4, 32);
    // Row 5: divider after alert
    dash.getRange("A5:X5").setBackground(divider);
    dash.setRowHeight(5, 2);
  } else {
    // Row 3: normal thin divider
    dash.getRange("A3:X3").setBackground(divider);
    dash.setRowHeight(3, 2);
  }

  // Dynamic row offset — everything below shifts down by 2 when alert is showing
  const alertOffset = showAlert ? 2 : 0;

  // ═══════════════════════════════════════════════
  // KPI panels (numbers on top, labels below)
  // ═══════════════════════════════════════════════

  // Column H & P: spacers between columns
  dash.setColumnWidth(8, 20);   // H spacer
  dash.setColumnWidth(16, 20);  // P spacer

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
  dash.getRange(`M${k1}:O${k1}`).merge().setBackground(cardBg)
    .setValue(aircall.totalOutbound).setNumberFormat("0")
    .setFontSize(18).setFontWeight("bold").setFontColor(darkText)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");

  // Social KPIs — row 1: big numbers
  const metaUnread = (meta && meta.unreadDMs) || 0;
  const metaConversations = (meta && meta.recentConversations) || [];
  const metaComments = (meta && meta.recentComments) || [];
  dash.getRange(`Q${k1}:R${k1}`).merge().setBackground(cardBg)
    .setValue(metaUnread).setNumberFormat("0")
    .setFontSize(28).setFontWeight("bold").setFontColor(socialColor)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");
  dash.getRange(`S${k1}:T${k1}`).merge().setBackground(cardBg)
    .setValue(Math.round(socialOldestWaitMin / 60)).setNumberFormat("0")
    .setFontSize(18).setFontWeight("bold").setFontColor(darkText)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");
  dash.getRange(`U${k1}:X${k1}`).merge().setBackground(cardBg)
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
  dash.getRange(`M${k2}:O${k2}`).merge().setBackground(cardBg)
    .setValue("Outbound")
    .setFontSize(8).setFontColor(label).setVerticalAlignment("top");

  // Social labels — row 2
  dash.getRange(`Q${k2}:R${k2}`).merge().setBackground(cardBg)
    .setValue("Unread DMs")
    .setFontSize(8).setFontColor(label).setVerticalAlignment("top");
  dash.getRange(`S${k2}:T${k2}`).merge().setBackground(cardBg)
    .setValue("Oldest DM Wait (h)")
    .setFontSize(8).setFontColor(label).setVerticalAlignment("top");
  dash.getRange(`U${k2}:X${k2}`).merge().setBackground(cardBg)
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
  dash.getRange(`K${k3}:O${k3}`).merge().setBackground(cardBg);

  // Social row 3: Total conversations
  const totalConversations = metaConversations.length > 0 ? metaConversations.length : 0;
  dash.getRange(`Q${k3}:R${k3}`).merge().setBackground(cardBg)
    .setValue(totalConversations).setNumberFormat("0")
    .setFontSize(16).setFontWeight("bold").setFontColor(darkText)
    .setVerticalAlignment("bottom").setHorizontalAlignment("left");
  dash.getRange(`S${k3}:X${k3}`).merge().setBackground(cardBg);

  // --- ROW 7: Labels under queue counts ---
  const k4 = 7 + alertOffset;
  dash.setRowHeight(k4, 16);

  // Open label always shows SAS sub-count
  const sasColor = zendesk.sasTickets > 0 ? amber : navy;
  const sasLabel = `Open · ${zendesk.sasTickets} New SAS`;
  dash.getRange(`A${k4}:B${k4}`).merge().setBackground(cardBg)
    .setRichTextValue(
      SpreadsheetApp.newRichTextValue()
        .setText(sasLabel)
        .setTextStyle(0, 4, SpreadsheetApp.newTextStyle().setFontSize(8).setForegroundColor(navy).build())
        .setTextStyle(4, sasLabel.length,
          SpreadsheetApp.newTextStyle().setFontSize(8).setForegroundColor(sasColor).build())
        .build()
    ).setVerticalAlignment("top");
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
  dash.getRange(`K${k4}:O${k4}`).merge().setBackground(cardBg);

  dash.getRange(`Q${k4}:R${k4}`).merge().setBackground(cardBg)
    .setValue("Conversations (24h)")
    .setFontSize(8).setFontColor(navy).setVerticalAlignment("top");
  dash.getRange(`S${k4}:X${k4}`).merge().setBackground(cardBg);

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
  dash.getRange(`I${k5}:O${k5}`).setBackground(phoneAccent);
  dash.getRange(`P${k5}`).setBackground(bg);
  dash.getRange(`Q${k5}:X${k5}`).setBackground(socialAccent);

  // Card borders around KPI panels for visual grouping
  const kpiBorder = { style: SpreadsheetApp.BorderStyle.SOLID, color: divider };
  dash.getRange(`A${k1}:G${k4}`).setBorder(true, true, true, true, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange(`I${k1}:O${k4}`).setBorder(true, true, true, true, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
  dash.getRange(`Q${k1}:X${k4}`).setBorder(true, true, true, true, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);

  // Spacer row after KPIs
  const spacerRow = k5 + 1;
  dash.setRowHeight(spacerRow, 10);
  dash.getRange(`A${spacerRow}:X${spacerRow}`).setBackground(bg);

  // Gap column backgrounds for full height
  const maxBodyRow = 2000; // will update after we know final row
  dash.getRange(`H1:H${maxBodyRow}`).setBackground(bg);
  dash.getRange(`P1:P${maxBodyRow}`).setBackground(bg);

  // ═══════════════════════════════════════════════
  // EMAIL TABLES (Columns A-G)
  // ═══════════════════════════════════════════════

  // ─── Email Queue by Owner ───
  const eqRow = 10 + alertOffset;
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
  dash.getRange(`C${ethRow}`).setValue("Assigned")
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
  if (zendesk.agentCounts["Other"] && zendesk.agentCounts["Other"].assigned > 0) {
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
  const paHeaderRow = 10 + alertOffset;
  dash.getRange(`I${paHeaderRow}:O${paHeaderRow}`).merge()
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
  // Outbound group header — spans K-O with tinted background
  dash.getRange(`K${pahRow1}:O${pahRow1}`).merge().setValue("Outbound")
    .setFontWeight("bold").setFontSize(9).setFontColor(navy).setBackground(outboundBg)
    .setHorizontalAlignment("center");
  dash.getRange(`K${pahRow1}:O${pahRow1}`).setBorder(true, true, false, true, false, false, navy, SpreadsheetApp.BorderStyle.SOLID);

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
  dash.getRange(`O${pahRow2}`).setBackground(outboundBg);
  dash.getRange(`K${pahRow2}:O${pahRow2}`).setBorder(false, true, true, true, false, false, navy, SpreadsheetApp.BorderStyle.SOLID);

  let paRow = pahRow2 + 1;
  CONFIG.agents.forEach(agent => {
    const stats = aircall.agentStats[agent] || { answered: 0, outbound: 0, outboundConnected: 0, outboundShort: 0, outboundLong: 0 };
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
    dash.getRange(`O${paRow}`).setBackground(outboundBg);
    // Borders
    dash.getRange(`K${paRow}:O${paRow}`).setBorder(false, true, false, true, false, false, navy, SpreadsheetApp.BorderStyle.SOLID);
    dash.getRange(`I${paRow}:O${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    paRow++;
  });

  // Notes row (after-hours + SMS note)
  const phoneNotes = [];
  if (aircall.afterHoursCalls > 0) {
    phoneNotes.push(`${aircall.afterHoursCalls} call(s) outside biz hrs excluded`);
  }
  // SMS tracking now available via Aircall webhook
  dash.getRange(`I${paRow}:O${paRow}`).merge()
    .setValue(phoneNotes.join("  ·  "))
    .setFontColor(gray).setFontSize(8).setFontStyle("italic").setBackground(bg);
  paRow++;

  // ─── Missed Calls (detail table) ───
  paRow++; // spacer
  const missedDetails = aircall.missedCallDetails || [];
  dash.getRange(`I${paRow}:O${paRow}`).merge()
    .setValue("Missed Calls")
    .setBackground(navy).setFontColor("#FFFFFF")
    .setFontSize(11).setFontWeight("bold").setVerticalAlignment("middle");
  dash.setRowHeight(paRow, 26);
  paRow++;

  if (missedDetails.length === 0) {
    // Quiet empty state
    dash.setRowHeight(paRow, 22);
    dash.getRange(`I${paRow}:O${paRow}`).merge()
      .setValue("No missed calls today")
      .setFontSize(9).setFontColor(gray).setFontStyle("italic").setBackground(cardBg);
    dash.getRange(`I${paRow}:O${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
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
    dash.getRange(`M${paRow}:O${paRow}`).merge().setValue("Reason")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`I${paRow}:O${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
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
      dash.getRange(`M${paRow}:O${paRow}`).merge()
        .setValue(detail.reason).setFontSize(9).setFontColor(amber).setBackground(cardBg);
      dash.getRange(`I${paRow}:O${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
      paRow++;
    });
  }

  // ─── Phone CSAT Survey — PostCall (last 24h) ───
  paRow++; // spacer
  const pcResponses = (postCall && postCall.responses) || [];
  dash.getRange(`I${paRow}:O${paRow}`).merge()
    .setValue("Phone CSAT Survey")
    .setBackground(navy).setFontColor("#FFFFFF")
    .setFontSize(11).setFontWeight("bold").setVerticalAlignment("middle");
  dash.setRowHeight(paRow, 26);
  paRow++;

  if (pcResponses.length === 0) {
    dash.setRowHeight(paRow, 22);
    dash.getRange(`I${paRow}:O${paRow}`).merge()
      .setValue("No surveys submitted in the last 24 hours")
      .setFontSize(9).setFontColor(gray).setFontStyle("italic").setBackground(cardBg);
    dash.getRange(`I${paRow}:O${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    paRow++;
  } else {
    // Summary line
    const pcScoreStr = postCall.score !== null ? postCall.score + "%" : "—";
    const pcSumColor = postCall.score >= 90 ? green : (postCall.score >= 80 ? amber : risk);
    dash.setRowHeight(paRow, 20);
    dash.getRange(`I${paRow}:J${paRow}`).merge()
      .setValue(pcScoreStr + " CSAT")
      .setFontSize(9).setFontWeight("bold").setFontColor(postCall.score !== null ? pcSumColor : gray).setBackground(bg);
    dash.getRange(`K${paRow}:O${paRow}`).merge()
      .setValue(`${postCall.satisfied} of ${postCall.total} satisfied`)
      .setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`I${paRow}:O${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    paRow++;

    // Column headers
    dash.setRowHeight(paRow, 20);
    dash.getRange(`I${paRow}`).setValue("Score")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`J${paRow}:K${paRow}`).merge().setValue("Customer")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`L${paRow}`).setValue("Agent")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`M${paRow}:O${paRow}`).merge().setValue("Time")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
    dash.getRange(`I${paRow}:O${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
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
      dash.getRange(`M${paRow}:O${paRow}`).merge().setValue(r.dateStr + " " + r.timeStr)
        .setFontSize(9).setBackground(pcRowBg).setHorizontalAlignment("right");
      dash.getRange(`I${paRow}:O${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
      paRow++;
    });
  }

  // ─── SMS Activity Today ───
  const smsData = sms || { totalToday: 0, inbound: 0, outbound: 0, agentStats: {}, messages: [] };
  paRow++; // spacer
  dash.getRange(`I${paRow}:O${paRow}`).merge()
    .setValue("SMS Activity Today")
    .setBackground(navy).setFontColor("#FFFFFF")
    .setFontSize(11).setFontWeight("bold").setVerticalAlignment("middle");
  dash.setRowHeight(paRow, 26);
  paRow++;

  if (smsData.totalToday === 0) {
    dash.setRowHeight(paRow, 22);
    dash.getRange(`I${paRow}:O${paRow}`).merge()
      .setValue("No SMS activity today")
      .setFontSize(9).setFontColor(gray).setFontStyle("italic").setBackground(cardBg);
    dash.getRange(`I${paRow}:O${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    paRow++;
  } else {
    // Summary line: In X · Out X · Total X
    const smsSumColor = smsData.totalToday > 0 ? darkText : gray;
    dash.setRowHeight(paRow, 20);
    dash.getRange(`I${paRow}:O${paRow}`).merge()
      .setRichTextValue(
        SpreadsheetApp.newRichTextValue()
          .setText(`In: ${smsData.inbound}  ·  Out: ${smsData.outbound}  ·  Total: ${smsData.totalToday}`)
          .setTextStyle(SpreadsheetApp.newTextStyle().setFontSize(9).setForegroundColor(darkText).build())
          .build()
      ).setBackground(bg);
    dash.getRange(`I${paRow}:O${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
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
      dash.getRange(`M${paRow}:O${paRow}`).merge().setValue("Received")
        .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
      dash.getRange(`I${paRow}:O${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
      paRow++;

      smsAgentsWithActivity.forEach(agent => {
        const s = smsData.agentStats[agent];
        dash.setRowHeight(paRow, 22);
        dash.getRange(`I${paRow}:J${paRow}`).merge().setValue(agent)
          .setFontSize(9).setBackground(cardBg);
        dash.getRange(`K${paRow}:L${paRow}`).merge().setValue(s.sent)
          .setFontSize(9).setBackground(cardBg).setHorizontalAlignment("right");
        dash.getRange(`M${paRow}:O${paRow}`).merge().setValue(s.received)
          .setFontSize(9).setBackground(cardBg).setHorizontalAlignment("right");
        dash.getRange(`I${paRow}:O${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
        paRow++;
      });
    }

    // Recent messages (last 10)
    const recentSMS = smsData.messages.slice(0, 10);
    if (recentSMS.length > 0) {
      // Section sub-header
      dash.setRowHeight(paRow, 20);
      dash.getRange(`I${paRow}:O${paRow}`).merge().setValue("Recent Messages")
        .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
      dash.getRange(`I${paRow}:O${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
      paRow++;

      recentSMS.forEach(m => {
        const isOut = m.direction === "outbound";
        const dirIcon = isOut ? "→ Out" : "← In";
        const dirColor = isOut ? BRAND.airBlueDark : green;

        // Row 1: direction, agent/contact info, time
        let description = "";
        if (isOut) {
          const agentShort = m.agent ? m.agent.split(" ")[0] : "?";
          const contactStr = m.contact || m.phone || "Unknown";
          description = `${agentShort} → ${contactStr}`;
          if (m.lineName) description += ` via ${m.lineName}`;
        } else {
          const contactStr = m.contact || m.phone || "Unknown";
          description = `${contactStr}`;
          if (m.lineName) description += ` → ${m.lineName}`;
        }

        dash.setRowHeight(paRow, 20);
        dash.getRange(`I${paRow}`).setValue(dirIcon)
          .setFontSize(9).setFontColor(dirColor).setFontWeight("bold").setBackground(cardBg);
        dash.getRange(`J${paRow}:M${paRow}`).merge().setValue(description)
          .setFontSize(9).setBackground(cardBg);
        dash.getRange(`N${paRow}`).setValue(m.timeStr)
          .setFontSize(9).setBackground(cardBg).setHorizontalAlignment("right");
        paRow++;

        // Row 2: message body (truncated to fit, lighter color)
        if (m.body) {
          const truncBody = m.body.length > 120 ? m.body.substring(0, 117) + "..." : m.body;
          dash.setRowHeight(paRow, 18);
          dash.getRange(`I${paRow}:O${paRow}`).merge().setValue(truncBody)
            .setFontSize(8).setFontColor(gray).setFontStyle("italic").setBackground(cardBg)
            .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
        } else {
          dash.setRowHeight(paRow, 4);
          dash.getRange(`I${paRow}:O${paRow}`).merge().setBackground(cardBg);
        }
        dash.getRange(`I${paRow}:O${paRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
        paRow++;
      });
    }
  }


  // ═══════════════════════════════════════════════
  // SOCIAL TABLES (Columns Q-X)
  // ═══════════════════════════════════════════════
  let sRow = 10 + alertOffset;  // Social row counter (same starting row as phone/email)

  // ─── Social — Meta Business Suite (Messenger + Instagram) ───
  sRow++; // spacer to align with other sections
  dash.getRange(`Q${sRow}:X${sRow}`).merge()
    .setValue("Social — Meta Business Suite")
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
  dash.getRange(`Q${sRow}:S${sRow}`).merge()
    .setValue(metaSummary)
    .setFontSize(9).setFontWeight("bold").setFontColor(metaStatusColor).setBackground(bg);
  dash.getRange(`T${sRow}:X${sRow}`).merge()
    .setValue(commentSummary)
    .setFontSize(9).setFontColor(metaComments.length > 0 ? amber : gray).setBackground(bg);
  dash.getRange(`Q${sRow}:X${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
  sRow++;

  // Token expiry warning
  if (meta && meta.tokenWarning) {
    dash.setRowHeight(sRow, 20);
    dash.getRange(`Q${sRow}:X${sRow}`).merge()
      .setValue(meta.tokenWarning)
      .setFontSize(9).setFontWeight("bold").setFontColor(risk).setBackground(riskLt);
    dash.getRange(`Q${sRow}:X${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    sRow++;
  }

  // ── DMs sub-header ──
  dash.setRowHeight(sRow, 20);
  dash.getRange(`Q${sRow}:X${sRow}`).merge()
    .setValue("Direct Messages (Last 24h / Unread)")
    .setFontWeight("bold").setFontSize(9).setFontColor(darkText).setBackground(bg);
  dash.getRange(`Q${sRow}:X${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
  sRow++;

  if (metaConversations.length === 0) {
    dash.setRowHeight(sRow, 22);
    dash.getRange(`Q${sRow}:X${sRow}`).merge()
      .setValue("No recent conversations")
      .setFontSize(9).setFontColor(gray).setFontStyle("italic").setBackground(cardBg);
    dash.getRange(`Q${sRow}:X${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    sRow++;
  } else {
    // Column headers
    dash.setRowHeight(sRow, 20);
    dash.getRange(`Q${sRow}`).setValue("Customer")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`R${sRow}`).setValue("Via")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`S${sRow}:U${sRow}`).merge().setValue("Last Message")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`V${sRow}`).setValue("From")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`W${sRow}:X${sRow}`).merge().setValue("Time")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
    dash.getRange(`Q${sRow}:X${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    sRow++;

    // Show up to 10 conversations
    metaConversations.slice(0, 10).forEach((convo, cIdx) => {
      const rowBg = convo.unread > 0 ? amberLt : (cIdx % 2 === 1 ? altRow : cardBg);

      // Row 1: Customer name (hyperlinked), platform badge, from, time
      dash.setRowHeight(sRow, 20);
      dash.getRange(`Q${sRow}`)
        .setFormula(`=HYPERLINK("${convo.inboxUrl}","${convo.customerName.replace(/"/g, '""')}")`)
        .setFontSize(9).setFontColor("#1155CC").setBackground(rowBg);
      // Platform badge
      const platformShort = convo.platform === "Instagram" ? "IG" : "FB";
      const platformColor = convo.platform === "Instagram" ? "#C13584" : "#1877F2";
      dash.getRange(`R${sRow}`).setValue(platformShort)
        .setFontSize(8).setFontWeight("bold").setFontColor(platformColor).setBackground(rowBg);
      dash.getRange(`S${sRow}:U${sRow}`).merge()
        .setBackground(rowBg);
      dash.getRange(`V${sRow}`).setValue(convo.lastMessageFrom)
        .setFontSize(8).setFontColor(gray).setBackground(rowBg);

      // Format time
      let timeDisplay = "";
      if (convo.time) {
        try {
          const d = new Date(convo.time);
          timeDisplay = Utilities.formatDate(d, Session.getScriptTimeZone(), "MMM d h:mm a");
        } catch (e) { timeDisplay = ""; }
      }
      dash.getRange(`W${sRow}:X${sRow}`).merge().setValue(timeDisplay)
        .setFontSize(8).setFontColor(gray).setBackground(rowBg).setHorizontalAlignment("right");
      sRow++;

      // Row 2: message excerpt
      dash.setRowHeight(sRow, 18);
      const excerpt = convo.lastMessage || "";
      dash.getRange(`Q${sRow}:X${sRow}`).merge().setValue(excerpt)
        .setFontSize(8).setFontColor(gray).setFontStyle("italic").setBackground(rowBg)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      dash.getRange(`Q${sRow}:X${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
      sRow++;
    });
  }

  // ── Comments & Mentions sub-header (last 24h) ──
  sRow++; // spacer
  dash.setRowHeight(sRow, 20);
  dash.getRange(`Q${sRow}:X${sRow}`).merge()
    .setValue("Comments & Mentions (Last 24h)")
    .setFontWeight("bold").setFontSize(9).setFontColor(darkText).setBackground(bg);
  dash.getRange(`Q${sRow}:X${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
  sRow++;

  if (metaComments.length === 0) {
    dash.setRowHeight(sRow, 22);
    dash.getRange(`Q${sRow}:X${sRow}`).merge()
      .setValue("No new comments or mentions")
      .setFontSize(9).setFontColor(gray).setFontStyle("italic").setBackground(cardBg);
    dash.getRange(`Q${sRow}:X${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    sRow++;
  } else {
    // Column headers
    dash.setRowHeight(sRow, 20);
    dash.getRange(`Q${sRow}`).setValue("Author")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`R${sRow}`).setValue("Source")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`S${sRow}:U${sRow}`).merge().setValue("Comment")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`V${sRow}`).setValue("On")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg);
    dash.getRange(`W${sRow}:X${sRow}`).merge().setValue("Time")
      .setFontWeight("bold").setFontSize(9).setFontColor(gray).setBackground(bg).setHorizontalAlignment("right");
    dash.getRange(`Q${sRow}:X${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
    sRow++;

    metaComments.slice(0, 8).forEach((c, cmIdx) => {
      const isMention = c.type === "mention";
      const rowBg = isMention ? amberLt : (cmIdx % 2 === 1 ? altRow : cardBg);

      dash.setRowHeight(sRow, 20);
      // Author (hyperlinked to source)
      const safeAuthor = c.author.replace(/"/g, '""');
      dash.getRange(`Q${sRow}`)
        .setFormula(`=HYPERLINK("${c.url}","${safeAuthor}")`)
        .setFontSize(9).setFontColor("#1155CC").setBackground(rowBg);
      // Platform + type badge
      const platformShort = c.platform === "Instagram" ? "IG" : "FB";
      const platformColor = c.platform === "Instagram" ? "#C13584" : "#1877F2";
      const label = isMention ? platformShort + " tag" : platformShort;
      dash.getRange(`R${sRow}`).setValue(label)
        .setFontSize(8).setFontWeight("bold").setFontColor(platformColor).setBackground(rowBg);
      dash.getRange(`S${sRow}:U${sRow}`).merge()
        .setBackground(rowBg);
      // Post snippet in "On" column
      dash.getRange(`V${sRow}`).setValue(c.postSnippet || "")
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
      dash.getRange(`W${sRow}:X${sRow}`).merge().setValue(timeDisplay)
        .setFontSize(8).setFontColor(gray).setBackground(rowBg).setHorizontalAlignment("right");
      sRow++;

      // Row 2: comment text
      dash.setRowHeight(sRow, 18);
      dash.getRange(`Q${sRow}:X${sRow}`).merge().setValue(c.text || "")
        .setFontSize(8).setFontColor(gray).setFontStyle("italic").setBackground(rowBg)
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      dash.getRange(`Q${sRow}:X${sRow}`).setBorder(false, false, true, false, false, false, divider, SpreadsheetApp.BorderStyle.SOLID);
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
  dash.setColumnWidth(9, 105);  // I — Agent name / Customer
  dash.setColumnWidth(10, 48);  // J — In / Via badge
  dash.setColumnWidth(11, 55);  // K — Dialed
  dash.setColumnWidth(12, 55);  // L — No Ans
  dash.setColumnWidth(13, 50);  // M — <90s
  dash.setColumnWidth(14, 50);  // N — 90s+
  dash.setColumnWidth(15, 50);  // O — extra phone col
  dash.setColumnWidth(16, 20);  // P spacer
  dash.setColumnWidth(17, 110); // Q — Customer/Author
  dash.setColumnWidth(18, 40);  // R — Via/Source badge
  dash.setColumnWidth(19, 70);  // S — Message col 1
  dash.setColumnWidth(20, 70);  // T — Message col 2
  dash.setColumnWidth(21, 60);  // U — Message col 3
  dash.setColumnWidth(22, 60);  // V — From/On
  dash.setColumnWidth(23, 50);  // W — Time col 1
  dash.setColumnWidth(24, 50);  // X — Time col 2

  // Fill remaining
  const lastRow = Math.max(tRow, paRow, sRow) + 2;
  dash.getRange(`A${lastRow}:X${lastRow + 5}`).setBackground(bg);

  // Footer — version & goals
  dash.getRange(`A${lastRow}:X${lastRow}`).merge()
    .setValue(`CS Command Center v1.5.0  ·  Refreshes every 5 min  ·  Goal: reply within ${slaHours} business hours · answer ${CONFIG.thresholds.phoneAnswerRate.green}%+ inbound calls · Mon–Fri 6a–5p PST`)
    .setFontColor(gray).setFontSize(8).setFontStyle("italic")
    .setHorizontalAlignment("center").setBackground(bg);

  // Footer — legend & logic explanation
  const legendRow = lastRow + 1;
  const legendLines = [
    `Email status: Healthy = 0 tickets past SLA, Watch = 1–5 past SLA, At Risk = 6+ past SLA  ·  `
    + `Phone status: Healthy = answer rate ≥ ${CONFIG.thresholds.phoneAnswerRate.green}%, Watch = ${CONFIG.thresholds.phoneAnswerRate.yellow}–${CONFIG.thresholds.phoneAnswerRate.green - 1}%, At Risk = < ${CONFIG.thresholds.phoneAnswerRate.yellow}%  ·  `
    + `Social status: Healthy = oldest DM wait ≤ 2h, Watch = 2–6h, At Risk = > 6h`,
    `Wait times are business hours only (Mon–Fri 6a–5p PST)  ·  Past SLA = waiting > ${CONFIG.thresholds.oldestUnanswered.green} biz hrs without a reply  ·  `
    + `Oldest Waiting table: top 3 past SLA highlighted red, others past SLA amber, within SLA green`,
    `CSAT % = (satisfied ÷ total) × 100  ·  Satisfied = 4+ out of 5  ·  Phone answer rate = answered inbound ÷ total inbound (biz hrs only)  ·  `
    + `Open tickets exclude ${(CONFIG.excludeAgents || []).join(", ")} (not on CS team)`,
    `Social: FB = Facebook Messenger, IG = Instagram DM  ·  Comments & Mentions show last 24h from FB + IG posts  ·  `
    + `Amber highlight = unread DM or @mention  ·  Meta token expiry warning appears at 7 days`,
    `API limitations: IG DMs not available (requires OAuth flow not supported by Graph API Explorer)  ·  `
    + `Missed call customer info shows "Check Aircall #" when contact data is not exposed in the API payload`,
  ];
  for (let li = 0; li < legendLines.length; li++) {
    dash.getRange(`A${legendRow + li}:X${legendRow + li}`).merge()
      .setValue(legendLines[li])
      .setFontColor(gray).setFontSize(7).setFontStyle("italic")
      .setHorizontalAlignment("center").setBackground(bg).setWrap(false);
  }
}

// --- SOCIAL OLDEST WAIT HELPER ---
function computeSocialOldestWait(meta) {
  if (!meta || !meta.recentConversations) return 0;
  const now = new Date();
  let oldestMin = 0;
  meta.recentConversations.forEach(c => {
    if (c.unread > 0 && c.time) {
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
    const event = payload.event || "";

    // Route based on event type
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
  const contactPhone = d.external_number || "";
  const contactEmail = (d.contact && d.contact.emails
    && d.contact.emails[0] && d.contact.emails[0].value) || "";

  // Agent / user info (present on message.sent, absent on message.received)
  const agentName = (d.user && d.user.name) || "";
  const agentEmail = (d.user && d.user.email) || "";

  // Aircall line info: data.number.name / digits
  const lineName = (d.number && d.number.name) || "";
  const lineNumber = (d.number && d.number.digits) || "";

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

// Also handle GET (PostCall may ping the URL to verify)
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: "ok", service: "CS Command Center PostCall Webhook" }))
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

  // ─── Step 2: Instagram DMs (same endpoint, platform filter) ───
  const igDMData = graphGet(
    `${pageId}/conversations?platform=instagram&fields=id,updated_time,unread_count,participants,messages.limit(1){message,from,created_time}&limit=15`
  );
  const igConvos = parseConversations(igDMData, "Instagram");

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
  try {
    const debugData = graphGet(`debug_token?input_token=${pageToken}`);
    if (debugData && debugData.data && debugData.data.expires_at) {
      const expiresAt = new Date(debugData.data.expires_at * 1000);
      const daysLeft = Math.floor((expiresAt - new Date()) / (1000 * 60 * 60 * 24));
      if (daysLeft <= 7) {
        tokenWarning = `Meta token expires in ${daysLeft} day${daysLeft !== 1 ? "s" : ""}!`;
      }
    }
  } catch (e) {
    Logger.log("Token debug check failed: " + e.toString());
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

// --- HELPER FUNCTIONS ---

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

// Format business minutes as business hours (e.g. "2.5h", "125h")
function formatBizMinutes(min) {
  if (min === 0) return "-";
  const hours = min / 60;
  if (hours < 1) return Math.round(min) + "m";
  if (hours < 10) return hours.toFixed(1) + "h";
  return Math.round(hours) + "h";
}

