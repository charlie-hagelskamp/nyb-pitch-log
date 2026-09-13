const ADMIN_EMAIL = "charlie.hagelskamp@gmail.com";

/*
Games tab headers:
GameID | Date | Team | Opponent | Notes | SubmissionTime

Pitching tab headers:
GameID | Date | Team | Pitcher | Pitches | RestDays | EligibleDate
*/

const GC_COMPARE_START_DATE = "2026-04-18";
const DEFAULT_SEASON_YEAR = 2026;
const FALL_SEASON_START_DATE = "2026-08-01";
const MISSING_OVERRIDE_TYPE = "Ignore Missing";
const FALL_PRACTICE_SOURCE_ID = "1KpbgC-0iugmxSMpjsXFx1nPUyPRrW32QHp6Bd_G41eU";
const FALL_PRACTICE_SOURCE_TAB = "Master";
const FALL_PRACTICE_TEAMS = ["7B", "8B", "8G", "10B", "10G", "11B", "11G", "12G"];
const APP_BASE_URL = "https://charlie-hagelskamp.github.io/nyb-pitch-log";
const APP_LOGO_URL = APP_BASE_URL + "/nyb-logo.png";
const WEEKLY_REPORT_TIMEZONE = "America/New_York";
const WEEKLY_RECIPIENTS_SHEET = "Weekly_Email_Recipients";
const WEEKLY_REPORT_LOG_SHEET = "Weekly_Report_Log";
const PITCH_LOG_SPREADSHEET_ID = "1I52gbuegk6fYZOVo0lPjoJHCiVIqdBXUggz3oq3Twxc";

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const params = (e && e.parameter) || {};
  const callback = params.callback || "dashboardCallback";

  // =========================
  // Anonymous site traffic summary
  // ?siteAnalyticsSummary=1&callback=...
  // =========================
  if (params.siteAnalyticsSummary) {
    try {
      return outputJsonp(callback, buildSiteAnalyticsSummary_());
    } catch (err) {
      return outputJsonp(callback, {
        error: true,
        message: err.message || String(err)
      });
    }
  }

  // =========================
  // Performance summary endpoint
  // ?performanceSummary=1&includeFall=1&callback=...
  // =========================
  if (params.performanceSummary) {
    try {
      const includeFall = isTruthy_(params.includeFall);
      return outputJsonp(callback, buildPerformanceSummary_(includeFall));
    } catch (err) {
      return outputJsonp(callback, {
        error: true,
        message: err.message || String(err)
      });
    }
  }

  // =========================
  // Read-only weekly email preview
  // ?weeklyReportPreview=1&callback=...
  // =========================
  if (params.weeklyReportPreview) {
    try {
      const referenceDate = params.weekStart
        ? parseLocalDate(String(params.weekStart))
        : new Date();
      const report = buildWeeklyReportData_(referenceDate);
      return outputJsonp(callback, {
        report: report,
        recipientCount: getActiveWeeklyRecipients_().length,
        automationInstalled: isSmartSyncTriggerInstalled_(),
        html: buildWeeklyReportEmailHtml_(report, APP_LOGO_URL)
      });
    } catch (err) {
      return outputJsonp(callback, {
        error: true,
        message: err.message || String(err)
      });
    }
  }

  // =========================
  // Current week schedule endpoint
  // ?currentWeekSchedule=1&callback=...
  // =========================
  if (params.currentWeekSchedule) {
    try {
      return outputJsonp(callback, buildCurrentWeekSchedule_());
    } catch (err) {
      return outputJsonp(callback, {
        error: true,
        message: err.message || String(err)
      });
    }
  }

  // =========================
  // Read-only fall practice endpoint
  // ?fallPractices=1&callback=...
  // =========================
  if (params.fallPractices) {
    try {
      return outputJsonp(callback, buildFallPracticeSchedule_());
    } catch (err) {
      return outputJsonp(callback, {
        error: true,
        message: err.message || String(err)
      });
    }
  }

  // =========================
  // Coach drills library
  // ?drills=1&callback=...
  // =========================
  if (params.drills) {
    try {
      return outputJsonp(callback, buildDrills_());
    } catch (err) {
      return outputJsonp(callback, {
        error: true,
        message: err.message || String(err)
      });
    }
  }

  // =========================
  // Coach-facing GameChanger game cards
  // ?gcGameCards=1&teamName=11U%20Black&date=2026-06-01&callback=...
  // Reads only the local GC_Games_Sync cache; it never calls GameChanger while
  // a coach is waiting for the picker to load.
  // =========================
  if (params.gcGameCards) {
    try {
      return outputJsonp(callback, buildGameCards_(
        String(params.teamName || "").trim(),
        String(params.date || "").trim()
      ));
    } catch (err) {
      return outputJsonp(callback, {
        games: [],
        meta: { error: true, message: err.message || String(err) }
      });
    }
  }

  // =========================
  // Verified submission result polling endpoint
  // ?submissionResult=1&submissionId=...&callback=...
  // =========================
  if (params.submissionResult) {
    try {
      const result = getSubmissionResult_(String(params.submissionId || "").trim());
      return outputJsonp(callback, result ? { found: true, result: result } : { found: false });
    } catch (err) {
      return outputJsonp(callback, { found: false, error: true, message: err.message || String(err) });
    }
  }

  // =========================
  // Submission summary endpoint
  // ?submissionSummary=1&callback=...
  // =========================
  if (params.submissionSummary) {
    try {
      return outputJsonp(callback, buildSubmissionSummary_());
    } catch (err) {
      return outputJsonp(callback, {
        error: true,
        message: err.message || String(err)
      });
    }
  }

  // =========================
  // Admin missing submissions endpoint
  // ?adminMissing=1&callback=...
  // =========================
  if (params.adminMissing) {
    try {
      return outputJsonp(callback, buildAdminMissingSubmissions_());
    } catch (err) {
      return outputJsonp(callback, []);
    }
  }

  // =========================
  // Add missing override endpoint
  // ?addMissingOverride=1&team=...&date=...&reason=...&callback=...
  // =========================
  if (params.addMissingOverride) {
    try {
      const team = String(params.team || "").trim();
      const date = String(params.date || "").trim();
      const reason = String(params.reason || "").trim();

      if (!team || !date) {
        throw new Error("Missing team or date");
      }

      if (!reason) {
        throw new Error("Reason is required");
      }

      addMissingOverride_(team, date, reason);

      return outputJsonp(callback, {
        success: true,
        message: "Missing submission ignored",
        team: team,
        date: date
      });
    } catch (err) {
      return outputJsonp(callback, {
        success: false,
        message: err.message || String(err)
      });
    }
  }

  // =========================
  // Game Notes Logs endpoint
  // ?notesTeam=8U%20Black&callback=...
  // =========================
  if (params.notesTeam) {
    try {
      return outputJsonp(callback, buildTeamGameNotesReview_(
        String(params.notesTeam || "").trim(),
        String(params.startDate || "").trim(),
        String(params.endDate || formatDateOnly(new Date())).trim()
      ));
    } catch (err) {
      return outputJsonp(callback, {
        error: true,
        message: err.message || String(err)
      });
    }
  }

  // =========================
  // GC Submission Check endpoint
  // ?submissionCheck=1&callback=...
  // Optional filters:
  // ?submissionCheck=1&status=Missing%20Submission
  // ?submissionCheck=1&teamName=11U%20Red
  // =========================
  if (params.submissionCheck) {
    const checkSheet = ss.getSheetByName("GC_Submission_Check");

    if (!checkSheet) {
      return outputJsonp(callback, []);
    }

    const rows = checkSheet.getDataRange().getValues();

    if (rows.length <= 1) {
      return outputJsonp(callback, []);
    }

    const headers = rows.shift();

    const teamCol = headers.indexOf("Team");
    const dateCol = headers.indexOf("Date");
    const gcCol = headers.indexOf("GC_Completed_Games");
    const ntbCol = headers.indexOf("NTB_Submissions");
    const missingCol = headers.indexOf("Missing_Count");
    const extraCol = headers.indexOf("Extra_Count");
    const statusCol = headers.indexOf("Check_Status");
    const detailsCol = headers.indexOf("GC_Game_Details");

    let results = rows.map(r => ({
      team: r[teamCol] || "",
      date: normalizeDateStringGC_(r[dateCol]),
      gcCompletedGames: Number(r[gcCol]) || 0,
      ntbSubmissions: Number(r[ntbCol]) || 0,
      missingCount: Number(r[missingCol]) || 0,
      extraCount: Number(r[extraCol]) || 0,
      checkStatus: r[statusCol] || "",
      gcGameDetails: r[detailsCol] || ""
    }));

    if (params.status) {
      results = results.filter(r => r.checkStatus === params.status);
    }

    if (params.teamName) {
      results = results.filter(r => r.team === params.teamName);
    }

    results.sort((a, b) => String(b.date + b.team).localeCompare(String(a.date + a.team)));

    return outputJsonp(callback, results);
  }

  // =========================
  // GameChanger games sync view endpoint
  // ?gcGames=1&callback=...
  // Optional:
  // ?gcGames=1&teamName=11U%20Red
  // =========================
  if (params.gcGames) {
    const syncSheet = ss.getSheetByName("GC_Games_Sync");

    if (!syncSheet) {
      return outputJsonp(callback, []);
    }

    const rows = syncSheet.getDataRange().getValues();

    if (rows.length <= 1) {
      return outputJsonp(callback, []);
    }

    const headers = rows.shift();

    const teamCol = headers.indexOf("Team");
    const gcTeamIdCol = headers.indexOf("GC_Team_ID");
    const gcGameIdCol = headers.indexOf("GC_Game_ID");
    const dateCol = headers.indexOf("Date");
    const timeCol = headers.indexOf("Start_Time");
    const oppCol = headers.indexOf("Opponent");
    const statusCol = headers.indexOf("Game_Status");
    const teamScoreCol = headers.indexOf("Team_Score");
    const oppScoreCol = headers.indexOf("Opponent_Score");
    const resultCol = headers.indexOf("Result");
    const syncTimeCol = headers.indexOf("Sync_Time");

    let results = rows.map(r => ({
      team: r[teamCol] || "",
      gcTeamId: r[gcTeamIdCol] || "",
      gcGameId: r[gcGameIdCol] || "",
      date: normalizeDateStringGC_(r[dateCol]),
      startTime: formatTimeForDisplayGC_(r[timeCol]),
      opponent: r[oppCol] || "",
      gameStatus: r[statusCol] || "",
      teamScore: r[teamScoreCol] === "" ? "" : r[teamScoreCol],
      opponentScore: r[oppScoreCol] === "" ? "" : r[oppScoreCol],
      result: r[resultCol] || "",
      syncTime: r[syncTimeCol] ? formatDateTime(new Date(r[syncTimeCol])) : ""
    }));

    if (params.teamName) {
      results = results.filter(r => r.team === params.teamName);
    }

    results.sort((a, b) =>
      String(b.date + " " + b.startTime + b.team).localeCompare(
        String(a.date + " " + a.startTime + a.team)
      )
    );

    return outputJsonp(callback, results);
  }

  // =========================
  // Manual sync endpoint
  // ?runGcSync=1&callback=...
  // =========================
  if (params.runGcSync) {
    try {
      syncGameChangerAndSubmissionCheck();

      return outputJsonp(callback, {
        success: true,
        message: "GameChanger sync and submission check completed",
        ranAt: formatDateTime(new Date())
      });
    } catch (err) {
      return outputJsonp(callback, {
        success: false,
        message: err.message || String(err)
      });
    }
  }

  // =========================
  // Pitching dashboard endpoint
  // ?team=8U%20Black&callback=...
  // =========================
  const pitchingSheet = ss.getSheetByName("Pitching");

  if (!pitchingSheet) {
    return outputJsonp(callback, []);
  }

  const team = params.team;
  if (!team) {
    return outputJsonp(callback, []);
  }

  const rows = pitchingSheet.getDataRange().getValues();
  if (rows.length <= 1) {
    return outputJsonp(callback, []);
  }

  const headers = rows.shift();

  const teamCol = headers.indexOf("Team");
  const pitcherCol = headers.indexOf("Pitcher");
  const eligibleCol = headers.indexOf("EligibleDate");
  const dateCol = headers.indexOf("Date");
  const pitchesCol = headers.indexOf("Pitches");

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const sevenDaysAgo = new Date(today);
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 6);

  const statsByPitcher = {};

  rows.forEach(r => {
    if (r[teamCol] !== team) return;

    const pitcher = r[pitcherCol];
    const gameDate = parseSheetDate(r[dateCol]);
    const eligibleDate = parseSheetDate(r[eligibleCol]);
    const pitches = Number(r[pitchesCol]) || 0;

    if (!pitcher || !gameDate) return;

    if (!statsByPitcher[pitcher]) {
      statsByPitcher[pitcher] = {
        name: pitcher,
        latestGameDate: null,
        latestEligibleDate: null,
        totalPitches: 0,
        pitchesLast7: 0
      };
    }

    statsByPitcher[pitcher].totalPitches += pitches;

    if (gameDate >= sevenDaysAgo && gameDate <= today) {
      statsByPitcher[pitcher].pitchesLast7 += pitches;
    }

    if (
      !statsByPitcher[pitcher].latestGameDate ||
      gameDate > statsByPitcher[pitcher].latestGameDate
    ) {
      statsByPitcher[pitcher].latestGameDate = gameDate;
      statsByPitcher[pitcher].latestEligibleDate = eligibleDate;
    }
  });

  const result = Object.keys(statsByPitcher).map(name => {
    const p = statsByPitcher[name];
    return {
      name: p.name,
      latestGameDate: p.latestGameDate ? formatDateOnly(p.latestGameDate) : "",
      eligibleDate: p.latestEligibleDate ? formatDateOnly(p.latestEligibleDate) : "",
      totalPitches: p.totalPitches,
      pitchesLast7: p.pitchesLast7
    };
  });

  return outputJsonp(callback, result);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  if (data.eventType === "site_visit") {
    recordSiteVisit_(data);
    return jsonOutput({ result: "success" });
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const submissionId = String(data.submissionId || Utilities.getUuid()).trim();
    const existing = getSubmissionResult_(submissionId);
    if (existing) return jsonOutput(existing);

    if (!data.date || !data.team || !data.opponent) {
      throw new Error("Date, team, and opponent are required");
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const gamesSheet = ss.getSheetByName("Games");
    const pitchingSheet = ss.getSheetByName("Pitching");
    if (!gamesSheet || !pitchingSheet) throw new Error("Missing Games or Pitching sheet");

    const gameDate = parseLocalDate(data.date);
    const submissionTime = new Date();
    const gameID = new Date().getTime();
    const gameHeaders = ensureSheetHeaders_(gamesSheet, [
      "GameID", "Date", "Team", "Opponent", "Notes", "SubmissionTime",
      "SubmissionID", "GC_Game_ID", "GameStartTime", "GameSequence",
      "ScheduleMatched", "ScheduleLastSynced", "SubmittedBy"
    ]);
    appendObjectRows_(gamesSheet, gameHeaders, [{
      GameID: gameID,
      Date: formatDateOnly(gameDate),
      Team: data.team || "",
      Opponent: data.opponent || "",
      Notes: data.notes || "",
      SubmissionTime: submissionTime,
      SubmissionID: submissionId,
      GC_Game_ID: data.gcGameId || "",
      GameStartTime: data.gameStartTime || "",
      GameSequence: data.gameSequence || "",
      ScheduleMatched: data.scheduleMatched === true,
      ScheduleLastSynced: data.scheduleLastSynced || "",
      SubmittedBy: data.submittedBy || "Coach"
    }]);

    const pitchingHeaders = ensureSheetHeaders_(pitchingSheet, [
      "GameID", "Date", "Team", "Pitcher", "Pitches", "RestDays", "EligibleDate",
      "DailyTotal", "AlertLevel", "GameStartTime", "GameSequence", "SubmissionID"
    ]);
    const pitchingValues = pitchingSheet.getDataRange().getValues();
    pitchingValues.shift();
    const allOutings = pitchingValues.map((row, index) => rowToPitchOuting_(row, pitchingHeaders, index + 2)).filter(Boolean);
    const submittedPitchers = Array.isArray(data.pitchers) ? data.pitchers : [];
    const submittedSummary = [];
    const rowsToAppend = [];

    submittedPitchers.forEach(p => {
      const pitches = Number(p.pitches);
      const pitcher = String(p.name || "").trim();
      if (!pitcher || isNaN(pitches) || pitches < 0) return;

      const current = {
        gameID: gameID,
        team: data.team,
        pitcher: pitcher,
        gameDate: new Date(gameDate),
        pitches: pitches,
        gameStartTime: data.gameStartTime || "",
        gameSequence: data.gameSequence || "",
        submittedAt: submissionTime
      };
      const evaluation = evaluatePitchSubmission_(allOutings, current);

      rowsToAppend.push({
        GameID: gameID,
        Date: formatDateOnly(gameDate),
        Team: data.team,
        Pitcher: pitcher,
        Pitches: pitches,
        RestDays: evaluation.restDays,
        EligibleDate: evaluation.eligibleDate,
        DailyTotal: evaluation.dailyTotal,
        AlertLevel: evaluation.level,
        GameStartTime: data.gameStartTime || "",
        GameSequence: data.gameSequence || "",
        SubmissionID: submissionId
      });

      updateExistingSameDayRest_(pitchingSheet, pitchingHeaders, allOutings, current, evaluation);
      allOutings.push(current);
      submittedSummary.push({
        name: pitcher,
        pitches: pitches,
        dailyTotal: evaluation.dailyTotal,
        restDays: evaluation.restDays,
        eligibleDate: evaluation.eligibleDate,
        level: evaluation.level,
        alerts: evaluation.alerts
      });
    });

    appendObjectRows_(pitchingSheet, pitchingHeaders, rowsToAppend);

    const level = overallSubmissionLevel_(submittedSummary);
    const result = {
      result: "success",
      submissionId: submissionId,
      gameId: gameID,
      date: formatDateOnly(gameDate),
      team: data.team || "",
      opponent: data.opponent || "",
      gcGameId: data.gcGameId || "",
      level: level,
      pitchers: submittedSummary,
      confirmedAt: formatDateTime(submissionTime)
    };
    recordSubmissionResult_(result);

    try {
      sendSubmissionEmail(data, gameID, gameDate, submissionTime, submittedSummary);
      const alerts = flattenSubmissionAlerts_(data.team, gameDate, submittedSummary);
      if (alerts.length) sendRuleAlertEmail_(alerts);
    } catch (mailError) {
      console.error("Submission saved, but email failed: " + mailError);
    }

    return jsonOutput(result);
  } finally {
    lock.releaseLock();
  }
}

function sendSubmissionEmail(data, gameID, gameDate, submissionTime, submittedPitchers) {
  let body =
    `A game submission was recorded.\n\n` +
    `GameID: ${gameID}\n` +
    `Date: ${formatDateOnly(gameDate)}\n` +
    `Team: ${data.team || "N/A"}\n` +
    `Opponent: ${data.opponent || "N/A"}\n` +
    `GameChanger Game ID: ${data.gcGameId || "Manual entry"}\n` +
    `Game Time: ${data.gameStartTime || "Not provided"}\n` +
    `Submission Time: ${formatDateTime(submissionTime)}\n\n` +
    `Notes:\n${data.notes || "N/A"}\n\nPitchers Submitted:\n`;

  if (!submittedPitchers.length) body += "None\n";
  submittedPitchers.forEach(p => {
    body +=
      `• ${p.name}\n` +
      `  This Game: ${p.pitches}\n` +
      `  Total Today: ${p.dailyTotal}\n` +
      `  Rest Days: ${p.restDays}\n` +
      `  Eligible Again: ${p.eligibleDate}\n` +
      `  Result: ${String(p.level).toUpperCase()}\n`;
  });

  MailApp.sendEmail(ADMIN_EMAIL, "NYB Game Submission", body);
}

function flattenSubmissionAlerts_(team, gameDate, pitchers) {
  const alerts = [];
  pitchers.forEach(p => (p.alerts || []).forEach(alert => alerts.push({
    team: team,
    date: formatDateOnly(gameDate),
    pitcher: p.name,
    dailyTotal: p.dailyTotal,
    severity: alert.severity,
    code: alert.code,
    message: alert.message
  })));
  return alerts;
}

function sendRuleAlertEmail_(alerts) {
  const hasViolation = alerts.some(a => a.severity === "violation");
  const subject = hasViolation ? "NYB PitchSmart Violation" : "NYB PitchSmart Warning";
  let body = hasViolation ? "PitchSmart violation(s) detected.\n\n" : "PitchSmart warning detected.\n\n";
  alerts.forEach((alert, index) => {
    body +=
      `${index + 1}. ${String(alert.severity).toUpperCase()} — ${alert.pitcher}\n` +
      `Team: ${alert.team}\nDate: ${alert.date}\nDaily Total: ${alert.dailyTotal}\n` +
      `${alert.message}\n\n`;
  });
  MailApp.sendEmail(ADMIN_EMAIL, subject, body);
}

function rowToPitchOuting_(row, headers, rowNumber) {
  const get = name => {
    const index = headers.indexOf(name);
    return index >= 0 ? row[index] : "";
  };
  const team = get("Team");
  const pitcher = get("Pitcher");
  const gameDate = parseSheetDate(get("Date"));
  if (!team || !pitcher || !gameDate) return null;
  return {
    rowNumber: rowNumber,
    gameID: get("GameID"),
    team: team,
    pitcher: pitcher,
    gameDate: gameDate,
    pitches: Number(get("Pitches")) || 0,
    gameStartTime: get("GameStartTime") || "",
    gameSequence: get("GameSequence") || "",
    submittedAt: null
  };
}

function evaluatePitchSubmission_(history, current) {
  const playerHistory = history.filter(o => o.team === current.team && o.pitcher === current.pitcher);
  const sameDay = playerHistory.filter(o => datesEqual(o.gameDate, current.gameDate)).concat([current]);
  sameDay.sort(compareGameOrder_);
  const dailyTotal = sameDay.reduce((sum, o) => sum + (Number(o.pitches) || 0), 0);
  const restDays = calcRestDays(dailyTotal);
  const eligibleDateObj = addDaysToDate_(current.gameDate, restDays + 1);
  const eligibleDate = formatDateOnly(eligibleDateObj);
  const alerts = [];

  if (sameDay.length > 1) {
    const firstGamePitches = Number(sameDay[0].pitches) || 0;
    const warningOnly = sameDay.length === 2 && firstGamePitches <= 20;
    alerts.push({
      code: "same_day",
      severity: warningOnly ? "warning" : "violation",
      message: warningOnly
        ? `Same-day pitching warning: first game was ${firstGamePitches} pitches (20 or fewer).`
        : `Same-day pitching violation: ${sameDay.length} appearances; first game was ${firstGamePitches} pitches.`
    });
  }

  const dailyTotals = buildDailyTotals_(playerHistory);
  const currentDateText = formatDateOnly(current.gameDate);
  const earlierDates = Object.keys(dailyTotals).filter(date => date < currentDateText).sort();
  if (earlierDates.length) {
    const priorDate = earlierDates[earlierDates.length - 1];
    const priorTotal = dailyTotals[priorDate];
    const priorEligible = addDaysToDate_(parseLocalDate(priorDate), calcRestDays(priorTotal) + 1);
    if (current.gameDate < priorEligible) {
      alerts.push({
        code: "rest",
        severity: "violation",
        message: `Required rest was not complete. ${priorTotal} pitches on ${priorDate} made the pitcher eligible on ${formatDateOnly(priorEligible)}.`
      });
    }
  }

  const futureDates = Object.keys(dailyTotals).filter(date => date > currentDateText).sort();
  if (futureDates.length && parseLocalDate(futureDates[0]) < eligibleDateObj) {
    alerts.push({
      code: "rest",
      severity: "violation",
      message: `This backdated entry requires rest through ${eligibleDate}; an outing already exists on ${futureDates[0]}.`
    });
  }

  const maximum = dailyMaximumForTeam_(current.team);
  if (maximum && dailyTotal > maximum) {
    alerts.push({
      code: "daily_max",
      severity: "violation",
      message: `Daily maximum exceeded: ${dailyTotal} pitches (limit ${maximum}).`
    });
  }

  const pitchingDates = {};
  Object.keys(dailyTotals).forEach(date => pitchingDates[date] = true);
  pitchingDates[currentDateText] = true;
  const yesterday = formatDateOnly(addDaysToDate_(current.gameDate, -1));
  const twoDaysAgo = formatDateOnly(addDaysToDate_(current.gameDate, -2));
  const tomorrow = formatDateOnly(addDaysToDate_(current.gameDate, 1));
  const twoDaysAhead = formatDateOnly(addDaysToDate_(current.gameDate, 2));
  if ((pitchingDates[twoDaysAgo] && pitchingDates[yesterday]) ||
      (pitchingDates[yesterday] && pitchingDates[tomorrow]) ||
      (pitchingDates[tomorrow] && pitchingDates[twoDaysAhead])) {
    alerts.push({ code: "three_days", severity: "violation", message: "Pitcher has appeared on three consecutive days." });
  }

  const level = alerts.some(a => a.severity === "violation")
    ? "violation" : (alerts.some(a => a.severity === "warning") ? "warning" : "clear");
  return { dailyTotal: dailyTotal, restDays: restDays, eligibleDate: eligibleDate, level: level, alerts: alerts };
}

function compareGameOrder_(a, b) {
  const aSequence = Number(a.gameSequence);
  const bSequence = Number(b.gameSequence);
  if (aSequence && bSequence && aSequence !== bSequence) return aSequence - bSequence;
  const aMinutes = parseTimeMinutes_(a.gameStartTime);
  const bMinutes = parseTimeMinutes_(b.gameStartTime);
  if (aMinutes < 24 * 60 && bMinutes < 24 * 60 && aMinutes !== bMinutes) return aMinutes - bMinutes;
  if (a.rowNumber && !b.rowNumber) return -1;
  if (b.rowNumber && !a.rowNumber) return 1;
  const aSubmitted = a.submittedAt ? new Date(a.submittedAt).getTime() : 0;
  const bSubmitted = b.submittedAt ? new Date(b.submittedAt).getTime() : 0;
  return aSubmitted - bSubmitted;
}

function parseTimeMinutes_(value) {
  const match = String(value || "").trim().match(/^(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);
  if (!match) return 24 * 60;
  let hour = Number(match[1]);
  const minute = Number(match[2]);
  const period = String(match[3] || "").toUpperCase();
  if (period === "AM" && hour === 12) hour = 0;
  if (period === "PM" && hour !== 12) hour += 12;
  return hour * 60 + minute;
}

function buildDailyTotals_(outings) {
  const totals = {};
  outings.forEach(o => {
    const date = formatDateOnly(o.gameDate);
    totals[date] = (totals[date] || 0) + (Number(o.pitches) || 0);
  });
  return totals;
}

function dailyMaximumForTeam_(team) {
  const match = String(team || "").match(/(\d{1,2})\s*U/i);
  const age = match ? Number(match[1]) : null;
  if (age === 9 || age === 10) return 75;
  if (age === 11 || age === 12) return 85;
  return null;
}

function addDaysToDate_(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  result.setHours(0, 0, 0, 0);
  return result;
}

function calcRestDays(pitches) {
  if (pitches <= 20) return 0;
  if (pitches <= 35) return 1;
  if (pitches <= 50) return 2;
  if (pitches <= 65) return 3;
  return 4;
}

function overallSubmissionLevel_(pitchers) {
  if (pitchers.some(p => p.level === "violation")) return "violation";
  if (pitchers.some(p => p.level === "warning")) return "warning";
  return "clear";
}

function updateExistingSameDayRest_(sheet, headers, history, current, evaluation) {
  const restCol = headers.indexOf("RestDays") + 1;
  const eligibleCol = headers.indexOf("EligibleDate") + 1;
  const dailyTotalCol = headers.indexOf("DailyTotal") + 1;
  history.filter(o => o.rowNumber && o.team === current.team && o.pitcher === current.pitcher && datesEqual(o.gameDate, current.gameDate))
    .forEach(o => {
      sheet.getRange(o.rowNumber, restCol).setValue(evaluation.restDays);
      sheet.getRange(o.rowNumber, eligibleCol).setValue(evaluation.eligibleDate);
      sheet.getRange(o.rowNumber, dailyTotalCol).setValue(evaluation.dailyTotal);
    });
}

function ensureSheetHeaders_(sheet, requiredHeaders) {
  let headers = sheet.getLastRow() ? sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0] : [];
  if (!headers.length || headers.every(value => !value)) headers = [];
  requiredHeaders.forEach(header => { if (headers.indexOf(header) < 0) headers.push(header); });
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  return headers;
}

function appendObjectRows_(sheet, headers, objects) {
  if (!objects.length) return;
  const rows = objects.map(object => headers.map(header => object[header] === undefined ? "" : object[header]));
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

function ensureSubmissionResultsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Submission_Results");
  if (!sheet) sheet = ss.insertSheet("Submission_Results");
  ensureSheetHeaders_(sheet, ["SubmissionID", "Result_JSON", "Created_At"]);
  return sheet;
}

function recordSubmissionResult_(result) {
  const sheet = ensureSubmissionResultsSheet_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  appendObjectRows_(sheet, headers, [{
    SubmissionID: result.submissionId,
    Result_JSON: JSON.stringify(result),
    Created_At: new Date()
  }]);
  CacheService.getScriptCache().put("submission:" + result.submissionId, JSON.stringify(result), 21600);
}

function getSubmissionResult_(submissionId) {
  if (!submissionId) return null;
  const cached = CacheService.getScriptCache().get("submission:" + submissionId);
  if (cached) return JSON.parse(cached);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Submission_Results");
  if (!sheet || sheet.getLastRow() < 2) return null;
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  const idCol = headers.indexOf("SubmissionID");
  const resultCol = headers.indexOf("Result_JSON");
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][idCol]) === submissionId) return JSON.parse(values[i][resultCol]);
  }
  return null;
}

function outputJsonp(callback, data) {
  const safeCallback = /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(String(callback || ""))
    ? String(callback)
    : "dashboardCallback";
  return ContentService
    .createTextOutput(`${safeCallback}(${JSON.stringify(data)})`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function recordSiteVisit_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Site_Visits");
    if (!sheet) throw new Error("Missing Site_Visits sheet");

    sheet.getRange(sheet.getLastRow() + 1, 1, 1, 12).setValues([[
      new Date(),
      safeAnalyticsText_(data.page, 80),
      safeAnalyticsText_(data.path, 200),
      safeAnalyticsText_(data.referrer, 200),
      safeAnalyticsText_(data.device, 30),
      safeAnalyticsText_(data.browser, 40),
      safeAnalyticsText_(data.operatingSystem, 40),
      safeAnalyticsText_(data.screen, 30),
      safeAnalyticsText_(data.viewport, 30),
      safeAnalyticsText_(data.language, 20),
      safeAnalyticsText_(data.timezone, 80),
      safeAnalyticsText_(data.sessionId, 80)
    ]]);
  } finally {
    lock.releaseLock();
  }
}

function safeAnalyticsText_(value, maxLength) {
  return String(value == null ? "" : value).trim().slice(0, maxLength);
}

function buildSiteAnalyticsSummary_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Site_Visits");
  const rows = !sheet || sheet.getLastRow() < 2
    ? []
    : sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
  return buildSiteAnalyticsSummaryFromRows_(rows, new Date());
}

function buildSiteAnalyticsSummaryFromRows_(rows, now) {
  const todayKey = analyticsDateKey_(now);
  const sevenDayStart = analyticsDateKey_(analyticsDaysAgo_(now, 6));
  const thirtyDayStart = analyticsDateKey_(analyticsDaysAgo_(now, 29));
  const fourteenDayStart = analyticsDateKey_(analyticsDaysAgo_(now, 13));
  const records = [];

  (rows || []).forEach((row, index) => {
    const timestamp = row[0] instanceof Date ? new Date(row[0].getTime()) : new Date(row[0]);
    if (isNaN(timestamp.getTime())) return;
    const date = analyticsDateKey_(timestamp);
    records.push({
      date: date,
      page: safeAnalyticsText_(row[1], 80) || "Unknown page",
      referrer: safeAnalyticsText_(row[3], 200) || "Direct / unknown",
      device: safeAnalyticsText_(row[4], 30) || "Unknown",
      browser: safeAnalyticsText_(row[5], 40) || "Unknown",
      sessionId: safeAnalyticsText_(row[11], 80) || "row-" + index
    });
  });

  const today = records.filter(record => record.date === todayKey);
  const sevenDays = records.filter(record => record.date >= sevenDayStart && record.date <= todayKey);
  const thirtyDays = records.filter(record => record.date >= thirtyDayStart && record.date <= todayKey);
  const dailyRecords = records.filter(record => record.date >= fourteenDayStart && record.date <= todayKey);

  return {
    generatedAt: formatDateTime(now),
    today: analyticsWindowStats_(today),
    sevenDays: analyticsWindowStats_(sevenDays),
    thirtyDays: analyticsWindowStats_(thirtyDays),
    allTime: analyticsWindowStats_(records),
    topPages: analyticsGroupedCounts_(thirtyDays, "page", true).slice(0, 5),
    devices: analyticsGroupedCounts_(thirtyDays, "device", false).slice(0, 5),
    browsers: analyticsGroupedCounts_(thirtyDays, "browser", false).slice(0, 5),
    referrers: analyticsGroupedCounts_(thirtyDays, "referrer", false).slice(0, 5),
    daily: analyticsDailySeries_(dailyRecords, now)
  };
}

function analyticsDateKey_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function analyticsDaysAgo_(date, days) {
  const result = new Date(date.getTime());
  result.setHours(12, 0, 0, 0);
  result.setDate(result.getDate() - days);
  return result;
}

function analyticsWindowStats_(records) {
  const sessions = new Set();
  records.forEach(record => { sessions.add(record.sessionId); });
  return { views: records.length, sessions: sessions.size };
}

function analyticsGroupedCounts_(records, field, includeSessions) {
  const groups = Object.create(null);

  records.forEach(record => {
    const label = record[field] || "Unknown";
    if (!groups[label]) groups[label] = { label: label, count: 0, sessions: new Set() };
    groups[label].count += 1;
    groups[label].sessions.add(record.sessionId);
  });

  return Object.keys(groups).map(label => {
    const group = groups[label];
    if (includeSessions) {
      return {
        label: group.label,
        views: group.count,
        sessions: group.sessions.size
      };
    }
    return { label: group.label, count: group.count };
  }).sort((a, b) => {
    const aCount = includeSessions ? a.views : a.count;
    const bCount = includeSessions ? b.views : b.count;
    return bCount - aCount || String(a.label).localeCompare(String(b.label));
  });
}

function analyticsDailySeries_(records, now) {
  const grouped = Object.create(null);
  records.forEach(record => {
    if (!grouped[record.date]) grouped[record.date] = { views: 0, sessions: new Set() };
    grouped[record.date].views += 1;
    grouped[record.date].sessions.add(record.sessionId);
  });

  const series = [];
  for (let offset = 13; offset >= 0; offset--) {
    const date = analyticsDateKey_(analyticsDaysAgo_(now, offset));
    const item = grouped[date] || { views: 0, sessions: new Set() };
    series.push({
      date: date,
      views: item.views,
      sessions: item.sessions.size
    });
  }
  return series;
}

function parseLocalDate(yyyyMmDd) {
  const parts = yyyyMmDd.split("-");
  const d = new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2])
  );
  d.setHours(0, 0, 0, 0);
  return d;
}

function parseSheetDate(value) {
  if (!value) return null;
  const d = new Date(value);
  if (isNaN(d.getTime())) return null;
  d.setHours(0, 0, 0, 0);
  return d;
}

function formatDateOnly(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function formatDateTime(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

function normalizeDateString(value) {
  const d = parseSheetDate(value);
  return d ? formatDateOnly(d) : "";
}

function datesEqual(a, b) {
  return a.getTime() === b.getTime();
}

function isTruthy_(value) {
  const s = String(value || "").toLowerCase();
  return s === "1" || s === "true" || s === "yes" || s === "y";
}

/* =========================
   GAME RESULTS + NOTES REVIEW
========================= */

function buildTeamGameNotesReview_(team, startDate, endDate) {
  if (!team) throw new Error("Team is required");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gcSheet = ss.getSheetByName("GC_Games_Sync");
  const gamesSheet = ss.getSheetByName("Games");
  if (!gcSheet || !gamesSheet) throw new Error("Missing GC_Games_Sync or Games sheet");

  const gcValues = gcSheet.getDataRange().getValues();
  const gameValues = gamesSheet.getDataRange().getValues();
  const gcHeaders = gcValues.length ? gcValues.shift() : [];
  const gameHeaders = gameValues.length ? gameValues.shift() : [];
  const built = buildGameReviewFromRows_(gcValues, gcHeaders, gameValues, gameHeaders, {
    team: team,
    startDate: startDate,
    endDate: endDate
  });

  return {
    team: team,
    startDate: startDate || "",
    endDate: endDate || "",
    summary: summarizeReviewGames_(built.games),
    games: built.games,
    lastSynced: built.lastSynced
  };
}

function buildGameReviewFromRows_(gcRows, gcHeaders, gameRows, gameHeaders, options) {
  const filters = options || {};
  const gcCol = name => gcHeaders.indexOf(name);
  const gameCol = name => gameHeaders.indexOf(name);
  const submissionsByGcId = Object.create(null);
  const submissions = [];

  (gameRows || []).forEach((row, index) => {
    const team = String(row[gameCol("Team")] || "").trim();
    const date = normalizeDateStringGC_(row[gameCol("Date")]);
    if (!team || !date) return;
    const item = {
      rowIndex: index,
      team: team,
      date: date,
      opponent: row[gameCol("Opponent")] || "",
      notes: String(row[gameCol("Notes")] || "").trim(),
      gcGameId: String(row[gameCol("GC_Game_ID")] || "").trim(),
      startTime: formatTimeForDisplayGC_(row[gameCol("GameStartTime")]),
      submittedBy: row[gameCol("SubmittedBy")] || "",
      submissionTime: row[gameCol("SubmissionTime")]
    };
    submissions.push(item);

    if (item.gcGameId) {
      const existing = submissionsByGcId[item.gcGameId];
      const existingTime = existing && existing.submissionTime ? new Date(existing.submissionTime).getTime() : 0;
      const itemTime = item.submissionTime ? new Date(item.submissionTime).getTime() : 0;
      if (!existing || itemTime >= existingTime) submissionsByGcId[item.gcGameId] = item;
    }
  });

  const games = [];
  const usedSubmissionRows = Object.create(null);
  let latestSync = null;

  (gcRows || []).forEach(row => {
    const team = String(row[gcCol("Team")] || "").trim();
    const date = normalizeDateStringGC_(row[gcCol("Date")]);
    if (!gameReviewMatchesFilters_(team, date, filters)) return;

    const gcGameId = String(row[gcCol("GC_Game_ID")] || "").trim();
    const submission = gcGameId ? submissionsByGcId[gcGameId] : null;
    if (submission) {
      submissions.forEach(item => {
        if (item.gcGameId === gcGameId) usedSubmissionRows[item.rowIndex] = true;
      });
    }
    const status = String(row[gcCol("Game_Status")] || "").trim();
    const teamScore = row[gcCol("Team_Score")] === "" ? "" : row[gcCol("Team_Score")];
    const opponentScore = row[gcCol("Opponent_Score")] === "" ? "" : row[gcCol("Opponent_Score")];
    const result = row[gcCol("Result")] || buildResultTextGC_(status, teamScore, opponentScore);
    const syncValue = row[gcCol("Sync_Time")];
    const syncDate = syncValue ? new Date(syncValue) : null;
    if (syncDate && !isNaN(syncDate.getTime()) && (!latestSync || syncDate > latestSync)) latestSync = syncDate;

    const notes = submission ? submission.notes : "";
    const completed = status.toLowerCase() === "completed";
    games.push({
      team: team,
      gcGameId: gcGameId,
      date: date,
      startTime: formatTimeForDisplayGC_(row[gcCol("Start_Time")]),
      opponent: row[gcCol("Opponent")] || "",
      gameStatus: status,
      teamScore: teamScore,
      opponentScore: opponentScore,
      result: result,
      notes: notes,
      hasNotes: !!notes,
      notesMissing: completed && !notes,
      hasSubmission: !!submission,
      submittedBy: submission ? submission.submittedBy : "",
      submissionTime: submission && submission.submissionTime
        ? formatDateTime(new Date(submission.submissionTime))
        : "",
      matchStatus: submission ? "matched" : "gamechanger_only"
    });
  });

  submissions.forEach(submission => {
    if (usedSubmissionRows[submission.rowIndex]) return;
    if (!gameReviewMatchesFilters_(submission.team, submission.date, filters)) return;
    games.push({
      team: submission.team,
      gcGameId: submission.gcGameId,
      date: submission.date,
      startTime: submission.startTime,
      opponent: submission.opponent,
      gameStatus: "submitted",
      teamScore: "",
      opponentScore: "",
      result: "",
      notes: submission.notes,
      hasNotes: !!submission.notes,
      notesMissing: false,
      hasSubmission: true,
      submittedBy: submission.submittedBy,
      submissionTime: submission.submissionTime
        ? formatDateTime(new Date(submission.submissionTime))
        : "",
      matchStatus: submission.gcGameId ? "gamechanger_unavailable" : "manual"
    });
  });

  games.sort((a, b) => {
    const byDate = String(b.date).localeCompare(String(a.date));
    if (byDate) return byDate;
    return parseTimeMinutes_(b.startTime) - parseTimeMinutes_(a.startTime);
  });

  return {
    games: games,
    lastSynced: latestSync ? formatDateTime(latestSync) : "Never"
  };
}

function gameReviewMatchesFilters_(team, date, filters) {
  if (!team || !date) return false;
  if (filters.team && team !== filters.team) return false;
  if (filters.startDate && date < filters.startDate) return false;
  if (filters.endDate && date > filters.endDate) return false;
  return true;
}

function summarizeReviewGames_(games) {
  const summary = createEmptyStatLine_();
  (games || []).forEach(game => {
    if (String(game.gameStatus || "").toLowerCase() !== "completed") return;
    if (game.teamScore === "" || game.opponentScore === "") return;
    const teamScore = Number(game.teamScore);
    const opponentScore = Number(game.opponentScore);
    if (!isNaN(teamScore) && !isNaN(opponentScore)) applyGameToStatLine_(summary, teamScore, opponentScore);
  });
  summary.missingNotes = (games || []).filter(game => game.notesMissing).length;
  return summary;
}

function buildWeeklyReportData_(referenceDate) {
  const window = getWeekWindowForDate_(referenceDate || new Date());
  const ss = getPitchLogSpreadsheet_();
  const gcSheet = ss.getSheetByName("GC_Games_Sync");
  const gamesSheet = ss.getSheetByName("Games");
  if (!gcSheet || !gamesSheet) throw new Error("Missing GC_Games_Sync or Games sheet");

  const gcValues = gcSheet.getDataRange().getValues();
  const gameValues = gamesSheet.getDataRange().getValues();
  const gcHeaders = gcValues.length ? gcValues.shift() : [];
  const gameHeaders = gameValues.length ? gameValues.shift() : [];
  const startDate = formatDateOnly(window.startDate);
  const endDate = formatDateOnly(window.endDate);
  const built = buildGameReviewFromRows_(gcValues, gcHeaders, gameValues, gameHeaders, {
    startDate: startDate,
    endDate: endDate
  });

  return buildWeeklyReportFromGames_(built.games, startDate, endDate, built.lastSynced);
}

function buildWeeklyReportFromGames_(games, startDate, endDate, lastSynced) {
  const grouped = Object.create(null);
  (games || []).forEach(game => {
    if (!grouped[game.team]) grouped[game.team] = [];
    grouped[game.team].push(game);
  });

  const teams = Object.keys(grouped).sort((a, b) => a.localeCompare(b)).map(team => ({
    team: team,
    summary: summarizeReviewGames_(grouped[team]),
    games: grouped[team]
  }));

  return {
    generatedAt: formatDateTime(new Date()),
    startDate: startDate,
    endDate: endDate,
    lastSynced: lastSynced || "Never",
    summary: summarizeReviewGames_(games),
    teams: teams
  };
}

/* =========================
   PERFORMANCE + ADMIN HELPERS
========================= */

function buildPerformanceSummary_(includeFall) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gcSheet = ss.getSheetByName("GC_Games_Sync");

  if (!gcSheet) {
    throw new Error("Missing GC_Games_Sync sheet");
  }

  const rows = gcSheet.getDataRange().getValues();
  const windows = getWeekWindows_();

  if (rows.length <= 1) {
    return {
      seasonYear: DEFAULT_SEASON_YEAR,
      includeFall: includeFall,
      overallOrg: createEmptyStatLine_(),
      currentWeekOrg: buildWindowStatLine_(createEmptyStatLine_(), windows.current.startDate, windows.current.endDate),
      previousWeekOrg: buildWindowStatLine_(createEmptyStatLine_(), windows.previous.startDate, windows.previous.endDate),
      teams: []
    };
  }

  const headers = rows.shift();

  const teamCol = headers.indexOf("Team");
  const dateCol = headers.indexOf("Date");
  const statusCol = headers.indexOf("Game_Status");
  const teamScoreCol = headers.indexOf("Team_Score");
  const oppScoreCol = headers.indexOf("Opponent_Score");

  const overallOrg = createEmptyStatLine_();
  const currentWeekOrg = buildWindowStatLine_(createEmptyStatLine_(), windows.current.startDate, windows.current.endDate);
  const previousWeekOrg = buildWindowStatLine_(createEmptyStatLine_(), windows.previous.startDate, windows.previous.endDate);

  const teamStats = {};

  rows.forEach(r => {
    const status = String(r[statusCol] || "").toLowerCase();
    if (status !== "completed") return;

    const team = r[teamCol] || "";
    const dateStr = normalizeDateStringGC_(r[dateCol]);
    if (!team || !dateStr) return;

    const gameDate = parseLocalDate(dateStr);
    if (!shouldIncludeGameForSeason_(gameDate, includeFall)) return;

    const teamScore = Number(r[teamScoreCol]);
    const oppScore = Number(r[oppScoreCol]);

    if (isNaN(teamScore) || isNaN(oppScore)) return;

    if (!teamStats[team]) {
      teamStats[team] = {
        team: team,
        overall: createEmptyStatLine_(),
        currentWeek: buildWindowStatLine_(createEmptyStatLine_(), windows.current.startDate, windows.current.endDate),
        previousWeek: buildWindowStatLine_(createEmptyStatLine_(), windows.previous.startDate, windows.previous.endDate)
      };
    }

    applyGameToStatLine_(overallOrg, teamScore, oppScore);
    applyGameToStatLine_(teamStats[team].overall, teamScore, oppScore);

    if (isDateInRange_(gameDate, windows.current.startDate, windows.current.endDate)) {
      applyGameToStatLine_(currentWeekOrg, teamScore, oppScore);
      applyGameToStatLine_(teamStats[team].currentWeek, teamScore, oppScore);
    }

    if (isDateInRange_(gameDate, windows.previous.startDate, windows.previous.endDate)) {
      applyGameToStatLine_(previousWeekOrg, teamScore, oppScore);
      applyGameToStatLine_(teamStats[team].previousWeek, teamScore, oppScore);
    }
  });

  return {
    seasonYear: DEFAULT_SEASON_YEAR,
    includeFall: includeFall,
    overallOrg: overallOrg,
    currentWeekOrg: currentWeekOrg,
    previousWeekOrg: previousWeekOrg,
    teams: Object.keys(teamStats)
      .sort((a, b) => a.localeCompare(b))
      .map(k => teamStats[k])
  };
}

function buildCurrentWeekSchedule_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gcSheet = ss.getSheetByName("GC_Games_Sync");

  if (!gcSheet) {
    throw new Error("Missing GC_Games_Sync sheet");
  }

  const rows = gcSheet.getDataRange().getValues();
  const windows = getWeekWindows_();

  if (rows.length <= 1) {
    return {
      startDate: formatDateOnly(windows.current.startDate),
      endDate: formatDateOnly(windows.current.endDate),
      totalRowsRead: 0,
      totalRowsMatched: 0,
      teams: []
    };
  }

  const headers = rows.shift();

  const teamCol = headers.indexOf("Team");
  const dateCol = headers.indexOf("Date");
  const timeCol = headers.indexOf("Start_Time");
  const oppCol = headers.indexOf("Opponent");
  const statusCol = headers.indexOf("Game_Status");
  const resultCol = headers.indexOf("Result");
  const teamScoreCol = headers.indexOf("Team_Score");
  const oppScoreCol = headers.indexOf("Opponent_Score");

  const grouped = {};
  let matchedCount = 0;

  rows.forEach(r => {
    const team = String(r[teamCol] || "").trim();
    if (!team) return;

    const rawDate = r[dateCol];
    const normalizedDate = normalizeDateStringGC_(rawDate);
    if (!normalizedDate) return;

    const gameDate = parseLocalDate(normalizedDate);
    if (!gameDate) return;

    if (!isDateInRange_(gameDate, windows.current.startDate, windows.current.endDate)) {
      return;
    }

    matchedCount += 1;

    const time = r[timeCol] || "";
    const opponent = r[oppCol] || "";
    const status = r[statusCol] || "";
    const result = r[resultCol] || "";
    const teamScore = r[teamScoreCol] === "" ? "" : r[teamScoreCol];
    const opponentScore = r[oppScoreCol] === "" ? "" : r[oppScoreCol];

    if (!grouped[team]) {
      grouped[team] = {
        team: team,
        summary: createEmptyStatLine_(),
        games: []
      };
    }

    if (String(status).toLowerCase() === "completed" && teamScore !== "" && opponentScore !== "") {
      const t = Number(teamScore);
      const o = Number(opponentScore);

      if (!isNaN(t) && !isNaN(o)) {
        applyGameToStatLine_(grouped[team].summary, t, o);
      }
    }

    grouped[team].games.push({
      date: normalizedDate,
      startTime: time,
      opponent: opponent,
      gameStatus: status,
      result: result,
      teamScore: teamScore,
      opponentScore: opponentScore
    });
  });

  const teams = Object.keys(grouped)
    .sort((a, b) => a.localeCompare(b))
    .map(teamName => {
      const item = grouped[teamName];
      item.games.sort((a, b) =>
        String(a.date + " " + a.startTime + " " + a.opponent).localeCompare(
          String(b.date + " " + b.startTime + " " + b.opponent)
        )
      );
      return item;
    });

  return {
    startDate: formatDateOnly(windows.current.startDate),
    endDate: formatDateOnly(windows.current.endDate),
    totalRowsRead: rows.length,
    totalRowsMatched: matchedCount,
    teams: teams
  };
}

function buildFallPracticeSchedule_() {
  // This connection is intentionally read-only. It never calls a source-sheet
  // setter, append, clear, delete, or batch-update method.
  const source = SpreadsheetApp.openById(FALL_PRACTICE_SOURCE_ID);
  const master = source.getSheetByName(FALL_PRACTICE_SOURCE_TAB);

  if (!master) {
    throw new Error("The Master schedule tab was not found.");
  }

  const values = master.getRange(1, 1, 120, Math.min(master.getLastColumn(), 377)).getDisplayValues();
  const dayLayouts = [
    { dateRow: 2, nextDateRow: 9 },
    { dateRow: 9, nextDateRow: 16 },
    { dateRow: 16, nextDateRow: 23 },
    { dateRow: 23, nextDateRow: 30 },
    { dateRow: 30, nextDateRow: 37 },
    { dateRow: 37, nextDateRow: 43 },
    { dateRow: 43, nextDateRow: 121 }
  ];
  const allowedTeams = {};
  FALL_PRACTICE_TEAMS.forEach(team => allowedTeams[team] = true);

  const practices = [];
  const seen = {};

  dayLayouts.forEach(layout => {
    const dateIndex = layout.dateRow - 1;
    const fieldIndex = dateIndex + 1;
    const firstTimeIndex = dateIndex + 2;
    const lastTimeIndex = Math.min(layout.nextDateRow - 2, values.length - 1);

    values[dateIndex].forEach((rawDate, dateColumn) => {
      const parsedDate = parseFallMasterDate_(rawDate);
      if (!parsedDate) return;

      const fieldStart = dateColumn - 3;
      const timeColumn = fieldStart - 1;
      if (fieldStart < 0 || timeColumn < 0) return;

      for (let row = firstTimeIndex; row <= lastTimeIndex; row += 1) {
        const time = String(values[row][timeColumn] || "").trim();
        if (!time) continue;

        for (let fieldOffset = 0; fieldOffset < 8; fieldOffset += 1) {
          const rawTeam = String(values[row][fieldStart + fieldOffset] || "").trim();
          const team = rawTeam.toUpperCase();
          if (!allowedTeams[team]) continue;

          const fieldLabel = String(values[fieldIndex][fieldStart + fieldOffset] || (fieldOffset + 1)).trim();
          const key = [parsedDate.iso, time, fieldLabel, team].join("|");
          if (seen[key]) continue;
          seen[key] = true;

          practices.push({
            date: parsedDate.iso,
            displayDate: parsedDate.display,
            day: parsedDate.day,
            month: parsedDate.month,
            time: time,
            field: fieldLabel,
            team: team
          });
        }
      }
    });
  });

  practices.sort((a, b) => {
    const byDate = a.date.localeCompare(b.date);
    if (byDate) return byDate;
    const byTime = practiceTimeSortValue_(a.time) - practiceTimeSortValue_(b.time);
    if (byTime) return byTime;
    return Number(a.field) - Number(b.field) || a.team.localeCompare(b.team);
  });

  return {
    sourceTitle: source.getName(),
    sourceTab: FALL_PRACTICE_SOURCE_TAB,
    readOnly: true,
    refreshedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMM d, yyyy h:mm a"),
    teams: FALL_PRACTICE_TEAMS,
    practices: practices
  };
}

function buildDrills_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Drills");

  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet.getRange(1, 1, sheet.getLastRow(), 6).getDisplayValues();
  const headers = values[0].map(value => String(value || "").trim().toLowerCase());
  const columnIndex = name => headers.indexOf(name.toLowerCase());
  const indexes = {
    title: columnIndex("Title"),
    category: columnIndex("Category"),
    focus: columnIndex("Focus"),
    source: columnIndex("Source"),
    url: columnIndex("URL"),
    notes: columnIndex("Notes")
  };

  if (indexes.title < 0) {
    throw new Error("The Drills tab must include a Title column.");
  }

  return values.slice(1).map(row => ({
    title: String(row[indexes.title] || "").trim(),
    category: indexes.category < 0 ? "" : String(row[indexes.category] || "").trim(),
    focus: indexes.focus < 0 ? "" : String(row[indexes.focus] || "").trim(),
    source: indexes.source < 0 ? "" : String(row[indexes.source] || "").trim(),
    url: indexes.url < 0 ? "" : String(row[indexes.url] || "").trim(),
    notes: indexes.notes < 0 ? "" : String(row[indexes.notes] || "").trim()
  })).filter(drill => drill.title);
}

function parseFallMasterDate_(value) {
  const match = String(value || "").trim().match(/^(8|9|10)\/(\d{1,2})(?:\/(\d{2,4}))?$/);
  if (!match) return null;

  const month = Number(match[1]);
  const day = Number(match[2]);
  let year = Number(match[3] || DEFAULT_SEASON_YEAR);
  if (year < 100) year += 2000;

  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) {
    return null;
  }

  const monthName = ["", "", "", "", "", "", "", "", "August", "September", "October"][month];
  return {
    iso: Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd"),
    display: Utilities.formatDate(date, Session.getScriptTimeZone(), "MMM d"),
    day: Utilities.formatDate(date, Session.getScriptTimeZone(), "EEE"),
    month: monthName
  };
}

function practiceTimeSortValue_(value) {
  const match = String(value || "").trim().match(/^(\d{1,2})(?::(\d{2}))?\s*(AM|PM)?$/i);
  if (!match) return 9999;
  let hour = Number(match[1]);
  const minute = Number(match[2] || 0);
  const meridiem = String(match[3] || "PM").toUpperCase();
  if (meridiem === "PM" && hour < 12) hour += 12;
  if (meridiem === "AM" && hour === 12) hour = 0;
  return hour * 60 + minute;
}

function buildSubmissionSummary_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const checkSheet = ss.getSheetByName("GC_Submission_Check");
  const ignoredKeys = getIgnoredMissingKeys_();

  const result = {
    compareStartDate: GC_COMPARE_START_DATE,
    completedGames: 0,
    ntbSubmissions: 0,
    missingSubmissions: 0,
    extraSubmissions: 0,
    rowsChecked: 0,
    ignoredMissingRows: 0
  };

  if (!checkSheet) {
    return result;
  }

  const rows = checkSheet.getDataRange().getValues();
  if (rows.length <= 1) {
    return result;
  }

  const headers = rows.shift();

  const teamCol = headers.indexOf("Team");
  const dateCol = headers.indexOf("Date");
  const gcCol = headers.indexOf("GC_Completed_Games");
  const ntbCol = headers.indexOf("NTB_Submissions");
  const missingCol = headers.indexOf("Missing_Count");
  const extraCol = headers.indexOf("Extra_Count");

  rows.forEach(r => {
    const team = r[teamCol] || "";
    const date = normalizeDateStringGC_(r[dateCol]);
    const key = buildTeamDateKey_(team, date);
    const isIgnored = ignoredKeys[key] === true;

    result.completedGames += Number(r[gcCol]) || 0;
    result.ntbSubmissions += Number(r[ntbCol]) || 0;
    result.extraSubmissions += Number(r[extraCol]) || 0;
    result.rowsChecked += 1;

    const missing = Number(r[missingCol]) || 0;
    if (isIgnored) {
      result.ignoredMissingRows += missing > 0 ? 1 : 0;
    } else {
      result.missingSubmissions += missing;
    }
  });

  return result;
}

function buildAdminMissingSubmissions_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gcSheet = ss.getSheetByName("GC_Games_Sync");
  const checkSheet = ss.getSheetByName("GC_Submission_Check");
  const ignoredKeys = getIgnoredMissingKeys_();

  if (!gcSheet || !checkSheet) {
    return [];
  }

  const gcRows = gcSheet.getDataRange().getValues();
  const checkRows = checkSheet.getDataRange().getValues();

  if (gcRows.length <= 1 || checkRows.length <= 1) {
    return [];
  }

  const gcHeaders = gcRows.shift();
  const checkHeaders = checkRows.shift();

  const gcTeamCol = gcHeaders.indexOf("Team");
  const gcDateCol = gcHeaders.indexOf("Date");
  const gcOppCol = gcHeaders.indexOf("Opponent");
  const gcResultCol = gcHeaders.indexOf("Result");
  const gcStatusCol = gcHeaders.indexOf("Game_Status");

  const checkTeamCol = checkHeaders.indexOf("Team");
  const checkDateCol = checkHeaders.indexOf("Date");
  const checkMissingCol = checkHeaders.indexOf("Missing_Count");
  const checkStatusCol = checkHeaders.indexOf("Check_Status");

  const missingKeys = {};

  checkRows.forEach(r => {
    const status = r[checkStatusCol] || "";
    const missing = Number(r[checkMissingCol]) || 0;
    const team = r[checkTeamCol] || "";
    const date = normalizeDateStringGC_(r[checkDateCol]);
    const key = buildTeamDateKey_(team, date);

    if (status !== "Missing Submission" || missing <= 0 || !team || !date) return;
    if (ignoredKeys[key]) return;

    missingKeys[key] = true;
  });

  const results = [];

  gcRows.forEach(r => {
    const status = String(r[gcStatusCol] || "").toLowerCase();
    if (status !== "completed") return;

    const team = r[gcTeamCol] || "";
    const date = normalizeDateStringGC_(r[gcDateCol]);
    const key = buildTeamDateKey_(team, date);

    if (!missingKeys[key]) return;

    results.push({
      team: team,
      date: date,
      opponent: r[gcOppCol] || "",
      result: r[gcResultCol] || ""
    });
  });

  results.sort((a, b) => String(b.date + b.team + b.opponent).localeCompare(String(a.date + a.team + a.opponent)));
  return results;
}

function createEmptyStatLine_() {
  return {
    wins: 0,
    losses: 0,
    ties: 0,
    runsScored: 0,
    runsAllowed: 0,
    completedGames: 0
  };
}

function buildWindowStatLine_(line, startDate, endDate) {
  line.startDate = formatDateOnly(startDate);
  line.endDate = formatDateOnly(endDate);
  return line;
}

function applyGameToStatLine_(line, teamScore, oppScore) {
  line.completedGames += 1;
  line.runsScored += teamScore;
  line.runsAllowed += oppScore;

  if (teamScore > oppScore) {
    line.wins += 1;
  } else if (teamScore < oppScore) {
    line.losses += 1;
  } else {
    line.ties += 1;
  }
}

function getWeekWindows_() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const current = getWeekWindowForDate_(today);

  const previousWeekStart = new Date(current.startDate);
  previousWeekStart.setDate(previousWeekStart.getDate() - 7);

  const previousWeekEnd = new Date(previousWeekStart);
  previousWeekEnd.setDate(previousWeekEnd.getDate() + 6);

  return {
    current: current,
    previous: {
      startDate: previousWeekStart,
      endDate: previousWeekEnd
    }
  };
}

function getWeekWindowForDate_(date) {
  const startDate = getMondayOfWeek_(date);
  const endDate = new Date(startDate);
  endDate.setDate(endDate.getDate() + 6);
  return { startDate: startDate, endDate: endDate };
}

function getMondayOfWeek_(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);

  const day = d.getDay();
  const diff = day === 0 ? -6 : 1 - day;

  d.setDate(d.getDate() + diff);
  return d;
}

function isDateInRange_(date, startDate, endDate) {
  return date >= startDate && date <= endDate;
}

function shouldIncludeGameForSeason_(gameDate, includeFall) {
  const yearStart = parseLocalDate(DEFAULT_SEASON_YEAR + "-01-01");
  const yearEnd = parseLocalDate(DEFAULT_SEASON_YEAR + "-12-31");
  const fallStart = parseLocalDate(FALL_SEASON_START_DATE);

  if (gameDate < yearStart || gameDate > yearEnd) return false;
  if (!includeFall && gameDate >= fallStart) return false;

  return true;
}

function buildTeamDateKey_(team, date) {
  return String(team || "").trim() + "||" + String(date || "").trim();
}

function getIgnoredMissingKeys_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("GC_Submission_Overrides");
  const results = {};

  if (!sheet) {
    return results;
  }

  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) {
    return results;
  }

  const headers = rows.shift();
  const teamCol = headers.indexOf("Team");
  const dateCol = headers.indexOf("Date");
  const typeCol = headers.indexOf("Override_Type");

  rows.forEach(r => {
    const team = r[teamCol] || "";
    const date = normalizeDateStringGC_(r[dateCol]);
    const type = String(r[typeCol] || "").trim();

    if (!team || !date) return;
    if (type !== MISSING_OVERRIDE_TYPE) return;

    results[buildTeamDateKey_(team, date)] = true;
  });

  return results;
}

function ensureSubmissionOverridesSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("GC_Submission_Overrides");

  if (!sheet) {
    sheet = ss.insertSheet("GC_Submission_Overrides");
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Team", "Date", "Override_Type", "Reason", "Created_At"]);
  } else if (sheet.getLastRow() === 1 && sheet.getLastColumn() === 1 && !sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, 5).setValues([["Team", "Date", "Override_Type", "Reason", "Created_At"]]);
  }

  return sheet;
}

function addMissingOverride_(team, date, reason) {
  const sheet = ensureSubmissionOverridesSheet_();
  const rows = sheet.getDataRange().getValues();

  if (rows.length > 1) {
    const headers = rows.shift();
    const teamCol = headers.indexOf("Team");
    const dateCol = headers.indexOf("Date");
    const typeCol = headers.indexOf("Override_Type");

    for (let i = 0; i < rows.length; i++) {
      const existingTeam = rows[i][teamCol] || "";
      const existingDate = normalizeDateStringGC_(rows[i][dateCol]);
      const existingType = String(rows[i][typeCol] || "").trim();

      if (
        buildTeamDateKey_(existingTeam, existingDate) === buildTeamDateKey_(team, date) &&
        existingType === MISSING_OVERRIDE_TYPE
      ) {
        return;
      }
    }
  }

  sheet.appendRow([
    team,
    date,
    MISSING_OVERRIDE_TYPE,
    reason,
    new Date()
  ]);
}

/* =========================
   WEEKLY ORGANIZATION EMAIL
========================= */

function buildWeeklyReportEmailHtml_(report, logoSource) {
  const summary = report.summary || createEmptyStatLine_();
  const teams = Array.isArray(report.teams) ? report.teams : [];
  const dateRange = weeklyEmailDateRange_(report.startDate, report.endDate);
  const logo = escapeHtml_(logoSource || APP_LOGO_URL);
  const reviewUrl = APP_BASE_URL + "/index.html?view=notes";

  let teamHtml = "";
  teams.forEach(team => {
    const stats = team.summary || createEmptyStatLine_();
    let gamesHtml = "";
    (team.games || []).forEach(game => {
      const result = String(game.result || "").trim();
      const status = String(game.gameStatus || "").trim();
      const outcomeColor = result.indexOf("W ") === 0
        ? "#16794b"
        : result.indexOf("L ") === 0 ? "#b42318" : "#6b6250";
      const score = game.teamScore !== "" && game.opponentScore !== ""
        ? escapeHtml_(String(game.teamScore)) + "–" + escapeHtml_(String(game.opponentScore))
        : escapeHtml_(result || weeklyTitleCase_(status || "Score pending"));
      const noteBody = game.hasNotes
        ? escapeHtml_(game.notes).replace(/\n/g, "<br>")
        : game.notesMissing
          ? '<span style="color:#8a5a00;font-weight:bold;">Game notes have not been submitted.</span>'
          : '<span style="color:#6b6250;">No notes recorded.</span>';
      const matchLabel = game.matchStatus === "manual"
        ? '<span style="color:#6b6250;font-size:12px;">Manual entry</span>'
        : "";

      gamesHtml += `
        <div style="border:1px solid #e1d8c4;border-radius:8px;padding:12px;margin:0 0 10px;background:#fff;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>
            <td style="vertical-align:top;padding-right:10px;">
              <div style="font-size:12px;color:#6b6250;margin-bottom:4px;">${escapeHtml_(weeklyEmailGameDate_(game.date))}${game.startTime ? " • " + escapeHtml_(game.startTime) : ""}</div>
              <div style="font-size:15px;font-weight:bold;color:#171717;">vs ${escapeHtml_(game.opponent || "Opponent unavailable")}</div>
              ${matchLabel}
            </td>
            <td align="right" style="vertical-align:top;white-space:nowrap;">
              <div style="font-size:20px;font-weight:bold;color:${outcomeColor};">${score}</div>
              <div style="font-size:11px;color:#6b6250;text-transform:uppercase;">${escapeHtml_(result || status)}</div>
            </td>
          </tr></table>
          <div style="border-top:1px solid #eee7d7;margin-top:10px;padding-top:10px;font-size:14px;line-height:1.5;color:#2c2922;">${noteBody}</div>
        </div>`;
    });

    const teamUrl = reviewUrl + "&team=" + encodeURIComponent(team.team || "");
    teamHtml += `
      <div style="margin:18px 0 0;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#171717;border-radius:8px 8px 0 0;"><tr>
          <td style="padding:12px 14px;color:#fff;font-size:18px;font-weight:bold;">${escapeHtml_(team.team || "")}</td>
          <td align="right" style="padding:12px 14px;color:#d5bf76;font-size:13px;">${escapeHtml_(weeklyRecordText_(stats))} • RS ${Number(stats.runsScored || 0)} • RA ${Number(stats.runsAllowed || 0)}</td>
        </tr></table>
        <div style="border:1px solid #d9cfb8;border-top:none;border-radius:0 0 8px 8px;padding:12px;background:#fbf8f0;">
          ${gamesHtml || '<div style="color:#6b6250;">No games this week.</div>'}
          <a href="${escapeHtml_(teamUrl)}" style="display:inline-block;color:#70591d;font-weight:bold;text-decoration:none;font-size:13px;">Review ${escapeHtml_(team.team || "")} game notes →</a>
        </div>
      </div>`;
  });

  return `
    <div style="margin:0;padding:22px 10px;background:#f4f1ea;font-family:Arial,Helvetica,sans-serif;color:#171717;">
      <div style="max-width:680px;margin:0 auto;">
        <div style="background:#050505;border-radius:10px;padding:18px;text-align:center;">
          <img src="${logo}" width="82" height="82" alt="Noblesville Travel Baseball" style="display:block;margin:0 auto 10px;border-radius:50%;border:2px solid #b8a05c;">
          <div style="color:#c8b46c;font-size:12px;font-weight:bold;text-transform:uppercase;letter-spacing:.7px;">Noblesville Travel Baseball</div>
          <div style="color:#fff;font-size:24px;font-weight:bold;margin-top:5px;">Weekly Game Report</div>
          <div style="color:#ddd5c0;font-size:14px;margin-top:5px;">${escapeHtml_(dateRange)}</div>
        </div>

        <table role="presentation" width="100%" cellpadding="0" cellspacing="8" style="margin:10px -8px 0;width:calc(100% + 16px);"><tr>
          ${weeklyMetricCell_("Record", weeklyRecordText_(summary))}
          ${weeklyMetricCell_("Completed Games", Number(summary.completedGames || 0))}
        </tr><tr>
          ${weeklyMetricCell_("Runs Scored", Number(summary.runsScored || 0))}
          ${weeklyMetricCell_("Runs Allowed", Number(summary.runsAllowed || 0))}
        </tr></table>

        ${teamHtml || '<div style="background:#fff;border:1px solid #d9cfb8;border-radius:8px;padding:18px;margin-top:16px;color:#6b6250;">No games are currently listed for this week.</div>'}

        <div style="text-align:center;color:#6b6250;font-size:11px;line-height:1.5;padding:18px 8px;">
          GameChanger data last synced ${escapeHtml_(report.lastSynced || "Never")}.<br>
          <a href="${escapeHtml_(reviewUrl)}" style="color:#70591d;">Open Game Notes</a>
        </div>
      </div>
    </div>`;
}

function weeklyMetricCell_(label, value) {
  return `<td width="50%" style="background:#fff;border:1px solid #ddd2ba;border-radius:8px;padding:13px;text-align:center;">
    <div style="font-size:11px;color:#6b6250;text-transform:uppercase;">${escapeHtml_(label)}</div>
    <div style="font-size:24px;font-weight:bold;margin-top:4px;">${escapeHtml_(String(value))}</div>
  </td>`;
}

function buildWeeklyReportPlainText_(report) {
  const summary = report.summary || createEmptyStatLine_();
  let body = "NOBLESVILLE TRAVEL BASEBALL\nWEEKLY GAME REPORT\n" +
    weeklyEmailDateRange_(report.startDate, report.endDate) + "\n\n" +
    "Organization: " + weeklyRecordText_(summary) +
    " | Games " + Number(summary.completedGames || 0) +
    " | RS " + Number(summary.runsScored || 0) +
    " | RA " + Number(summary.runsAllowed || 0) + "\n";

  (report.teams || []).forEach(team => {
    const stats = team.summary || createEmptyStatLine_();
    body += "\n" + team.team + " — " + weeklyRecordText_(stats) +
      " | RS " + Number(stats.runsScored || 0) +
      " | RA " + Number(stats.runsAllowed || 0) + "\n";
    (team.games || []).forEach(game => {
      body += "• " + game.date + (game.startTime ? " " + game.startTime : "") +
        " vs " + (game.opponent || "Opponent unavailable") +
        " — " + (game.result || game.gameStatus || "Score pending") + "\n" +
        "  Notes: " + (game.notes || (game.notesMissing ? "MISSING" : "None")) + "\n";
    });
  });
  body += "\nOpen Game Notes: " + APP_BASE_URL + "/index.html?view=notes";
  return body;
}

function weeklyRecordText_(stats) {
  const record = Number(stats.wins || 0) + "-" + Number(stats.losses || 0);
  return Number(stats.ties || 0) ? record + "-" + Number(stats.ties || 0) : record;
}

function weeklyEmailDateRange_(startDate, endDate) {
  const start = parseLocalDate(startDate);
  const end = parseLocalDate(endDate);
  return Utilities.formatDate(start, WEEKLY_REPORT_TIMEZONE, "MMM d") + "–" +
    Utilities.formatDate(end, WEEKLY_REPORT_TIMEZONE, "MMM d, yyyy");
}

function weeklyEmailGameDate_(date) {
  return Utilities.formatDate(parseLocalDate(date), WEEKLY_REPORT_TIMEZONE, "EEE, MMM d");
}

function weeklyTitleCase_(value) {
  const text = String(value || "").trim();
  return text ? text.charAt(0).toUpperCase() + text.slice(1).toLowerCase() : "";
}

function escapeHtml_(value) {
  return String(value == null ? "" : value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getActiveWeeklyRecipients_() {
  const sheet = getPitchLogSpreadsheet_().getSheetByName(WEEKLY_RECIPIENTS_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getRange(1, 1, sheet.getLastRow(), 3).getValues();
  const headers = values.shift();
  const nameCol = headers.indexOf("Name");
  const emailCol = headers.indexOf("Email");
  const activeCol = headers.indexOf("Active");
  const seen = Object.create(null);

  return values.map(row => ({
    name: nameCol < 0 ? "" : String(row[nameCol] || "").trim(),
    email: emailCol < 0 ? "" : String(row[emailCol] || "").trim(),
    active: activeCol >= 0 && isTruthy_(row[activeCol])
  })).filter(recipient => {
    const key = recipient.email.toLowerCase();
    if (!recipient.active || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(recipient.email) || seen[key]) return false;
    seen[key] = true;
    return true;
  });
}

function sendWeeklyReportTest() {
  const report = buildWeeklyReportData_(new Date());
  sendWeeklyReportToRecipients_([{ name: "Admin", email: ADMIN_EMAIL }], report);
  logWeeklyReport_(report, 1, "Test", "Test report sent to " + ADMIN_EMAIL);
  return "Test report sent to " + ADMIN_EMAIL;
}

function sendWeeklyReportNow() {
  const recipients = getActiveWeeklyRecipients_();
  if (!recipients.length) throw new Error("No active weekly email recipients are configured");
  syncGameChangerAndSubmissionCheck();
  const report = buildWeeklyReportData_(new Date());
  sendWeeklyReportToRecipients_(recipients, report);
  markWeeklyReportSent_(report.startDate);
  logWeeklyReport_(report, recipients.length, "Sent", "Manual weekly report sent");
  return "Weekly report sent to " + recipients.length + " recipient(s)";
}

function sendWeeklyReportToRecipients_(recipients, report) {
  if (MailApp.getRemainingDailyQuota() < recipients.length) {
    throw new Error("Not enough email recipient quota remains for the weekly report");
  }

  let logoSource = APP_LOGO_URL;
  let logoBlob = null;
  try {
    const response = UrlFetchApp.fetch(APP_LOGO_URL, { muteHttpExceptions: true });
    if (response.getResponseCode() >= 200 && response.getResponseCode() < 300) {
      logoBlob = response.getBlob().setName("nyb-logo.png");
      logoSource = "cid:nybLogo";
    }
  } catch (err) {
    console.warn("Unable to inline the app logo: " + err);
  }

  const subject = "NYB Council Weekly Game Report • " + weeklyEmailDateRange_(report.startDate, report.endDate);
  const body = buildWeeklyReportPlainText_(report);
  const htmlBody = buildWeeklyReportEmailHtml_(report, logoSource);

  recipients.forEach(recipient => {
    const message = {
      to: recipient.email,
      subject: subject,
      body: body,
      htmlBody: htmlBody,
      name: "Noblesville Travel Baseball"
    };
    if (logoBlob) message.inlineImages = { nybLogo: logoBlob };
    MailApp.sendEmail(message);
  });
}

function maybeSendWeeklyReport_(now, alreadySynced) {
  const day = Number(Utilities.formatDate(now, WEEKLY_REPORT_TIMEZONE, "u"));
  const hour = Number(Utilities.formatDate(now, WEEKLY_REPORT_TIMEZONE, "H"));
  if (day !== 7 || hour < 20) return { status: "not_due" };

  const window = getWeekWindowForDate_(now);
  const weekStart = formatDateOnly(window.startDate);
  if (weeklyReportWasSent_(weekStart)) return { status: "already_sent", weekStart: weekStart };

  const recipients = getActiveWeeklyRecipients_();
  if (!recipients.length) return { status: "no_recipients", weekStart: weekStart };

  let report = null;
  try {
    if (!alreadySynced) syncGameChangerAndSubmissionCheck();
    report = buildWeeklyReportData_(now);
    if (!Number(report.summary.completedGames || 0)) {
      markWeeklyReportSent_(weekStart);
      logWeeklyReport_(report, recipients.length, "Skipped", "No completed games this week");
      return { status: "no_games", weekStart: weekStart };
    }
    sendWeeklyReportToRecipients_(recipients, report);
    markWeeklyReportSent_(weekStart);
    logWeeklyReport_(report, recipients.length, "Sent", "Automatic Sunday report sent");
    return { status: "sent", weekStart: weekStart, recipients: recipients.length };
  } catch (err) {
    const failedReport = report || {
      startDate: weekStart,
      endDate: formatDateOnly(window.endDate),
      summary: { completedGames: 0 }
    };
    try {
      logWeeklyReport_(failedReport, recipients.length, "Error", err.message || String(err));
    } catch (logErr) {
      console.error("Unable to log weekly report failure: " + logErr);
    }
    console.error("Weekly report failed: " + err);
    return { status: "error", weekStart: weekStart, message: err.message || String(err) };
  }
}

function weeklyReportWasSent_(weekStart) {
  return PropertiesService.getScriptProperties().getProperty("WEEKLY_REPORT_SENT_" + weekStart) === "true";
}

function markWeeklyReportSent_(weekStart) {
  PropertiesService.getScriptProperties().setProperty("WEEKLY_REPORT_SENT_" + weekStart, "true");
}

function logWeeklyReport_(report, recipientCount, status, message) {
  const ss = getPitchLogSpreadsheet_();
  let sheet = ss.getSheetByName(WEEKLY_REPORT_LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(WEEKLY_REPORT_LOG_SHEET);
    sheet.appendRow(["Week Start", "Week End", "Sent At", "Recipient Count", "Completed Games", "Status", "Message"]);
  }
  sheet.appendRow([
    parseLocalDate(report.startDate),
    parseLocalDate(report.endDate),
    new Date(),
    recipientCount,
    Number((report.summary || {}).completedGames || 0),
    status,
    message || ""
  ]);
}

function getPitchLogSpreadsheet_() {
  return SpreadsheetApp.openById(PITCH_LOG_SPREADSHEET_ID);
}

/* =========================
   GAMECHANGER SYNC + COUNT-BASED SUBMISSION CHECK
========================= */

function buildGameCards_(team, date) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const syncSheet = ss.getSheetByName("GC_Games_Sync");
  if (!syncSheet || syncSheet.getLastRow() < 2) {
    return {
      games: [],
      meta: {
        lastSynced: "Never",
        syncMode: "quiet period • daily source check",
        message: "No cached GameChanger games are available."
      }
    };
  }

  const values = syncSheet.getDataRange().getValues();
  const headers = values.shift();
  const col = name => headers.indexOf(name);
  const submittedIds = {};
  const gamesSheet = ss.getSheetByName("Games");
  if (gamesSheet && gamesSheet.getLastRow() > 1) {
    const gameValues = gamesSheet.getDataRange().getValues();
    const gameHeaders = gameValues.shift();
    const gcIdCol = gameHeaders.indexOf("GC_Game_ID");
    if (gcIdCol >= 0) gameValues.forEach(row => { if (row[gcIdCol]) submittedIds[String(row[gcIdCol])] = true; });
  }

  let latestSync = null;
  const allGames = values.map(row => {
    const syncValue = row[col("Sync_Time")];
    const syncDate = syncValue ? new Date(syncValue) : null;
    if (syncDate && !isNaN(syncDate.getTime()) && (!latestSync || syncDate > latestSync)) latestSync = syncDate;
    const gcGameId = String(row[col("GC_Game_ID")] || "");
    return {
      team: row[col("Team")] || "",
      gcTeamId: row[col("GC_Team_ID")] || "",
      gcGameId: gcGameId,
      date: normalizeDateStringGC_(row[col("Date")]),
      startTime: formatTimeForDisplayGC_(row[col("Start_Time")]),
      opponent: row[col("Opponent")] || "",
      gameStatus: row[col("Game_Status")] || "",
      teamScore: row[col("Team_Score")] === "" ? "" : row[col("Team_Score")],
      opponentScore: row[col("Opponent_Score")] === "" ? "" : row[col("Opponent_Score")],
      result: row[col("Result")] || "",
      submitted: !!submittedIds[gcGameId]
    };
  });
  const games = allGames.filter(game => (!team || game.team === team) && (!date || game.date === date));

  games.sort((a, b) => parseTimeMinutes_(a.startTime) - parseTimeMinutes_(b.startTime));
  games.forEach((game, index) => game.gameSequence = index + 1);
  const syncPolicy = getGameChangerSyncPolicy_();
  return {
    games: games,
    missingSubmissions: buildCoachMissingSubmissions_(ss, team, allGames, submittedIds),
    meta: {
      lastSynced: latestSync ? formatDateTime(latestSync) : "Never",
      syncMode: syncPolicy.label,
      sourceRefreshMinutes: syncPolicy.minutes,
      loadsFromCache: true
    }
  };
}

function buildCoachMissingSubmissions_(ss, team, allGames, submittedIds) {
  if (!team) return [];
  const checkSheet = ss.getSheetByName("GC_Submission_Check");
  if (!checkSheet || checkSheet.getLastRow() < 2) return [];

  const ignoredKeys = getIgnoredMissingKeys_();
  const values = checkSheet.getDataRange().getValues();
  const headers = values.shift();
  const teamCol = headers.indexOf("Team");
  const dateCol = headers.indexOf("Date");
  const missingCol = headers.indexOf("Missing_Count");
  const statusCol = headers.indexOf("Check_Status");
  const today = formatDateOnly(new Date());
  const results = [];

  values.forEach(row => {
    const rowTeam = row[teamCol] || "";
    const rowDate = normalizeDateStringGC_(row[dateCol]);
    const missingCount = Number(row[missingCol]) || 0;
    const status = row[statusCol] || "";
    const key = buildTeamDateKey_(rowTeam, rowDate);
    if (rowTeam !== team || !rowDate || rowDate >= today || missingCount <= 0) return;
    if (status !== "Missing Submission" || ignoredKeys[key]) return;

    const possibleGames = allGames.filter(game =>
      game.team === rowTeam &&
      game.date === rowDate &&
      String(game.gameStatus || "").toLowerCase() === "completed" &&
      !submittedIds[game.gcGameId]
    ).map(game => ({
      gcGameId: game.gcGameId,
      startTime: game.startTime,
      opponent: game.opponent,
      teamScore: game.teamScore,
      opponentScore: game.opponentScore,
      result: game.result
    }));

    // GC_Submission_Check refreshes with the background sync, while the game
    // picker refreshes immediately after a coach submits. Stable GC game IDs
    // let this notice reduce/clear immediately without another source sync.
    const effectiveMissingCount = Math.min(missingCount, possibleGames.length);
    if (effectiveMissingCount > 0) {
      results.push({ date: rowDate, missingCount: effectiveMissingCount, games: possibleGames });
    }
  });

  results.sort((a, b) => String(b.date).localeCompare(String(a.date)));
  return results.slice(0, 5);
}

function getGameChangerSyncPolicy_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const syncSheet = ss.getSheetByName("GC_Games_Sync");
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  let nearestDays = null;

  if (syncSheet && syncSheet.getLastRow() > 1) {
    const values = syncSheet.getDataRange().getValues();
    const headers = values.shift();
    const dateCol = headers.indexOf("Date");
    values.forEach(row => {
      const gameDate = parseSheetDate(row[dateCol]);
      if (!gameDate) return;
      const difference = Math.round((gameDate.getTime() - today.getTime()) / 86400000);
      if (nearestDays === null || Math.abs(difference) < Math.abs(nearestDays)) nearestDays = difference;
    });
  }

  if (nearestDays !== null && nearestDays >= -1 && nearestDays <= 1) {
    return { minutes: 60, label: "game window • hourly source refresh" };
  }
  if (nearestDays !== null && nearestDays >= -2 && nearestDays <= 7) {
    return { minutes: 120, label: "active week • source refresh every 2 hours" };
  }
  return { minutes: 1440, label: "quiet period • daily source check" };
}

function smartGameChangerSync() {
  const properties = PropertiesService.getScriptProperties();
  const policy = getGameChangerSyncPolicy_();
  const lastRunValue = properties.getProperty("GC_LAST_SOURCE_SYNC_AT");
  const lastRun = lastRunValue ? new Date(lastRunValue) : null;
  const minutesSince = lastRun && !isNaN(lastRun.getTime()) ? (new Date().getTime() - lastRun.getTime()) / 60000 : Infinity;
  let ran = false;
  if (minutesSince < policy.minutes) {
    return {
      ran: false,
      reason: "not_due",
      nextPolicy: policy,
      weeklyReport: maybeSendWeeklyReport_(new Date(), false)
    };
  }
  syncGameChangerAndSubmissionCheck();
  ran = true;
  return {
    ran: ran,
    reason: "due",
    nextPolicy: getGameChangerSyncPolicy_(),
    weeklyReport: maybeSendWeeklyReport_(new Date(), true)
  };
}

function installSmartGameChangerSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "smartGameChangerSync") ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger("smartGameChangerSync").timeBased().everyHours(1).create();
  return "Installed hourly policy check. GameChanger itself is contacted hourly near games, every 2 hours during active weeks, and daily in quiet periods. The first Sunday check after 8:00 PM Eastern sends the weekly report when active recipients exist.";
}

function isSmartSyncTriggerInstalled_() {
  return ScriptApp.getProjectTriggers().some(trigger => trigger.getHandlerFunction() === "smartGameChangerSync");
}

function syncGameChangerAndSubmissionCheck() {
  syncGameChangerGames();
  refreshSubmissionCheck();
  PropertiesService.getScriptProperties().setProperty("GC_LAST_SOURCE_SYNC_AT", new Date().toISOString());
}

function syncGameChangerGames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const teamsSheet = ss.getSheetByName("GC_Teams");
  const syncSheet = ss.getSheetByName("GC_Games_Sync");

  if (!teamsSheet || !syncSheet) {
    throw new Error("Missing GC_Teams or GC_Games_Sync sheet");
  }

  const teamRows = teamsSheet.getDataRange().getValues();
  if (teamRows.length < 2) {
    throw new Error("GC_Teams has no team rows");
  }

  const teamHeaders = teamRows.shift();
  const teamCol = teamHeaders.indexOf("Team");
  const idCol = teamHeaders.indexOf("GC_Team_ID");
  const activeCol = teamHeaders.indexOf("Active");

  const output = [[
    "Team",
    "GC_Team_ID",
    "GC_Game_ID",
    "Date",
    "Start_Time",
    "Opponent",
    "Game_Status",
    "Team_Score",
    "Opponent_Score",
    "Result",
    "Sync_Time"
  ]];

  const syncTime = new Date();

  teamRows.forEach(r => {
    const team = r[teamCol];
    const gcTeamId = r[idCol];
    const active = String(r[activeCol]).toUpperCase() === "TRUE";

    if (!team || !gcTeamId || !active) return;

    const games = fetchGameChangerGamesForTeam_(gcTeamId);

    games.forEach(g => {
      const startTs = g.start_ts || "";
      const dateOnly = normalizeDateStringGC_(startTs);
      const startTime = formatTimeForDisplayGC_(startTs);
      const opponent = g.opponent_team && g.opponent_team.name ? g.opponent_team.name : "";
      const gameStatus = g.game_status || "";
      const teamScore = g.score && g.score.team != null ? g.score.team : "";
      const opponentScore = g.score && g.score.opponent_team != null ? g.score.opponent_team : "";
      const result = buildResultTextGC_(gameStatus, teamScore, opponentScore);

      output.push([
        team,
        gcTeamId,
        g.id || "",
        dateOnly,
        startTime,
        opponent,
        gameStatus,
        teamScore,
        opponentScore,
        result,
        syncTime
      ]);
    });
  });

  syncSheet.clearContents();
  syncSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}

function refreshSubmissionCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gcSheet = ss.getSheetByName("GC_Games_Sync");
  const gamesSheet = ss.getSheetByName("Games");
  const checkSheet = ss.getSheetByName("GC_Submission_Check");

  if (!gcSheet || !gamesSheet || !checkSheet) {
    throw new Error("Missing GC_Games_Sync, Games, or GC_Submission_Check sheet");
  }

  const gcRows = gcSheet.getDataRange().getValues();
  const gamesRows = gamesSheet.getDataRange().getValues();

  const output = [[
    "Team",
    "Date",
    "GC_Completed_Games",
    "NTB_Submissions",
    "Missing_Count",
    "Extra_Count",
    "Check_Status",
    "GC_Game_Details"
  ]];

  if (gcRows.length < 2) {
    checkSheet.clearContents();
    checkSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
    return;
  }

  const compareStartDate = parseLocalDate(GC_COMPARE_START_DATE);

  const gcHeaders = gcRows.shift();
  const gamesHeaders = gamesRows.length > 0 ? gamesRows.shift() : [];

  const gcTeamCol = gcHeaders.indexOf("Team");
  const gcDateCol = gcHeaders.indexOf("Date");
  const gcStatusCol = gcHeaders.indexOf("Game_Status");
  const gcOppCol = gcHeaders.indexOf("Opponent");
  const gcResultCol = gcHeaders.indexOf("Result");
  const gcStartTimeCol = gcHeaders.indexOf("Start_Time");

  const gameTeamCol = gamesHeaders.indexOf("Team");
  const gameDateCol = gamesHeaders.indexOf("Date");

  const gcCounts = {};
  const submissionCounts = {};

  gcRows.forEach(r => {
    const status = String(r[gcStatusCol] || "").toLowerCase();
    if (status !== "completed") return;

    const team = r[gcTeamCol] || "";
    const date = normalizeDateStringGC_(r[gcDateCol]);
    if (!team || !date) return;

    const gameDateObj = parseLocalDate(date);
    if (gameDateObj < compareStartDate) return;

    const key = buildTeamDateKey_(team, date);

    if (!gcCounts[key]) {
      gcCounts[key] = {
        team: team,
        date: date,
        count: 0,
        details: []
      };
    }

    gcCounts[key].count += 1;

    const opp = r[gcOppCol] || "";
    const result = r[gcResultCol] || "";
    const startTime = r[gcStartTimeCol] || "";

    let detail = "";
    if (startTime) detail += startTime + " - ";
    if (result) {
      detail += result;
      if (opp) detail += " vs " + opp;
    } else {
      detail += opp || "Completed game";
    }

    gcCounts[key].details.push(detail);
  });

  if (gamesRows.length > 0 && gameTeamCol >= 0 && gameDateCol >= 0) {
    gamesRows.forEach(r => {
      const team = r[gameTeamCol] || "";
      const date = normalizeDateStringGC_(r[gameDateCol]);
      if (!team || !date) return;

      const gameDateObj = parseLocalDate(date);
      if (gameDateObj < compareStartDate) return;

      const key = buildTeamDateKey_(team, date);
      submissionCounts[key] = (submissionCounts[key] || 0) + 1;
    });
  }

  const keys = Object.keys(gcCounts).sort((a, b) => {
    const aObj = gcCounts[a];
    const bObj = gcCounts[b];
    return (bObj.date + bObj.team).localeCompare(aObj.date + aObj.team);
  });

  keys.forEach(key => {
    const gc = gcCounts[key];
    const ntb = submissionCounts[key] || 0;
    const missing = Math.max(gc.count - ntb, 0);
    const extra = Math.max(ntb - gc.count, 0);

    let status = "Complete";
    if (missing > 0) {
      status = "Missing Submission";
    } else if (extra > 0) {
      status = "Extra Submission";
    }

    output.push([
      gc.team,
      gc.date,
      gc.count,
      ntb,
      missing,
      extra,
      status,
      gc.details.join(" | ")
    ]);
  });

  checkSheet.clearContents();
  checkSheet.getRange(1, 1, output.length, output[0].length).setValues(output);
}

function fetchGameChangerGamesForTeam_(gcTeamId) {
  const url = "https://api.team-manager.gc.com/public/teams/" + encodeURIComponent(gcTeamId) + "/games";

  const response = UrlFetchApp.fetch(url, {
    method: "get",
    muteHttpExceptions: true,
    headers: {
      "Accept": "application/json"
    }
  });

  const code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error("GameChanger fetch failed for " + gcTeamId + " with HTTP " + code);
  }

  const data = JSON.parse(response.getContentText());
  return Array.isArray(data) ? data : [];
}

function normalizeDateStringGC_(value) {
  if (!value) return "";

  const d = new Date(value);
  if (isNaN(d.getTime())) {
    const parsed = parseSheetDate(value);
    return parsed ? formatDateOnly(parsed) : "";
  }

  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function formatTimeForDisplayGC_(value) {
  if (!value) return "";

  const text = String(value).trim();
  if (/^\d{1,2}:\d{2}\s*(AM|PM)$/i.test(text)) return text;

  const d = new Date(value);
  if (isNaN(d.getTime())) return "";

  return Utilities.formatDate(d, Session.getScriptTimeZone(), "h:mm a");
}

function buildResultTextGC_(gameStatus, teamScore, opponentScore) {
  if (String(gameStatus).toLowerCase() !== "completed") return "";

  const hasTeam = teamScore !== "" && teamScore != null;
  const hasOpp = opponentScore !== "" && opponentScore != null;

  if (!hasTeam || !hasOpp) return "Completed";

  const t = Number(teamScore);
  const o = Number(opponentScore);

  if (isNaN(t) || isNaN(o)) return "Completed";

  if (t > o) return "W " + t + "-" + o;
  if (t < o) return "L " + t + "-" + o;
  return "T " + t + "-" + o;
}
