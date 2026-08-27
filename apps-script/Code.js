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

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const params = (e && e.parameter) || {};
  const callback = params.callback || "dashboardCallback";

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
    const gamesSheet = ss.getSheetByName("Games");

    if (!gamesSheet) {
      return outputJsonp(callback, []);
    }

    const team = params.notesTeam;
    const rows = gamesSheet.getDataRange().getValues();

    if (rows.length <= 1) {
      return outputJsonp(callback, []);
    }

    const headers = rows.shift();

    const dateCol = headers.indexOf("Date");
    const teamCol = headers.indexOf("Team");
    const opponentCol = headers.indexOf("Opponent");
    const notesCol = headers.indexOf("Notes");

    const results = rows
      .filter(r => r[teamCol] === team)
      .map(r => ({
        date: normalizeDateString(r[dateCol]),
        opponent: r[opponentCol] || "",
        notes: r[notesCol] || ""
      }))
      .sort((a, b) => String(b.date).localeCompare(String(a.date)));

    return outputJsonp(callback, results);
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
      startTime: r[timeCol] || "",
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
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const data = JSON.parse(e.postData.contents);
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

  const currentWeekStart = getMondayOfWeek_(today);
  const currentWeekEnd = new Date(currentWeekStart);
  currentWeekEnd.setDate(currentWeekEnd.getDate() + 6);

  const previousWeekStart = new Date(currentWeekStart);
  previousWeekStart.setDate(previousWeekStart.getDate() - 7);

  const previousWeekEnd = new Date(previousWeekStart);
  previousWeekEnd.setDate(previousWeekEnd.getDate() + 6);

  return {
    current: {
      startDate: currentWeekStart,
      endDate: currentWeekEnd
    },
    previous: {
      startDate: previousWeekStart,
      endDate: previousWeekEnd
    }
  };
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
      startTime: row[col("Start_Time")] || "",
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
  if (minutesSince < policy.minutes) {
    return { ran: false, reason: "not_due", nextPolicy: policy };
  }
  syncGameChangerAndSubmissionCheck();
  return { ran: true, reason: "due", nextPolicy: getGameChangerSyncPolicy_() };
}

function installSmartGameChangerSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "smartGameChangerSync") ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger("smartGameChangerSync").timeBased().everyHours(1).create();
  return "Installed hourly policy check. GameChanger itself is contacted hourly near games, every 2 hours during active weeks, and daily in quiet periods.";
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
