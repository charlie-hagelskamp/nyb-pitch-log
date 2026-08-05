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
  const data = JSON.parse(e.postData.contents);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const gamesSheet = ss.getSheetByName("Games");
  const pitchingSheet = ss.getSheetByName("Pitching");

  if (!gamesSheet || !pitchingSheet) {
    throw new Error("Missing Games or Pitching sheet");
  }

  const gameDate = parseLocalDate(data.date);
  const submissionTime = new Date();
  const gameID = new Date().getTime();

  gamesSheet.appendRow([
    gameID,
    formatDateOnly(gameDate),
    data.team || "",
    data.opponent || "",
    data.notes || "",
    submissionTime
  ]);

  const submittedPitchers = Array.isArray(data.pitchers) ? data.pitchers : [];

  if (submittedPitchers.length === 0) {
    sendSubmissionEmail(data, gameID, gameDate, submissionTime, [], []);
    return jsonOutput({ result: "success" });
  }

  const pitchingRows = pitchingSheet.getDataRange().getValues();
  const headers = pitchingRows.shift();

  const teamCol = headers.indexOf("Team");
  const pitcherCol = headers.indexOf("Pitcher");
  const eligibleCol = headers.indexOf("EligibleDate");
  const dateCol = headers.indexOf("Date");
  const pitchesCol = headers.indexOf("Pitches");

  const allOutings = pitchingRows
    .map(r => rowToOuting(r, teamCol, pitcherCol, eligibleCol, dateCol, pitchesCol))
    .filter(Boolean);

  const violations = [];
  const submittedSummary = [];

  submittedPitchers.forEach(p => {
    const pitches = Number(p.pitches);
    if (isNaN(pitches) || pitches < 0 || !p.name) return;

    const restDays = calcRestDays(pitches);

    const eligibleDate = new Date(gameDate);
    eligibleDate.setDate(eligibleDate.getDate() + restDays + 1);
    eligibleDate.setHours(0, 0, 0, 0);

    const newOuting = {
      gameID: gameID,
      team: data.team,
      pitcher: p.name,
      gameDate: new Date(gameDate),
      pitches: pitches,
      restDays: restDays,
      eligibleDate: new Date(eligibleDate),
      opponent: data.opponent || "N/A",
      isNewSubmission: true
    };

    const sameDayOutings = allOutings.filter(o =>
      o.team === data.team &&
      o.pitcher === p.name &&
      datesEqual(o.gameDate, gameDate)
    );

    if (sameDayOutings.length > 0) {
      const outingsForEmail = sameDayOutings
        .map(o => ({
          date: new Date(o.gameDate),
          pitches: o.pitches,
          eligibleDate: new Date(o.eligibleDate),
          opponent: o.opponent || "N/A"
        }))
        .concat([{
          date: new Date(newOuting.gameDate),
          pitches: newOuting.pitches,
          eligibleDate: new Date(newOuting.eligibleDate),
          opponent: newOuting.opponent
        }]);

      violations.push({
        type: "same_day",
        team: data.team,
        pitcher: p.name,
        date: new Date(gameDate),
        outings: outingsForEmail
      });
    }

    const mostRecentEarlier = findMostRecentEarlierOuting(
      allOutings,
      data.team,
      p.name,
      gameDate
    );

    if (mostRecentEarlier && gameDate < mostRecentEarlier.eligibleDate) {
      violations.push({
        type: "rest",
        team: data.team,
        pitcher: p.name,
        sourceOuting: {
          date: new Date(mostRecentEarlier.gameDate),
          pitches: mostRecentEarlier.pitches,
          eligibleDate: new Date(mostRecentEarlier.eligibleDate),
          opponent: mostRecentEarlier.opponent || "N/A"
        },
        violatingOuting: {
          date: new Date(newOuting.gameDate),
          pitches: newOuting.pitches,
          eligibleDate: new Date(newOuting.eligibleDate),
          opponent: newOuting.opponent
        }
      });
    }

    const laterOutings = allOutings.filter(o =>
      o.team === data.team &&
      o.pitcher === p.name &&
      o.gameDate > gameDate
    );

    laterOutings.forEach(later => {
      if (later.gameDate < newOuting.eligibleDate) {
        violations.push({
          type: "rest",
          team: data.team,
          pitcher: p.name,
          sourceOuting: {
            date: new Date(newOuting.gameDate),
            pitches: newOuting.pitches,
            eligibleDate: new Date(newOuting.eligibleDate),
            opponent: newOuting.opponent
          },
          violatingOuting: {
            date: new Date(later.gameDate),
            pitches: later.pitches,
            eligibleDate: new Date(later.eligibleDate),
            opponent: later.opponent || "N/A"
          }
        });
      }
    });

    pitchingSheet.appendRow([
      gameID,
      formatDateOnly(gameDate),
      data.team,
      p.name,
      pitches,
      restDays,
      formatDateOnly(eligibleDate)
    ]);

    allOutings.push({
      gameID: gameID,
      team: data.team,
      pitcher: p.name,
      gameDate: new Date(gameDate),
      pitches: pitches,
      restDays: restDays,
      eligibleDate: new Date(eligibleDate),
      opponent: data.opponent || "N/A",
      isNewSubmission: true
    });

    submittedSummary.push({
      name: p.name,
      pitches: pitches,
      restDays: restDays,
      eligibleDate: new Date(eligibleDate)
    });
  });

  sendSubmissionEmail(data, gameID, gameDate, submissionTime, submittedSummary, violations);

  if (violations.length > 0) {
    sendViolationEmail(violations);
  }

  return jsonOutput({ result: "success" });
}

function sendSubmissionEmail(data, gameID, gameDate, submissionTime, submittedPitchers, violations) {
  const subject = "NYB Game Submission";

  let body =
    `A game submission was recorded.\n\n` +
    `GameID: ${gameID}\n` +
    `Date: ${formatDateOnly(gameDate)}\n` +
    `Team: ${data.team || "N/A"}\n` +
    `Opponent: ${data.opponent || "N/A"}\n` +
    `Submission Time: ${formatDateTime(submissionTime)}\n\n` +
    `Notes:\n${data.notes || "N/A"}\n\n`;

  body += `Pitchers Submitted:\n`;

  if (submittedPitchers.length === 0) {
    body += `None\n`;
  } else {
    submittedPitchers.forEach(p => {
      body +=
        `• ${p.name}\n` +
        `  Pitch Count: ${p.pitches}\n` +
        `  Rest Days: ${p.restDays}\n` +
        `  Eligible Again: ${formatDateOnly(p.eligibleDate)}\n`;
    });
  }

  body += `\nViolation Count: ${violations.length}\n`;

  MailApp.sendEmail(ADMIN_EMAIL, subject, body);
}

function sendViolationEmail(violations) {
  const subject = "NYB PitchSmart Violation";

  let body = "PitchSmart violation(s) detected.\n\n";

  violations.forEach((v, i) => {
    body += `Violation ${i + 1}\n`;
    body += `Pitcher: ${v.pitcher}\n`;
    body += `Team: ${v.team}\n`;

    if (v.type === "same_day") {
      body += `Issue: Pitcher appeared more than 1 time on the same day\n`;
      body += `Date: ${formatDateOnly(v.date)}\n`;
      body += `Outings that day:\n`;

      v.outings
        .sort((a, b) => a.pitches - b.pitches)
        .forEach(o => {
          body +=
            `• Pitch Count: ${o.pitches}\n` +
            `  Next Eligible: ${formatDateOnly(o.eligibleDate)}\n`;
        });

    } else if (v.type === "rest") {
      body +=
        `Required rest was not satisfied between these outings:\n\n` +
        `Outing that started rest:\n` +
        `• Date: ${formatDateOnly(v.sourceOuting.date)}\n` +
        `• Pitch Count: ${v.sourceOuting.pitches}\n` +
        `• Next Eligible: ${formatDateOnly(v.sourceOuting.eligibleDate)}\n\n` +
        `Outing that violated rest:\n` +
        `• Date: ${formatDateOnly(v.violatingOuting.date)}\n` +
        `• Pitch Count: ${v.violatingOuting.pitches}\n`;
    }

    body += `\n--------------------------\n\n`;
  });

  MailApp.sendEmail(ADMIN_EMAIL, subject, body);
}

function rowToOuting(r, teamCol, pitcherCol, eligibleCol, dateCol, pitchesCol) {
  const team = r[teamCol];
  const pitcher = r[pitcherCol];
  const gameDate = parseSheetDate(r[dateCol]);
  const eligibleDate = parseSheetDate(r[eligibleCol]);
  const pitches = Number(r[pitchesCol]) || 0;

  if (!team || !pitcher || !gameDate || !eligibleDate) return null;

  return {
    team: team,
    pitcher: pitcher,
    gameDate: gameDate,
    eligibleDate: eligibleDate,
    pitches: pitches,
    opponent: "N/A",
    isNewSubmission: false
  };
}

function findMostRecentEarlierOuting(outings, team, pitcher, gameDate) {
  let latest = null;

  outings.forEach(o => {
    if (o.team !== team || o.pitcher !== pitcher) return;
    if (o.gameDate >= gameDate) return;

    if (!latest || o.gameDate > latest.gameDate) {
      latest = o;
    }
  });

  return latest;
}

function calcRestDays(pitches) {
  if (pitches <= 20) return 0;
  if (pitches <= 35) return 1;
  if (pitches <= 50) return 2;
  if (pitches <= 65) return 3;
  return 4;
}

function outputJsonp(callback, data) {
  return ContentService
    .createTextOutput(`${callback}(${JSON.stringify(data)})`)
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

function syncGameChangerAndSubmissionCheck() {
  syncGameChangerGames();
  refreshSubmissionCheck();
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