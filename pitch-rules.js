(function(root, factory){
  const api = factory();
  if(typeof module !== "undefined" && module.exports) module.exports = api;
  root.PitchRules = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function(){
  function restDaysForPitches(pitches){
    const total = Number(pitches) || 0;
    if(total <= 20) return 0;
    if(total <= 35) return 1;
    if(total <= 50) return 2;
    if(total <= 65) return 3;
    return 4;
  }

  function addDays(dateString, days){
    const parts = String(dateString).split("-").map(Number);
    const date = new Date(parts[0], parts[1] - 1, parts[2]);
    date.setDate(date.getDate() + days);
    return date.getFullYear() + "-" +
      String(date.getMonth() + 1).padStart(2, "0") + "-" +
      String(date.getDate()).padStart(2, "0");
  }

  function teamAge(team){
    const match = String(team || "").match(/(\d{1,2})\s*U/i);
    return match ? Number(match[1]) : null;
  }

  function dailyMaximum(team){
    const age = teamAge(team);
    if(age === 9 || age === 10) return 75;
    if(age === 11 || age === 12) return 85;
    return null;
  }

  function groupDaily(history){
    const totals = {};
    (history || []).forEach(item => {
      if(!item || !item.date) return;
      totals[item.date] = (totals[item.date] || 0) + (Number(item.pitches) || 0);
    });
    return totals;
  }

  function timeMinutes(value){
    const match = String(value || "").trim().match(/^(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);
    if(!match) return 24 * 60;
    let hour = Number(match[1]);
    const minute = Number(match[2]);
    const period = String(match[3] || "").toUpperCase();
    if(period === "AM" && hour === 12) hour = 0;
    if(period === "PM" && hour !== 12) hour += 12;
    return hour * 60 + minute;
  }

  function evaluatePitcherSubmission(history, current){
    const previous = (history || []).filter(item =>
      item && item.team === current.team && item.pitcher === current.pitcher
    );
    const sameDay = previous.filter(item => item.date === current.date);
    const allSameDay = sameDay.concat([current]).sort((a,b) => {
      const timeCompare = timeMinutes(a.gameStartTime) - timeMinutes(b.gameStartTime);
      if(timeCompare) return timeCompare;
      return String(a.submittedAt || "").localeCompare(String(b.submittedAt || ""));
    });
    const dailyTotal = allSameDay.reduce((sum,item) => sum + (Number(item.pitches) || 0), 0);
    const restDays = restDaysForPitches(dailyTotal);
    const eligibleDate = addDays(current.date, restDays + 1);
    const alerts = [];

    if(allSameDay.length > 1){
      const firstGamePitches = Number(allSameDay[0].pitches) || 0;
      const isWarning = allSameDay.length === 2 && firstGamePitches <= 20;
      alerts.push({
        code:"same_day",
        severity:isWarning ? "warning" : "violation",
        message:isWarning
          ? `Same-day pitching warning: first game was ${firstGamePitches} pitches (20 or fewer).`
          : `Same-day pitching violation: ${allSameDay.length} appearances; first game was ${firstGamePitches} pitches.`
      });
    }

    const dailyTotals = groupDaily(previous);
    const earlierDates = Object.keys(dailyTotals).filter(date => date < current.date).sort();
    if(earlierDates.length){
      const priorDate = earlierDates[earlierDates.length - 1];
      const priorTotal = dailyTotals[priorDate];
      const priorEligible = addDays(priorDate, restDaysForPitches(priorTotal) + 1);
      if(current.date < priorEligible){
        alerts.push({
          code:"rest",
          severity:"violation",
          message:`Required rest was not complete. ${priorTotal} pitches on ${priorDate} made the pitcher eligible on ${priorEligible}.`
        });
      }
    }

    const maximum = dailyMaximum(current.team);
    if(maximum && dailyTotal > maximum){
      alerts.push({
        code:"daily_max",
        severity:"violation",
        message:`Daily maximum exceeded: ${dailyTotal} pitches (limit ${maximum}).`
      });
    }

    const pitchingDates = new Set(Object.keys(dailyTotals).concat([current.date]));
    const yesterday = addDays(current.date, -1);
    const twoDaysAgo = addDays(current.date, -2);
    const tomorrow = addDays(current.date, 1);
    const twoDaysAhead = addDays(current.date, 2);
    const threeDayViolation =
      (pitchingDates.has(twoDaysAgo) && pitchingDates.has(yesterday)) ||
      (pitchingDates.has(yesterday) && pitchingDates.has(tomorrow)) ||
      (pitchingDates.has(tomorrow) && pitchingDates.has(twoDaysAhead));
    if(threeDayViolation){
      alerts.push({
        code:"three_days",
        severity:"violation",
        message:"Pitcher has appeared on three consecutive days."
      });
    }

    const level = alerts.some(a => a.severity === "violation")
      ? "violation"
      : alerts.some(a => a.severity === "warning") ? "warning" : "clear";

    return {dailyTotal, restDays, eligibleDate, level, alerts};
  }

  return {restDaysForPitches, addDays, teamAge, dailyMaximum, evaluatePitcherSubmission};
});
