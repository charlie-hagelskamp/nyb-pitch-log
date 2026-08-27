const test = require("node:test");
const assert = require("node:assert/strict");
const rules = require("../pitch-rules.js");

function outing(overrides = {}){
  return Object.assign({team:"11U Black",pitcher:"Test Pitcher",date:"2026-06-01",pitches:20,gameStartTime:"10:00 AM"},overrides);
}

test("rest is based on cumulative pitches for the day",()=>{
  const result = rules.evaluatePitcherSubmission([outing()],outing({pitches:16,gameStartTime:"2:00 PM"}));
  assert.equal(result.dailyTotal,36);
  assert.equal(result.restDays,2);
  assert.equal(result.eligibleDate,"2026-06-04");
});

test("second appearance is warning when first game is 20 or fewer",()=>{
  const result = rules.evaluatePitcherSubmission([outing({pitches:20})],outing({pitches:10,gameStartTime:"2:00 PM"}));
  assert.equal(result.level,"warning");
  assert.equal(result.alerts[0].code,"same_day");
});

test("second appearance is violation when first game is over 20",()=>{
  const result = rules.evaluatePitcherSubmission([outing({pitches:21})],outing({pitches:5,gameStartTime:"2:00 PM"}));
  assert.equal(result.level,"violation");
  assert.equal(result.alerts[0].severity,"violation");
});

test("third appearance is always a violation",()=>{
  const history = [outing({pitches:10}),outing({pitches:5,gameStartTime:"2:00 PM"})];
  const result = rules.evaluatePitcherSubmission(history,outing({pitches:1,gameStartTime:"5:30 PM"}));
  assert.equal(result.level,"violation");
});

test("prior-day rest uses that day's total",()=>{
  const history = [
    outing({date:"2026-06-01",pitches:20}),
    outing({date:"2026-06-01",pitches:16,gameStartTime:"2:00 PM"})
  ];
  const result = rules.evaluatePitcherSubmission(history,outing({date:"2026-06-03",pitches:5}));
  assert.equal(result.alerts.some(a=>a.code === "rest"),true);
});

test("three consecutive days creates a violation",()=>{
  const history = [outing({date:"2026-06-01"}),outing({date:"2026-06-02"})];
  const result = rules.evaluatePitcherSubmission(history,outing({date:"2026-06-03"}));
  assert.equal(result.alerts.some(a=>a.code === "three_days"),true);
});

test("11U daily maximum is 85",()=>{
  const result = rules.evaluatePitcherSubmission([],outing({pitches:86}));
  assert.equal(result.alerts.some(a=>a.code === "daily_max"),true);
});
