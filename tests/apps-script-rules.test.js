const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");

const code = fs.readFileSync(require.resolve("../apps-script/Code.js"),"utf8");
const context = {
  console,
  Date,
  JSON,
  Math,
  Number,
  String,
  Object,
  Array,
  RegExp,
  isNaN,
  Utilities:{
    formatDate(date,timezone,pattern){
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2,"0");
      const day = String(date.getDate()).padStart(2,"0");
      if(pattern === "yyyy-MM-dd") return `${year}-${month}-${day}`;
      if(pattern === "h:mm a"){
        const hour = date.getHours();
        return `${hour % 12 || 12}:${String(date.getMinutes()).padStart(2,"0")} ${hour >= 12 ? "PM" : "AM"}`;
      }
      return `${year}-${month}-${day} 12:00:00`;
    }
  },
  Session:{getScriptTimeZone(){ return "America/Indiana/Indianapolis"; }}
};
vm.createContext(context);
vm.runInContext(code,context);

function outing(overrides = {}){
  return Object.assign({
    team:"11U Black",
    pitcher:"Test Pitcher",
    gameDate:new Date(2026,5,1),
    pitches:20,
    gameStartTime:"10:00 AM",
    gameSequence:1,
    submittedAt:new Date(2026,5,1,11)
  },overrides);
}

test("Apps Script uses the cumulative daily total for rest",()=>{
  const result = context.evaluatePitchSubmission_([outing()],outing({pitches:16,gameStartTime:"2:00 PM",gameSequence:2}));
  assert.equal(result.dailyTotal,36);
  assert.equal(result.restDays,2);
  assert.equal(result.eligibleDate,"2026-06-04");
});

test("Apps Script emits warning then violation same-day levels",()=>{
  const warning = context.evaluatePitchSubmission_([outing()],outing({pitches:10,gameStartTime:"2:00 PM",gameSequence:2}));
  assert.equal(warning.level,"warning");
  const violation = context.evaluatePitchSubmission_([outing({pitches:21})],outing({pitches:5,gameStartTime:"2:00 PM",gameSequence:2}));
  assert.equal(violation.level,"violation");
});

test("Apps Script formats Google Sheets time-valued cells for game cards",()=>{
  assert.equal(context.formatTimeForDisplayGC_(new Date(1899,11,30,16,0)),"4:00 PM");
  assert.equal(context.formatTimeForDisplayGC_("6:15 PM"),"6:15 PM");
});
