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

test("Apps Script reads drills from the Drills tab",()=>{
  const rows = [
    ["Title","Category","Focus","Source","URL","Notes"],
    ["Infield EveryDays","Infield","Fundamentals","YouTube","https://example.com/1",""],
    ["","Hitting","Ignored blank title","YouTube","https://example.com/2",""]
  ];
  context.SpreadsheetApp = {
    getActiveSpreadsheet(){
      return {
        getSheetByName(name){
          assert.equal(name,"Drills");
          return {
            getLastRow(){ return rows.length; },
            getRange(row,column,rowCount,columnCount){
              assert.deepEqual([row,column,rowCount,columnCount],[1,1,3,6]);
              return {getDisplayValues(){ return rows; }};
            }
          };
        }
      };
    }
  };

  const drills = context.buildDrills_();
  assert.equal(drills.length,1);
  assert.equal(drills[0].title,"Infield EveryDays");
  assert.equal(drills[0].source,"YouTube");
});

test("Apps Script summarizes anonymous traffic by rolling windows",()=>{
  const rows = [
    [new Date(2026,7,30,9),"Pitch Log","/index.html","Direct / unknown","Mobile","Safari","iOS","390x844","390x700","en-US","America/Indiana/Indianapolis","session-a"],
    [new Date(2026,7,30,10),"Drills","/drills.html","Internal","Mobile","Safari","iOS","390x844","390x700","en-US","America/Indiana/Indianapolis","session-a"],
    [new Date(2026,7,25,18),"Pitch Log","/index.html","google.com","Desktop","Chrome","Windows","1920x1080","1280x720","en-US","America/Indiana/Indianapolis","session-b"],
    [new Date(2026,6,30,12),"Pitch Log","/index.html","Direct / unknown","Desktop","Edge","Windows","1920x1080","1280x720","en-US","America/Indiana/Indianapolis","session-c"]
  ];

  const summary = context.buildSiteAnalyticsSummaryFromRows_(rows,new Date(2026,7,30,20));
  assert.deepEqual(JSON.parse(JSON.stringify(summary.today)),{views:2,sessions:1});
  assert.deepEqual(JSON.parse(JSON.stringify(summary.sevenDays)),{views:3,sessions:2});
  assert.deepEqual(JSON.parse(JSON.stringify(summary.thirtyDays)),{views:3,sessions:2});
  assert.deepEqual(JSON.parse(JSON.stringify(summary.allTime)),{views:4,sessions:3});
  assert.equal(summary.topPages[0].label,"Pitch Log");
  assert.equal(summary.topPages[0].views,2);
  assert.equal(summary.daily.length,14);
  assert.deepEqual(JSON.parse(JSON.stringify(summary.daily[13])),{date:"2026-08-30",views:2,sessions:1});
});

test("Apps Script appends a sanitized site visit row",()=>{
  let appended;
  context.LockService = {
    getScriptLock(){
      return {waitLock(ms){ assert.equal(ms,5000); },releaseLock(){}};
    }
  };
  context.SpreadsheetApp = {
    getActiveSpreadsheet(){
      return {
        getSheetByName(name){
          assert.equal(name,"Site_Visits");
          return {
            getLastRow(){ return 1; },
            getRange(row,column,rowCount,columnCount){
              assert.deepEqual([row,column,rowCount,columnCount],[2,1,1,12]);
              return {setValues(values){ appended = values[0]; }};
            }
          };
        }
      };
    }
  };

  context.recordSiteVisit_({
    page:"P".repeat(100),
    path:"/index.html",
    referrer:"Direct / unknown",
    device:"Mobile",
    browser:"Safari",
    operatingSystem:"iOS",
    screen:"390x844",
    viewport:"390x700",
    language:"en-US",
    timezone:"America/Indiana/Indianapolis",
    sessionId:"session-test"
  });

  assert.equal(appended.length,12);
  assert.equal(appended[1].length,80);
  assert.equal(appended[4],"Mobile");
  assert.equal(appended[11],"session-test");
});
