import http from "node:http";
import fs from "node:fs";
import path from "node:path";
import {fileURLToPath} from "node:url";
import {createRequire} from "node:module";

const require = createRequire(import.meta.url);
const PitchRules = require("./pitch-rules.js");
const root = path.dirname(fileURLToPath(import.meta.url));
const port = Number(process.env.NYB_DEV_PORT || 8787);
const host = "127.0.0.1";
const submissions = new Map();
const history = [];
const siteVisits = [];
let lastSynced = new Date();

function localDate(offsetDays = 0){
  const date = new Date();
  date.setDate(date.getDate() + offsetDays);
  return date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2,"0") + "-" + String(date.getDate()).padStart(2,"0");
}

const teams = ["12U Gold","11U Black","11U Gold","10U Black","10U Gold","8U Black","8U Gold","7U Black"];
const opponents = ["Westfield Rocks","Fishers Tigers","Carmel Pups"];

function mockGames(){
  const games = [];
  teams.forEach((team, teamIndex)=>{
    [
      {startTime:"5:30 PM",opponent:"Brownsburg Bulldogs",teamScore:7,opponentScore:4,result:"W 7-4"},
      {startTime:"8:00 PM",opponent:"Avon Orioles",teamScore:3,opponentScore:5,result:"L 3-5"}
    ].forEach((completed,index)=>{
      games.push({
        team,
        gcTeamId:"local-team-" + teamIndex,
        gcGameId:`local-${localDate(-1)}-${teamIndex}-${index + 1}`,
        date:localDate(-1),
        startTime:completed.startTime,
        gameSequence:index + 1,
        opponent:completed.opponent,
        gameStatus:"completed",
        teamScore:completed.teamScore,
        opponentScore:completed.opponentScore,
        result:completed.result,
        submitted:false
      });
    });
    ["10:00 AM","2:00 PM","5:30 PM"].forEach((startTime,index)=>{
      games.push({
        team,
        gcTeamId:"local-team-" + teamIndex,
        gcGameId:`local-${localDate()}-${teamIndex}-${index + 1}`,
        date:localDate(),
        startTime,
        gameSequence:index + 1,
        opponent:opponents[index],
        gameStatus:"scheduled",
        submitted:false
      });
    });
    games.push({
      team,
      gcTeamId:"local-team-" + teamIndex,
      gcGameId:`local-${localDate(1)}-${teamIndex}-1`,
      date:localDate(1),
      startTime:"6:00 PM",
      gameSequence:1,
      opponent:"Zionsville Eagles",
      gameStatus:"scheduled",
      submitted:false
    });
  });
  return games;
}

function sendJson(res, value, status = 200){
  const body = JSON.stringify(value);
  res.writeHead(status, {"Content-Type":"application/json; charset=utf-8","Cache-Control":"no-store"});
  res.end(body);
}

function sendJsonp(res, callback, value){
  const safeCallback = /^[A-Za-z_$][\w$]*$/.test(callback || "") ? callback : "callback";
  const body = `${safeCallback}(${JSON.stringify(value)})`;
  res.writeHead(200, {"Content-Type":"application/javascript; charset=utf-8","Cache-Control":"no-store"});
  res.end(body);
}

function readBody(req){
  return new Promise((resolve,reject)=>{
    let body = "";
    req.on("data", chunk=>{
      body += chunk;
      if(body.length > 1_000_000) reject(new Error("Request too large"));
    });
    req.on("end",()=>resolve(body));
    req.on("error",reject);
  });
}

function overallLevel(pitchers){
  if(pitchers.some(p=>p.level === "violation")) return "violation";
  if(pitchers.some(p=>p.level === "warning")) return "warning";
  return "clear";
}

function mockTrafficSummary(){
  const dailyViews = [2,4,3,6,5,0,8,4,7,9,3,6,11,8];
  const daily = dailyViews.map((views,index)=>({
    date:localDate(index - 13),
    views,
    sessions:Math.max(0,Math.ceil(views * .7))
  }));
  return {
    generatedAt:new Date().toLocaleString(),
    today:{views:8,sessions:6},
    sevenDays:{views:48,sessions:31},
    thirtyDays:{views:126,sessions:74},
    allTime:{views:126,sessions:74},
    topPages:[
      {label:"Pitch Log",views:79,sessions:51},
      {label:"Drills",views:29,sessions:18},
      {label:"Fall Practice",views:18,sessions:13}
    ],
    devices:[{label:"Mobile",count:88},{label:"Desktop",count:31},{label:"Tablet",count:7}],
    browsers:[{label:"Safari",count:72},{label:"Chrome",count:45},{label:"Edge",count:9}],
    referrers:[{label:"Direct / unknown",count:94},{label:"Internal",count:25},{label:"groupme.com",count:7}],
    daily
  };
}

function mockTeamNotes(team){
  const games = mockGames().filter(game=>game.team === team && game.gameStatus === "completed").map((game,index)=>
    Object.assign({},game,{
      notes:index === 0 ? "Strong pitching and good situational hitting. Defense stayed composed late." : "",
      hasNotes:index === 0,
      notesMissing:index !== 0,
      hasSubmission:index === 0,
      matchStatus:index === 0 ? "matched" : "gamechanger_only",
      submissionTime:index === 0 ? new Date().toLocaleString() : ""
    })
  );
  return {
    team,
    startDate:"",
    endDate:localDate(),
    lastSynced:lastSynced.toLocaleString(),
    summary:{wins:1,losses:1,ties:0,completedGames:2,runsScored:10,runsAllowed:9,missingNotes:1},
    games
  };
}

function mockWeeklyReportPreview(){
  return {
    recipientCount:0,
    automationInstalled:true,
    report:{startDate:localDate(-5),endDate:localDate(1)},
    html:`<div style="margin:0;padding:22px 10px;background:#f4f1ea;font-family:Arial;color:#171717;"><div style="max-width:680px;margin:auto;"><div style="background:#050505;border-radius:10px;padding:18px;text-align:center;"><img src="nyb-logo.png" width="82" height="82" style="border-radius:50%;border:2px solid #b8a05c;"><div style="color:#c8b46c;font-size:12px;font-weight:bold;margin-top:8px;">NOBLESVILLE TRAVEL BASEBALL</div><div style="color:#fff;font-size:24px;font-weight:bold;margin-top:5px;">Weekly Game Report</div><div style="color:#ddd5c0;margin-top:5px;">Monday–Sunday</div></div><div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:10px;"><div style="background:#fff;border:1px solid #ddd2ba;border-radius:8px;padding:13px;text-align:center;"><small>RECORD</small><div style="font-size:24px;font-weight:bold;">5-3</div></div><div style="background:#fff;border:1px solid #ddd2ba;border-radius:8px;padding:13px;text-align:center;"><small>COMPLETED GAMES</small><div style="font-size:24px;font-weight:bold;">8</div></div></div><div style="background:#171717;color:#fff;border-radius:8px 8px 0 0;padding:12px 14px;margin-top:18px;font-size:18px;font-weight:bold;">${teamSafeHtml("11U Black")}</div><div style="background:#fff;border:1px solid #ddd2ba;padding:14px;">W 7-4 vs Brownsburg Bulldogs<br><div style="margin-top:8px;color:#444;">Strong pitching and good situational hitting.</div></div></div></div>`
  };
}

function teamSafeHtml(value){
  return String(value).replace(/[&<>"']/g,char=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[char]));
}

async function handleApi(req,res,url){
  if(req.method === "POST"){
    try{
      const data = JSON.parse(await readBody(req));
      if(data.eventType === "site_visit"){
        siteVisits.push(Object.assign({timestamp:new Date().toISOString()},data));
        sendJson(res,{result:"success"});
        return;
      }
      if(data.submissionId && submissions.has(data.submissionId)){
        sendJson(res, submissions.get(data.submissionId));
        return;
      }
      const submittedAt = new Date().toISOString();
      const pitchers = (data.pitchers || []).map(pitcher=>{
        const current = {
          team:data.team,
          pitcher:pitcher.name,
          date:data.date,
          pitches:Number(pitcher.pitches) || 0,
          gameStartTime:data.gameStartTime || "",
          submittedAt
        };
        const evaluation = PitchRules.evaluatePitcherSubmission(history,current);
        history.push(current);
        return Object.assign({name:pitcher.name,pitches:current.pitches},evaluation);
      });
      const result = {
        result:"success",
        submissionId:data.submissionId,
        gameId:Date.now(),
        date:data.date,
        team:data.team,
        opponent:data.opponent,
        gcGameId:data.gcGameId || "",
        level:overallLevel(pitchers),
        pitchers,
        confirmedAt:submittedAt
      };
      submissions.set(data.submissionId,result);
      sendJson(res,result);
    }catch(error){
      sendJson(res,{result:"error",message:error.message},400);
    }
    return;
  }

  const callback = url.searchParams.get("callback") || "callback";
  let value = {};
  if(url.searchParams.has("weeklyReportPreview")){
    value = mockWeeklyReportPreview();
  }else if(url.searchParams.has("siteAnalyticsSummary")){
    value = mockTrafficSummary();
  }else if(url.searchParams.has("drills")){
    value = JSON.parse(fs.readFileSync(path.join(root,"drills.json"),"utf8"));
  }else if(url.searchParams.has("gcGameCards")){
    const team = url.searchParams.get("teamName") || "";
    const date = url.searchParams.get("date") || "";
    const submittedGameIds = new Set(Array.from(submissions.values()).map(item=>item.gcGameId).filter(Boolean));
    const allGames = mockGames();
    const missingGames = allGames.filter(game=>
      game.team === team &&
      game.date < localDate() &&
      game.gameStatus === "completed" &&
      !submittedGameIds.has(game.gcGameId)
    );
    const missingByDate = {};
    missingGames.forEach(game=>{
      if(!missingByDate[game.date]) missingByDate[game.date] = {date:game.date,missingCount:0,games:[]};
      missingByDate[game.date].missingCount += 1;
      missingByDate[game.date].games.push(game);
    });
    value = {
      games:allGames.filter(game=>game.team === team && game.date === date).map(game=>
        Object.assign({},game,{submitted:submittedGameIds.has(game.gcGameId)})
      ),
      missingSubmissions:Object.values(missingByDate).sort((a,b)=>b.date.localeCompare(a.date)),
      meta:{
        lastSynced:lastSynced.toLocaleString(),
        syncMode:"local demo • active-window cache",
        nextBackgroundCheck:"Hourly; source sync only when due"
      }
    };
  }else if(url.searchParams.has("submissionResult")){
    const id = url.searchParams.get("submissionId") || "";
    value = submissions.has(id) ? {found:true,result:submissions.get(id)} : {found:false};
  }else if(url.searchParams.has("runGcSync")){
    lastSynced = new Date();
    value = {success:true,message:"Local GameChanger schedule refreshed",ranAt:lastSynced.toLocaleString()};
  }else if(url.searchParams.has("team")){
    value = [];
  }else if(url.searchParams.has("notesTeam")){
    value = mockTeamNotes(url.searchParams.get("notesTeam") || "");
  }
  sendJsonp(res,callback,value);
}

const mimeTypes = {
  ".html":"text/html; charset=utf-8",
  ".js":"application/javascript; charset=utf-8",
  ".json":"application/json; charset=utf-8",
  ".png":"image/png",
  ".jpg":"image/jpeg",
  ".jpeg":"image/jpeg"
};

const server = http.createServer(async (req,res)=>{
  const url = new URL(req.url,`http://${host}:${port}`);
  if(url.pathname === "/api"){
    await handleApi(req,res,url);
    return;
  }
  const requested = url.pathname === "/" ? "/index.html" : url.pathname;
  const filePath = path.resolve(root,"." + decodeURIComponent(requested));
  if(!filePath.startsWith(root + path.sep)){
    res.writeHead(403); res.end("Forbidden"); return;
  }
  fs.readFile(filePath,(error,data)=>{
    if(error){ res.writeHead(404); res.end("Not found"); return; }
    res.writeHead(200,{"Content-Type":mimeTypes[path.extname(filePath)] || "application/octet-stream","Cache-Control":"no-store"});
    res.end(data);
  });
});

server.listen(port,host,()=>{
  process.stdout.write(`NYB local test app: http://${host}:${port}\n`);
  process.stdout.write("Local submissions are held in memory and reset when this server stops.\n");
});
