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

async function handleApi(req,res,url){
  if(req.method === "POST"){
    try{
      const data = JSON.parse(await readBody(req));
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
  if(url.searchParams.has("gcGameCards")){
    const team = url.searchParams.get("teamName") || "";
    const date = url.searchParams.get("date") || "";
    const submittedGameIds = new Set(Array.from(submissions.values()).map(item=>item.gcGameId).filter(Boolean));
    value = {
      games:mockGames().filter(game=>game.team === team && game.date === date).map(game=>
        Object.assign({},game,{submitted:submittedGameIds.has(game.gcGameId)})
      ),
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
    value = [];
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
