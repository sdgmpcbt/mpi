
// === SDGMP SmartCBT - FINAL v5 - MATCH HEADER ASLI ===
const SPREADSHEET_ID = "1OINhkLqDdpGyYD5_vuOy9JIg9Ea8EuxKL6_ncc3xTK0";
const SHEET_SISWA = "Siswa";
const SHEET_GURU = "Guru";
const SHEET_QUIZ = "Quizzes"; // header kamu: timestamp | quizId | judul | mapel | kelas | jumlahSoal | json
const SHEET_ATTEMPT = "Attempts"; // header kamu: timestamp | attemptId | studentName | kelas | quizId | quizJudul | score | correct | total | date | answersJson
const SHEET_PROGRESS = "Progress"; // header kamu: timestamp | studentName | kelas | type | coinsReward | totalCoins | lastSpin | streak | bonusInfo

function getSheet(name){
  try{
    if(!SPREADSHEET_ID || SPREADSHEET_ID.indexOf("GANTI")>=0){
      // kalau ID belum diganti, coba buka sheet aktif (jika script ter-bound)
      try{ return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }catch(e){ return null; }
    }
    return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
  }catch(e){
    Logger.log("getSheet error "+name+": "+e);
    try{ return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }catch(e2){ return null; }
  }
}

function getHeaders(sheet){
  if(!sheet) return [];
  let vals = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  return vals.map(h=>h.toString().trim().toLowerCase());
}

// ===== SISWA & GURU =====
function readSiswaData(){
  let sh = getSheet(SHEET_SISWA);
  if(!sh) return [];
  let values = sh.getDataRange().getValues();
  let headers = values[0].map(h=>h.toString().trim().toLowerCase());
  let idxNama = headers.indexOf("nama");
  let idxKelas = headers.indexOf("kelas");
  let idxNisn = headers.indexOf("nisn");
  let idxNis = headers.indexOf("nis");
  if(idxNama<0) idxNama=0;
  if(idxKelas<0) idxKelas=1;
  if(idxNisn<0) idxNisn=2;
  let out=[];
  for(let i=1;i<values.length;i++){
    let r = values[i];
    if(!r[idxNama]) continue;
    out.push({
      nama: (r[idxNama]+"").trim(),
      kelas: (r[idxKelas]+"").trim() || "1",
      nisn: (r[idxNisn]+"").trim(),
      nis: idxNis>=0 ? (r[idxNis]+"").trim() : ""
    });
  }
  return out;
}

function readGuruData(){
  let sh = getSheet(SHEET_GURU);
  if(!sh) return [];
  let values = sh.getDataRange().getValues();
  let headers = values[0].map(h=>h.toString().trim().toLowerCase());
  let idxNama = headers.indexOf("nama");
  let idxKelas = headers.indexOf("kelas");
  let idxKey = headers.indexOf("key");
  let idxKunci = headers.indexOf("kunci");
  if(idxNama<0) idxNama=0;
  if(idxKelas<0) idxKelas=1;
  if(idxKey<0) idxKey = idxKunci>=0?idxKunci:2;
  let out=[];
  for(let i=1;i<values.length;i++){
    let r=values[i];
    if(!r[idxNama]) continue;
    out.push({
      nama: (r[idxNama]+"").trim(),
      kelas: (r[idxKelas]+"").trim() || "-",
      key: (r[idxKey]+"").trim()
    });
  }
  return out;
}

// ===== HELPER PROGRESS =====
function getLatestProgressForStudent(studentName){
  let sh = getSheet(SHEET_PROGRESS);
  if(!sh) return null;
  let vals = sh.getDataRange().getValues();
  let headers = vals[0].map(h=>h.toString().trim().toLowerCase());
  let idxStudent = headers.indexOf("studentname");
  let idxKelas = headers.indexOf("kelas");
  let idxTotal = headers.indexOf("totalcoins");
  let idxLastSpin = headers.indexOf("lastspin");
  let idxStreak = headers.indexOf("streak");
  let idxBonus = headers.indexOf("bonusinfo");
  let idxType = headers.indexOf("type");
  if(idxStudent<0) idxStudent=1;
  let latest=null;
  let shopItems=[];
  for(let i=1;i<vals.length;i++){
    let name = (vals[i][idxStudent]+"").trim();
    if(name.toLowerCase() !== studentName.trim().toLowerCase()) continue;
    // kumpulkan shop items
    let type = idxType>=0? (vals[i][idxType]+"").toLowerCase() : "";
    let bonus = idxBonus>=0? (vals[i][idxBonus]+"") : "";
    if(type.indexOf("shop")>=0 || type.indexOf("buy")>=0){
      if(bonus && shopItems.indexOf(bonus)==-1) shopItems.push(bonus);
    }
    latest = {
      row: i,
      kelas: vals[i][idxKelas]||"",
      totalCoins: Number(vals[i][idxTotal]||0),
      lastSpin: idxLastSpin>=0? (vals[i][idxLastSpin]+"") : "",
      streak: idxStreak>=0? Number(vals[i][idxStreak]||0) : 0,
      bonusInfo: bonus,
      timestamp: vals[i][0]
    };
  }
  if(latest){
    latest.shopItems = shopItems;
  }
  return latest;
}

function doGet(e){
  let actionRaw = (e.parameter.action||"").toString();
  let action = actionRaw.toLowerCase().trim();

  if(action==="getsiswapublic"){
    let data = readSiswaData().map(s=>({nama:s.nama, kelas:s.kelas}));
    return jsonResponse({status:"ok", siswa:data, count:data.length});
  }
  if(action==="getgurupublic"){
    let data = readGuruData().map(g=>({nama:g.nama, kelas:g.kelas}));
    return jsonResponse({status:"ok", guru:data, count:data.length});
  }
  if(action==="getquizzes"){
    let sh = getSheet(SHEET_QUIZ);
    let quizzes=[];
    let debug={sheetFound: sh?"yes":"no", header:[]};
    if(sh){
      let vals = sh.getDataRange().getValues();
      if(vals.length>0) debug.header = vals[0];
      let headers = vals[0].map(h=>h.toString().trim().toLowerCase());
      let idxJson = headers.indexOf("json");
      let idxQuizId = headers.indexOf("quizid");
      if(idxJson<0) idxJson = headers.length-1; // last col
      debug.idxJson=idxJson;
      debug.rows=vals.length-1;
      for(let i=1;i<vals.length;i++){
        if(!vals[i][idxQuizId] && !vals[i][idxJson]) continue;
        try{
          let jsonStr = vals[i][idxJson]||"";
          if(!jsonStr) continue;
          let obj = JSON.parse(jsonStr);
          // pastikan id konsisten dengan kolom quizId
          if(vals[i][idxQuizId] && !obj.id) obj.id = vals[i][idxQuizId];
          quizzes.push(obj);
        }catch(err){
          debug.parseError = (debug.parseError||"") + " row"+i+":"+err+";";
        }
      }
    }
    if(quizzes.length===0){
      try{
        let prop = PropertiesService.getScriptProperties().getProperty("quizzes");
        if(prop){
          quizzes = JSON.parse(prop);
          debug.fromProp = quizzes.length;
        }
      }catch(err){ debug.propError=err+""; }
    }
    return jsonResponse({status:"ok", quizzes:quizzes, count:quizzes.length, debug:debug});
  }
  // === ATTEMPTS ===
  if(["getattempts","getallattempts","get_attempts","get_all_attempts","getattempt","attempts"].indexOf(action)>=0){
    let sh = getSheet(SHEET_ATTEMPT);
    let attempts=[];
    let debug={sheetFound: sh?"yes":"no", header:[]};
    if(sh){
      let vals = sh.getDataRange().getValues();
      if(vals.length>0) debug.header=vals[0];
      let headers = vals[0].map(h=>h.toString().trim().toLowerCase());
      let idxAttemptId = headers.indexOf("attemptid");
      let idxStudent = headers.indexOf("studentname");
      let idxKelas = headers.indexOf("kelas");
      let idxQuizId = headers.indexOf("quizid");
      let idxQuizJudul = headers.indexOf("quizjudul");
      let idxScore = headers.indexOf("score");
      let idxCorrect = headers.indexOf("correct");
      let idxTotal = headers.indexOf("total");
      let idxDate = headers.indexOf("date");
      let idxAnswers = headers.indexOf("answersjson");
      debug.rows=vals.length-1;
      for(let i=1;i<vals.length;i++){
        let row = vals[i];
        if(!row[idxAttemptId] && !row[idxStudent]) continue;
        try{
          let attemptId = row[idxAttemptId]||"";
          let studentName = row[idxStudent]||"";
          let kelas = row[idxKelas]||"";
          let quizId = row[idxQuizId]||"";
          let quizJudul = row[idxQuizJudul]||"";
          let score = Number(row[idxScore]||0);
          let correct = Number(row[idxCorrect]||0);
          let total = Number(row[idxTotal]||0);
          let date = row[idxDate]||"";
          let answersJson = row[idxAnswers]||"";
          let answers=null;
          try{ answers = JSON.parse(answersJson); }catch{ answers = answersJson; }
          let obj={
            id: attemptId,
            attemptId: attemptId,
            studentName: studentName,
            nama: studentName,
            kelas: kelas,
            quizId: quizId,
            quizJudul: quizJudul,
            score: score,
            correct: correct,
            total: total,
            date: date,
            answers: answers,
            answersJson: answersJson
          };
          if(e.parameter.quizId && obj.quizId != e.parameter.quizId) continue;
          if(e.parameter.studentName && (obj.studentName||"").toLowerCase() != e.parameter.studentName.toLowerCase()) continue;
          attempts.push(obj);
        }catch(err){ debug.parseError = (debug.parseError||"") + " row"+i+":"+err+";"; }
      }
    }
    if(attempts.length===0){
      try{
        let prop = PropertiesService.getScriptProperties().getProperty("attempts");
        if(prop){
          let arr = JSON.parse(prop);
          attempts = attempts.concat(arr);
          debug.fromProp = arr.length;
        }
      }catch(err){ debug.propError=err+""; }
    }
    return jsonResponse({status:"ok", attempts:attempts, count:attempts.length, debug:debug});
  }
  // === PROGRESS SINGLE ===
  if(action==="getprogress"){
    let studentName = e.parameter.studentName||"";
    if(!studentName) return jsonResponse({status:"error", message:"studentName required"});
    let sh = getSheet(SHEET_PROGRESS);
    let result = {status:"ok", studentName:studentName, totalCoins:0, streak:0, lastSpin:"", shopItems:[], found:false};
    let debug={sheetFound: sh?"yes":"no"};
    if(sh){
      let vals = sh.getDataRange().getValues();
      debug.header = vals[0];
      let headers = vals[0].map(h=>h.toString().trim().toLowerCase());
      let idxStudent = headers.indexOf("studentname");
      let idxKelas = headers.indexOf("kelas");
      let idxTotal = headers.indexOf("totalcoins");
      let idxLastSpin = headers.indexOf("lastspin");
      let idxStreak = headers.indexOf("streak");
      let idxBonus = headers.indexOf("bonusinfo");
      let idxType = headers.indexOf("type");
      let idxCoinsReward = headers.indexOf("coinsreward");
      let latest=null;
      let shopItems=[];
      for(let i=1;i<vals.length;i++){
        let name = (vals[i][idxStudent]+"").trim();
        if(name.toLowerCase() !== studentName.trim().toLowerCase()) continue;
        let type = idxType>=0? (vals[i][idxType]+"").toLowerCase() : "";
        let bonus = idxBonus>=0? (vals[i][idxBonus]+"") : "";
        if(type.indexOf("shop")>=0 || type.indexOf("buy")>=0){
          if(bonus && shopItems.indexOf(bonus)==-1) shopItems.push(bonus);
        }
        latest={
          kelas: vals[i][idxKelas]||"",
          totalCoins: Number(vals[i][idxTotal]||0),
          lastSpin: idxLastSpin>=0? (vals[i][idxLastSpin]+"") : "",
          streak: idxStreak>=0? Number(vals[i][idxStreak]||0) : 0,
          bonusInfo: bonus,
          type: type,
          coinsReward: idxCoinsReward>=0? Number(vals[i][idxCoinsReward]||0):0,
          timestamp: vals[i][0]
        };
      }
      if(latest){
        result.kelas = latest.kelas;
        result.totalCoins = latest.totalCoins;
        result.lastSpin = latest.lastSpin;
        result.streak = latest.streak;
        result.bonusInfo = latest.bonusInfo;
        result.shopItems = shopItems;
        result.found=true;
        result.timestamp = latest.timestamp;
      }
      debug.rows=vals.length-1;
    }
    if(!result.found){
      try{
        let key = "progress_"+studentName.toLowerCase();
        let prop = PropertiesService.getScriptProperties().getProperty(key);
        if(prop){
          let p = JSON.parse(prop);
          result.totalCoins = p.totalCoins||0;
          result.streak = p.streak||0;
          result.lastSpin = p.lastSpin||"";
          result.shopItems = p.shopItems||[];
          result.fromProp=true;
          result.found=true;
        }
      }catch{}
    }
    result.debug=debug;
    return jsonResponse(result);
  }
  // === PROGRESS ALL ===
  if(action==="getallprogress" || action==="get_all_progress" || action==="getprogressall"){
    let sh = getSheet(SHEET_PROGRESS);
    let map={}; // nama lower -> latest
    let debug={sheetFound: sh?"yes":"no"};
    if(sh){
      let vals = sh.getDataRange().getValues();
      debug.header=vals[0];
      debug.rows=vals.length-1;
      let headers = vals[0].map(h=>h.toString().trim().toLowerCase());
      let idxStudent = headers.indexOf("studentname");
      let idxKelas = headers.indexOf("kelas");
      let idxTotal = headers.indexOf("totalcoins");
      let idxLastSpin = headers.indexOf("lastspin");
      let idxStreak = headers.indexOf("streak");
      let idxBonus = headers.indexOf("bonusinfo");
      let idxType = headers.indexOf("type");
      for(let i=1;i<vals.length;i++){
        let name = (vals[i][idxStudent]+"").trim();
        if(!name) continue;
        let lower = name.toLowerCase();
        let type = idxType>=0? (vals[i][idxType]+"").toLowerCase() : "";
        let bonus = idxBonus>=0? (vals[i][idxBonus]+"") : "";
        if(!map[lower]){
          map[lower]={nama:name, kelas: vals[i][idxKelas]||"", totalCoins:0, streak:0, lastSpin:"", shopItems:[]};
        }
        map[lower].totalCoins = Number(vals[i][idxTotal]|| map[lower].totalCoins);
        if(vals[i][idxLastSpin]) map[lower].lastSpin = vals[i][idxLastSpin];
        if(vals[i][idxStreak]) map[lower].streak = Number(vals[i][idxStreak]||0);
        map[lower].kelas = vals[i][idxKelas]|| map[lower].kelas;
        if(type.indexOf("shop")>=0 && bonus){
          if(map[lower].shopItems.indexOf(bonus)==-1) map[lower].shopItems.push(bonus);
        }
      }
    }
    let all = Object.values(map);
    return jsonResponse({status:"ok", progress:all, count:all.length, debug:debug});
  }
  if(action==="checkspin"){
    let studentName = e.parameter.studentName||"";
    let date = e.parameter.date||"";
    let latest = getLatestProgressForStudent(studentName);
    let already=false;
    let lastSpin = latest? latest.lastSpin : "";
    if(lastSpin && date && lastSpin===date) already=true;
    return jsonResponse({status:"ok", alreadySpun:already, lastSpin:lastSpin, studentName:studentName});
  }

  if(!action){ return jsonResponse({status:"ok", message:"SDGMP SmartCBT API aktif. Gunakan ?action=getQuizzes | getAllAttempts | getProgress?studentName=... | getAllProgress | getSiswaPublic | getGuruPublic | checkSpin", available:["getSiswaPublic","getGuruPublic","getQuizzes","getAllAttempts","getProgress","getAllProgress","checkSpin"], timestamp:new Date()}); } return jsonResponse({status:"error", message:"Unknown action: "+actionRaw, available:["getSiswaPublic","getGuruPublic","getQuizzes","getAllAttempts","getProgress?studentName=...","getAllProgress","checkSpin"]});
}

function doPost(e){
  let body="";
  try{ body = e.postData.contents; }catch{}
  let payload={};
  try{ payload = JSON.parse(body); }catch{ payload = e.parameter; }
  let actionRaw = payload.action||"";
  let action = actionRaw.toString().toLowerCase().trim();

  if(action==="loginsiswa"){
    let nama = (payload.nama||"").trim();
    let nisn = (payload.nisn||"").trim();
    if(!nama || !nisn) return jsonResponse({status:"error", message:"Nama & NISN wajib"});
    let siswa = readSiswaData().find(s=>s.nama.toLowerCase()===nama.toLowerCase());
    if(!siswa) return jsonResponse({status:"error", message:"Nama tidak ditemukan"});
    if(siswa.nisn===nisn || siswa.nis===nisn){
      return jsonResponse({status:"ok", nama:siswa.nama, kelas:siswa.kelas});
    }else{
      return jsonResponse({status:"error", message:"NISN tidak cocok"});
    }
  }
  if(action==="loginguru"){
    let nama = (payload.nama||"").trim();
    let key = (payload.key||"").trim();
    if(!nama || !key) return jsonResponse({status:"error", message:"Nama & Key wajib"});
    let guru = readGuruData().find(g=>g.nama.toLowerCase()===nama.toLowerCase());
    if(!guru) return jsonResponse({status:"error", message:"Nama guru tidak ditemukan"});
    if(guru.key.toUpperCase()===key.toUpperCase()){
      return jsonResponse({status:"ok", nama:guru.nama, kelas:guru.kelas});
    }else{
      return jsonResponse({status:"error", message:"Key guru tidak valid"});
    }
  }
  // === SAVE QUIZ - SESUAI HEADER: timestamp | quizId | judul | mapel | kelas | jumlahSoal | json ===
  if(action==="savequiz"){
    let quiz = payload.quiz;
    if(!quiz || !quiz.id) return jsonResponse({status:"error", message:"Quiz invalid"});
    let sh = getSheet(SHEET_QUIZ);
    if(!sh){
      try{
        sh = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(SHEET_QUIZ);
        sh.appendRow(["timestamp","quizId","judul","mapel","kelas","jumlahSoal","json"]);
      }catch(err){
        let arr=[];
        try{
          let prop = PropertiesService.getScriptProperties().getProperty("quizzes");
          if(prop) arr = JSON.parse(prop);
        }catch{}
        let idx = arr.findIndex(q=>q.id===quiz.id);
        if(idx>=0) arr[idx]=quiz; else arr.unshift(quiz);
        PropertiesService.getScriptProperties().setProperty("quizzes", JSON.stringify(arr));
        return jsonResponse({status:"ok", savedTo:"prop"});
      }
    }
    let headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>h.toString().trim().toLowerCase());
    let idxQuizId = headers.indexOf("quizid");
    let idxJudul = headers.indexOf("judul");
    let idxMapel = headers.indexOf("mapel");
    let idxKelas = headers.indexOf("kelas");
    let idxJumlah = headers.indexOf("jumlahsoal");
    let idxJson = headers.indexOf("json");
    let vals = sh.getDataRange().getValues();
    let found=false;
    for(let i=1;i<vals.length;i++){
      if((vals[i][idxQuizId]+"").trim() === quiz.id){
        if(idxJudul>=0) sh.getRange(i+1, idxJudul+1).setValue(quiz.judul||"");
        if(idxMapel>=0) sh.getRange(i+1, idxMapel+1).setValue(quiz.mapel||"");
        if(idxKelas>=0) sh.getRange(i+1, idxKelas+1).setValue(quiz.kelas||"");
        if(idxJumlah>=0) sh.getRange(i+1, idxJumlah+1).setValue((quiz.questions||[]).length);
        if(idxJson>=0) sh.getRange(i+1, idxJson+1).setValue(JSON.stringify(quiz));
        sh.getRange(i+1,1).setValue(new Date());
        found=true; break;
      }
    }
    if(!found){
      sh.appendRow([new Date(), quiz.id, quiz.judul||"", quiz.mapel||"", quiz.kelas||"", (quiz.questions||[]).length, JSON.stringify(quiz)]);
    }
    return jsonResponse({status:"ok", savedTo:"sheet"});
  }
  if(action==="deletequiz"){
    let quizId = payload.quizId;
    let sh = getSheet(SHEET_QUIZ);
    if(sh){
      let headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>h.toString().trim().toLowerCase());
      let idxQuizId = headers.indexOf("quizid");
      let vals = sh.getDataRange().getValues();
      for(let i=1;i<vals.length;i++){
        if((vals[i][idxQuizId]+"").trim()===quizId){ sh.deleteRow(i+1); break; }
      }
    }
    try{
      let prop = PropertiesService.getScriptProperties().getProperty("quizzes");
      if(prop){
        let arr = JSON.parse(prop);
        arr = arr.filter(q=>q.id!==quizId);
        PropertiesService.getScriptProperties().setProperty("quizzes", JSON.stringify(arr));
      }
    }catch{}
    return jsonResponse({status:"ok"});
  }
  // === SAVE ATTEMPT - SESUAI HEADER: timestamp | attemptId | studentName | kelas | quizId | quizJudul | score | correct | total | date | answersJson ===
  if(action==="saveattempt"){
    let attempt = payload.attempt;
    if(!attempt) return jsonResponse({status:"error", message:"attempt kosong"});
    let sh = getSheet(SHEET_ATTEMPT);
    if(!sh){
      try{
        sh = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(SHEET_ATTEMPT);
        sh.appendRow(["timestamp","attemptId","studentName","kelas","quizId","quizJudul","score","correct","total","date","answersJson"]);
      }catch(err){
        try{
          let prop = PropertiesService.getScriptProperties().getProperty("attempts");
          let arr = prop? JSON.parse(prop):[];
          arr.push(attempt);
          PropertiesService.getScriptProperties().setProperty("attempts", JSON.stringify(arr));
        }catch{}
        return jsonResponse({status:"ok", savedTo:"prop"});
      }
    }
    let answersJson = "";
    try{ answersJson = JSON.stringify(attempt.answers||attempt.answersJson||[]); }catch{ answersJson = (attempt.answersJson||"").toString(); }
    sh.appendRow([new Date(), attempt.id||attempt.attemptId||"", attempt.studentName||"", attempt.kelas||"", attempt.quizId||"", attempt.quizJudul||"", attempt.score||0, attempt.correct||0, attempt.total||0, attempt.date||new Date().toLocaleString("id-ID"), answersJson]);
    return jsonResponse({status:"ok", savedTo:"sheet"});
  }
  if(action==="resetattempt"){
    let studentName = payload.studentName;
    let quizId = payload.quizId;
    let sh = getSheet(SHEET_ATTEMPT);
    if(sh){
      let headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>h.toString().trim().toLowerCase());
      let idxStudent = headers.indexOf("studentname");
      let idxQuizId = headers.indexOf("quizid");
      let vals = sh.getDataRange().getValues();
      for(let i=vals.length-1;i>=1;i--){
        let matchStudent = !studentName || (vals[i][idxStudent]+"").toLowerCase()===studentName.toLowerCase();
        let matchQuiz = !quizId || (vals[i][idxQuizId]+"")==quizId;
        if(matchStudent && matchQuiz) sh.deleteRow(i+1);
      }
    }
    return jsonResponse({status:"ok"});
  }
  if(action==="resetquiz"){
    let quizId = payload.quizId;
    let sh = getSheet(SHEET_ATTEMPT);
    if(sh){
      let headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>h.toString().trim().toLowerCase());
      let idxQuizId = headers.indexOf("quizid");
      let vals = sh.getDataRange().getValues();
      for(let i=vals.length-1;i>=1;i--){
        if((vals[i][idxQuizId]+"")==quizId) sh.deleteRow(i+1);
      }
    }
    return jsonResponse({status:"ok"});
  }
  if(action==="resetall"){
    let sh = getSheet(SHEET_ATTEMPT);
    if(sh){ 
      sh.clearContents(); 
      sh.appendRow(["timestamp","attemptId","studentName","kelas","quizId","quizJudul","score","correct","total","date","answersJson"]);
    }
    return jsonResponse({status:"ok"});
  }
  // === SAVE PROGRESS - SESUAI HEADER: timestamp | studentName | kelas | type | coinsReward | totalCoins | lastSpin | streak | bonusInfo ===
  if(["saveprogress","shopbuy","spin","savecoin","savecoins"].indexOf(action)>=0){
    let studentName = payload.studentName;
    if(!studentName) return jsonResponse({status:"error", message:"studentName wajib"});
    let sh = getSheet(SHEET_PROGRESS);
    if(!sh){
      try{
        sh = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(SHEET_PROGRESS);
        sh.appendRow(["timestamp","studentName","kelas","type","coinsReward","totalCoins","lastSpin","streak","bonusInfo"]);
      }catch(err){
        try{
          let key = "progress_"+studentName.toLowerCase();
          let prop = PropertiesService.getScriptProperties().getProperty(key);
          let cur = prop? JSON.parse(prop):{totalCoins:0, streak:0, lastSpin:"", shopItems:[]};
          if(payload.totalCoins!==undefined) cur.totalCoins=Number(payload.totalCoins);
          else if(payload.coins) cur.totalCoins = (cur.totalCoins||0)+Number(payload.coins);
          if(payload.streak!==undefined) cur.streak=Number(payload.streak);
          if(payload.lastSpin) cur.lastSpin=payload.lastSpin;
          if(payload.item){
            cur.shopItems = cur.shopItems||[];
            if(cur.shopItems.indexOf(payload.item)==-1) cur.shopItems.push(payload.item);
            if(payload.price) cur.totalCoins = (cur.totalCoins||0)-Number(payload.price);
          }
          PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(cur));
        }catch{}
        return jsonResponse({status:"ok", savedTo:"prop"});
      }
    }
    let headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h=>h.toString().trim().toLowerCase());
    let idxStudent = headers.indexOf("studentname");
    let idxKelas = headers.indexOf("kelas");
    let idxType = headers.indexOf("type");
    let idxCoinsReward = headers.indexOf("coinsreward");
    let idxTotal = headers.indexOf("totalcoins");
    let idxLastSpin = headers.indexOf("lastspin");
    let idxStreak = headers.indexOf("streak");
    let idxBonus = headers.indexOf("bonusinfo");

    let type = payload.type||action;
    let coinsReward = Number(payload.coinsReward||payload.coins||payload.reward||0);
    let lastSpin = payload.lastSpin||"";
    let streak = payload.streak!==undefined? Number(payload.streak) : null;
    let bonusInfo = payload.bonusInfo||payload.item||payload.shopItems? (typeof payload.shopItems==="string"? payload.shopItems : JSON.stringify(payload.shopItems||payload.item||"")) : "";

    // Anti-cheat spin
    if(type==="spin" && lastSpin){
      let latest = getLatestProgressForStudent(studentName);
      if(latest && latest.lastSpin===lastSpin){
        return jsonResponse({status:"error", error:"ALREADY_SPUN", message:"Sudah spin hari ini", lastSpin:latest.lastSpin});
      }
    }

    // Hitung totalCoins terbaru
    let latest = getLatestProgressForStudent(studentName);
    let prevTotal = latest? Number(latest.totalCoins||0) : 0;
    let newTotal = payload.totalCoins!==undefined? Number(payload.totalCoins) : (prevTotal + coinsReward);
    // kalau shopBuy dengan price, kurangi
    if(type.indexOf("shop")>=0 && payload.price){
      newTotal = prevTotal - Number(payload.price);
      coinsReward = -Number(payload.price);
      bonusInfo = payload.item||bonusInfo;
    }

    sh.appendRow([new Date(), studentName, payload.kelas|| (latest?latest.kelas:""), type, coinsReward, newTotal, lastSpin || (latest?latest.lastSpin:""), streak!==null? streak : (latest?latest.streak:0), bonusInfo]);

    return jsonResponse({status:"ok", savedTo:"sheet", totalCoins:newTotal});
  }

  if(!action){ return jsonResponse({status:"ok", message:"SDGMP SmartCBT API POST aktif. Kirim JSON {action: ...}", received:payload, timestamp:new Date()}); } return jsonResponse({status:"error", message:"Unknown POST action "+actionRaw});
}

function jsonResponse(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
