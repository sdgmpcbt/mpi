
// SD Global Mandiri - SmartCBT Cloud Backend V7 - Import Excel + Anti-Farming Quiz + Spin
const SHEET_NAMES = {
  QUIZZES: "Quizzes",
  ATTEMPTS: "Attempts",
  PROGRESS: "Progress",
  SHOP: "ShopLog",
  DEBUG: "DebugLog",
  SISWA: "Siswa",
  GURU: "Guru"
};

function getOrCreateSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) { sheet = ss.insertSheet(name); sheet.appendRow(headers); }
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  return sheet;
}

function getServerToday() {
  // Use server time in Asia/Jakarta
  let now = new Date();
  // Convert to YYYY-MM-DD in Jakarta timezone
  let jakarta = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd");
  return jakarta;
}

function doPost(e) {
  let data = {};
  try {
    if (e.postData) {
      try { data = JSON.parse(e.postData.contents); } 
      catch { 
        try {
          let c = e.postData.contents.trim();
          if(c.startsWith("{")) data = JSON.parse(c);
          else data = e.parameter || {};
        } catch { data = e.parameter || {}; }
      }
    } else data = e.parameter || {};
    
    const action = data.action || "";
    const timestamp = new Date().toLocaleString("id-ID", {timeZone: "Asia/Jakarta"});
    const serverToday = getServerToday();
    if (!action) return jsonResponse({status:"error", message:"No action"});

    // === ANTI-FARMING SPIN - pakai server date, bukan client date ===
    if (action === "saveProgress" && data.type === "spin") {
      const progressSheet = getOrCreateSheet(SHEET_NAMES.PROGRESS, ["timestamp","studentName","kelas","type","coinsReward","totalCoins","lastSpin","streak","bonusInfo"]);
      const rows = progressSheet.getDataRange().getValues();
      // Pakai serverToday untuk validasi, bukan client lastSpin
      let clientLastSpin = (data.lastSpin||"").toString().trim();
      let todayToCheck = serverToday; // Server yang tentukan hari ini
      
      // Jika client kirim tanggal beda jauh dari server (manipulasi tanggal HP), tolak
      // Allow 1 day tolerance for timezone
      let normalizedName = (data.studentName || "").trim().toLowerCase();
      
      for (let i = 1; i < rows.length; i++) {
        let rowName = (rows[i][1] || "").toString().trim().toLowerCase();
        let rowLastSpin = (rows[i][6] || "").toString().trim();
        let rowType = (rows[i][3] || "").toString().toLowerCase();
        if (rowName === normalizedName && rowLastSpin === todayToCheck && rowType === "spin") {
          return jsonResponse({
            status:"error", 
            message:"Sudah spin hari ini", 
            error:"ALREADY_SPUN",
            lastSpin: todayToCheck,
            studentName: data.studentName,
            serverDate: serverToday
          });
        }
        // Also check if clientLastSpin == today but server already has different entry today
        if (rowName === normalizedName && rowLastSpin === clientLastSpin && rowType === "spin" && clientLastSpin===todayToCheck) {
          return jsonResponse({
            status:"error", 
            message:"Sudah spin hari ini (client date)", 
            error:"ALREADY_SPUN",
            lastSpin: clientLastSpin,
            studentName: data.studentName
          });
        }
      }
      // Simpan dengan serverToday sebagai lastSpin yang valid
      progressSheet.appendRow([timestamp, data.studentName || "", data.kelas || "", "spin", data.coins || 0, data.totalCoins || "", todayToCheck, data.streak || 0, (data.streak % 7 === 0 && data.streak > 0) ? "BONUS 200" : ""]);
      return jsonResponse({status:"ok", action:"saveProgress", message:"Spin saved", lastSpin: todayToCheck, serverDate: serverToday});
    }

    if (action === "saveQuiz") {
      const sheet = getOrCreateSheet(SHEET_NAMES.QUIZZES, ["timestamp","quizId","judul","mapel","kelas","jumlahSoal","json"]);
      const q = data.quiz || {}; const rows = sheet.getDataRange().getValues();
      let found = -1; for (let i = 1; i < rows.length; i++) if (rows[i][1] === q.id) { found = i+1; break; }
      const rowData = [timestamp, q.id, q.judul, q.mapel, (q.kelas||[]).join(","), (q.questions||[]).length, JSON.stringify(q)];
      if (found > 0) sheet.getRange(found, 1, 1, rowData.length).setValues([rowData]); else sheet.appendRow(rowData);
    } 
    else if (action === "deleteQuiz") {
      const sheet = getOrCreateSheet(SHEET_NAMES.QUIZZES, ["timestamp","quizId","judul","mapel","kelas","jumlahSoal","json"]);
      const rows = sheet.getDataRange().getValues();
      for (let i = rows.length-1; i >= 1; i--) if (rows[i][1] === data.quizId) { sheet.deleteRow(i+1); break; }
    }
    else if (action === "saveAttempt") {
      const sheet = getOrCreateSheet(SHEET_NAMES.ATTEMPTS, ["timestamp","attemptId","studentName","kelas","quizId","quizJudul","score","correct","total","date","answersJson"]);
      const a = data.attempt || {};
      const rows = sheet.getDataRange().getValues();
      let normalizedName = (a.studentName||"").toString().trim().toLowerCase();
      for (let i = 1; i < rows.length; i++) {
        if ((rows[i][2]||"").toString().trim().toLowerCase() === normalizedName && rows[i][4] === a.quizId) {
          return jsonResponse({status:"error", message:"Sudah pernah kerjakan quiz ini", error:"ALREADY_ATTEMPTED", quizId: a.quizId, studentName: a.studentName});
        }
      }
      sheet.appendRow([timestamp, a.id || "", a.studentName || "", a.kelas || "", a.quizId || "", a.quizJudul || "", a.score || 0, a.correct || 0, a.total || 0, a.date || timestamp, JSON.stringify(a.answers || []).substring(0, 5000)]);
    }
    else if (action === "resetAttempt") {
      const sheet = getOrCreateSheet(SHEET_NAMES.ATTEMPTS, ["timestamp","attemptId","studentName","kelas","quizId","quizJudul","score","correct","total","date","answersJson"]);
      const rows = sheet.getDataRange().getValues();
      let normalizedName = (data.studentName||"").toString().trim().toLowerCase();
      for (let i = rows.length-1; i >=1; i--) if ((rows[i][2]||"").toString().trim().toLowerCase() === normalizedName && rows[i][4] === data.quizId) sheet.deleteRow(i+1);
    }
    else if (action === "resetQuiz") {
      const sheet = getOrCreateSheet(SHEET_NAMES.ATTEMPTS, ["timestamp","attemptId","studentName","kelas","quizId","quizJudul","score","correct","total","date","answersJson"]);
      const rows = sheet.getDataRange().getValues();
      for (let i = rows.length-1; i >=1; i--) if (rows[i][4] === data.quizId) sheet.deleteRow(i+1);
    }
    else if (action === "resetAll") {
      const sheet = getOrCreateSheet(SHEET_NAMES.ATTEMPTS, ["timestamp","attemptId","studentName","kelas","quizId","quizJudul","score","correct","total","date","answersJson"]);
      if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow()-1, sheet.getLastColumn()).clearContent();
    }
    else if (action === "saveProgress" && data.type !== "spin") {
      const sheet = getOrCreateSheet(SHEET_NAMES.PROGRESS, ["timestamp","studentName","kelas","type","coinsReward","totalCoins","lastSpin","streak","bonusInfo"]);
      sheet.appendRow([timestamp, data.studentName || "", data.kelas || "", data.type || "manual", data.coins || 0, data.totalCoins || "", data.lastSpin || "", data.streak || 0, data.bonusInfo || ""]);
    }
    else if (action === "shopBuy") {
      const sheet = getOrCreateSheet(SHEET_NAMES.SHOP, ["timestamp","studentName","item","price","remainingCoins"]);
      const rows = sheet.getDataRange().getValues();
      let normalizedName = (data.studentName||"").toString().trim().toLowerCase();
      for (let i = 1; i < rows.length; i++) {
        if ((rows[i][1]||"").toString().trim().toLowerCase() === normalizedName && rows[i][2] === data.item) {
          return jsonResponse({status:"error", message:"Sudah memiliki item ini", error:"ALREADY_OWNED"});
        }
      }
      sheet.appendRow([timestamp, data.studentName || "", data.item || "", data.price || 0, data.coins || 0]);
      const progSheet = getOrCreateSheet(SHEET_NAMES.PROGRESS, ["timestamp","studentName","kelas","type","coinsReward","totalCoins","lastSpin","streak","bonusInfo"]);
      progSheet.appendRow([timestamp, data.studentName || "", "", "shopBuy", -Math.abs(data.price || 0), data.coins || 0, "", "", `Beli ${data.item}`]);
    }

    return jsonResponse({status:"ok", action: action, serverDate: serverToday});
  } catch (err) {
    return jsonResponse({status:"error", message: err.toString()});
  }
}

function doGet(e) {
  try {
    const action = (e.parameter.action || "").toLowerCase();
    const studentName = e.parameter.studentName || e.parameter.nama || "";
    const quizId = e.parameter.quizId || "";
    const serverToday = getServerToday();
    
    if (action === "checkattempt" && studentName && quizId) {
      const attemptsSheet = getOrCreateSheet(SHEET_NAMES.ATTEMPTS, ["timestamp","attemptId","studentName","kelas","quizId","quizJudul","score","correct","total","date","answersJson"]);
      const rows = attemptsSheet.getDataRange().getValues();
      let normalizedName = studentName.trim().toLowerCase();
      for (let i = 1; i < rows.length; i++) {
        if ((rows[i][2]||"").toString().trim().toLowerCase() === normalizedName && rows[i][4] === quizId) {
          return jsonResponse({status:"ok", alreadyAttempted:true, quizId: quizId, studentName: studentName, score: rows[i][6], date: rows[i][9]});
        }
      }
      return jsonResponse({status:"ok", alreadyAttempted:false, quizId: quizId, studentName: studentName});
    }
    
    if (action === "checkspin" && studentName) {
      const progressSheet = getOrCreateSheet(SHEET_NAMES.PROGRESS, ["timestamp","studentName","kelas","type","coinsReward","totalCoins","lastSpin","streak","bonusInfo"]);
      const rows = progressSheet.getDataRange().getValues();
      let todayStr = e.parameter.date || serverToday;
      // Always use serverToday for anti-farming, but allow param for testing
      let checkDate = serverToday;
      if (e.parameter.date) checkDate = e.parameter.date; // for manual check
      // For strict anti-farming, always check serverToday
      let normalizedName = studentName.trim().toLowerCase();
      for (let i = 1; i < rows.length; i++) {
        let rowName = (rows[i][1] || "").toString().trim().toLowerCase();
        let rowLastSpin = (rows[i][6] || "").toString().trim();
        let rowType = (rows[i][3] || "").toString().toLowerCase();
        if (rowName === normalizedName && rowLastSpin === serverToday && rowType === "spin") {
          return jsonResponse({status:"ok", alreadySpun:true, lastSpin: serverToday, studentName: studentName, serverDate: serverToday, message:"Sudah spin hari ini"});
        }
      }
      return jsonResponse({status:"ok", alreadySpun:false, lastSpin:"", studentName: studentName, serverDate: serverToday, message:"Belum spin hari ini"});
    }

    if (action === "getprogress" && studentName) {
      const progressSheet = getOrCreateSheet(SHEET_NAMES.PROGRESS, ["timestamp","studentName","kelas","type","coinsReward","totalCoins","lastSpin","streak","bonusInfo"]);
      const shopSheet = getOrCreateSheet(SHEET_NAMES.SHOP, ["timestamp","studentName","item","price","remainingCoins"]);
      const attemptsSheet = getOrCreateSheet(SHEET_NAMES.ATTEMPTS, ["timestamp","attemptId","studentName","kelas","quizId","quizJudul","score","correct","total","date","answersJson"]);
      const pRows = progressSheet.getDataRange().getValues();
      const sRows = shopSheet.getDataRange().getValues();
      const aRows = attemptsSheet.getDataRange().getValues();
      let totalCoins = 0, lastSpin = "", streak = 0, kelas = "", history = [];
      let normalizedName = studentName.trim().toLowerCase();
      for (let i = 1; i < pRows.length; i++) {
        if ((pRows[i][1]||"").toString().trim().toLowerCase() === normalizedName) {
          totalCoins += parseInt(pRows[i][4] || 0, 10) || 0;
          if (pRows[i][6]) lastSpin = pRows[i][6];
          if (pRows[i][7]) streak = parseInt(pRows[i][7],10) || streak;
          if (pRows[i][2]) kelas = pRows[i][2];
          history.push({timestamp: pRows[i][0], type: pRows[i][3], coins: pRows[i][4], lastSpin: pRows[i][6], streak: pRows[i][7]});
        }
      }
      let shopItems = [];
      for (let i = 1; i < sRows.length; i++) if ((sRows[i][1]||"").toString().trim().toLowerCase() === normalizedName) shopItems.push(sRows[i][2]);
      shopItems = [...new Set(shopItems)];
      let attempts = [];
      for (let i = 1; i < aRows.length; i++) if ((aRows[i][2]||"").toString().trim().toLowerCase() === normalizedName) {
        attempts.push({
          attemptId: aRows[i][1], id: aRows[i][1], studentName: aRows[i][2], kelas: aRows[i][3],
          quizId: aRows[i][4], quizJudul: aRows[i][5], score: aRows[i][6], correct: aRows[i][7], total: aRows[i][8], date: aRows[i][9],
          answers: (()=>{try{return JSON.parse(aRows[i][10]||"[]")}catch{return []}})()
        });
      }
      let attemptsCount = attempts.length;
      let xp = (attemptsCount * 120) % 1000;
      let level = Math.min(10, Math.floor(attemptsCount/2)+1);
      let badgeCount = Math.min(13, attemptsCount+1);
      let alreadySpunToday = false;
      for (let i = 1; i < pRows.length; i++) {
        if ((pRows[i][1]||"").toString().trim().toLowerCase() === normalizedName && (pRows[i][6]||"") === serverToday && (pRows[i][3]||"").toLowerCase() === "spin") {
          alreadySpunToday = true; break;
        }
      }
      return jsonResponse({
        status:"ok", studentName: studentName, kelas: kelas,
        totalCoins: totalCoins, lastSpin: lastSpin, streak: streak,
        shopItems: shopItems, attemptsCount: attemptsCount,
        xp: xp, level: level, badgeCount: badgeCount,
        alreadySpunToday: alreadySpunToday,
        serverDate: serverToday,
        attempts: attempts, history: history.slice(-20)
      });
    }

    if (action === "getleaderboard") {
      const progressSheet = getOrCreateSheet(SHEET_NAMES.PROGRESS, ["timestamp","studentName","kelas","type","coinsReward","totalCoins","lastSpin","streak","bonusInfo"]);
      const attemptsSheet = getOrCreateSheet(SHEET_NAMES.ATTEMPTS, ["timestamp","attemptId","studentName","kelas","quizId","quizJudul","score","correct","total","date","answersJson"]);
      const pRows = progressSheet.getDataRange().getValues();
      const aRows = attemptsSheet.getDataRange().getValues();
      let map = {};
      for (let i = 1; i < pRows.length; i++) {
        let name = pRows[i][1]; if (!name) continue;
        let norm = name.toString().trim().toLowerCase();
        if (!map[norm]) map[norm] = {studentName: name, totalCoins: 0, streak: 0, kelas: pRows[i][2] || "", attempts:0, originalName: name};
        map[norm].totalCoins += parseInt(pRows[i][4] || 0, 10) || 0;
        if (pRows[i][7]) map[norm].streak = parseInt(pRows[i][7],10) || map[norm].streak;
      }
      for (let i = 1; i < aRows.length; i++) {
        let name = aRows[i][2]; if (!name) continue;
        let norm = name.toString().trim().toLowerCase();
        if (!map[norm]) map[norm] = {studentName: name, totalCoins: 0, streak: 0, kelas: aRows[i][3] || "", attempts:0, originalName: name};
        map[norm].attempts = (map[norm].attempts||0)+1;
        map[norm].xp = (map[norm].attempts * 120) % 1000;
        map[norm].level = Math.min(10, Math.floor(map[norm].attempts/2)+1);
        map[norm].badgeCount = Math.min(13, map[norm].attempts+1);
      }
      let leaderboard = Object.values(map).sort((a,b) => b.totalCoins - a.totalCoins).slice(0,20);
      return jsonResponse({status:"ok", leaderboard: leaderboard, serverDate: serverToday});
    }

    if (action === "getquizzes") {
      const sheet = getOrCreateSheet(SHEET_NAMES.QUIZZES, ["timestamp","quizId","judul","mapel","kelas","jumlahSoal","json"]);
      const rows = sheet.getDataRange().getValues();
      let quizzes = [];
      for (let i = 1; i < rows.length; i++) { try { let q = JSON.parse(rows[i][6]); if(q) quizzes.push(q); } catch {} }
      return jsonResponse({status:"ok", count: quizzes.length, quizzes: quizzes, serverDate: serverToday});
    }

    return jsonResponse({status:"ok", message:"V7 Anti-Farming Quiz+Spin + Import Excel Active - Server date: "+serverToday, time: new Date().toISOString(), serverDate: serverToday});
  } catch (err) {
    return jsonResponse({status:"error", message: err.toString()});
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function doOptions(e) { return jsonResponse({status:"ok"}); }
function setupSheets() {
  getOrCreateSheet(SHEET_NAMES.QUIZZES, ["timestamp","quizId","judul","mapel","kelas","jumlahSoal","json"]);
  getOrCreateSheet(SHEET_NAMES.ATTEMPTS, ["timestamp","attemptId","studentName","kelas","quizId","quizJudul","score","correct","total","date","answersJson"]);
  getOrCreateSheet(SHEET_NAMES.PROGRESS, ["timestamp","studentName","kelas","type","coinsReward","totalCoins","lastSpin","streak","bonusInfo"]);
  getOrCreateSheet(SHEET_NAMES.SHOP, ["timestamp","studentName","item","price","remainingCoins"]);
  getOrCreateSheet(SHEET_NAMES.DEBUG, ["timestamp","payload","error"]);
}
