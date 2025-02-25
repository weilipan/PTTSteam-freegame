function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("PTT Steam 限免遊戲查詢")
    .addItem("📄 建立空白表格", "initializeSheet")
    .addItem(✉️ 設定 Email", "setEmail")
    .addItem("🔄 設定 maxPages", "setMaxPages")
    .addItem("🔍 查詢並寄送相關資訊", "checkPTTSteamFreeGames")
    .addToUi();
}

function initializeSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheetId = sheet.getId();

  // **📌 先建立需要的表格**
  var trackingSheet = sheet.getSheetByName("PTTSteam Tracking");
  var settingsSheet = sheet.getSheetByName("Settings");

  if (!trackingSheet) {
    trackingSheet = sheet.insertSheet("PTTSteam Tracking");
    trackingSheet.appendRow(["Article Links", "Sent Time"]);
  }

  if (!settingsSheet) {
    settingsSheet = sheet.insertSheet("Settings");
    settingsSheet.appendRow(["email", "maxPages", "sheet_id"]);
    settingsSheet.appendRow(["your_email@example.com", "5", sheetId]); // 預設 Email, maxPages, Sheet ID
  } else {
    settingsSheet.getRange("C2").setValue(sheetId);
  }

  // **🗑️ 刪除多餘的工作表**
  var sheets = sheet.getSheets();
  sheets.forEach(function(s) {
    var name = s.getName();
    if (name !== "PTTSteam Tracking" && name !== "Settings") {
      sheet.deleteSheet(s);
    }
  });

  SpreadsheetApp.getUi().alert("試算表初始化完成，所有多餘的工作表已刪除！");
}

function getSheetById() {
  var settingsSheet = SpreadsheetApp.openById(SpreadsheetApp.getActiveSpreadsheet().getId()).getSheetByName("Settings");
  var sheetId = settingsSheet.getRange("C2").getValue();
  return SpreadsheetApp.openById(sheetId);
}

function setEmail() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("請輸入 Email 地址（用於接收通知）", ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var email = response.getResponseText();
    var sheet = getSheetById().getSheetByName("Settings");
    sheet.getRange("A2").setValue(email);
    ui.alert("Email 已設定為：" + email);
  }
}

function setMaxPages() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt("請輸入要爬取的頁數（預設 5 頁）", ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() == ui.Button.OK) {
    var maxPages = parseInt(response.getResponseText(), 10);
    if (isNaN(maxPages) || maxPages <= 0) {
      ui.alert("請輸入有效的數字！");
      return;
    }
    var sheet = getSheetById().getSheetByName("Settings");
    sheet.getRange("B2").setValue(maxPages);
    ui.alert("maxPages 已設定為：" + maxPages);
  }
}
function cleanupOldEntries(trackingSheet) {
  var today = new Date();
  var twoWeeksAgo = new Date();
  twoWeeksAgo.setDate(today.getDate() - 14); // 設定為 14 天前

  var data = trackingSheet.getDataRange().getValues(); // 讀取整個表格
  var rowsToDelete = [];

  // **從最後一行往上刪除，避免刪除時索引錯誤**
  for (var i = data.length - 1; i > 0; i--) {
    var rowDate = new Date(data[i][1]); // 第二欄是 Send Time
    if (rowDate < twoWeeksAgo) {
      rowsToDelete.push(i + 1); // Google Sheets 索引從 1 開始
    }
  }

  // **批量刪除過期資料**
  rowsToDelete.forEach(function(rowIndex) {
    trackingSheet.deleteRow(rowIndex);
    Logger.log("🗑️ 已刪除超過 2 週的資料，行數：" + rowIndex);
  });
}

function checkPTTSteamFreeGames() {
  var sheet = getSheetById();
  var trackingSheet = sheet.getSheetByName("PTTSteam Tracking");
  var settingsSheet = sheet.getSheetByName("Settings");

  var email = settingsSheet.getRange("A2").getValue();
  var maxPages = parseInt(settingsSheet.getRange("B2").getValue(), 10);
  var existingLinks = trackingSheet.getRange("A:A").getValues().flat().filter(String);
  var baseUrl = "https://www.ptt.cc";
  var startUrl = baseUrl + "/bbs/Steam/index.html";

  var freeGamesList = [];
  var newEntries = []; // ✅ 用來儲存要一次性寫入 Google Sheets 的資料

  // ✅ 先清理超過 2 週的舊記錄
  cleanupOldEntries(trackingSheet);

  for (var i = 0; i < maxPages; i++) {
    Logger.log("🔍 正在爬取：" + startUrl);
    
    var response = UrlFetchApp.fetch(startUrl, {
      "method": "get",
      "headers": { 
        "Cookie": "over18=1",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
      },
      "muteHttpExceptions": true
    });

    var html = response.getContentText("UTF-8");

    // 📝 找到標題內含 `[限免]` 的文章
    var pattern = /<div class="title">\s*<a href="([^"]+)"[^>]*>\[限免\]([^<]+)<\/a>/g;
    var match;
    
    while ((match = pattern.exec(html)) !== null) {
      var link = match[1];
      var title = `[限免] ` + match[2].trim();
      
      if (!title || !link) continue;
      
      var fullUrl = baseUrl + link;
      
      if (!existingLinks.includes(fullUrl)) {
        freeGamesList.push({ title: title, url: fullUrl });
        newEntries.push([fullUrl, new Date()]); // ✅ 暫存到陣列，最後一次性寫入
      }
    }

    // 🛑 找到「上頁」按鈕
    var nextPageMatch = html.match(/<a class="btn wide" href="(\/bbs\/Steam\/index\d+\.html)">.*?上頁<\/a>/);
    Logger.log(nextPageMatch);
    if (nextPageMatch) {
      startUrl = baseUrl + nextPageMatch[1];
      Logger.log("🔄 發現上頁連結，繼續爬取：" + startUrl);
    } else {
      Logger.log("🚫 找不到上頁，停止爬取。");
      break;
    }
  }

  // ✅ 批量寫入 Google Sheets，避免最後一筆漏寫
  if (newEntries.length > 0) {
    trackingSheet.getRange(trackingSheet.getLastRow() + 1, 1, newEntries.length, 2).setValues(newEntries);
  }

  // ✉️ 若有新 `[限免]` 文章，寄送 Email
  if (freeGamesList.length > 0) {
    var subject = "PTT Steam 限免遊戲通知 (爬取 " + maxPages + " 頁)";
    var body = "以下是今天新發現的 [限免] 遊戲：\n\n";
    
    freeGamesList.forEach(function(item) {
      body += "標題: " + item.title + "\n" +
              "連結: " + item.url + "\n\n";
    });

    MailApp.sendEmail({
      to: email,
      subject: subject,
      body: body
    });

    SpreadsheetApp.getUi().alert("查詢完成，已發送 Email 通知！");
  } else {
    SpreadsheetApp.getUi().alert("未找到新的 [限免] 遊戲。");
  }
}
