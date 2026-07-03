/**
 * Google Apps Script for Lunchbox Ordering System
 * 
 * 部署步驟：
 * 1. 在 Google 試算表中，點擊選單的「擴充功能」 -> 「Apps Script」。
 * 2. 清空原本的程式碼，將此檔案的所有內容貼上。
 * 3. 點擊右上角的「部署」 -> 「新增部署作業」。
 * 4. 點擊「選取類型」齒輪，選擇「網頁應用程式 (Web App)」。
 * 5. 設定：
 *    - 說明：隨意填寫（例如：訂便當系統 API）
 *    - 專案執行身分：選擇「您的帳戶 (me)」
 *    - 誰有存取權：選擇「所有人 (Anyone)」
 * 6. 點擊「部署」，並複製產生的「網頁應用程式 URL」（Web App URL）。
 * 7. 將此 URL 複製到點餐系統的管理者設定中。
 */

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuSheet = getOrCreateSheet(sheet, "Menu");
  var ordersSheet = getOrCreateSheet(sheet, "Orders");
  var configSheet = getOrCreateSheet(sheet, "Config");
  
  // 讀取截止時間與指定店家
  var deadline = "";
  var activeStore = "";
  if (configSheet.getLastRow() >= 2) {
    deadline = configSheet.getRange("A2").getValue();
    activeStore = configSheet.getRange("B2").getValue().toString();
  }
  
  // 讀取菜單
  var menu = [];
  if (menuSheet.getLastRow() >= 2) {
    var menuData = menuSheet.getRange(2, 1, menuSheet.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < menuData.length; i++) {
      if (menuData[i][0] && menuData[i][1]) {
        menu.push({
          store: menuData[i][0].toString(),
          item: menuData[i][1].toString(),
          price: parseFloat(menuData[i][2]) || 0
        });
      }
    }
  }
  
  // 讀取所有訂單
  var orders = [];
  if (ordersSheet.getLastRow() >= 2) {
    var ordersData = ordersSheet.getRange(2, 1, ordersSheet.getLastRow() - 1, 5).getValues();
    for (var i = 0; i < ordersData.length; i++) {
      if (ordersData[i][1]) {
        orders.push({
          timestamp: ordersData[i][0],
          name: ordersData[i][1].toString(),
          store: ordersData[i][2].toString(),
          item: ordersData[i][3].toString(),
          price: parseFloat(ordersData[i][4]) || 0
        });
      }
    }
  }
  
  var result = {
    deadline: deadline,
    activeStore: activeStore,
    menu: menu,
    orders: orders
  };
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var result = { success: false };
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("無效的 POST 請求內容");
    }
    
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    
    if (data.action === "updateMenu") {
      var menuSheet = getOrCreateSheet(sheet, "Menu");
      menuSheet.clearContents();
      
      // 寫入標頭
      menuSheet.getRange(1, 1, 1, 3).setValues([["Store", "Item", "Price"]]);
      
      if (data.menu && data.menu.length > 0) {
        var rows = [];
        for (var i = 0; i < data.menu.length; i++) {
          rows.push([
            data.menu[i].store,
            data.menu[i].item,
            parseFloat(data.menu[i].price) || 0
          ]);
        }
        menuSheet.getRange(2, 1, rows.length, 3).setValues(rows);
      }
      result.success = true;
      
    } else if (data.action === "setDeadline") {
      var configSheet = getOrCreateSheet(sheet, "Config");
      configSheet.getRange(1, 1, 1, 2).setValues([["Deadline", "ActiveStore"]]);
      configSheet.getRange("A2").setValue(data.deadline || "");
      result.success = true;
      
    } else if (data.action === "setActiveStore") {
      var configSheet = getOrCreateSheet(sheet, "Config");
      configSheet.getRange(1, 1, 1, 2).setValues([["Deadline", "ActiveStore"]]);
      configSheet.getRange("B2").setValue(data.activeStore || "");
      result.success = true;
      
    } else if (data.action === "submitOrder") {
      var ordersSheet = getOrCreateSheet(sheet, "Orders");
      
      // 檢查是否已截止
      var configSheet = getOrCreateSheet(sheet, "Config");
      var deadlineStr = "";
      if (configSheet.getLastRow() >= 2) {
        deadlineStr = configSheet.getRange("A2").getValue();
      }
      
      if (deadlineStr) {
        var deadline = new Date(deadlineStr);
        var now = new Date();
        if (now > deadline) {
          throw new Error("已過截止點餐時間！無法新增訂單。");
        }
      }
      
      var timestamp = new Date();
      var rows = [];
      for (var j = 0; j < data.items.length; j++) {
        rows.push([
          timestamp,
          data.name,
          data.store,
          data.items[j].item,
          parseFloat(data.items[j].price) || 0
        ]);
      }
      
      if (rows.length > 0) {
        ordersSheet.getRange(ordersSheet.getLastRow() + 1, 1, rows.length, 5).setValues(rows);
      }
      result.success = true;
      
    } else if (data.action === "clearOrders") {
      var ordersSheet = getOrCreateSheet(sheet, "Orders");
      ordersSheet.clearContents();
      ordersSheet.getRange(1, 1, 1, 5).setValues([["Timestamp", "Name", "Store", "Item", "Price"]]);
      result.success = true;
      
    } else {
      throw new Error("未知的 action 操作");
    }
  } catch(err) {
    result.error = err.toString();
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(sheet, name) {
  var s = sheet.getSheetByName(name);
  if (!s) {
    s = sheet.insertSheet(name);
    if (name === "Menu") {
      s.getRange(1, 1, 1, 3).setValues([["Store", "Item", "Price"]]);
    } else if (name === "Orders") {
      s.getRange(1, 1, 1, 5).setValues([["Timestamp", "Name", "Store", "Item", "Price"]]);
    } else if (name === "Config") {
      s.getRange(1, 1, 1, 2).setValues([["Deadline", "ActiveStore"]]);
    }
  }
  return s;
}
