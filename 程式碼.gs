/**
 * COMET 系統核心 V6.1 - 規格智慧判定與模式 4 增強版
 */
const SS = SpreadsheetApp.getActiveSpreadsheet();
const ADMIN_EMAIL = "huxiu2148@gmail.com"; 
const API_KEY = "b594b6d36a9f8ba1e40ddf26"; 

function doGet() {
  const userEmail = Session.getActiveUser().getEmail();
  const templateName = (userEmail === ADMIN_EMAIL) ? 'Index' : 'Order';
  return HtmlService.createTemplateFromFile(templateName).evaluate()
      .setTitle('COMET小小代購💫💟')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0');
}

/* --- [使用者端功能] --- */

function getUserEmail() { return Session.getActiveUser().getEmail() || "anonymous"; }

/**
 * 抓取當前團務（用於填單頁與管理控制台）
 * 包含自動判定「已收單」的邏輯
 */
function getActiveGroups() {
  sortManagementSheetRealTime();
  SpreadsheetApp.flush(); 
  const sheet = SS.getSheetByName('團務管理');
  if (!sheet) return [];

  const displayData = sheet.getDataRange().getDisplayValues();
  const realData = sheet.getDataRange().getValues();

  displayData.shift(); 
  realData.shift();
  
  const nowTime = new Date().getTime(); 

  // 1. 先將原始資料轉換為物件，並判定「當前真正狀態」
  let groups = displayData.map((row, index) => {
    try {
      if (!row[0] || !row[1]) return null;

      let createTime = new Date(realData[index][5]).getTime(); // 假設「建立時間」在 F 欄 (索引 5)
      let deadlineDate = new Date(realData[index][2]);
      let deadlineTime = deadlineDate.getTime();
      
      let rawStatus = row[4] ? row[4].toString().trim() : "團務進行中";
      let displayStatus = rawStatus;
      let canOrder = false;

      // 自動收單判斷
      if (rawStatus === "團務進行中") {
        if (nowTime > deadlineTime) {
          displayStatus = "已收單";
          canOrder = false;
        } else {
          displayStatus = "團務進行中";
          canOrder = true;
        }
      } else {
        displayStatus = rawStatus;
        canOrder = false;
      }

      return {
        id: row[0],
        name: row[1],
        deadlineStr: row[2], 
        shippingTime: row[3] || "待更新",
        mode: row[6] || "1",
        link: row[7] || "", 
        status: displayStatus, 
        canOrder: canOrder,
        createTime: createTime // 為了排序暫存
      };
    } catch (e) { return null; }
  }).filter(g => g !== null);

  // 2. 定義狀態權重順序
  const statusWeight = {
    '團務進行中': 1,
    '已收單': 2,
    '官方出貨中': 3,
    '到集運': 4,
    '運回中': 5,
    '抵台寄出中': 6,
    '團務結束': 7
  };

  // 3. 執行自定義排序
  groups.sort((a, b) => {
    let weightA = statusWeight[a.status] || 99;
    let weightB = statusWeight[b.status] || 99;

    if (weightA !== weightB) {
      return weightA - weightB; // 狀態不同，按權重排 (1 > 2 > 3...)
    } else {
      return b.createTime - a.createTime; // 狀態相同，按建立時間排 (新到舊)
    }
  });

  return groups;
}

function sortManagementSheetRealTime() {
  const sheet = SS.getSheetByName('團務管理');
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return; // 只有標題就不排

  const range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn());
  const data = range.getValues();

  // 定義狀態權重
  const statusWeight = {
    '團務進行中': 1,
    '已收單': 2,
    '官方出貨中': 3,
    '到集運': 4,
    '運回中': 5,
    '抵台寄出中': 6,
    '團務結束': 7
  };

  // 執行排序
  data.sort((a, b) => {
    // a[4] 是狀態欄 (E欄), a[5] 是建立時間 (F欄)
    let weightA = statusWeight[a[4]] || 99;
    let weightB = statusWeight[b[4]] || 99;

    if (weightA !== weightB) {
      return weightA - weightB;
    } else {
      // 狀態相同，按建立時間排 (新到舊)
      return new Date(b[5]) - new Date(a[5]);
    }
  });

  // 把排好的資料寫回試算表
  range.setValues(data);
}

function getGroupProducts(groupId) {
  const sheet = SS.getSheetByName('商品資料');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  return data.filter(r => r[0] === groupId).map(r => ({ name: r[2], price: r[3] }));
}

function submitOrderToSheet(orderData) {
  let sheet = SS.getSheetByName('訂單資料') || SS.insertSheet('訂單資料');
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["時間", "ID", "團務名稱", "LINE暱稱", "聯絡方式", "明細", "金額", "狀態", "系統Email", "使用者填寫Email", "末五碼"]);
  }
  
  const sysEmail = getUserEmail();
  sheet.insertRowBefore(2); 
  
  const rowData = [
    new Date(), orderData.groupId, orderData.groupName, orderData.lineName, 
    orderData.contact, orderData.detail, orderData.total, "未核對", 
    sysEmail, orderData.userEmail, ""
  ];
  
  sheet.getRange(2, 1, 1, rowData.length).setValues([rowData]);

  // --- 1. 處理護照姓名提取 & 寄送 Email ---
  if (!orderData.alreadyProcessed) {
    let extractedName = "";
    if (orderData.detail.includes("護照姓名")) {
       const match = orderData.detail.match(/【護照姓名[：:]\s*([^】]+)】/);
       extractedName = match ? match[1].trim() : "";
    }
    orderData.passportName = extractedName; 
    
    sendOrderConfirmEmail(orderData);
    orderData.alreadyProcessed = true; 
  }
  
  // --- 2. Telegram 通知內容分流 ---
  const mode = String(orderData.mode);

  // 💡 只有模式 4 不發 Telegram，其餘 (1, 2, 3) 都發
  if (mode !== "4") {
    let tgIcon = "🔔";
    let tgType = "新訂單通知";
    // ✨ 所有模式預設都顯示金額
    let tgDetail = "💰 金額：NT$ " + Number(orderData.total).toLocaleString();

    if (mode === "2") {
      tgIcon = "✈️";
      tgType = "韓國代收通知";
      // 模式二額外加上護照名
      tgDetail += "\n📛 護照：" + (orderData.passportName || "未填");
    } else if (mode === "3") {
      tgIcon = "📝";
      tgType = "新登記通知";
      // 模式三在金額後加上提醒
      tgDetail += "\n📌 狀態：僅供登記";
    }

    const tgMsg = tgIcon + " <b>【" + tgType + "】</b>\n" +
                  "━━━━━━━━━━━━━━━━━━━━\n" +
                  "📦 團務：" + orderData.groupName + "\n" +
                  "👤 暱稱：" + orderData.lineName + "\n" +
                  tgDetail;

    sendTelegramNotification(tgMsg);
  }

  // --- 3. 儲存屬性 ---
  PropertiesService.getUserProperties().setProperties({
    "last_ln": orderData.lineName,
    "last_ct": orderData.contact,
    "last_em": orderData.userEmail
  });

  return { success: true, rowIndex: 2 };
}

function getMyHistoryOrders(email) {
  if (!email) return [];
  const cleanEmail = email.toString().trim().toLowerCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const orderSheet = ss.getSheetByName("訂單資料");
  const groupSheet = ss.getSheetByName("團務管理");
  
  const orderData = orderSheet.getDataRange().getValues();
  const groupData = groupSheet.getDataRange().getValues();

  const groupInfoMap = {};
  for (let j = 1; j < groupData.length; j++) {
    const gName = groupData[j][1] ? groupData[j][1].toString().trim() : "";
    if (gName) {
      groupInfoMap[gName] = {
        status: groupData[j][4] ? groupData[j][4].toString() : "團務進行中",
        mode: groupData[j][6] ? groupData[j][6].toString() : "1",
        link: groupData[j][7] ? groupData[j][7].toString() : ""
      };
    }
  }

  const results = [];
  for (let i = 1; i < orderData.length; i++) {
    const rowEmail = orderData[i][9] ? orderData[i][9].toString().trim().toLowerCase() : ""; 
    if (rowEmail === cleanEmail) {
      const gName = orderData[i][2] ? orderData[i][2].toString().trim() : "";
      const info = groupInfoMap[gName] || { status: "團務進行中", mode: "1", link: "" };
      results.push({
        rowIndex: i + 1,
        time: orderData[i][0] ? Utilities.formatDate(new Date(orderData[i][0]), "GMT+8", "MM/dd HH:mm") : "",
        groupName: gName,
        detail: orderData[i][5] || "",
        total: orderData[i][6] || 0,
        status: orderData[i][7] || "未核對",
        groupStatus: info.status,
        mode: info.mode,
        link: info.link,
        remitCode: orderData[i][10] || "" 
      });
    }
  }
  return results.reverse();
}

function getUserInfo() {
  const props = PropertiesService.getUserProperties();
  return {
    lineName: props.getProperty("last_ln") || "",
    contact: props.getProperty("last_ct") || "",
    userEmail: props.getProperty("last_em") || ""
  };
}

// 即時權限檢查：掃描試算表
function checkOrderAuth(groupId) {
  SpreadsheetApp.flush();
  const sysEmail = getUserEmail();
  const props = PropertiesService.getUserProperties();
  const manualEmail = props.getProperty("last_em") || ""; 
  
  const sheet = SS.getSheetByName('訂單資料');
  if (!sheet || sheet.getLastRow() <= 1) return { hasOrder: false };
  
  const data = sheet.getDataRange().getValues();
  // 注意：這裡不要 shift()，直接用迴圈倒著找比較好抓正確的行號
  
  for (let i = data.length - 1; i >= 1; i--) { // 從最後一行往回找
    const r = data[i];
    const matchEmail = (sysEmail !== "anonymous" && r[8] === sysEmail) || (r[9] === manualEmail);
    const matchGroup = r[1].toString() === groupId.toString();
    
    if (matchEmail && matchGroup) {
      return { 
        hasOrder: true, 
        rowIndex: i + 1 // 陣列索引從 0 開始，所以行號要 +1
      };
    }
  }

  return { hasOrder: false };
}

/* --- [管理者端功能] --- */

function getExchangeRate(baseCurrency) {
  try {
    const currency = baseCurrency || "KRW";
    const url = `https://v6.exchangerate-api.com/v6/${API_KEY}/latest/${currency}`;
    const response = UrlFetchApp.fetch(url);
    const data = JSON.parse(response.getContentText());
    return data.result === "success" ? data.conversion_rates.TWD : 0;
  } catch (e) { return 0; }
}

/* --- [管理者端功能：自動與手動狀態管理] --- */

/**
 * ✅ 需求：新增團務改為「置頂插入」(第 2 列)
 */
function addAdminGroup(groupName, endTime, shippingTime, themePrefix, mode, link) {
  let sheet = SS.getSheetByName('團務管理') || SS.insertSheet('團務管理');
  const dateStr = Utilities.formatDate(new Date(), "GMT+8", "yyMMdd");
  const data = sheet.getDataRange().getValues();
  
  let count = 0;
  const searchKey = `${themePrefix}-${dateStr}`;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().includes(searchKey)) count++;
  }
  
  const finalId = `${themePrefix}-${dateStr}-${(count + 1).toString().padStart(2, '0')}`;
  
  // 插入新列於標題下方
  sheet.insertRowBefore(2);
  
  const rowData = [
    finalId, 
    groupName, 
    new Date(endTime), 
    shippingTime || "待更新", 
    "團務進行中", 
    new Date(), 
    mode || "1", 
    link || "",
    themePrefix
  ];

  sheet.getRange(2, 1, 1, rowData.length).setValues([rowData]);
  return { success: true, newId: finalId };
}



function batchAddProducts(productArray) {
  let sheet = SS.getSheetByName('商品資料') || SS.insertSheet('商品資料');
  let rows = productArray.map((p, index) => [
    p.groupId, (index + 1).toString().padStart(2, '0'), p.name, p.twd, p.foreign, p.master, p.profit
  ]);
  if (rows.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 7).setValues(rows);
  return { success: true };
}

/**
 * 1. 管理者：獲取所有訂單 (純訂單，不混團務狀態)
 */
function getAllOrdersForAdmin() {
  try {
    const sheet = SS.getSheetByName("訂單資料");
    if (!sheet) return []; 
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return []; 

    return data.slice(1).map((r, i) => ({
      rowIndex: i + 2,
      time: r[0] ? Utilities.formatDate(new Date(r[0]), "GMT+8", "MM/dd HH:mm") : "",
      groupName: r[2] || "未分類", 
      lineName: r[3] || "未知",  
      contact: r[4] || "",       
      detail: r[5] || "",        
      total: r[6] || 0,          
      status: r[7] || "未核對",  
      email: r[9] || "",         
      remitCode: r[10] || ""     
    }));
  } catch (e) { return []; }
}
/**
 * 2. 管理者：手動變更「團務管理」工作表的狀態 (對應 E 欄)
 */
function updateGroupStatusByName(groupName, newStatus) {
  const sheet = SS.getSheetByName('團務管理');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().trim() === groupName.toString().trim()) {
      sheet.getRange(i + 1, 5).setValue(newStatus); // 修改 E 欄
      return { success: true };
    }
  }
  return { success: false, error: "找不到團務名稱" };
}

/**
 * 3. 管理者：切換「訂單資料」單筆核對狀態 (對應 H 欄)
 */
function toggleOrderStatus(rowIndex, currentStatus) {
  const sheet = SS.getSheetByName('訂單資料');
  const nextStatus = (currentStatus === '已核對') ? '未核對' : '已核對';
  sheet.getRange(rowIndex, 8).setValue(nextStatus); // 修改 H 欄
  return { success: true, newStatus: nextStatus };
}

/**
 * 自動檢查收單時間（建議設定每分鐘觸發一次）
 */
function autoCloseGroups() {
  const sheet = SS.getSheetByName('團務管理');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < data.length; i++) {
    const deadline = new Date(data[i][2]);
    const currentStatus = data[i][4];
    if (currentStatus === "團務進行中" && deadline && now > deadline) {
      sheet.getRange(i + 1, 5).setValue("已收單");
    }
  }
}
function createForm(data) {
  try {
    var form = FormApp.create(data.title);
    form.setCollectEmail(true); 
    form.setConfirmationMessage("已收到！\n請於收單期限內進行匯款\n\n帳號：\n國泰 013 - 699507161336\n永豐 807 - 20401800319484\n中信 822 - 901567858153\n\n匯款完成才會協助下單！\n記得到 LINE 社群記事本留下末五碼對帳\n社群連結：https://reurl.cc/EQ1rLk");
    var productSummary = data.productList.map((p, index) => (index + 1) + ". " + p.name + " NT$" + p.price).join('\n');
    var fullDescription = "💟 商品金額💰\n" + productSummary + "\n\n收單時間：" + data.deadline + "\n官方出貨時間：" + data.shippingTime + "\n\n" + (data.extraNote ? "📝 注意事項：\n" + data.extraNote + "\n\n" : "") +"填完表單後會有匯款帳號 請在收單期限內匯款\n匯款完成才會協助下單 逾時不候\n\n"+ "✓ 以上皆需二補\n𖦹跟團前請先詳閱記事本重要貼文裡的注意事項\n𖦹填單即視同已閱讀並同意所有內容\n\n有任何問題都歡迎詢問 謝謝！";
    form.setDescription(fullDescription);
    form.addParagraphTextItem().setTitle("聯絡用帳號(FACEBOOK/INSTAGRAM)").setRequired(true);
    form.addParagraphTextItem().setTitle("在LINE社群裡的暱稱").setRequired(true);
    data.productList.forEach(function(p) { form.addParagraphTextItem().setTitle(p.name + " NT$" + p.price).setRequired(false); });
    form.addParagraphTextItem().setTitle("💰總金額").setRequired(true);
    return { success: true, url: form.getEditUrl(), viewUrl: form.getPublishedUrl() };
  } catch (e) { return { success: false, error: e.toString() }; }
}
function updateOrderRemitCode(rowIndex, code) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("訂單資料");
    const targetRow = parseInt(rowIndex);
    
    if (targetRow > 1) {
      // 寫入 K 欄
      sheet.getRange(targetRow, 11).setValue(code.toString().trim());
      
      // 取得該行資訊用來發送通知 (假設 D 欄是暱稱，即 index 4)
      const name = sheet.getRange(targetRow, 4).getValue(); 
      const group = sheet.getRange(targetRow, 3).getValue(); // C 欄團名
      
      // 發送 Telegram 通知給你
      const tgMsg = "🏦 <b>【資料待核對】</b>\n" +
                    "━━━━━━━━━━━━━━━━━━━━\n" +
                    "📦 團務：" + group + "\n" +
                    "👤 暱稱：" + name + "\n" +
                    "ℹ️ 資訊：<b>" + code + "</b>\n" +
                    "🔗 請進入管理端核對";
      sendTelegramNotification(tgMsg);
      
      return true;
    }
    return false;
  } catch (e) {
    return false;
  }
}


// 填入你的 Telegram 資訊
const TG_CONFIG = {
  token: "8302610604:AAHxXu2pYS6aEG2rtkjSHZy7qatbgxq8LIs",
  chatId: "7857837091"
};

/**
 * 共通發送 Telegram 函數
 */
function sendTelegramNotification(msg) {
  const url = "https://api.telegram.org/bot" + TG_CONFIG.token + "/sendMessage";
  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify({ "chat_id": TG_CONFIG.chatId, "text": msg, "parse_mode": "HTML" })
  };
  try { UrlFetchApp.fetch(url, options); } catch (e) { console.error("TG失敗: " + e.toString()); }
}

/**
 * 發送下單副本 Email
 */
function sendOrderConfirmEmail(data) {
  const mode = String(data.mode);
  
  // 💡 模式 4：賣貨便，不需要寄信
  if (mode === "4") return;

  let subject = "【跟團成功確認】COMET小小代購💫💟｜" + data.groupName;
  let body = "您好，已收到下單資料！以下是您的副本留存：\n";
  const divider = "━━━━━━━━━━━━━━━━━━━━\n";
  
  // 1. 基礎內容 (所有模式共用：團務、暱稱、帳號、金額)
  let content = 
    divider +
    "📦 團務名稱： " + data.groupName + "\n" +
    "👤 社群暱稱： " + data.lineName + "\n" +
    "📱 聯絡帳號： " + data.contact + "\n" +
    "💰 訂單總金額： NT$ " + Number(data.total).toLocaleString() + "\n"; // ✨ 移到這裡，所有模式都會顯示

  // 2. 根據模式補充特定資訊
  if (mode === "2") {
    // 模式二：直寄團
    subject = "【代收資訊確認】COMET小小代購💫💟｜" + data.groupName;
    var pName = data.passportName || "（請參照您填寫的護照姓名）";
    content += 
      "📛 護照姓名： " + pName + "\n" +
      "📝 代收明細：\n" + data.detail + "\n" +
      divider +
      "⚠️ 韓國集運地址請到網頁上查看\n💡 官方出貨後，請到原網頁點擊'查看資訊'回填資料，以利核對～謝謝💟";
  } 
  else if (mode === "3") {
    // 模式三：僅登記
    subject = "【登記成功確認】COMET小小代購💫💟｜" + data.groupName;
    content += 
      "📝 登記明細：\n" + data.detail + "\n" +
      divider +
      "💡 此團務目前僅供登記，後續請留意LINE社群通知～謝謝💟";
  } 
  else {
    // 模式一：國內團
    content += 
      "📝 訂單明細：\n" + data.detail + "\n" +
      divider +
      "💡 匯款完成後，請到原網頁點擊'查看資訊'回填資料，以利核對～謝謝💟";
  }

  try {
    MailApp.sendEmail({ 
      to: data.userEmail, 
      subject: subject, 
      body: body + content,
      name: "COMET小小代購💫💟" 
    });
    return true;
  } catch (e) { 
    console.error("發信失敗: " + e.toString()); 
  }
}