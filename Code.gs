/**
 * สุวรรณ การค้า - Google Apps Script (Backend API + Daily Email Report)
 * Target Email: sasithorn.klaysuwan@gmail.com
 * Time Trigger: 21:00 น. ของทุกวัน
 */

const TARGET_EMAIL = 'sasithorn.klaysuwan@gmail.com';

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
  
  const headers = data[0];
  const items = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const item = {};
    headers.forEach((h, index) => {
      item[h] = row[index];
    });
    items.push(item);
  }
  return ContentService.createTextOutput(JSON.stringify(items)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const contents = JSON.parse(e.postData.contents);
    sendDailyReportEmail(contents);
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: 'Report sent via email' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * ฟังก์ชันส่งอีเมลรายงานสรุปยอดขายประจำวัน
 */
function sendDailyReportEmail(reportData) {
  const todayStr = new Date().toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' });
  
  const data = reportData || {
    storeName: 'สุวรรณ การค้า',
    date: todayStr,
    todayTotal: 0,
    billCount: 0,
    cashSum: 0,
    transferSum: 0,
    thongfahSum: 0,
    thaiChueaySum: 0
  };

  const subject = `📊 รายงานสรุปยอดขายประจำวัน ร้านสุวรรณ การค้า (${data.date || todayStr})`;
  
  const htmlBody = `
    <div style="font-family: 'Prompt', Arial, sans-serif; max-width: 600px; margin: 0 auto; background-color: #ffffff; border: 2px solid #0284c7; border-radius: 20px; overflow: hidden; color: #0f172a;">
      <div style="background-color: #0f172a; color: #ffffff; padding: 24px; text-align: center;">
        <h1 style="margin: 0; font-size: 26px;">🏪 สุวรรณ การค้า</h1>
        <p style="margin: 6px 0 0 0; color: #38bdf8; font-size: 16px;">รายงานสรุปยอดขายประจำวัน</p>
        <p style="margin: 4px 0 0 0; font-size: 14px; color: #94a3b8;">วันที่: ${data.date || todayStr}</p>
      </div>

      <div style="padding: 24px;">
        <div style="background-color: #f0f9ff; border: 2px solid #0284c7; border-radius: 16px; padding: 20px; text-align: center; margin-bottom: 20px;">
          <span style="font-size: 14px; color: #0369a1; display: block;">ยอดขายรวมทั้งหมด</span>
          <span style="font-size: 38px; font-weight: bold; color: #0284c7;">฿${Number(data.todayTotal || 0).toLocaleString('th-TH', { minimumFractionDigits: 2 })}</span>
          <span style="font-size: 14px; color: #475569; display: block; margin-top: 4px;">จำนวนบิลทั้งหมด: <strong>${data.billCount || 0} บิล</strong></span>
        </div>

        <h3 style="font-size: 18px; color: #0f172a; border-bottom: 2px solid #e2e8f0; padding-bottom: 8px; margin-bottom: 16px;">
          💳 รายละเอียดแยกตามช่องทางชำระเงิน:
        </h3>

        <table style="width: 100%; border-collapse: collapse; font-size: 16px;">
          <tr style="border-bottom: 1px solid #e2e8f0;">
            <td style="padding: 12px 0;">💵 <strong>เงินสด (Cash):</strong></td>
            <td style="padding: 12px 0; text-align: right; color: #0284c7; font-weight: bold;">฿${Number(data.cashSum || 0).toLocaleString('th-TH', { minimumFractionDigits: 2 })}</td>
          </tr>
          <tr style="border-bottom: 1px solid #e2e8f0;">
            <td style="padding: 12px 0;">📲 <strong>โอนเงิน / QR Code:</strong></td>
            <td style="padding: 12px 0; text-align: right; color: #0284c7; font-weight: bold;">฿${Number(data.transferSum || 0).toLocaleString('th-TH', { minimumFractionDigits: 2 })}</td>
          </tr>
          <tr style="border-bottom: 1px solid #e2e8f0;">
            <td style="padding: 12px 0;">🏷️ <strong>ธงฟ้าประชารัฐ:</strong></td>
            <td style="padding: 12px 0; text-align: right; color: #0284c7; font-weight: bold;">฿${Number(data.thongfahSum || 0).toLocaleString('th-TH', { minimumFractionDigits: 2 })}</td>
          </tr>
          <tr style="border-bottom: 1px solid #e2e8f0;">
            <td style="padding: 12px 0;">🇹🇭 <strong>โครงการไทยช่วยไทย:</strong></td>
            <td style="padding: 12px 0; text-align: right; color: #0284c7; font-weight: bold;">฿${Number(data.thaiChueaySum || 0).toLocaleString('th-TH', { minimumFractionDigits: 2 })}</td>
          </tr>
        </table>
      </div>

      <div style="background-color: #f8fafc; padding: 16px; text-align: center; border-top: 1px solid #e2e8f0; font-size: 13px; color: #64748b;">
        ระบบ POS ร้านค้าผู้สูงอายุ สุวรรณ การค้า | ส่งอัตโนมัติประจำวันเวลา 21:00 น.
      </div>
    </div>
  `;

  MailApp.sendEmail({
    to: TARGET_EMAIL,
    subject: subject,
    htmlBody: htmlBody
  });
}

/**
 * ตั้งเวลาส่งรายงานอัตโนมัติทุกวัน เวลา 21:00 น. (Time-driven Trigger)
 * (กด้รันฟังก์ชันนี้ครั้งเดียวใน Apps Script เพื่อเปิดใช้งาน Trigger)
 */
function setupDailyReportTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'sendDailyReportEmailAt2100') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('sendDailyReportEmailAt2100')
    .timeBased()
    .everyDays(1)
    .atHour(21)
    .create();

  Logger.log('ตั้งเวลาส่งรายงานสรุปยอดขายเข้าอีเมลทุกวันเวลา 21:00 น. สำเร็จ!');
}

function sendDailyReportEmailAt2100() {
  sendDailyReportEmail();
}
