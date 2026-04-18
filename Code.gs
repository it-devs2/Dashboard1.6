/**
 * สคริปต์นี้ดึงข้อมูลจาก Google Sheets ด้วยวิธี getDisplayValues (ปลอดภัยที่สุด)
 */

const CACHE_KEY = 'dashboard_data_v3';
const CACHE_DURATION = 120; // 2 นาที

const TH_MONTHS = ['ม.ค.','ก.พ.','มี.ค.','เม.ย.','พ.ค.','มิ.ย.','ก.ค.','ส.ค.','ก.ย.','ต.ค.','พ.ย.','ธ.ค.'];

function doGet(e) {
  // *** ตรวจสอบชื่อ Sheet ให้ตรงเป๊ะที่นี่ ***
  const sheetName = 'เจ้าหนี้อื่นๆ-รถเจาะไทย';
  const forceRefresh = e && e.parameter && e.parameter.refresh === 'true';

  try {
    const cache = CacheService.getScriptCache();
    if (!forceRefresh) {
      const cached = cache.get(CACHE_KEY);
      if (cached) return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({
        status: 'error',
        message: 'ไม่พบหน้า Sheet: ' + sheetName + '. กรุณาตรวจสอบชื่อแถบด้านล่าง!'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ใช้ getDisplayValues() เพื่อดึงค่า "ตามที่ตาเห็น" เป็นข้อความทั้งหมด (ตัดปัญหาวันที่และตัวเลข)
    const dataValues = sheet.getRange(2, 1, lastRow - 1, 17).getDisplayValues();
    const jsonData = [];
    
    for (let i = 0; i < dataValues.length; i++) {
        const row = dataValues[i];
        
        // row[5] คือจำนวนเงิน (คอลัมน์ F)
        const amountStr = row[5].replace(/[^0-9.-]/g, '');
        const amount = parseFloat(amountStr) || 0;
        if (amount === 0) continue;

        // แยก วัน/เดือน/ปี จากคอลัมน์ G (row[6])
        // โดยปกติจะมาเป็น dd/mm/yyyy หรือ yyyy-mm-dd ตามที่ตั้งในเครื่อง
        let dayDue = '', monthDue = '', yearDue = '';
        const dateStr = row[6].trim();

        if (dateStr.includes('/') || dateStr.includes('-')) {
          const sep = dateStr.includes('/') ? '/' : '-';
          const parts = dateStr.split(sep);
          
          if (parts.length === 3) {
            // กรณี dd/mm/yyyy
            if (parts[0].length <= 2) {
                dayDue = parts[0].padStart(2, '0');
                const mIdx = parseInt(parts[1]) - 1;
                monthDue = TH_MONTHS[mIdx] || '';
                yearDue = parts[2];
            } 
            // กรณี yyyy-mm-dd
            else {
                yearDue = parts[0];
                const mIdx = parseInt(parts[1]) - 1;
                monthDue = TH_MONTHS[mIdx] || '';
                dayDue = parts[2].padStart(2, '0');
            }
          }
        }

        // Fallback ไปใช้คอลัมน์ O, P, Q ถ้าใน G ไม่มีวันที่ขัดเจน
        if (!dayDue) dayDue = row[14].trim();
        if (!monthDue) monthDue = row[15].trim();
        if (!yearDue) yearDue = row[16].trim();

        // แยก วัน/เดือน/ปี จากคอลัมน์ H (row[7]) = วันที่ทำเอกสารจ่าย
        let payDocDay = '', payDocMonth = '', payDocYear = '';
        const payDocDateStr = row[7].trim();

        if (payDocDateStr.includes('/') || payDocDateStr.includes('-')) {
          const pSep = payDocDateStr.includes('/') ? '/' : '-';
          const pParts = payDocDateStr.split(pSep);
          
          if (pParts.length === 3) {
            if (pParts[0].length <= 2) {
                payDocDay = pParts[0].padStart(2, '0');
                const pMIdx = parseInt(pParts[1]) - 1;
                payDocMonth = TH_MONTHS[pMIdx] || '';
                payDocYear = pParts[2];
            } else {
                payDocYear = pParts[0];
                const pMIdx = parseInt(pParts[1]) - 1;
                payDocMonth = TH_MONTHS[pMIdx] || '';
                payDocDay = pParts[2].padStart(2, '0');
            }
          }
        }

        jsonData.push({
            creditor: row[3].trim(),
            description: row[2].trim(),
            docNo: row[1].trim(),
            amount: amount,
            paymentStatus: row[9].trim(),
            category: row[12].trim(),
            status: row[13].trim(),
            overdueDays: parseFloat(row[10].replace(/[^0-9.-]/g, '')) || 0,
            dayDue,
            monthDue,
            yearDue,
            payDocDay,
            payDocMonth,
            payDocYear
        });
    }

    const result = JSON.stringify({ status: 'success', data: jsonData });
    
    // พยายามเก็บ cache, ถ้าข้อมูลใหญ่เกิน 100KB ก็ข้ามไป
    try {
      cache.put(CACHE_KEY, result, CACHE_DURATION);
    } catch (cacheError) {
      // ข้อมูลใหญ่เกินไปสำหรับ cache — ไม่เป็นไร ทำงานต่อได้
    }
    
    return ContentService.createTextOutput(result).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'ข้อผิดพลาดภายใน: ' + error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
