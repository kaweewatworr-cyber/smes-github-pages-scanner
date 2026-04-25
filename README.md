# SMEs GitHub Pages Scanner

ระบบเช็กอิน QR ที่ใช้งานผ่านเว็บได้ทันที โดยให้ `GitHub Pages` เป็น frontend และ `Google Apps Script Web App` เป็น backend

แนวคิดหลักของชุดนี้คือ:
- หน้าเว็บถูก deploy ครั้งเดียว แล้วส่งต่อเป็นลิงก์ให้ทีมหน้างานใช้ได้เลย
- backend ของแต่ละงานยังแยกกันได้ตาม Google Sheet / Apps Script ของทีมนั้น
- หน้าเว็บสร้างลิงก์พร้อมใช้งานจาก `apiUrl` อัตโนมัติ และสามารถแนบ `staffName` เข้าไปในลิงก์ได้

## ไฟล์หลัก
- `index.html` หน้า scanner หลักสำหรับใช้งานหน้างาน
- `app.js` logic ฝั่งเว็บสำหรับเชื่อม backend, สแกน QR, และสร้าง launch link
- `style.css` UI ของหน้าเว็บ
- `welcome.html` หน้า launch portal สำหรับเปิดหรือคัดลอกลิงก์พร้อมใช้งาน
- `Code.gs` backend มาตรฐานบน Google Apps Script

## วิธีใช้งานแบบลิงก์เดียว

เมื่อหน้าเว็บรู้ `Apps Script Web App URL` แล้ว ระบบจะสร้างลิงก์ใช้งานแบบพร้อมเปิดได้ทันที เช่น

```text
https://your-site.github.io/smes-github-pages-scanner/?apiUrl=https%3A%2F%2Fscript.google.com%2Fmacros%2Fs%2Fxxx%2Fexec
```

ถ้าต้องการระบุเจ้าหน้าที่หรือจุดสแกนเพิ่ม:

```text
https://your-site.github.io/smes-github-pages-scanner/?apiUrl=https%3A%2F%2Fscript.google.com%2Fmacros%2Fs%2Fxxx%2Fexec&staffName=Desk%20A
```

เมื่อผู้ใช้เปิดลิงก์นี้:
- หน้าเว็บจะเชื่อม backend ให้เอง
- ไม่ต้องวาง URL ซ้ำที่เครื่องปลายทาง
- ถ้ามี `staffName` จะส่งไปตอน check-in ด้วย

## Deploy Frontend ขึ้น GitHub Pages
1. สร้าง repository แล้ว push ไฟล์ชุดนี้ขึ้น GitHub
2. เปิด `Settings > Pages`
3. เลือก branch ที่ต้องการ publish
4. รอจนได้ public URL เช่น `https://your-site.github.io/smes-github-pages-scanner/`
5. เปิดหน้าเว็บนั้นแล้ววาง `Apps Script Web App URL` ที่หน้า `index.html`
6. ใช้ปุ่ม `คัดลอกลิงก์` หรือ `แชร์ลิงก์` เพื่อส่งต่อให้ทีมหน้างาน

## ตั้งค่า Backend ใน `Code.gs`

ค่าแนะนำที่ควรตั้ง:
- `spreadsheetId`
- `sheetName`
- `apiName`
- `eventName`
- `eventDateText`
- `eventTimeText`
- `eventLocationText`
- `defaultCheckedInBy`

ค่าใหม่ที่เกี่ยวกับการใช้งานผ่านเว็บ:
- `frontendBaseUrl`
  ใส่ URL ของ GitHub Pages เช่น `https://your-site.github.io/smes-github-pages-scanner/`
- `checkinWebAppUrl`
  optional สำหรับ override เป็นลิงก์หน้า scanner แบบตายตัว ถ้าปล่อยว่าง ระบบจะพยายามสร้างลิงก์จาก `frontendBaseUrl` + URL ของ Apps Script ที่ deploy อยู่

## สิ่งที่ backend ส่งกลับมาเพิ่ม

`GET ?action=status` จะตอบข้อมูลสำหรับหน้าเว็บ เช่น
- `eventName`
- `eventDateText`
- `eventTimeText`
- `eventLocationText`
- `scannerBaseUrl`
- `scannerUrl`
- `defaultStaffName`

ตัวอย่าง response แบบย่อ:

```json
{
  "ok": true,
  "apiName": "Adaptive Intelligence 2026 Check-in API",
  "eventName": "Adaptive Intelligence 2026",
  "scannerBaseUrl": "https://your-site.github.io/smes-github-pages-scanner/",
  "scannerUrl": "https://your-site.github.io/smes-github-pages-scanner/?apiUrl=https%3A%2F%2Fscript.google.com%2Fmacros%2Fs%2Fxxx%2Fexec",
  "actions": ["status", "checkin", "lookup"]
}
```

## Deploy Apps Script
1. เปิด Apps Script
2. วางโค้ดจาก `Code.gs`
3. แก้ `CONFIG`
4. Deploy > New deployment > Web app
5. Execute as: `Me`
6. Who has access: `Anyone`
7. คัดลอก URL ที่ลงท้ายด้วย `/exec`

## Flow แนะนำสำหรับใช้งานจริง
1. Deploy `Code.gs` เป็น Web App
2. Deploy frontend ขึ้น GitHub Pages
3. ใส่ `frontendBaseUrl` ใน `Code.gs`
4. เปิดหน้า `index.html` แล้วทดสอบการเชื่อมต่อ
5. คัดลอกลิงก์พร้อมใช้งานจากหน้าเว็บหรือเปิด `welcome.html`
6. ส่งลิงก์ให้เจ้าหน้าที่หน้างานเปิดจากมือถือ/แท็บเล็ต

## ข้อควรทดสอบก่อนวันงาน
- เปิดผ่านมือถือจริงและยอมรับสิทธิ์กล้อง
- เช็กอินสำเร็จ
- เช็กอินซ้ำ
- ไม่พบข้อมูลผู้ลงทะเบียน
- สแกนจากรูปภาพ
- กรอก `Registration ID` เอง
