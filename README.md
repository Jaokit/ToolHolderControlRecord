# Tool Holder Control Record — Excel VBA Macros

มาโครควบคุมการกรอก/สิทธิ์บนชีท **Tool Holder Control Record**  
ทำงานตั้งแต่แถว `7` ไปจนสุดชีท โดยอัปเดตล็อก/สี/การล้างค่าแบบเรียลไทม์ และรีเซ็ตสิทธิ์อัตโนมัติเมื่อเปิดไฟล์หรือกลับเข้าชีท

## Features
- **Start row 7 → ∞** (ปรับได้ด้วย `START_ROW`)
- **Per-row locking** ที่ไม่ผูกชื่อชีท (sheet-agnostic)
- **Auto clear & highlights** ตามกติกา
- **K=NO helper:** แจ้งเตือนภาษาอังกฤษ + ไฮไลท์ L/M แบบรายเซลล์ และหายเหลืองทันทีเมื่อกรอก

## Rules (Business Logic)
- `C = NG`
  - ล็อกทั้งแถว `B:M` **ยกเว้น C**
  - แถวเป็น **สีเทา**
  - เคลียร์และ **ล็อก** `K:L:M`
- `C = RP`
  - พฤติกรรม **เหมือน NG ทุกอย่าง** แต่แถวเป็น **สีเหลือง**
- `C = (อื่นๆ)`
  - พิมพ์ได้ที่ `C/K/L/M`
  - `K = YES` → เคลียร์ **L/M** และ **ล็อก L/M**
  - `K = NO` → แจ้งเตือน “Input required” และ **ไฮไลท์ L/M** ทีละเซลล์จนกว่าจะกรอก

## Installation
1. เปิด VBA Editor (`Alt+F11`)
2. วาง **Sheet Code** ลงในชีทที่ใช้งาน (แท็บที่มีตาราง)
3. `Insert → Module` วาง **Module Code**
4. (แนะนำ) วางโค้ดใน **ThisWorkbook** เพื่อรีเซ็ตสิทธิ์เมื่อเปิดไฟล์/เข้าแผ่นงาน
5. รัน `ApplyPermissionsAll` หนึ่งครั้งเพื่อเซ็ตสิทธิ์เริ่มต้น

## Configuration
ปรับค่าที่ส่วนบนของ Module:
```vba
Public Const PWD As String        = "1234"  ' รหัสสำหรับ Protect/Unprotect ชีท
Public Const START_ROW As Long    = 7       ' แถวเริ่มทำงาน
