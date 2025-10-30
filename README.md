# Tool Holder Control Record — Excel VBA Macros

ระบบ Excel VBA สำหรับควบคุมการกรอกและสิทธิ์บนชีท **Tool Holder Control Record**  
ทำงานตั้งแต่แถว **7 → ∞** โดยอัปเดต “ล็อก/สี/ล้างค่า” แบบเรียลไทม์ และรีเซ็ตสิทธิ์อัตโนมัติเมื่อเปิดไฟล์หรือกลับเข้าชีท

> 🔐 **VBA Project Lock Code (รหัสเข้าดู/แก้โค้ด VBA): `5678`**  
> 🔒 **Sheet Protect Password (ค่าเริ่มต้นในโค้ด): `1234`** — เปลี่ยนได้ที่ตัวแปร `PWD`

---

## Features
- **Start at row 7 → ∞** (กำหนดได้ด้วย `START_ROW`)
- **Sheet-agnostic**: ไม่ผูกชื่อชีท ใช้กับแท็บปัจจุบันได้ทันที
- **Per-row locking**: ปรับสิทธิ์เฉพาะแถวที่ถูกแก้ไข ลดผลข้างเคียง
- **Auto clear & highlights** ตามกติกา (NG/RP/YES/NO)
- **K = NO helper**: แจ้งเตือนภาษาอังกฤษ + ไฮไลท์ L/M แบบรายเซลล์ และหายเหลืองทันทีเมื่อมีการกรอก

---

## Rules (Business Logic)
- **C = NG**
  - ล็อกทั้งแถว `B:M` **ยกเว้น C**
  - แถวเป็น **สีเทา**
  - เคลียร์และ **ล็อก** `K:L:M`
- **C = RP**
  - พฤติกรรม **เหมือน NG ทุกอย่าง** แต่แถวเป็น **สีเหลือง**
- **C = ค่าอื่น**
  - โดยปกติพิมพ์ได้ที่ `C/K/L/M`
  - **K = YES** → เคลียร์ **L/M** และ **ล็อก L/M**
  - **K = NO** → แจ้งเตือน “Input required” และ **ไฮไลท์ L/M** เป็นรายเซลล์จนกว่าจะกรอก

> หมายเหตุ: เงื่อนไขทั้งหมดเริ่มมีผลตั้งแต่ **แถว 7** เป็นต้นไป

---

## Installation
1. เปิด **VBA Editor** (`Alt+F11`)
2. วาง **Sheet Code** ลงในชีทที่ใช้งาน (แท็บที่มีตาราง)
3. ไปที่ `Insert → Module` แล้ววาง **Module Code**
4. (แนะนำ) วางโค้ดใน **ThisWorkbook** เพื่อรีเซ็ตสิทธิ์เมื่อเปิดไฟล์/เข้าแผ่นงาน
5. กลับ Excel แล้วรันแมโคร `ApplyPermissionsAll` หนึ่งครั้ง

---

## Configuration
แก้ค่าที่ส่วนบนของ **Module**:
```vba
Public Const PWD As String        = "1234"  ' รหัส Protect/Unprotect ชีท
Public Const START_ROW As Long    = 7       ' แถวเริ่มทำงาน
' VBA Project Lock Code (รหัสเข้าดู/แก้โค้ดใน VBA Editor) = 5678
