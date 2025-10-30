# Tool Holder Control Record – Excel VBA Macros

มาโครสำหรับควบคุมการกรอกข้อมูลและคุมสิทธิ์บนชีท **Tool Holder Control Record**  
โฟกัสช่วงทำงาน `B5:M16` โดยใช้ตรรกะธุรกิจกับคอลัมน์ `C, K, L, M`.

## Features
- **Per-row locking:** โดยปกติพิมพ์ได้ที่ C/K/L/M; เมื่อ `C=NG` ล็อกทั้งแถว B:M ยกเว้น C
- **Auto clear:** เมื่อ `C=NG` เคลียร์ค่า `K:L:M` อัตโนมัติ
- **K=NO helper:** เตือนเป็นภาษาอังกฤษ และทำ **ไฮไลท์เหลืองรายเซลล์** ที่ L และ/หรือ M; พิมพ์ช่องไหน ช่องนั้นหายเหลืองทันที
- **K=YES rule:** ล้างค่า L/M และ **ล็อก L/M** ในแถวนั้น
- **Re-apply on open/activate:** เซ็ตสิทธิ์ให้ถูกต้องทุกครั้งที่เปิดไฟล์หรือกลับมาที่ชีท

## Rules (สรุปกติกา)
- `C=NG` ⇒ แถวนั้นย้อมเทา, เคลียร์ `K:L:M`, ล็อกทั้ง `B:M` **ยกเว้น C**  
- `K=NO` ⇒ แจ้งเตือน “Input required”, L/M เหลืองทีละเซลล์จนกว่าจะกรอก  
- `K=YES` ⇒ ล้างค่า L/M และ **ล็อก L/M**  
- โดยปกติผู้ใช้ **พิมพ์ได้เฉพาะ C/K/L/M** (แถว 5–16) ตั้งแต่เปิดชีท

## Installation
1. เปิด VBA Editor (`Alt+F11`)
2. คลิกขวาแท็บชีทที่มีตาราง → **View Code** → วางไฟล์ *Sheet Code* (จากโปรเจกต์นี้)
3. `Insert → Module` → วางไฟล์ *Module Code*
4. ดับเบิลคลิก **ThisWorkbook** → วาง *ThisWorkbook code* (ตัวเลือก)
5. รันแมโคร `ApplyPermissionsAll` หนึ่งครั้ง

> ค่าเริ่มต้นรหัสผ่าน: `1234` • ชื่อชีท: `For Lathe Tooling`

## Usage
- กรอกข้อมูลที่ C/K/L/M ตามปกติ  
- เปลี่ยน `C` เป็น `NG` จะล็อกทั้งแถว (เว้น C) และล้าง `K:L:M` ทันที  
- ตั้ง `K=NO` จะขึ้นแจ้งเตือน และ L/M เหลืองจนกรอกอย่างน้อยหนึ่งช่อง  
- ตั้ง `K=YES` จะล้างและล็อก L/M  
- หากสิทธิ์รวน ให้รัน `ApplyPermissionsAll` อีกครั้ง

## Macros
- `ApplyPermissionsAll` – เซ็ตสิทธิ์ทั้งชีทตามค่าปัจจุบัน
- `UnlockSheet` – ปลดล็อกชีทชั่วคราว (ถามรหัสผ่าน)
- `ProtectSheet` – ล็อกชีทกลับ (UserInterfaceOnly)
- **ใน Sheet Code (อัตโนมัติ):** ตรวจจับแก้ไข, อัปเดตสี/ล็อก, เตือน K=NO

## Configuration
แก้ใน **Module Code**
```vba
Public Const PWD As String = "1234"
Public Const SHEET_NAME As String = "For Lathe Tooling"
