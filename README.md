# RPA Auto Test Login

ระบบนี้ใช้ Playwright สำหรับทดสอบการล็อกอินอัตโนมัติ และบันทึกผลการรันลงไฟล์ Excel โดยเก็บเวลา `DateTime` ของการล็อกอินทุกครั้ง

## ไฟล์สำคัญ

- `config/login.config.example.json` ตัวอย่าง config สำหรับหน้า login
- `scripts/auto-test-login.mjs` สคริปต์หลักสำหรับรัน RPA
- `run-auto-test-login.ps1` ตัวช่วยรันด้วย bundled Node.js
- `outputs/auto-test-login/login-log.xlsx` ไฟล์ Excel log ที่ระบบจะสร้างให้อัตโนมัติ

## วิธีตั้งค่า

ติดตั้ง bootstrap ให้พร้อมก่อน:

```powershell
.\setup-auto-test-login.ps1
```

1. สคริปต์ setup จะสร้าง `config/login.config.json` ให้อัตโนมัติถ้ายังไม่มี
2. ใส่ค่า URL, username, password และ selector ของฟอร์ม login ให้ตรงกับระบบจริง

ตัวอย่าง field หลัก:

- `loginUrl` หน้า login
- `usernameSelector` selector ของช่อง username
- `passwordSelector` selector ของช่อง password
- `submitSelector` selector ของปุ่ม login
- `successSelector` selector ที่จะมองเห็นหลัง login สำเร็จ
- `successUrlContains` ส่วนของ URL ที่คาดว่าจะเจอหลัง login สำเร็จ

## วิธีรัน

ทดสอบระบบแบบไม่ต้อง login จริง:

```powershell
.\run-auto-test-login.ps1 --self-test
```

คำสั่งนี้จะเช็กว่า script รันได้ แต่จะไม่เขียนข้อมูลลง Excel

รัน login test จริง:

```powershell
.\run-auto-test-login.ps1
```

## การบันทึก Excel

ระบบจะบันทึกข้อมูลลง Excel หลังจากรันหน้าเว็บจริงแล้วเท่านั้น โดยจะกรอกข้อมูลบนหน้าเว็บก่อน และค่อย append log ผลลัพธ์ลงไฟล์ Excel

## ข้อมูลที่บันทึกลง Excel

- Run ID
- Login DateTime
- Status
- Login URL
- Username
- Browser
- Duration (ms)
- Screenshot
- Message
