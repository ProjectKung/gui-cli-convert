# GUI Convert (เว็บรวมไฟล์ Log)

เว็บแอปสำหรับรวมไฟล์ 3 ส่วนให้เป็นไฟล์ผลลัพธ์เดียว (PDF หรือ DOCX):

1. ไฟล์ Config/FDO log
2. ไฟล์ APIC log
3. รูปภาพ Show Log

## ลำดับการรวมข้อมูล

1. เนื้อหา FDO/Config ตั้งแต่ต้นไฟล์จนถึงบรรทัดก่อน `show environment` ตัวแรก
2. เนื้อหา APIC ทั้งหมด
3. เนื้อหา FDO/Config ส่วนที่เหลือ
4. หน้ารูปภาพพร้อมหัวข้อ `Show log` แบบไฮไลต์

บรรทัดคำสั่งลักษณะ `...# show ...` จะถูกไฮไลต์สีเหลืองในไฟล์ PDF

ชื่อไฟล์ผลลัพธ์จะอิงจากชื่อไฟล์ Config/FDO อัตโนมัติ
ตัวอย่าง: `FDO25040LT2.log` -> `FDO25040LT2.pdf` หรือ `FDO25040LT2.docx`

## ไฟล์สำคัญในโปรเจกต์

- `app.py`: เว็บแอป Flask
- `merge_logs_to_pdf.py`: ตรรกะการรวมไฟล์และสร้าง PDF
- `templates/index.html`: หน้าอัปโหลดไฟล์
- `static/site.css`: สไตล์หน้าเว็บ
- `requirements.txt`: รายการ dependencies
- `setup_and_run.ps1`: สคริปต์ติดตั้งและรัน (Windows)
- `run_web.bat`: ไฟล์สำหรับดับเบิลคลิกเพื่อเริ่มโปรแกรม

## การใช้งานบนเครื่องใหม่ (Windows)

1. ติดตั้ง Python 3.11 ขึ้นไป
2. คัดลอกโฟลเดอร์โปรเจกต์นี้ไปยังเครื่อง
3. รัน `run_web.bat`
4. เปิด `http://127.0.0.1:5000`
5. อัปโหลดไฟล์ทั้ง 3 ไฟล์ แล้วกดสร้างไฟล์ผลลัพธ์

หากเครื่องมี `winget` ระบบใน `run_web.bat`/`setup_and_run.ps1` สามารถติดตั้ง Python ให้อัตโนมัติได้

## การใช้งานผ่าน CLI

```powershell
python merge_logs_to_pdf.py `
  --fdo "c:\path\FDO25040LT2.log" `
  --apic "c:\path\apic.log" `
  --image "c:\path\showlog.jpg" `
  --outdir ".\output"
```
