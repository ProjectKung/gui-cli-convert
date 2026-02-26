# Switch Converter (GUI + CLI)

เครื่องมือแปลงและรวมไฟล์ Log ของ Switch ให้เป็นไฟล์ผลลัพธ์เดียว โดยรองรับทั้ง:

- โหมดเว็บ GUI (`/gui`) สำหรับอัปโหลด 3 ไฟล์แล้วดาวน์โหลดผลลัพธ์ทันที
- โหมดเว็บ CLI (`/cli`) สำหรับแปลงไฟล์ข้อความแบบ interactive ในหน้าเดียว
- โหมดสคริปต์ Python (`merge_logs_to_pdf.py`) สำหรับใช้งานผ่าน command line

รองรับผลลัพธ์เป็น `PDF` และ `DOCX` พร้อมไฟล์ข้อความรวม (`.txt`) ในโฟลเดอร์ `output/`

## สรุปเร็ว (Quick Start)

### วิธีที่ง่ายสุดบน Windows

1. ดับเบิลคลิก `run_web.bat` (โหมด Console)
2. รอระบบเตรียม Python/venv/dependencies อัตโนมัติ
3. เปิด `http://127.0.0.1:5000`
4. เลือกโหมด `GUI` หรือ `CLI` จากหน้าแรก

### โหมดไม่มีหน้าต่าง Console (Tray)

- ดับเบิลคลิก `run_web.vbs`
- ระบบจะใช้ `run_web_launcher.ps1` เพื่อรันแบบพื้นหลังและควบคุมผ่าน tray icon

## โปรเจกต์นี้ทำอะไร

สำหรับงานรวมไฟล์ Log หลัก 3 ส่วน:

1. ไฟล์ Config/FDO (`.log/.txt/.cfg/.conf`)
2. ไฟล์ APIC (`.log/.txt/.cfg/.conf`)
3. รูป Show Log (`.png/.jpg/.jpeg/.bmp/.gif/.webp`)

และสร้างไฟล์ผลลัพธ์แบบ:

- `PDF` หรือ
- `DOCX`

โดยชื่อไฟล์ผลลัพธ์จะอ้างอิงจากชื่อไฟล์ FDO อัตโนมัติ  
ตัวอย่าง: `FDO25040LT2.log` -> `FDO25040LT2.pdf` หรือ `FDO25040LT2.docx`

## เงื่อนไขระบบ

- Windows 10/11 (สคริปต์ launcher ถูกออกแบบมาสำหรับ Windows)
- Python 3.11 ขึ้นไป
- อินเทอร์เน็ต (เฉพาะรอบแรก ถ้าต้องติดตั้ง Python/dependencies หรืออัปเดตโค้ด)

Dependencies หลักใน `requirements.txt`:

- `Flask==3.1.0`
- `Pillow==11.1.0`

## วิธีรันทั้งหมด

### 1) รันผ่าน `run_web.bat` (แนะนำ)

คำสั่งนี้จะเรียก `setup_and_run.ps1 -AutoInstallPython` โดยอัตโนมัติ ซึ่งทำงานดังนี้:

- ตรวจอัปเดตโค้ดจาก GitHub branch `main` (self-update)
- หา/ติดตั้ง Python 3.11+
- สร้างและซ่อมแซม virtual environment (`.venv`) ถ้าจำเป็น
- ติดตั้ง dependencies จาก `requirements.txt`
- สตาร์ต Flask ที่ `127.0.0.1:5000`
- เปิดเบราว์เซอร์ให้อัตโนมัติ

### 2) รันผ่าน Tray (`run_web.vbs`)

ใช้ `run_web_launcher.ps1` เพื่อ:

- รันแบบซ่อนหน้าต่าง console
- มีสถานะการเริ่มงานผ่าน tray icon
- คุมการเปิดเว็บ/หยุด server จากเมนู tray

### 3) รันด้วย PowerShell โดยตรง

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\setup_and_run.ps1 -AutoInstallPython
```

พารามิเตอร์สำคัญของ `setup_and_run.ps1`:

- `-NoBrowser` ไม่เปิดเบราว์เซอร์อัตโนมัติ
- `-AutoInstallPython` อนุญาตให้สคริปต์ติดตั้ง Python ให้เอง
- `-SkipSelfUpdate` ข้ามการดึงอัปเดตจาก GitHub
- `-StopOldServer` (ค่าเริ่มต้นเปิดอยู่) พยายามหยุดโปรเซสเก่าบนพอร์ต 5000

### 4) รันแบบ manual (กรณีไม่ใช้ launcher)

```powershell
python -m venv .venv
.\.venv\Scripts\python -m pip install -r requirements.txt
.\.venv\Scripts\python app.py
```

จากนั้นเปิด `http://127.0.0.1:5000`

## การใช้งานหน้าเว็บ

### เส้นทางหลัก

- `/` หน้าเลือกโหมด
- `/gui` หน้าอัปโหลด 3 ไฟล์เพื่อสร้าง PDF/DOCX
- `/cli` หน้าเครื่องมือ CLI แบบเว็บ (ไฟล์ static)
- `/health` เช็กสถานะเซิร์ฟเวอร์
- `/validation-report/<report_id>` ดึงรายงานตรวจสอบล่าสุด

### วิธีใช้โหมด GUI (`/gui`)

1. อัปโหลดไฟล์ FDO, APIC, และภาพ Show Log ให้ครบ
2. เลือกรูปแบบผลลัพธ์ (`PDF` หรือ `DOCX`)
3. (เลือกได้) เปิดโหมด `กำหนดวัน/เวลาเอง` เพื่อกำหนดช่วงเวลา clock
4. กดสร้างไฟล์ แล้วระบบจะดาวน์โหลดให้อัตโนมัติ
5. รายงาน Validation จะแสดงบนหน้าเว็บหลังสร้างไฟล์

ข้อจำกัดสำคัญ:

- รองรับไฟล์อัปโหลดรวมไม่เกิน `100 MB` ต่อ request (`MAX_CONTENT_LENGTH`)
- เซิร์ฟเวอร์ผูกที่ `127.0.0.1` (local only)

## Logic การประมวลผล Log

ระบบ preprocess ไฟล์ FDO ก่อนรวม:

1. ลบบรรทัดที่มีคำว่า `clear`
2. ใน section `show interface counters errors` จะปรับค่าตัวเลขเป็น `--`
3. ปรับเวลา `show clock` จุดที่ 1-3
4. ตรวจลำดับ/จำนวนคำสั่ง และสร้าง validation report

การปรับ `show clock`:

- ถ้าไม่เปิด custom mode: `clock#1` คงค่าเดิมจากไฟล์ต้นฉบับ
- ถ้าเปิด custom mode: สุ่ม `clock#1` ภายในวันที่/ช่วงเวลาที่ผู้ใช้กำหนด
- จากนั้นระบบกำหนด:
  - `clock#2 = clock#1 + 40..90 วินาที`
  - `clock#3 = clock#2 + 420..450 วินาที`

การรวม FDO + APIC:

- แทรกเนื้อหา APIC เข้าไปก่อน `show environment` ตัวแรกของ FDO (ตามเงื่อนไข fallback ในโค้ด)
- เติมส่วนที่เหลือของ FDO ต่อท้าย
- ตอน export PDF/DOCX จะมีหน้าหัวข้อ `Show log` พร้อมรูปภาพท้ายเอกสาร
- บรรทัดคำสั่งลักษณะ `...# show ...` ถูกทำไฮไลต์ใน PDF

## Validation Report ตรวจอะไรบ้าง

รายงานจะประเมิน `ผ่าน/ไม่ผ่าน` จากหลายเงื่อนไข เช่น:

- จำนวนคำสั่งที่ต้องมี:
  - `show clock` = 3
  - `show version` = 1
  - `show running-config` = 1
  - `show environment` = 1
  - `show interface counters errors` = 2
- ลำดับคำสั่งหลักตามแพทเทิร์นที่ระบบกำหนด
- ตำแหน่งของ `show interface counters errors` เทียบกับ clock#2 และ clock#3
- ช่วงเวลา `clock#1->#2` และ `clock#2->#3` อยู่ในกรอบที่กำหนด
- ความครบถ้วนของส่วน `Show log` ท้ายเอกสาร
- รายการบรรทัดที่ระบบแก้ไขจริง (เช่น interface rows)

## โหมด CLI แบบสคริปต์ Python

ไฟล์: `merge_logs_to_pdf.py`

```powershell
python .\merge_logs_to_pdf.py `
  --fdo "C:\path\FDO25040LT2.log" `
  --apic "C:\path\apic.log" `
  --image "C:\path\showlog.jpg" `
  --outdir ".\output" `
  --format pdf
```

พารามิเตอร์ที่รองรับ:

- `--fdo` path ไฟล์ FDO
- `--apic` path ไฟล์ APIC
- `--image` path ไฟล์รูป
- `--outdir` โฟลเดอร์ผลลัพธ์
- `--format` `pdf` หรือ `docx`
- `--pdf-name` ชื่อไฟล์ PDF (override)
- `--docx-name` ชื่อไฟล์ DOCX (override)
- `--text-name` ชื่อไฟล์ TXT (override)

ผลลัพธ์ที่ได้:

- สร้างไฟล์ `.txt` รวมเสมอ
- สร้างไฟล์ `.pdf` หรือ `.docx` ตาม `--format`

## โหมด CLI แบบหน้าเว็บ (`/cli`)

หน้า `/cli` เป็นเครื่องมือแปลงไฟล์ข้อความแบบ interactive โดยทำงานใน browser:

- รองรับเลือกหลายไฟล์
- แสดง diff ก่อน/หลังแปลง
- ดาวน์โหลดผลลัพธ์จากหน้าเว็บได้
- มีตัวเลือกกำหนดวัน/เวลา clock เองได้

หมายเหตุ: โหมดนี้แยกจากสคริปต์ `merge_logs_to_pdf.py` และถูกเสิร์ฟเป็น static HTML

## โครงสร้างไฟล์สำคัญ

- `app.py` Flask server และ API endpoint
- `merge_logs_to_pdf.py` แกน logic preprocess/merge/export
- `templates/home.html` หน้าเลือกโหมด
- `templates/index.html` หน้า GUI uploader + validation UI
- `static/cli/txt_log_converter_v20.html` หน้า CLI web tool
- `static/site.css` สไตล์หน้าเว็บ
- `setup_and_run.ps1` setup + run + self-update
- `run_web.bat` ตัวรันแบบ console
- `run_web.vbs` ตัวรันแบบซ่อน console/tray launcher
- `run_web_launcher.ps1` tray controller
- `output/` ไฟล์ผลลัพธ์และไฟล์สถานะ launcher/setup

## Troubleshooting

### รัน `.ps1` ไม่ได้เพราะ policy

ใช้คำสั่ง:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\setup_and_run.ps1 -AutoInstallPython
```

### เปิดเว็บไม่ได้ที่พอร์ต 5000

- ปิดโปรแกรมเก่าที่ใช้พอร์ต 5000
- รันใหม่ผ่าน `run_web.bat` (สคริปต์จะพยายามหยุด server เก่าให้อัตโนมัติ)

### เครื่องไม่มี Python

- ใช้ `run_web.bat` หรือ `setup_and_run.ps1 -AutoInstallPython`
- ถ้าเครือข่ายจำกัด ให้ติดตั้ง Python 3.11+ เองก่อน แล้วรันแบบ manual

### สร้างไฟล์ไม่ผ่าน

- ตรวจนามสกุลไฟล์ให้ตรงชนิดที่ระบบรองรับ
- ดูข้อความ error ในหน้าเว็บ
- ตรวจ `Validation Report` เพื่อหาเงื่อนไขที่ไม่ผ่าน

## หมายเหตุด้านการใช้งาน

- เครื่องมือนี้ออกแบบสำหรับใช้งานในเครื่องภายใน (localhost)
- หากต้อง deploy ให้ผู้ใช้หลายคนผ่านเครือข่ายจริง ควรเพิ่ม security/config เพิ่มเติมก่อนใช้งาน
