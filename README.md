# Bright Greek — แอปสั่ง Greek Yogurt

- **หน้าแอปสำหรับลูกค้า:** https://bright08913-cmyk.github.io/Greek-Yogurt-N-Bright/ (โฟลเดอร์ `docs/`)
- **หลังบ้าน:** Google Sheet "Bright Tracker" + Apps Script (`Code.gs`)

## ไฟล์
| ไฟล์ | ใช้ทำอะไร |
|---|---|
| `Code.gs` | วางใน Apps Script ของ Sheet — รับออเดอร์, เก็บสลิป, ส่งเมนูให้แอป |
| `Index.html` | หน้าแอป (วางใน Apps Script เป็นไฟล์ HTML ชื่อ `Index` ด้วย ใช้เป็นสำรอง) |
| `build.py` | สร้างหน้าเว็บใน `docs/` จาก `Index.html` + ลิงก์ใน `api_url.txt` |
| `api_url.txt` | ลิงก์ `/exec` ของ Apps Script |

## แก้โค้ดแล้ว
1. วาง `Code.gs` / `Index.html` ใน Apps Script → Deploy → Manage deployments → ✏️ → New version
2. `python3 build.py` แล้ว push — หน้าเว็บอัปเดตเองใน 1–2 นาที

## ใช้งานทุกวัน (แก้ใน Sheet ไม่ต้องแตะโค้ด)
| อยากทำ | แก้ที่ |
|---|---|
| เปิด/ปิดวันรับ | แท็บ **Rounds** ติ๊ก "เปิดรับ" (ป้าย "รอบส่ง" คำนวณเอง) |
| ปิดรับออเดอร์ทั้งหมด | Settings → `order_open` |
| ราคา / เมนู / ผลไม้หมด | แท็บ Menu / AddOns / Fruits |
| ผลไม้ใหม่ประจำสัปดาห์ | แท็บ Fruits พิมพ์ชื่อ + ติ๊ก "มีของ" — แอปเลือกรูปจากชื่อเอง (ถ้าไม่รู้จักจะเป็นลูกกลมสีพาสเทล) จะเลือก "รูปทรง"/"สี" เองก็ได้ |
| เช็คสลิป | แท็บ **Orders** คอลัมน์ "สลิป" (รูปย่อ) → เปลี่ยนสถานะจ่ายเป็น "จ่ายแล้ว" |
| ทำบิลขุนทอง | แท็บ **Summary** ช่อง A43 (เฉพาะคนที่เลือก "จ่ายตอนรับ") |

สลิปเก็บในโฟลเดอร์ Google Drive "Bright Greek Slips" (เปิดดูได้เฉพาะคนที่มีลิงก์)
