# main.py
# LINE + FastAPI + Google Sheets 訂單系統（商用整合版）
import os
import re
from datetime import datetime
from fastapi import FastAPI, Request
from fastapi.responses import PlainTextResponse
from dotenv import load_dotenv
from openpyxl import Workbook
from fastapi.responses import FileResponse
import gspread
from google.oauth2.service_account import Credentials
from linebot import LineBotApi
from linebot.models import TextSendMessage
load_dotenv()
print(os.getenv("GOOGLE_DRIVE_FOLDER_ID"))
app = FastAPI()
BASE_URL = os.getenv(
    "BASE_URL",
    "http://127.0.0.1:8000"
)
LINE_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
SHEET_ID = os.getenv("GOOGLE_SHEET_ID")

line_bot_api = LineBotApi(LINE_TOKEN)

scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

import cloudinary
import cloudinary.uploader
import tempfile

cloudinary.config(
    cloud_name=os.getenv("CLOUDINARY_CLOUD_NAME"),
    api_key=os.getenv("CLOUDINARY_API_KEY"),
    api_secret=os.getenv("CLOUDINARY_API_SECRET"),
    secure=True
)

import uuid

from io import BytesIO

def upload_excel_file(wb):

    filename = (
        f"orders_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )

    tmp_path = None

    try:

        with tempfile.NamedTemporaryFile(
            suffix=".xlsx",
            delete=False
        ) as tmp:

            tmp_path = tmp.name

        wb.save(tmp_path)

        result = cloudinary.uploader.upload(
            tmp_path,
            resource_type="raw",
            folder="order_exports",
            public_id=filename.replace(
                ".xlsx",
                ""
            ),
            overwrite=True
        )

        return result["secure_url"]

    finally:

        if tmp_path and os.path.exists(tmp_path):
            os.remove(tmp_path)

import json

service_account_info = json.loads(
    os.getenv("GOOGLE_CREDENTIALS")
)

creds = Credentials.from_service_account_info(
    service_account_info,
    scopes=scope
)


gc = gspread.authorize(creds)


sheet = gc.open_by_key(SHEET_ID).sheet1
customer_sheet = gc.open_by_key(SHEET_ID).worksheet("Customers")
settings_sheet = gc.open_by_key(SHEET_ID).worksheet("Settings")
DELIVERY_LIST = ["自取","自送","代送","業務自送","寄大榮","寄黑貓","寄順豐","寄梓華榮"]


def export_orders(keyword):

    rows = sheet.get_all_values()

    wb = Workbook()
    ws = wb.active

    ws.title = "Orders"

    ws.append([
        "單號",
        "建立時間",
        "日期",
        "客戶",
        "商品",
        "數量",
        "單位",
        "單價",
        "配送",
        "備註"
    ])

    count = 0

    for r in rows[1:]:

        status = r[11] if len(r) > 11 else ""

        if status == "已刪除":
            continue

        if keyword != "全部":

            if keyword not in " ".join(r):
                continue

        ws.append([
            r[0],
            r[1],
            r[2],
            r[3],
            r[4],
            r[5],
            r[6],
            r[7],
            r[8],
            r[9]
        ])

        count += 1

    if count == 0:
        return None
    return upload_excel_file(wb)




def get_user_name(user_id):

    try:
        profile = line_bot_api.get_profile(user_id)
        return profile.display_name

    except Exception:
        return user_id

def customer_setting(text):

    m = re.match(
        r"客戶設定\s+(\S+)\s+(\S+)\s+(\S+)\s+(\S+)",
        text
    )

    if not m:
        return "❌ 格式: 客戶設定 客戶 商品 數量 單價"

    customer, product, qty_unit, price = m.groups()

    rows = customer_sheet.get_all_values()

    for i, r in enumerate(rows[1:], start=2):

        if r[0] == customer:

            customer_sheet.update(
                f"A{i}:D{i}",
                [[customer, product, qty_unit, price]]
            )

            return "✅ 已更新客戶"

    customer_sheet.append_row([
        customer,
        product,
        qty_unit,
        price
    ])

    return "✅ 已新增客戶"

def customer_query(text):

    customer = text.replace("客戶", "", 1).strip()

    rows = customer_sheet.get_all_values()

    for r in rows[1:]:

        if r[0] == customer:

            return (
                f"客戶:{r[0]}\n"
                f"商品:{r[1]}\n"
                f"數量:{r[2]}\n"
                f"單價:{r[3]}"
            )

    return "❌ 找不到客戶"

from datetime import timedelta
import calendar

def create_schedule_order(text):

    text = text.strip()

    start_next_week = False

    if "下週開始" in text:
        start_next_week = True
        text = text.replace("下週開始", "")

    m = re.match(
        r"(\S+)\s+每週([一二三四五六日]+)到貨到月底",
        text
    )

    if not m:
        return "❌ 排程格式錯誤"

    customer = m.group(1)
    weekdays_text = m.group(2)

    weekday_map = {
        "一":0,
        "二":1,
        "三":2,
        "四":3,
        "五":4,
        "六":5,
        "日":6
    }

    target_days = [
        weekday_map[x]
        for x in weekdays_text
        if x in weekday_map
    ]

    rows = customer_sheet.get_all_values()

    customer_data = None

    for r in rows[1:]:
        if r[0] == customer:
            customer_data = r
            break

    if not customer_data:
        return "❌ 客戶未設定"

    product = customer_data[1]

    qty_match = re.search(r"(\d+(?:\.\d+)?)", customer_data[2])
    qty = float(qty_match.group(1))
    unit = parse_unit(customer_data[2])

    price = float(customer_data[3])

    today = datetime.now().date()

    if start_next_week:
        today += timedelta(days=7)

    last_day = calendar.monthrange(
        today.year,
        today.month
    )[1]

    end_date = today.replace(day=last_day)

    orders = []

    current = today

    while current <= end_date:

        if current.weekday() in target_days:

            orders.append({
                "date": f"{current.month}/{current.day}",
                "customer": customer,
                "product": product,
                "qty": qty,
                "unit": unit,
                "price": price,
                "delivery": "",
                "note": "固定排程"
            })

        current += timedelta(days=1)

    if not orders:
        return "❌ 沒有符合日期"

    count = save_orders_batch(
        orders,
        "SYSTEM"
    )

    return f"✅ 已建立 {count} 筆固定排程訂單"



def edit_order(text):

    rows = sheet.get_all_values()

    # 改單 單號 數量 20
    m = re.match(
        r"改單\s+(\d+)\s+(日期|商品|數量|單價|配送|備註)\s+(.+)",
        text
    )

    if m:

        order_id, field, value = m.groups()

        col_map = {
            "日期": 3,
            "商品": 5,
            "數量": 6,
            "單價": 8,
            "配送": 9,
            "備註": 10
        }

        for i, r in enumerate(rows[1:], start=2):

            if r[0] == order_id:

                sheet.update_cell(
                    i,
                    col_map[field],
                    value
                )

                return "✅ 已修改"

        return "❌ 找不到訂單"

    # 改單 6/1 阿振 單價 350

    m = re.match(
        r"改單\s+(\S+)\s+(\S+)\s+(日期|商品|數量|單價|配送|備註)\s+(.+)",
        text
    )

    if not m:
        return "❌ 格式錯誤"

    order_date, customer, field, value = m.groups()

    col_map = {
        "日期": 3,
        "商品": 5,
        "數量": 6,
        "單價": 8,
        "配送": 9,
        "備註": 10
    }

    updated = 0

    for i, r in enumerate(rows[1:], start=2):

        if r[2] == order_date and r[3] == customer:

            sheet.update_cell(
                i,
                col_map[field],
                value
            )

            updated += 1

    if updated == 0:
        return "❌ 找不到訂單"

    return f"✅ 已修改 {updated} 筆訂單"

def parse_unit(text):
    m = re.search(r"[a-zA-Z\u4e00-\u9fa5]+", text)
    return m.group(0) if m else "件"

def detect_delivery(text):
    for d in DELIVERY_LIST:
        if d in text:
            return d
    return ""

def generate_order_ids(count):

    today = datetime.now()

    prefix = (
        f"{today.year-1911}"
        f"{today.strftime('%m%d')}"
    )

    rows = sheet.get_all_values()

    today_ids = []

    for r in rows[1:]:

        if len(r) == 0:
            continue

        oid = r[0]

        if oid.startswith(prefix):
            today_ids.append(oid)

    seq = len(today_ids) + 1

    ids = []

    for i in range(count):

        ids.append(
            f"{prefix}"
            f"{str(seq+i).zfill(3)}"
        )

    return ids


def generate_order_id():
    return generate_order_ids(1)[0]

def save_order(data, user_id):
    oid = generate_order_id()
    sheet.append_row([
        oid,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        data["date"],
        data["customer"],
        data["product"],
        data["qty"],
        data["unit"],
        data["price"],
        data["delivery"],
        data["note"],
        user_id,
        "正常",
        "",
        "",
        "",
        str(uuid.uuid4())      # P欄
])
    return oid

def save_orders_batch(
    orders,
    user_id
):

    rows = []

    order_ids = generate_order_ids(
        len(orders)
    )

    for oid, order in zip(
        order_ids,
        orders
    ):

        rows.append([
            oid,
            datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
            ),
            order["date"],
            order["customer"],
            order["product"],
            order["qty"],
            order["unit"],
            order["price"],
            order["delivery"],
            order["note"],
            user_id,
            "正常",
            "",
            "",
            "",
            str(uuid.uuid4())
        ])

    sheet.append_rows(rows)

    return len(rows)

def parse_order_line(line):
    parts = line.split()
    if len(parts) < 4:
        return None

    date = parts[0]
    customer = parts[1]
    product = parts[2]

    qty_match = re.search(r"(\d+(?:\.\d+)?)", parts[3])
    if not qty_match:
        return None

    qty = float(qty_match.group(1))
    unit = parse_unit(parts[3])

    price_match = re.search(r"@(\d+(?:\.\d+)?)", line)
    price = float(price_match.group(1)) if price_match else 0

    delivery = detect_delivery(line)

    note = line
    note = re.sub(r"\d{1,2}/\d{1,2}", "", note)
    note = re.sub(r"@\d+(?:\.\d+)?", "", note)
    for d in DELIVERY_LIST:
        note = note.replace(d, "")
    for p in parts[:4]:
        note = note.replace(p, "")
    note = note.strip()

    return {
        "date": date,
        "customer": customer,
        "product": product,
        "qty": qty,
        "unit": unit,
        "price": price,
        "delivery": delivery,
        "note": note
    }

def parse_multi_customer_order(text):

    lines = [x.strip() for x in text.splitlines()]

    orders = []

    current_date = ""
    current_customer = ""

    for line in lines:

        if not line:
            continue

        m = re.match(
            r"^(\d{1,2}/\d{1,2})(?:\s+(.+))?$",
            line
        )

        if m:

            current_date = m.group(1)
            current_customer = m.group(2)

            continue

        if line.startswith("@"):
            continue

        if not current_date:
            continue

        m = re.match(
            r"(.+?)\s+(\d+(?:\.\d+)?)([^\d\s]+)?(?:\s+@(\d+(?:\.\d+)?))?(.*)",
            line
        )

        if not m:
            continue

        product = m.group(1).strip()

        qty = float(m.group(2))

        unit = m.group(3) or "件"

        price = float(m.group(4)) if m.group(4) else 0

        remain = m.group(5).strip()

        delivery = detect_delivery(remain)

        note = remain

        for d in DELIVERY_LIST:
            note = note.replace(d, "")

        note = note.strip()

        orders.append({
            "date": current_date,
            "customer": current_customer,
            "product": product,
            "qty": qty,
            "unit": unit,
            "price": price,
            "delivery": delivery,
            "note": note
        })

    return orders

def query_order(text):
    keyword = text.replace("查詢", "").strip()
    result = []

    for r in sheet.get_all_values()[1:]:

        # 跳過已刪除訂單
        status = r[11] if len(r) > 11 else ""

        if status == "已刪除":
            continue

        if keyword in " ".join(r):
            result.append(r)

    if not result:
        return "❌ 找不到訂單"

    lines = []
    for r in result[:50]:
        qty = r[5] if len(r) > 5 else ""
        unit = r[6] if len(r) > 6 else "件"
        delivery = r[8] if len(r) > 8 else ""
        note = r[9] if len(r) > 9 else ""
        lines.append(f"單號:{r[0]} 日期:{r[2]} 客戶:{r[3]} 商品:{r[4]} 數量:{qty}{unit} 配送:{delivery} 備註:{note}")
    return "\n".join(lines)

def delete_order(text, user_id):
    parts = text.split()
    if len(parts) < 2:
        return "❌ 刪單格式錯誤"

    rows = sheet.get_all_values()
    deleted = 0

    delete_batch = datetime.now().strftime("%Y%m%d%H%M%S")

    if parts[1].isdigit():
        for i in range(len(rows),1,-1):
            r = rows[i - 1]

            status = r[11] if len(r) > 11 else ""

            if status == "已刪除":
                continue

            if r[0] == parts[1]:

                sheet.update(
                    f"L{i}:O{i}",
                    [[
                        "已刪除",
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        user_id,
                        delete_batch
                    ]]
                )

                deleted += 1
        return f"✅ 已刪除 {deleted} 筆訂單" if deleted else "❌ 找不到訂單"

    dates = [x for x in parts[1:] if re.match(r"\d{1,2}/\d{1,2}", x)]
    customers = [x for x in parts[1:] if not re.match(r"\d{1,2}/\d{1,2}", x)]

# 不允許只輸入客戶
    if not dates:
        return "❌ 僅支援：刪單 單號、刪單 日期、刪單 日期 客戶"

# 日期
    if dates and not customers:

        for i in range(len(rows), 1, -1):

            r = rows[i - 1]

            status = r[11] if len(r) > 11 else ""

            if status == "已刪除":
               continue

            if r[2] in dates:

                sheet.update(
                    f"L{i}:O{i}",
                    [[
                        "已刪除",
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        user_id,
                        delete_batch
                    ]]
                )

                deleted += 1

        return f"✅ 已刪除 {deleted} 筆訂單" if deleted else "❌ 找不到訂單"

# 日期 + 客戶

    target_date = dates[0]
    target_customer = customers[0] if customers else ""

    for i in range(len(rows), 1, -1):

        r = rows[i - 1]

        status = r[11] if len(r) > 11 else ""

        if status == "已刪除":
            continue

        sheet_date = r[2].strip()
        sheet_customer = r[3].strip()

        if (
            target_date in sheet_date
            and target_customer == sheet_customer
        ):

            sheet.update(
                f"L{i}:O{i}",
                [[
                    "已刪除",
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    user_id,
                    delete_batch
                ]]
            )

            deleted += 1

    return f"✅ 已刪除 {deleted} 筆訂單" if deleted else "❌ 找不到訂單"

def restore_order(text):
    if text.strip() in ["復原最後刪除", "還原最後刪除"]:
        return restore_last_delete()
    parts = text.split()

    if len(parts) < 2:
        return "❌ 復原格式錯誤"

    rows = sheet.get_all_values()

    restored = 0

    # 復原單號
    if all(p.isdigit() for p in parts[1:]):

        order_ids = parts[1:]

        for i, r in enumerate(rows[1:], start=2):

            status = r[11] if len(r) > 11 else ""

            if status != "已刪除":
                continue

            if r[0] in order_ids:

                sheet.update(
                    f"L{i}:O{i}",
                    [[
                        "正常",
                        "",
                        "",
                        ""
                    ]]
                )

                restored += 1

        return (
            f"✅ 已復原 {restored} 筆訂單"
            if restored
            else "❌ 找不到已刪除訂單"
        )

    dates = [
        x for x in parts[1:]
        if re.match(r"\d{1,2}/\d{1,2}", x)
    ]

    customers = [
        x for x in parts[1:]
        if not re.match(r"\d{1,2}/\d{1,2}", x)
    ]

    if not dates:
        return "❌ 僅支援：復原 單號、復原 日期、復原 日期 客戶"

    # 日期
    if dates and not customers:

        for i, r in enumerate(rows[1:], start=2):

            status = r[11] if len(r) > 11 else ""

            if status != "已刪除":
                continue

            if r[2] in dates:

                sheet.update(
                    f"L{i}:O{i}",
                    [[
                        "正常",
                        "",
                        "",
                        ""
                    ]]
)

                restored += 1

        return (
            f"✅ 已復原 {restored} 筆訂單"
            if restored
            else "❌ 找不到已刪除訂單"
        )

    # 日期 + 客戶
    target_date = dates[0]
    target_customer = customers[0]

    for i, r in enumerate(rows[1:], start=2):

        status = r[11] if len(r) > 11 else ""

        if status != "已刪除":
            continue

        if (
            r[2] == target_date
            and r[3] == target_customer
        ):

            sheet.update(
                f"L{i}:O{i}",
                [[
                    "正常",
                    "",
                    "",
                    ""
                ]]
            )

            restored += 1

    return (
        f"✅ 已復原 {restored} 筆訂單"
        if restored
        else "❌ 找不到已刪除訂單"
    )

def restore_last_delete():

    rows = sheet.get_all_values()

    deleted_rows = []

    for i, r in enumerate(rows[1:], start=2):

        if len(r) < 15:
            continue

        status = r[11]

        if status != "已刪除":
            continue

        batch = r[14]

        if not batch:
            continue

        deleted_rows.append(
            (
                i,
                batch
            )
        )

    if not deleted_rows:
        return "❌ 沒有可復原的刪除紀錄"

    last_batch = max(
        x[1]
        for x in deleted_rows
    )

    restored = 0

    for i, batch in deleted_rows:

        if batch == last_batch:

            sheet.update(
                f"L{i}:O{i}",
                [[
                    "正常",
                    "",
                    "",
                    ""
                ]]
            )

            restored += 1

    return f"✅ 已復原最後一次刪除，共 {restored} 筆"

@app.get("/files")
def files():
    import os

    if not os.path.exists("exports"):
        return {"files": []}

    return {
        "files": os.listdir("exports")
    }


@app.post("/callback")
async def callback(request: Request):
    body = await request.json()

    for event in body["events"]:
        if event["type"] != "message":
            continue

        text = event["message"]["text"]
        user_id = event["source"]["userId"]
        user_name = get_user_name(user_id)

        if text.startswith("查詢"):
            reply = query_order(text)

        elif text.startswith("匯出"):

            keyword = text.replace("匯出", "").strip()

            if not keyword:
                keyword = "全部"

            url = export_orders(keyword)

            if not url:
                reply = "❌ 找不到資料"
            else:
                reply = (
                    f"✅ Excel匯出完成\n\n"
                    f"{url}"
                )

        elif text.startswith("刪單"):
            reply = delete_order(text, user_name)

        elif text.startswith("改單"):
            reply = edit_order(text)

        elif text.startswith("客戶設定"):
            reply = customer_setting(text)

        elif text.startswith("客戶 "):
            reply = customer_query(text)

        elif text.startswith("復原"):
            reply = restore_order(text)

        elif text.startswith("還原"):
            reply = restore_order(text)

        elif "每週" in text and "到貨" in text:
            reply = create_schedule_order(text)
            
        else:
            ids = []

            count = 0
            items = parse_multi_customer_order(text)
            if items:

                count = save_orders_batch(
                    items,
                    user_name
            )
            else:
                data = parse_order_line(text)
                if data:
                    save_order(data, user_name)
                    count = 1

            reply = f"✅ 已建立 {count} 筆訂單" if count > 0 else "❌ 格式錯誤"

        line_bot_api.reply_message(
            event["replyToken"],
            TextSendMessage(text=reply or "❌ 系統錯誤")
        )

    return PlainTextResponse("OK")


@app.get("/")
def home():
    return {"status":"running"}