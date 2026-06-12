# main.py
# LINE + FastAPI + Google Sheets 訂單系統（商用整合版）
import os
import re
from openai import OpenAI
import base64

from datetime import datetime
from tracemalloc import start
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
client = OpenAI(
    api_key=os.getenv("OPENAI_API_KEY")
)

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

google_credentials = os.getenv("GOOGLE_CREDENTIALS")

if not google_credentials:
    raise Exception("GOOGLE_CREDENTIALS 未設定")

service_account_info = json.loads(
    google_credentials
)

creds = Credentials.from_service_account_info(
    service_account_info,
    scopes=scope
)


gc = gspread.authorize(creds)

import time

def init_sheets():

    for retry in range(5):

        try:

            book = gc.open_by_key(SHEET_ID)

            return (
                book.sheet1,
                book.worksheet("Customers"),
                book.worksheet("Settings")
            )

        except Exception as e:

            print(
                f"Google Sheet連線失敗 "
                f"{retry+1}/5 : {e}"
            )

            time.sleep(5)

    raise Exception("Google Sheet無法連線")
sheet, customer_sheet, settings_sheet = init_sheets()
pending_orders = {}
DELIVERY_LIST = ["自取","自送","代送","業務自送","寄大榮","寄黑貓","寄順豐","寄梓華榮"]

def parse_order_image(image_bytes):

    image_base64 = base64.b64encode(
        image_bytes
    ).decode()

    response = client.chat.completions.create(
        model="gpt-4.1",
        messages=[
            {
                "role":"user",
                "content":[
                    {
                        "type":"text",
                        "text":"""
辨識這張訂單圖片。

請輸出：

[
 {
   "date":"",
   "customer":"",
   "product":"",
   "qty":0,
   "unit":"",
   "price":0,
   "delivery":"",
   "note":""
 }
]

只輸出JSON
禁止任何說明
禁止markdown
"""
                    },
                    {
                        "type":"image_url",
                        "image_url":{
                            "url":
                            f"data:image/jpeg;base64,{image_base64}"
                        }
                    }
                ]
            }
        ]
    )

    content = (
        response.choices[0]
        .message.content
        .strip()
    )

    print("========== GPT ==========")
    print(content)
    print("=========================")

    try:

        return json.loads(content)

    except Exception as e:

        print("JSON解析失敗:", e)

        return []


def parse_purchase_order_image(image_bytes):

    import json
    import base64

    image_base64 = base64.b64encode(
        image_bytes
    ).decode()

    response = client.chat.completions.create(
        model="gpt-4.1",
        messages=[
            {
                "role":"user",
                "content":[
                    {
                        "type":"text",
                        "text":"""
這是一張採購單。

請辨識表格資料。

採購公司通常為：
佳佳農產品實業有限公司

請擷取：

交貨日期 -> date
品名 -> product
數量 -> qty
單位 -> unit
單價 -> price
備註 -> note

customer固定填：

佳佳農產品實業有限公司

delivery固定空白

輸出格式：

[
 {
   "date":"",
   "customer":"",
   "product":"",
   "qty":0,
   "unit":"",
   "price":0,
   "delivery":"",
   "note":""
 }
]

規則：

1. 每個交貨日期算一筆
2. 保留完整商品名稱
3. qty只能輸出數字
4. price只能輸出數字
5. note保留完整內容
6. 只輸出JSON
7. 禁止markdown
8. 禁止任何說明文字
"""
                    },
                    {
                        "type":"image_url",
                        "image_url":{
                            "url":
                            f"data:image/jpeg;base64,{image_base64}"
                        }
                    }
                ]
            }
        ]
    )

    content = (
        response
        .choices[0]
        .message
        .content
        .strip()
    )

    print(content)

    try:
        return json.loads(content)

    except Exception as e:

        print("JSON解析失敗")
        print(content)

        return []
    
def detect_purchase_order(image_bytes):

    image_base64 = base64.b64encode(
        image_bytes
    ).decode()

    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {
                "role":"user",
                "content":[
                    {
                        "type":"text",
                        "text":"這是不是採購單？只回答 YES 或 NO"
                    },
                    {
                        "type":"image_url",
                        "image_url":{
                            "url":
                            f"data:image/jpeg;base64,{image_base64}"
                        }
                    }
                ]
            }
        ]
    )

    return (
        "YES"
        in response.choices[0]
        .message.content.upper()
    )

def format_orders(orders):

    lines = []

    for i, o in enumerate(
        orders,
        start=1
    ):

        lines.append(
            f"{i}. "
            f"{o['date']} "
            f"{o['customer']} "
            f"{o['product']} "
            f"{o['qty']}{o['unit']}"
        )

    return "\n".join(lines)

def edit_pending_order(
    user_id,
    text
):

    if user_id not in pending_orders:
        return "❌ 沒有待確認訂單"

    m = re.match(
        r"修改\s+(\d+)\s+(日期|客戶|商品|數量|單價|配送|備註)\s+(.+)",
        text
    )

    if not m:
        return "❌ 格式錯誤"

    idx = int(m.group(1)) - 1

    field = m.group(2)

    value = m.group(3)

    orders = pending_orders[user_id]

    if idx >= len(orders):
        return "❌ 找不到項目"

    order = orders[idx]

    field_map = {
        "日期":"date",
        "客戶":"customer",
        "商品":"product",
        "單價":"price",
        "配送":"delivery",
        "備註":"note"
    }

    if field == "數量":

        qty_match = re.match(
            r"(\d+(?:\.\d+)?)(.*)",
            value
        )

        if not qty_match:
            return "❌ 數量格式錯誤"

        order["qty"] = float(
            qty_match.group(1)
        )

        unit = qty_match.group(2).strip()

        if unit:
            order["unit"] = unit

    else:

        order[
            field_map[field]
        ] = value

    return (
        "✅ 已修改\n\n"
        + format_orders(orders)
    )

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

    summary = {}

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

        customer = r[3]
        product = r[4]
        unit = r[6]

        try:
            qty = float(r[5])
        except:
            qty = 0

        key = (
            customer,
            product,
            unit
        )

        summary[key] = summary.get(key, 0) + qty

        count += 1

    if count == 0:
        return None
    # =====================
    # 統計表
    # =====================

    ws.append([])
    ws.append([])
    ws.append(["統計表"])

    ws.append([
        "客戶",
        "商品",
        "單位",
        "數量合計"
    ])

    for (customer, product, unit), total_qty in sorted(summary.items()):

        ws.append([
            customer,
            product,
            unit,
            total_qty
        ])
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

            return "✅ 已更新"

    customer_sheet.append_row([
        customer,
        product,
        qty_unit,
        price
    ])

    return "✅ 已新增"

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

    lines = [
        x.strip()
        for x in text.splitlines()
        if x.strip()
    ]

    # 三行排程模式
    if len(lines) >= 3:

        customer = lines[0]

        rows = customer_sheet.get_all_values()

        customer_data = None

        for r in rows[1:]:

            if r[0] == customer:
                customer_data = r
                break

        if not customer_data:
            return "❌ 客戶未設定"

        weekday_map = {
            "一":0,
            "二":1,
            "三":2,
            "四":3,
            "五":4,
            "六":5,
            "日":6
        }

        product = customer_data[1]

        qty_match = re.search(
            r"(\d+(?:\.\d+)?)",
            customer_data[2]
        )

        qty = float(qty_match.group(1))

        unit = parse_unit(
            customer_data[2]
        )

        price = float(customer_data[3])

        today = datetime.now().date()

        orders = []

        # 下週四、五到貨

        m = re.search(
            r"下週([一二三四五六日、]+)",
            lines[1]
        )

        if m:

            target_days = [
                weekday_map[x]
                for x in re.findall(
                    r"[一二三四五六日]",
                    m.group(1)
                )
            ]

            next_week_start = today + timedelta(days=7)

            for i in range(7):

                d = next_week_start + timedelta(days=i)

                if d.weekday() in target_days:

                    orders.append({
                        "date": f"{d.month}/{d.day}",
                        "customer": customer,
                        "product": product,
                        "qty": qty,
                        "unit": unit,
                        "price": price,
                        "delivery": "",
                        "note": "特殊排程"
                    })

        # 之後每週二三四五到月底

        m = re.search(
            r"之後每週([一二三四五六日]+)",
            lines[2]
        )

        if m:

            target_days = [
                weekday_map[x]
                for x in m.group(1)
            ]

            start = today + timedelta(days=14)

            target_year = start.year
            target_month = start.month

            last_day = calendar.monthrange(
                target_year,
                target_month
            )[1]

            end_date = datetime(
                target_year,
                target_month,
                last_day
            ).date()

            current = start

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

        return f"✅ 已建立 {count} 筆"

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

    return f"✅ 已建立 {count} 筆"

    


UNIT_WHITELIST = [
    "包","袋","箱","件","桶",
    "公斤","斤","公克",
    "kg","KG","Kg",
    "g","G",
    "lb","LB","lbs"
]

def edit_order(text):

    rows = sheet.get_all_values()

    text = text.replace("\n", " ").strip()

    m = re.match(
        r"改單\s+(.+?)\s+(日期|商品|數量|單價|配送|備註)\s+(.+)",
        text
    )

    if not m:
        return "❌ 格式錯誤"

    target_text = m.group(1).strip()
    field = m.group(2).strip()
    value = m.group(3).strip()

    dates = re.findall(
        r"\d{1,2}/\d{1,2}",
        target_text
    )

    order_ids = re.findall(
        r"\d{9,}",
        target_text
    )

    customer = ""

    tokens = target_text.split()

    for t in tokens:

        if re.match(r"\d{1,2}/\d{1,2}", t):
            continue

        if re.match(r"\d{9,}", t):
            continue

        customer = t
        break

    qty = None
    unit = ""
    note = ""

    if field == "數量":

        qty_match = re.match(
            r"(\d+(?:\.\d+)?)(.*)",
            value
        )

        if not qty_match:
            return "❌ 數量格式錯誤"

        qty = qty_match.group(1)

        remain = qty_match.group(2).strip()

        if remain:

            for u in sorted(
                UNIT_WHITELIST,
                key=len,
                reverse=True
            ):

                if remain.startswith(u):

                    unit = u
                    note = remain[len(u):].strip()

                    break

            if not unit:

                note = remain

    updated = 0

    col_map = {
        "日期":3,
        "商品":5,
        "數量":6,
        "單價":8,
        "配送":9,
        "備註":10
    }

    for i, r in enumerate(
        rows[1:],
        start=2
    ):

        status = r[11] if len(r) > 11 else ""

        if status == "已刪除":
            continue

        match = False

        # 單號模式
        if order_ids:

            if r[0] in order_ids:
                match = True

        # 日期模式
        elif dates:

            if r[2] not in dates:
                continue

            if customer:

                if r[3] != customer:
                    continue

            match = True

        if not match:
            continue

        if field == "數量":

            sheet.update_cell(
                i,
                6,
                qty
            )

            if unit:

                sheet.update_cell(
                    i,
                    7,
                    unit
                )

            if note:

                sheet.update_cell(
                    i,
                    10,
                    note
                )

        else:

            sheet.update_cell(
                i,
                col_map[field],
                value
            )

        updated += 1

    if updated == 0:
        return "❌ 找不到符合訂單"

    return f"✅ 已修改 {updated} 筆"

def parse_unit(text):

    for u in sorted(
        UNIT_WHITELIST,
        key=len,
        reverse=True
    ):
        if u in text:
            return u

    return "件"

def detect_delivery(text):
    for d in DELIVERY_LIST:
        if d in text:
            return d
    return ""

def generate_order_ids(count):

    today = datetime.now()

    roc_date = (
        f"{today.year - 1911}"
        f"{today.strftime('%m%d')}"
    )

    last_date = settings_sheet.acell("B1").value
    last_seq = int(
        settings_sheet.acell("B2").value or "0"
    )

    # 換日歸零
    if last_date != roc_date:
        last_seq = 0

    start_seq = last_seq + 1

    # 寫回最新狀態
    settings_sheet.update(
        "B1:B2",
        [
            [roc_date],
            [str(last_seq + count)]
        ]
    )

    ids = []

    for i in range(count):

        seq = start_seq + i

        ids.append(
            f"{roc_date}"
            f"{str(seq).zfill(3)}"
        )

    return ids


def generate_order_id():
    return generate_order_ids(1)[0]

def save_order(data, user_id):

    oid = generate_order_ids(1)[0]

    seq_no = oid[-3:]

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
        str(uuid.uuid4()),  # P
        seq_no              # Q
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

        seq_no = oid[-3:]

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
            str(uuid.uuid4()),  # P
            seq_no              # Q
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

        unit_text = m.group(3) or ""

        unit = "件"
        remain_unit_note = unit_text

        for u in sorted(
            UNIT_WHITELIST,
            key=len,
            reverse=True
        ):
            if remain_unit_note.startswith(u):

                unit = u

                remain_unit_note = (
                    remain_unit_note[len(u):]
                    .strip()
                )

                break

        price = float(m.group(4)) if m.group(4) else 0

        remain = (
            remain_unit_note + " " + m.group(5)
        ).strip()

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
        return f"✅ 已刪除 {deleted} 筆" if deleted else "❌ 找不到訂單"

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

        return f"✅ 已刪除 {deleted} 筆" if deleted else "❌ 找不到訂單"

# 日期 + 客戶

    target_date = dates[0]
    target_customer = customers[0]

    for i in range(len(rows), 1, -1):

        r = rows[i - 1]

        status = r[11] if len(r) > 11 else ""

        if status == "已刪除":
            continue

        sheet_date = r[2].strip()
        sheet_customer = r[3].strip()

        if (
            sheet_date in dates
            and sheet_customer == target_customer
        ):

            sheet.update(
                f"L{i}:O{i}",
                [[
                    "已刪除",
                    datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),
                    user_id,
                    delete_batch
                ]]
            )

            deleted += 1

    return f"✅ 已刪除 {deleted} 筆" if deleted else "❌ 找不到訂單"

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
            f"✅ 已復原 {restored} 筆"
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
   # 日期

    if dates and not customers:

        for i in range(len(rows), 1, -1):

            r = rows[i - 1]

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
        f"✅ 已復原 {restored} 筆"
        if restored
        else "❌ 找不到已刪除訂單"
        )

    
    
    # 日期 + 客戶

    target_customer = customers[0]

    for i, r in enumerate(rows[1:], start=2):

        status = r[11] if len(r) > 11 else ""

        if status != "已刪除":
            continue

        sheet_date = r[2].strip()
        sheet_customer = r[3].strip()

        if (
            sheet_date in dates
            and sheet_customer == target_customer
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
        f"✅ 已復原 {restored} 筆"
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

    return f"✅ 已復原，共 {restored} 筆"



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

        if (
            event["type"] == "message"
            and
            event["message"]["type"] == "image"
        ):

            user_id = event["source"]["userId"]

            message_id = event["message"]["id"]

            content = line_bot_api.get_message_content(
                message_id
            )

            image_data = BytesIO()

            for chunk in content.iter_content():

                image_data.write(chunk)

            img = image_data.getvalue()

            if detect_purchase_order(img):

                orders = parse_purchase_order_image(
                    img
                )

            else:

                orders = parse_order_image(
                    img
                )

            pending_orders[user_id] = orders

            preview = []

            for i, o in enumerate(
                orders,
                start=1
            ):

                preview.append(
                    f"{i}."
                    f"\n日期:{o['date']}"
                    f"\n客戶:{o['customer']}"
                    f"\n商品:{o['product']}"
                    f"\n數量:{o['qty']}{o['unit']}"
                    f"\n單價:{o['price']}"
                    f"\n配送:{o['delivery']}"
                    f"\n備註:{o['note']}"
                )

            line_bot_api.reply_message(
                event["replyToken"],
                TextSendMessage(
                    text=
                    "📋 AI辨識完成\n\n"
                    + "\n\n".join(preview)
                    + "\n\n可輸入：\n修改 1 商品 高麗菜\n確認\n取消"
                )
            )

            continue

        if event["type"] != "message":
            continue

        text = event["message"]["text"]

    # ===== 移除 LINE Mention =====

        mentionees = []

        if "mention" in event["message"]:
            mentionees = event["message"]["mention"].get(
                "mentionees",
                []
            )

        for m in reversed(mentionees):

            start = m["index"]
            length = m["length"]

            text = (
                text[:start]
                + text[start + length:]
            )

        text = text.strip()

    # ============================

        user_id = event["source"]["userId"]
        user_name = get_user_name(user_id)
        
        if (
            "\n" in text
            and "下週" in text
            and "之後每週" in text
        ):

            reply = create_schedule_order(text)

            line_bot_api.reply_message(
                event["replyToken"],
                TextSendMessage(text=reply)
            )

            continue

        commands = [
            x.strip()
            for x in re.split(
                r"\n\s*\n",
                text
            )
            if x.strip()
        ]

        results = []

        for cmd in commands:

            if cmd.startswith("修改"):

                results.append(
                    edit_pending_order(
                        user_id,
                        cmd
                    )
                )

                continue

            elif cmd == "預覽":

                    if user_id not in pending_orders:

                        results.append(
                            "❌ 沒有待確認訂單"
                        )

                    else:

                        results.append(
                            format_orders(
                                pending_orders[user_id]
                            )
                        )

            elif cmd == "確認":

                if user_id not in pending_orders:

                    results.append(
                        "❌ 沒有待確認訂單"
                    )

                else:

                    count = save_orders_batch(
                        pending_orders[user_id],
                        user_name
                    )

                    del pending_orders[user_id]

                    results.append(
                        f"✅"
                    )

            elif cmd == "取消":

                pending_orders.pop(
                    user_id,
                    None
                )

                results.append(
                "✅"
                )

            if cmd.startswith("查詢"):

                results.append(
                    query_order(cmd)
                )

            elif cmd.startswith("匯出"):

                keyword = cmd.replace(
                    "匯出",
                    ""
                ).strip()

                if not keyword:
                    keyword = "全部"

                url = export_orders(keyword)

                if not url:
                    results.append(
                        "❌ 找不到資料"
                    )
                
                else:
                    results.append(
                        f"✅ Excel匯出完成\n{url}"
                    )

            elif cmd.startswith("刪單"):

                results.append(
                    delete_order(
                        cmd,
                        user_name
                    )   
                )

            elif cmd.startswith("改單"):

                results.append(
                    edit_order(cmd)
                )

            elif cmd.startswith("客戶設定"):

                results.append(
                    customer_setting(cmd)
                )

            elif cmd.startswith("客戶 "):

                results.append(
                    customer_query(cmd)
                )

            elif cmd.startswith("復原"):

                results.append(
                    restore_order(cmd)
                )

            elif cmd.startswith("還原"):

                results.append(
                    restore_order(cmd)
                )

            elif (
                "下週" in cmd
                and "之後每週" in cmd
            ):
                results.append(
                    create_schedule_order(cmd)
                )

            elif "每週" in cmd and "到貨" in cmd:

                results.append(
                    create_schedule_order(cmd)
                )

            else:

                count = 0

                items = parse_multi_customer_order(cmd)

                if items:

                    count = save_orders_batch(
                        items,
                        user_name
                    )

                else:

                    data = parse_order_line(cmd)

                    if data:

                        save_order(
                            data,
                            user_name
                        )

                        count = 1

                results.append(
                    f"✅"
                    if count > 0
                    else "❌ 格式錯誤"
                )

        reply = "\n\n".join(results)

        line_bot_api.reply_message(
            event["replyToken"],
            TextSendMessage(text=reply or "❌ 系統錯誤")
        )

    return PlainTextResponse("OK")


@app.get("/")
def home():
    return {"status":"running"}