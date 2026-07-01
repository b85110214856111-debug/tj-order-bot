# main.py
# LINE + FastAPI + Google Sheets 訂單系統（商用整合版）
import os
import re
from datetime import datetime
from tracemalloc import start
from fastapi import FastAPI, Request
from fastapi.responses import PlainTextResponse
from dotenv import load_dotenv
from openpyxl import Workbook
from fastapi.responses import FileResponse
import gspread

import zipfile
import shutil
import tempfile

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

def upload_zip_file(excel_path, image_folder):

    zip_name = f"orders_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"

    tmp_zip = None

    try:

        with tempfile.NamedTemporaryFile(
            suffix=".zip",
            delete=False
        ) as tmp:

            tmp_zip = tmp.name

        with zipfile.ZipFile(
            tmp_zip,
            "w",
            zipfile.ZIP_DEFLATED
        ) as zipf:

            # Excel
            zipf.write(
                excel_path,
                arcname="訂單.xlsx"
            )

            # 圖片
            if os.path.exists(image_folder):

                for file in os.listdir(image_folder):

                    full = os.path.join(
                        image_folder,
                        file
                    )

                    if os.path.isfile(full):

                        zipf.write(
                            full,
                            arcname=file
                        )

        result = cloudinary.uploader.upload(
            tmp_zip,
            resource_type="raw",
            folder="order_exports",
            public_id=zip_name.replace(
                ".zip",
                ""
            ),
            overwrite=True
        )

        return result["secure_url"]

    finally:

        if tmp_zip and os.path.exists(tmp_zip):
            os.remove(tmp_zip)

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
DELIVERY_LIST = ["自取","自送","代送","業務自送","寄大榮","寄黑貓","寄順豐","寄梓華榮"]


def export_orders(keyword="全部", start_date=None, end_date=None):

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

        # 日期區間
        if start_date and end_date:

            try:
                month, day = map(int, r[2].split("/"))
                order_num = month * 100 + day

                sm, sd = map(int, start_date.split("/"))
                em, ed = map(int, end_date.split("/"))

                start_num = sm * 100 + sd
                end_num = em * 100 + ed

                if order_num < start_num or order_num > end_num:
                    continue

            except:
                continue

        # 關鍵字(客戶)
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
    with tempfile.NamedTemporaryFile(
        suffix=".xlsx",
        delete=False
    ) as tmp:

        excel_path = tmp.name

    wb.save(excel_path)

    # 找圖片資料夾
    folder = ""

    if start_date:

        m, d = map(int, start_date.split("/"))

        folder = os.path.join(
            "uploads",
            f"{m:02d}{d:02d}"
        )

    url = upload_zip_file(
        excel_path,
        folder
    )

    os.remove(excel_path)

    return url




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
    added_dates = set()
    orders = []

    lines = [
        x.strip()
        for x in text.splitlines()
        if x.strip()
    ]

    # 三行排程模式
    if len(lines) >= 3:
        product = None
        qty = None
        unit = None
        price = 0
        delivery = ""
        note = ""

        order_text = lines[0]

        parts = order_text.split()

        customer = parts[0]

        rows = customer_sheet.get_all_values()

        customer_data = None

        for r in rows[1:]:
            if r[0] == customer:
                customer_data = r
                break

        if not customer_data:
            return "❌ 客戶未設定"


        # 商品使用輸入
        product = parts[1] if len(parts) >= 2 else None


        # 先保留輸入值
        input_qty = qty
        input_unit = unit
        input_price = price


        # 預設
        qty = 0
        unit = ""
        price = 0


        # 找 Customers 同商品補缺
        for r in rows[1:]:

            if len(r) < 4:
                continue

            if r[0] == customer and r[1] == product:

                qty_match = re.search(
                    r"(\d+(?:\.\d+)?)",
                    r[2]
                )

                if qty_match:
                    qty = float(qty_match.group(1))
                    unit = parse_unit(r[2])

                try:
                    price = float(r[3])
                except:
                    price = 0

                break


        # 輸入優先覆蓋

        if input_qty:
            qty = input_qty
            unit = input_unit

        if input_price:
            price = input_price


        weekday_map = {
            "一":0,
            "二":1,
            "三":2,
            "四":3,
            "五":4,
            "六":5,
            "日":6
        }    

        # 數量
        for p in parts[2:]:

            m = re.match(r"(\d+(?:\.\d+)?)(.*)", p)

            if not m:
                continue

            qty = float(m.group(1))

            remain = m.group(2).strip()

            if remain:
                unit = parse_unit(remain)

            break

        # 單價
        for p in parts:

            if p.startswith("@"):
                try:
                    price = float(p[1:])
                except:
                    pass

        # 配送
        delivery = detect_delivery(order_text)

        # 備註
        note = order_text

        note = note.replace(customer, "", 1)
        note = note.replace(product, "", 1)

        note = re.sub(
            r"@\d+(?:\.\d+)?",
            "",
            note
        )

        note = re.sub(
            r"\d+(?:\.\d+)?[^\s]*",
            "",
            note
        )

        note = note.replace("@", "")

        if delivery:
            note = note.replace(delivery, "")

        note = note.strip()

        today = datetime.now().date()

        
       
        

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

            # 下週一
            next_week_start = today + timedelta(
                days=(7 - today.weekday())
            )

            

            for i in range(7):

                d = next_week_start + timedelta(days=i)

                if d.weekday() in target_days:

                    date_str = f"{d.month}/{d.day}"

                    if date_str in added_dates:
                        continue

                    added_dates.add(date_str)

                    orders.append({
                        "date": date_str,
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

                    date_str = f"{current.month}/{current.day}"

                    if date_str in added_dates:
                        current += timedelta(days=1)
                        continue

                    added_dates.add(date_str)

                    orders.append({
                        "date": date_str,
                        "customer": customer,
                        "product": product,
                        "qty": qty,
                        "unit": unit,
                        "price": price,
                        "delivery": delivery,
                        "note": note if note else "固定排程"
                    })

                current += timedelta(days=1)

        if not orders:
            return "❌ 沒有符合日期"

        # ===== 檢查重複 =====
        rows = sheet.get_all_values()

        exist = {
            (r[2], r[3], r[4])
            for r in rows[1:]
            if len(r) > 11 and r[11] != "已刪除"
        }

        orders = [
            o for o in orders
            if (o["date"], o["customer"], o["product"]) not in exist
        ]

        if not orders:
            return "❌ 排程已存在"

        count = save_orders_batch(
            orders,
            "SYSTEM"
        )

        return f"✅ 已建立 {count} 筆固定排程"

    text = text.strip()

    start_next_week = False

    

    if "下週開始" in text:
        start_next_week = True
        text = text.replace("下週開始", "", 1).strip()

    m = re.match(
        r"(.+?)\s+每週([一二三四五六日]+)到貨到月底$",
        text
    )

    if not m:
        return "❌ 排程格式錯誤"

    order_text = m.group(1).strip()
    weekdays_text = m.group(2)

    parts = order_text.split()

    customer = parts[0]
    product = None
    qty = None
    unit = None
    price = 0

    delivery = ""
    note = ""

    # 商品先判斷
    product = None

    if len(parts) >= 2:
        if parts[1] not in DELIVERY_LIST:
            product = parts[1]
    # 數量
    for p in parts[1:]:

        m = re.match(r"(\d+(?:\.\d+)?)(.*)", p)

        if m:

            qty = float(m.group(1))

            if m.group(2):
                unit = parse_unit(m.group(2))

            break

    # 單價（輸入優先）

    for p in parts:

        m = re.match(
            r"@(\d+(?:\.\d+)?)",
            p
        )

        if m:
            price = float(m.group(1))
            break

    # 配送
    delivery = ""

    for d in DELIVERY_LIST:

        if d in order_text:
            delivery = d
            break

    # 備註
    note = order_text

    note = note.replace(customer, "", 1)

    if product:
        note = note.replace(product, "", 1)

    note = re.sub(
        r"\d+(?:\.\d+)?[^\s]*",
        "",
        note
    )

    # 移除單價符號
    note = re.sub(
        r"@\d+(?:\.\d+)?|@",
        "",
        note
    )

    for d in DELIVERY_LIST:
        note = note.replace(d, "")

    # 移除排程文字
    note = re.sub(
        r"每週[一二三四五六日]+到貨到月底",
        "",
        note
    )

    note = note.strip()

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


    #商品輸入優先
    if not product:
        product = customer_data[1]


    input_qty = qty
    input_unit = unit
    input_price = price


    # 預設
    qty = None
    unit = None
    price = None


    # 比對 Customers 同客戶+同商品
    for r in rows[1:]:

        if len(r) < 4:
            continue

        if r[0] == customer and r[1] == product:

            qty_match = re.search(
                r"(\d+(?:\.\d+)?)",
                r[2]
            )

            if qty_match:
                qty = float(qty_match.group(1))
                unit = parse_unit(r[2])

            try:
                price = float(r[3])
            except:
                price = 0

            break


    # 輸入優先覆蓋
    # 輸入優先覆蓋
    # 輸入有值才覆蓋，沒有就用基本檔

    if input_qty is not None:
        qty = input_qty

    if input_unit:
        unit = input_unit

    if input_price and input_price > 0:
        price = input_price

    today = datetime.now().date()

    if start_next_week:
    # 下週一開始
        today += timedelta(days=(7 - today.weekday()))

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
                "delivery": delivery,
                "note": note if note else "固定排程"
            })

        current += timedelta(days=1)

    if not orders:
        return "❌ 沒有符合日期"

    rows = sheet.get_all_values()

    exist = {
        (r[2], r[3], r[4])
        for r in rows[1:]
        if len(r) > 11 and r[11] != "已刪除"
    }

    orders = [
        o for o in orders
        if (o["date"], o["customer"], o["product"]) not in exist
    ]

    if not orders:
        return "❌ 排程已存在"

    count = save_orders_batch(
        orders,
        "SYSTEM"
    )

    return f"✅ 已建立 {count} 筆排程"

    


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
        r"改單\s+(.+?)\s+(日期|商品|數量|單價|配送|備註)\s+(.+)$",
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
    product = ""

    tokens = target_text.split()

    normal = []

    for t in tokens:

        if re.match(r"\d{1,2}/\d{1,2}", t):
            continue

        if re.match(r"\d{9,}", t):
            continue

        normal.append(t)

    if len(normal) >= 1:
        customer = normal[0]

    if len(normal) >= 2:
        product = normal[1]

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

            if customer and r[3] != customer:
                continue

            if product and r[4] != product:
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

    return f"✅ 已改 {updated} 筆"

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


def save_line_image(event):

    message_id = event["message"]["id"]

    today = datetime.now().strftime("%m%d")

    folder = os.path.join("uploads", today)

    os.makedirs(folder, exist_ok=True)

    content = line_bot_api.get_message_content(message_id)

    filename = os.path.join(folder, f"{message_id}.jpg")

    with open(filename, "wb") as f:
        for chunk in content.iter_content():
            f.write(chunk)

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
        return f"✅ 已刪 {deleted} 筆" if deleted else "❌ 找不到訂單"

    dates = [x for x in parts[1:] if re.match(r"\d{1,2}/\d{1,2}", x)]

    others = [x for x in parts[1:] if not re.match(r"\d{1,2}/\d{1,2}", x)]

    customer = others[0] if len(others) >= 1 else ""
    product = others[1] if len(others) >= 2 else ""

# 不允許只輸入客戶
    if not dates:
        return "❌ 僅支援：刪單 單號、刪單 日期、刪單 日期 客戶"

# 日期
    if dates and not customer:

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

        return f"✅ 已刪 {deleted} 筆" if deleted else "❌ 找不到訂單"

# 日期 + 客戶

    target_date = dates[0]

    for i in range(len(rows), 1, -1):

        r = rows[i - 1]

        status = r[11] if len(r) > 11 else ""

        if status == "已刪除":
            continue

        if r[2] != target_date:
            continue

        if customer and r[3] != customer:
            continue

        if product and r[4] != product:
            continue

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

    return f"✅ 已刪 {deleted} 筆" if deleted else "❌ 找不到訂單"

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

    others = [
        x for x in parts[1:]
        if not re.match(r"\d{1,2}/\d{1,2}", x)
    ]

    customer = others[0] if len(others) >= 1 else ""
    product = others[1] if len(others) >= 2 else ""

    if not dates:
        return "❌ 僅支援：復原 單號、復原 日期、復原 日期 客戶"

    # 日期
   # 日期

    if dates and not customer:

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

    # 日期 + 客戶 (+ 商品)

    target_date = dates[0]

    for i, r in enumerate(rows[1:], start=2):

        status = r[11] if len(r) > 11 else ""

        if status != "已刪除":
            continue

        if r[2].strip() != target_date:
            continue

        if customer and r[3].strip() != customer:
            continue

        if product and r[4].strip() != product:
            continue

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

        msg_type = event["message"]["type"]

        if msg_type == "image":
            save_line_image(event)
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

            if cmd.startswith("查詢"):

                results.append(
                    query_order(cmd)
                )

            elif cmd.startswith("匯出"):

                text = cmd.replace("匯出", "").strip()

                keyword = "全部"
                start_date = None
                end_date = None

                m = re.match(
                    r"(\d{1,2}/\d{1,2})\s*[~-]\s*(\d{1,2}/\d{1,2})(?:\s+(.*))?$",
                    text
                )

                if m:

                    start_date = m.group(1)
                    end_date = m.group(2)

                    if m.group(3):
                        keyword = m.group(3).strip()

                else:

                    if text:
                        keyword = text

                url = export_orders(
                    keyword,
                    start_date,
                    end_date
                )

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
                    f"✅ 已建 {count} 筆"
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