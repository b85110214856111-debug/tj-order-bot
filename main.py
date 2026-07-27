# main.py
# LINE + FastAPI + Google Sheets 訂單系統（商用整合版）
import cmd
import os
import re
import token
import requests
from datetime import datetime, date, timezone, timedelta
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
)
LINE_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
SHEET_ID = os.getenv("GOOGLE_SHEET_ID")

line_bot_api = LineBotApi(LINE_TOKEN)

scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

TW = timezone(timedelta(hours=8))

def now_tw():
    return datetime.now(TW)

DATE_FORMAT = "%Y/%m/%d"


def parse_date(text: str, base=None):
    """
    支援
    2026/12/31
    2026-12-31
    12/31
    1/2
    """

    text = text.strip().replace("-", "/")

    if base is None:
        base = now_tw().date()

    # YYYY/MM/DD
    m = re.match(r"^(\d{4})/(\d{1,2})/(\d{1,2})$", text)
    if m:
        return date(
            int(m.group(1)),
            int(m.group(2)),
            int(m.group(3))
        )

    # MM/DD
    m = re.match(r"^(\d{1,2})/(\d{1,2})$", text)

    if not m:
        raise ValueError(text)

    month = int(m.group(1))
    day = int(m.group(2))

    year = base.year

    d = date(year, month, day)

    # 自動跨年
    if (
        base.month >= 11
        and month <= 2
        and d < base
    ):
        d = date(year + 1, month, day)

    return d

def parse_date_range(start_text, end_text):
    """
    支援：
    12/25-1/5
    2026/12/25-2027/1/5
    """

    start = parse_date(start_text)
    end = parse_date(end_text, start)

    if end < start:
        end = date(end.year + 1, end.month, end.day)

    return start, end

def format_date(d):
    return d.strftime(DATE_FORMAT)

def expand_short_dates(text: str):
    """
    支援
    7/1、2、3、4
    7/1,2,3,4
    7/1，2，3，4
    7/1｀2｀3｀4
    ↓
    7/1 7/2 7/3 7/4
    """

    text = (
        text.replace("，", "、")
            .replace(",", "、")
            .replace("｀", "、")
    )

    pattern = r'(\d{1,2})/(\d{1,2}(?:、\d{1,2})+)'

    def repl(m):
        month = m.group(1)
        days = m.group(2).split("、")
        return " ".join(
            f"{month}/{d}"
            for d in days
        )

    return re.sub(pattern, repl, text)

def date_key(text):
    return parse_date(text)


def in_date_range(order_date, start_date, end_date):
    d = parse_date(order_date)
    s = parse_date(start_date, d)
    e = parse_date(end_date, s)

    if e < s:
        e = date(e.year + 1, e.month, e.day)

    if d < s:
        d = date(d.year + 1, d.month, d.day)

    return s <= d <= e

import cloudinary
import cloudinary.uploader
import tempfile

cloudinary.config(
    cloud_name=os.getenv("CLOUDINARY_CLOUD_NAME"),
    api_key=os.getenv("CLOUDINARY_API_KEY"),
    api_secret=os.getenv("CLOUDINARY_API_SECRET"),
    secure=True
)

import zipfile

def download_file(url, path):

    r = requests.get(url)

    if r.status_code == 200:

        with open(path, "wb") as f:

            f.write(r.content)

import uuid

from io import BytesIO

def download_line_image(message_id):

    headers = {
        "Authorization": f"Bearer {LINE_TOKEN}"
    }

    url = f"https://api-data.line.me/v2/bot/message/{message_id}/content"

    r = requests.get(url, headers=headers)

    if r.status_code != 200:
        return None

    return r.content

def upload_photo(image_bytes, filename):

    with tempfile.NamedTemporaryFile(
        suffix=".jpg",
        delete=False
    ) as tmp:

        tmp.write(image_bytes)

        tmp_path = tmp.name

    try:

        result = cloudinary.uploader.upload(
            tmp_path,
            folder="order_photos",
            public_id=filename,
            overwrite=True
        )

        return result["secure_url"]

    finally:

        os.remove(tmp_path)

def save_photo(user_name, message_id):

    image = download_line_image(message_id)

    if image is None:
        return

    filename = now_tw().strftime(
        "IMG_%Y%m%d_%H%M%S"
    )

    url = upload_photo(
        image,
        filename
    )

    photo_sheet.append_row([
        now_tw().strftime("%m/%d"),
        now_tw().strftime("%H:%M:%S"),
        user_name,
        message_id,
        url,
        filename + ".jpg"
    ])

def upload_excel_file(wb):

    filename = (
        f"orders_{now_tw().strftime('%Y%m%d_%H%M%S')}.xlsx"
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

def upload_zip(zip_bytes):

    filename = (
        f"Export_{now_tw().strftime('%Y%m%d_%H%M%S')}.zip"
    )

    result = cloudinary.uploader.upload(
        zip_bytes,
        resource_type="raw",
        folder="exports",
        public_id=filename.replace(".zip", ""),
        overwrite=True,
        format="zip"
    )

    return result["secure_url"]

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
                book.worksheet("Settings"),
                book.worksheet("Photos")
            )

        except Exception as e:

            print(
                f"Google Sheet連線失敗 "
                f"{retry+1}/5 : {e}"
            )

            time.sleep(5)

    raise Exception("Google Sheet無法連線")
sheet, customer_sheet, settings_sheet, photo_sheet = init_sheets()
DELIVERY_LIST = ["自取","自送","代送","業務自送","寄大榮","寄黑貓","寄順豐","寄梓華榮","送梓華榮","寄梓","寄大","送梓","司機自送","黑貓"]

def download_file_bytes(url):
    r = requests.get(url, timeout=30)
    if r.status_code == 200:
        return r.content
    return None


def export_orders(rows, keyword="全部", start_date=None, end_date=None):

    

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
        "備註",
        "狀態"
    ])

    count = 0

    summary = {}

    export_rows = []

    for r in rows[1:]:

        status = r[11] if len(r) > 11 else ""

        if status == "已刪除":
            continue

        # 日期區間（支援跨年）
        if start_date and end_date:

            try:

                order_date = parse_date(r[2])

                start_dt, end_dt = parse_date_range(
                    start_date,
                    end_date
                )

                check = order_date

                if check < start_dt:
                    check = check.replace(year=check.year + 1)

                if not (start_dt <= check <= end_dt):
                    continue

            except:
                continue

        # 關鍵字(客戶)
        if keyword != "全部":

            if keyword not in " ".join(r):
                continue

        export_rows.append(r)

        customer = r[3]
        product = r[4]
        unit = r[6]
        status = r[11]

        try:
            qty = float(r[5])
        except:
            qty = 0

        key = (
            customer,
            product,
            unit,
            status
        )

        summary[key] = summary.get(key, 0) + qty

        count += 1

    export_rows.sort(
        key=lambda r: (
            parse_date(r[2])
            if str(r[2]).strip() != "未定"
            else date.max,
            r[3],
            r[4]
        )
    )

    for r in export_rows:

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
            r[9],
            r[11]
        ])

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
        "狀態",
        "數量合計"
    ])

    for (customer, product, unit), total_qty in sorted(summary.items()):

        ws.append([
            customer,
            product,
            unit,
            status,
            total_qty
        ])
    excel_buffer = BytesIO()
    wb.save(excel_buffer)
    excel_buffer.seek(0)

    zip_buffer = BytesIO()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:

        # Excel
        z.writestr(
            "Orders.xlsx",
            excel_buffer.getvalue()
        )

        # Photos（🔥直接串流，不落地）
        photos = photo_sheet.get_all_values()

        index = 1

        for p in photos[1:]:

            if start_date and end_date:

                try:

                    photo_date = parse_date(p[0])

                    start_dt, end_dt = parse_date_range(
                        start_date,
                        end_date
                    )

                    check = photo_date

                    if check < start_dt:
                        check = check.replace(year=check.year + 1)

                    if not (start_dt <= check <= end_dt):
                        continue

                except:
                    continue

            img_bytes = download_file_bytes(p[4])
            if not img_bytes:
                continue

            filename = (
                f"{p[0].replace('/','-')}_"
                f"{p[2]}_"
                f"{index:03d}.jpg"
            )

            z.writestr(f"Photos/{filename}", img_bytes)

            index += 1


    zip_buffer.seek(0)

    url = upload_zip(zip_buffer.getvalue())

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
    text = text.replace("周", "週")
    added_dates = set()
    orders = []
    target_dates = []

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

        # 第一行：客戶
        customer = lines[0].strip()

        rows = customer_sheet.get_all_values()
        customer_rows = rows[1:]

        # 商品清單（第三行開始）
        product_lines = lines[2:]


        # 先保留輸入值
        input_qty = qty
        input_unit = unit
        input_price = price


        # 預設
        qty = 0
        unit = ""
        price = 0


        


        weekday_map = {
            "一":0,
            "二":1,
            "三":2,
            "四":3,
            "五":4,
            "六":5,
            "日":6
        }    

        

        today = now_tw().date()
       
        # ===============================
        # 下個月起，每個月初送（3個月）
        # ===============================
        if "下個月起" in lines[1] and "每個月初送" in lines[1]:

            year = today.year
            month = today.month + 1

            if month > 12:
                month = 1
                year += 1

            for _ in range(3):

                for day in range(1, 8):      # 每月1~7號

                    try:
                        d = date(year, month, day)

                        date_str = format_date(d)

                        if date_str not in added_dates:
                            added_dates.add(date_str)
                            target_dates.append(date_str)

                    except:
                        pass

                month += 1

                if month > 12:
                    month = 1
                    year += 1

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

                    date_str = format_date(d)

                    if date_str in added_dates:
                        continue

                    added_dates.add(date_str)

                    target_dates.append(date_str)

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

            days_to_next_monday = 7 - today.weekday()

            second_week_start = today + timedelta(
                days=days_to_next_monday + 7
            )

            target_year = second_week_start.year
            target_month = second_week_start.month

            last_day = calendar.monthrange(
                target_year,
                target_month
            )[1]

            end_date = datetime(
                target_year,
                target_month,
                last_day
            ).date()

            current = second_week_start

            while current <= end_date:

                if current.weekday() in target_days:

                    date_str = format_date(current)

                    if date_str not in added_dates:

                        added_dates.add(date_str)
                        target_dates.append(date_str)

                current += timedelta(days=1)

        if not product:

            # ⭐ 沒有客戶 → 全部客戶都排
            if not customer:
                customer_rows = rows[1:]
            else:
                customer_rows = [
                    r for r in rows[1:]
                    if r[0] == customer
                ]

            for r in customer_rows:

                qty_match = re.search(r"(\d+(?:\.\d+)?)", r[2])

                qty = float(qty_match.group(1)) if qty_match else 0

                unit = parse_unit(r[2])

                try:
                    price = float(r[3])
                except:
                    price = 0

                for date_str in target_dates:

                    orders.append({
                        "date": date_str,
                        "customer": customer,
                        "product": r[1],
                        "qty": qty,
                        "unit": unit,
                        "price": price,
                        "delivery": delivery,
                        "note": note if note else "固定排程"
                    })

        for order_text in product_lines:

            parts = order_text.split()

            if not parts:
                continue

            product = parts[0]

            # 先使用 Customers 預設值
            qty = 0
            unit = ""
            price = 0
            delivery = ""
            note = ""

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

            # 商品輸入覆蓋 Customers
            for p in parts[1:]:

                m = re.match(r"(\d+(?:\.\d+)?)(.*)", p)

                if m:
                    qty = float(m.group(1))

                    if m.group(2):
                        unit = parse_unit(m.group(2))

                    break

            for p in parts:

                if p.startswith("@"):
                    try:
                        price = float(p[1:])
                    except:
                        pass

            delivery = detect_delivery(order_text)

            note = order_text
            note = note.replace(product, "", 1)

            qty_text = (
                str(int(qty))
                if float(qty).is_integer()
                else str(qty)
            )

            note = note.replace(f"{qty_text}{unit}", "", 1)

            note = re.sub(r"@\d+(?:\.\d+)?", "", note)

            if delivery:
                note = note.replace(delivery, "")

            note = re.sub(r"\s+", " ", note).strip()

            for date_str in target_dates:

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

    customer = parts[0].strip() if parts else ""
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

    # 移除數量+單位（例如 5箱、10件）
    if qty is not None:
        qty_text = (
            str(int(qty))
            if float(qty).is_integer()
            else str(qty)
        )

        note = note.replace(
            f"{qty_text}{unit}",
            "",
            1
        )

    # 移除單價
    note = re.sub(
        r"@\d+(?:\.\d+)?|@",
        "",
        note
    )

    for d in DELIVERY_LIST:
        note = note.replace(d, "")

    note = re.sub(
        r"每週[一二三四五六日]+到貨到月底",
        "",
        note
    )

    note = re.sub(
        r"\s+",
        " ",
        note
    ).strip()

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


    #商品輸入優先
    


    input_qty = qty
    input_unit = unit
    input_price = price


    # 預設
    qty = 0
    unit = None
    price = 0


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
    # 輸入有值才覆蓋，沒有就用基本檔

    if input_qty is not None:
        qty = input_qty

    if input_unit:
        unit = input_unit

    if input_price is not None and input_price > 0:
        price = input_price

    today = now_tw().date()

    if start_next_week:
        today += timedelta(days=(7 - today.weekday()))

# 取得排程開始月份的月底
    target_year = today.year
    target_month = today.month

    last_day = calendar.monthrange(
        target_year,
        target_month
    )[1]

    end_date = date(
        target_year,
        target_month,
        last_day
    )

    orders = []

# 沒輸入商品 -> 建立客戶所有預設商品
    if not product:

        customer_rows = [
            r for r in rows[1:]
            if r[0] == customer
        ]

        current = today

        while current <= end_date:

            if current.weekday() in target_days:

                for r in customer_rows:

                    qty_match = re.search(r"(\d+(?:\.\d+)?)", r[2])

                    default_qty = float(qty_match.group(1)) if qty_match else 0
                    default_unit = parse_unit(r[2])

                    try:
                        default_price = float(r[3])
                    except:
                        default_price = 0

                    orders.append({
                        "date": format_date(current),
                        "customer": customer,
                        "product": r[1],
                        "qty": default_qty,
                        "unit": default_unit,
                        "price": default_price,
                        "delivery": delivery,
                        "note": note if note else "固定排程"
                    })

            current += timedelta(days=1)

    # 有輸入商品
    else:

        current = today

        while current <= end_date:

            if current.weekday() in target_days:

                orders.append({
                    "date": format_date(current),
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

    return f"✅ 已建 {count} 筆排程"

    
DATE_PATTERN = r"(?:\d{4}/)?\d{1,2}/\d{1,2}"

UNIT_WHITELIST = [
    "包","袋","箱","件","桶","噸","盒","籃","板",
    "公斤","斤","公克",
    "kg","KG","Kg",
    "g","G",
    "lb","LB","lbs",
    "K","k"
]

def schedule_order(text, rows, user_id):

    text = normalize_symbols(text)

    lines = [
        x.rstrip()
        for x in text.splitlines()
    ]

    if len(lines) < 3:
        return "❌ 排單格式錯誤"

    # ===== 第一行 =====

    current_customer = ""

    current_product = ""

    orders = []

    current_date = ""

    delivery = ""

    note = ""


    # ===== 開始解析 =====

    for i, line in enumerate(lines[1:]):

        # ===== 客戶 =====

        # 空白行
        if line == "":
            current_date = ""
            continue

        # ===== 日期 =====
        if re.fullmatch(f"{DATE_PATTERN}", line):

            current_date = format_date(
                parse_date(line)
            )

            delivery = ""
            note = ""

            continue

        # ===== 日期 + 數量（商品固定）=====
        m = re.match(
            rf"^({DATE_PATTERN})\s+(.+)$",
            line
        )

        if m and current_product:

            current_date = format_date(
                parse_date(m.group(1))
            )

            # 組回原本的商品格式
            line = current_product + " " + m.group(2)


        # ===== 配送 =====
        if line in DELIVERY_LIST:

            delivery = line

            continue


        # ===== 備註 =====
        if line.startswith("!"):

            note = line[1:].strip()

            continue


        parts = line.split()

        if not parts:
            continue


        # ===== 商品 =====
        # 第二個欄位像 5kg、2顆、3包...
        if (
            current_date
            and len(parts) >= 2
            and re.match(r"\d+(?:\.\d+)?", parts[1])
        ):
            pass

        # ===== 不是商品 =====
        else:

            # 如果下一行是 日期+數量，代表這行其實是商品
            if (
                current_customer
                and i + 2 < len(lines)
                and re.match(
                    rf"^{DATE_PATTERN}\s+",
                    lines[i + 2]
                )
            ):
                current_product = line
                continue

            # 否則就是新的客戶
            current_customer = line
            current_product = ""
            current_date = ""
            delivery = ""
            note = ""

            continue

        product = parts[0]

        qty = 0

        unit = ""

        price = ""

        remain = []

        for p in parts[1:]:

            # ===== 數量 =====

            m = re.match(
                r"(\d+(?:\.\d+)?)(.*)",
                p
            )

            if m and qty == 0:

                qty = float(m.group(1))

                unit = parse_unit(
                    m.group(2)
                )

                continue

            # ===== 單價 =====

            if p.startswith("@"):

                t = p[1:]

                if t in ("前價", "訂"):

                    price = "前價"

                else:

                    try:

                        v = float(t)

                        if v.is_integer():

                            price = str(int(v))

                        else:

                            price = str(v)

                    except:

                        price = t

                continue

            remain.append(p)

        if remain:

            if note:

                note += " "

            note += " ".join(remain)

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

    if not orders:
        return "❌ 沒有排單資料"

    new_orders = []

    reservation_updates = []

    for order in orders:

        found = False

        for row_no, r in enumerate(rows[1:], start=2):

            status = r[11] if len(r) > 11 else ""

            if status != "預約":
                continue

            if str(r[2]).strip() != "未定":
                continue

            if str(r[3]).strip() != order["customer"]:
                continue

            if str(r[4]).strip() != order["product"]:
                continue

            reserve_qty = float(r[5])

            schedule_qty = float(order["qty"])

            if schedule_qty > reserve_qty:

                return (
                    f"❌ {order['product']} "
                    f"排單數量({schedule_qty}) "
                    f"大於預約數量({reserve_qty})"
                )

            # =========================
            # 單價
            # =========================

            reserve_price = r[7]

            if order["price"] not in ("", None):
                price = order["price"]
            else:
                price = reserve_price

            new_orders.append({

                "date": order["date"],

                "customer": order["customer"],

                "product": order["product"],

                "qty": schedule_qty,

                "unit": order["unit"],

                "price": price,

                "delivery": order["delivery"] or r[8],

                "note": order["note"] or r[9]

            })

            remain = reserve_qty - schedule_qty

            # 更新記憶體中的剩餘數量
            r[5] = str(remain)

            reservation_updates.append({

                "row": row_no,

                "remain": remain

            })

            found = True

            break

        if not found:

            return (
                f"❌ 找不到預約："
                f"{order['customer']} "
                f"{order['product']}"
            )

    save_orders_batch(
        new_orders,
        user_id
    )
    
            # =========================
            # 更新預約
            # =========================
    
    batch_updates = []
    
    for item in reservation_updates:
    
        row = item["row"]
    
        remain = item["remain"]
    
                # 數量歸零
        if remain <= 0:
    
            batch_updates.append({
    
                "range": f"F{row}:L{row}",
    
                "values": [[
                    0,
                    "",
                    "",
                    "",
                    "",
                    "",
                    "已完成"
                ]]
    
            })
    
        # 還有剩餘
        else:
    
            batch_updates.append({
    
                "range": f"F{row}",
    
                "values": [[remain]]
    
            })
    
    if batch_updates:
    
        sheet.batch_update(batch_updates)
    
    return f"✅ 已排 {len(new_orders)} 筆"
        # =========================
        # 建立正式訂單
        # =========================

        

def edit_order(text, rows):
    text = expand_short_dates(text)
    
    text = text.strip()
    text = normalize_symbols(text)

    lines = [
        l.strip()
        for l in text.splitlines()
        if l.strip()
    ]

    blocks = []
    current = []

    for line in lines:

        # 遇到新的「改單」開始新區塊
        if line.startswith("改單") and current:
            blocks.append("\n".join(current))
            current = []

        current.append(line)

    if current:
        blocks.append("\n".join(current))

    updated_total = 0

    batch_updates = []

    for block in blocks:

        lines = [
            l.strip()
            for l in block.split("\n")
            if l.strip()
        ]

        if len(lines) < 2:
            continue

        # =========================
        # 1️⃣ 條件行
        # =========================
        head = lines[0].split()

        if head and head[0] == "改單":
            head = head[1:]

        dates = []
        rest = []

        for t in head:
            if re.match(f"^{DATE_PATTERN}$", t):
                dates.append(t.strip())
            else:
                rest.append(t)

        customer = rest[0] if rest else ""

        product_filter = [
            x.strip()
            for x in rest[1:]
            if x and not re.match(f"^{DATE_PATTERN}$", x)
        ]

        has_product_filter = len(product_filter) > 0

        # =========================
        # 2️⃣ 修改內容解析
        # =========================

        item_updates = []

        for line in lines[1:]:

            parts = line.split()

            if not parts:
                continue

            # 第一個一定是商品名稱
            updates = {}

            # 判斷第一個是不是修改指令
            if parts[0][:1] in ("@", "#", "*", "×", "+", "!"):
                # 沒指定商品，代表全部商品
                target_product = ""
                i = 0
            else:
                # 第一個是商品名稱
                target_product = parts[0]
                i = 1

            while i < len(parts):

                token = parts[i]

                # ===== 單價 =====
                if token.startswith("@"):

                    price_text = token[1:].strip()

                    # 允許 @前價
                    if price_text in ("前價", "訂"):
                        updates["單價"] = "前價"

                    # 允許 @80、@80.5
                    elif re.fullmatch(r"\d+(?:\.\d+)?", price_text):

                        price = float(price_text)

                        if price.is_integer():
                            updates["單價"] = str(int(price))
                        else:
                            updates["單價"] = str(price)

                    # 其它 (@、@成員...) 一律忽略
                    else:
                        i += 1
                        continue

                # ===== 日期 =====
                elif token.startswith("#"):
                    try:
                        updates["日期"] = format_date(parse_date(token[1:]))
                    except:
                        pass

                # ===== 商品名稱 =====
                elif token.startswith("*"):
                    updates["商品"] = token[1:]

                # ===== 數量 =====
                elif token.startswith("×"):

                    v = token[1:].strip()

                    if not v and i + 1 < len(parts):
                        i += 1
                        v = parts[i]

                    m = re.match(r"(\d+(?:\.\d+)?)(.*)", v)

                    if m:
                        updates["數量"] = m.group(1)

                        if m.group(2):
                            updates["單位"] = parse_unit(m.group(2))

                # ===== 配送 =====
                elif token.startswith("+"):
                    updates["配送"] = token[1:]

                # ===== 備註 =====
                elif token.startswith("!"):
                    updates["備註"] = " ".join(
                        [token[1:]] + parts[i+1:]
                    )
                    break

                i += 1

            # 沒有任何修改內容就不要加入
            if updates:
                item_updates.append((target_product, updates))

        # =========================
        # 3️⃣ 套用更新
        # =========================
        for row_no, r in enumerate(rows[1:], start=2):

            sheet_date = str(r[2]).strip()
            try:

                sheet_date_obj = parse_date(sheet_date)

            except:

                continue
            sheet_customer = str(r[3]).strip()
            sheet_product = str(r[4]).strip()

            matched_updates = None

            for target_product, u in item_updates:

                # 沒指定商品 -> 全部商品
                if target_product == "":
                    matched_updates = u
                    break

                # 指定商品
                if sheet_product.strip() == target_product.strip():
                    matched_updates = u
                    break

            if matched_updates is None:
                continue


            updates = matched_updates

            # 日期條件
            if dates:

                target = False

                for d in dates:

                    try:

                        if sheet_date_obj == parse_date(
                            d,
                            sheet_date_obj
                        ):

                            target = True

                            break

                    except:

                        pass

                if not target:

                    continue

            # 客戶條件
            if customer and customer not in sheet_customer:
                continue

            # 商品條件（🔥修正：如果你有要改商品，就不要被 filter 擋）
            if has_product_filter and "商品" not in updates:
                if not any(p in sheet_product for p in product_filter):
                    continue

            # =========================
            # 收集更新資料
            # =========================

            if "日期" in updates:
                batch_updates.append({
                    "range": f"C{row_no}",
                    "values": [[updates["日期"]]]
                })

            if "商品" in updates:
                batch_updates.append({
                    "range": f"E{row_no}",
                    "values": [[updates["商品"]]]
                })

            if "數量" in updates:
                batch_updates.append({
                    "range": f"F{row_no}",
                    "values": [[updates["數量"]]]
                })

            if "單位" in updates:
                batch_updates.append({
                    "range": f"G{row_no}",
                    "values": [[updates["單位"]]]
                })

            if "單價" in updates:
                batch_updates.append({
                    "range": f"H{row_no}",
                    "values": [[updates["單價"]]]
                })

            if "配送" in updates:
                batch_updates.append({
                    "range": f"I{row_no}",
                    "values": [[updates["配送"]]]
                })

            if "備註" in updates:
                batch_updates.append({
                    "range": f"J{row_no}",
                    "values": [[updates["備註"]]]
                })

            updated_total += 1

    if batch_updates:
        sheet.batch_update(batch_updates)

    if updated_total == 0:
        return "❌ 找不到符合訂單"

    return f"✅ 已改 {updated_total} 筆"

def normalize_symbols(text: str):
    table = str.maketrans({
        "×": "×",
        "×": "×",
        "＋": "+",
        "＋": "+",
        "！": "!",
        "＃": "#",
        "＆": "&"
    })
    return text.translate(table)

def parse_unit(text):

    for u in sorted(
        UNIT_WHITELIST,
        key=len,
        reverse=True
    ):
        if u in text:
            return u

    return ""

def detect_delivery(text):
    for d in DELIVERY_LIST:
        if d in text:
            return d
    return ""

def generate_order_ids(count):

    today = now_tw()

    roc_date = (
        f"{today.year - 1911}"
        f"{today.strftime('%m%d')}"
    )

    values = settings_sheet.batch_get(["B1:B2"])

    last_date = values[0][0][0] if values[0] else ""
    last_seq = int(values[0][1][0]) if len(values[0]) > 1 else 0

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

    now_str = now_tw().strftime("%Y-%m-%d %H:%M:%S")

    oid = generate_order_ids(1)[0]

    seq_no = oid[-3:]

    data["date"] = format_date(
        parse_date(data["date"])
    )

    sheet.append_row([
        oid,
        now_str,
        data["date"],
        data["customer"],
        data["product"],
        data["qty"],
        data["unit"],
        data["price"],
        data["delivery"],
        data["note"],
        user_id,
        data.get("status", "正常"),
        "",
        "",
        "",
        str(uuid.uuid4()),
        seq_no
    ])

    return oid

def save_orders_batch(
    orders,
    user_id
):

    rows = []

    now_str = now_tw().strftime("%Y-%m-%d %H:%M:%S")

    order_ids = generate_order_ids(
        len(orders)
    )

    for oid, order in zip(
        order_ids,
        orders
    ):

        if "/" in str(order["date"]):
            order["date"] = format_date(
                parse_date(order["date"])
            )

        seq_no = oid[-3:]

        rows.append([
            oid,
            now_str,
            order["date"],
            order["customer"],
            order["product"],
            order["qty"],
            order["unit"],
            order["price"],
            order["delivery"],
            order["note"],
            user_id,
            order.get("status", "正常"),
            "",
            "",
            "",
            str(uuid.uuid4()),
            seq_no
        ])

    sheet.append_rows(rows)

    return len(rows)

def is_header_line(line):

    line = (
        line.replace("、", " ")
            .replace("，", " ")
            .replace(",", " ")
    )

    tokens = line.split()

    if len(tokens) < 2:
        return False

    # 第一個是日期
    if re.match(r"\d{1,2}/\d{1,2}$", tokens[0]):

        # 找到第一個不是日期的位置
        i = 0
        while i < len(tokens) and re.match(r"\d{1,2}/\d{1,2}$", tokens[i]):
            i += 1

        # 只有日期+客戶
        return i == len(tokens) - 1

    # 第一個是客戶
    if (
        len(tokens) >= 2
        and re.match(r"\d{1,2}/\d{1,2}$", tokens[1])
    ):

        j = 1
        while j < len(tokens) and re.match(r"\d{1,2}/\d{1,2}$", tokens[j]):
            j += 1

        return j == len(tokens)

    return False

def parse_order_line(line):
    parts = line.split()
    if len(parts) < 4:
        return None

    line = line.strip()

    if line.startswith("追加"):
        line = line[2:].strip()

    # ===== 日期、客戶解析（支援多日期）=====
    dates = []
    i = 0

    # 日期在前
    if re.match(f"^{DATE_PATTERN}$", parts[0]):

        while i < len(parts) and re.match(f"^{DATE_PATTERN}$", parts[i]):
            dates.append(parts[i])
            i += 1

        if i >= len(parts):
            return None

        customer = parts[i]
        i += 1

    # 客戶在前
    else:

        customer = parts[0]
        i = 1

        while i < len(parts) and re.match(f"^{DATE_PATTERN}$", parts[i]):
            dates.append(parts[i])
            i += 1

        if not dates or i >= len(parts):
            return None

    product = parts[i]

    if i + 1 >= len(parts):
        return None

    # 支援「全出」
    if parts[i + 1] == "全出":
        qty = "全出"
        unit = ""
    else:
        qty_match = re.search(r"(\d+(?:\.\d+)?)", parts[i + 1])

        if not qty_match:
            return None

        qty = float(qty_match.group(1))
        unit = parse_unit(parts[i + 1])

    price_match = re.search(r"@([^\s]+)", line)

    if price_match:
        price_text = price_match.group(1)

        try:
            price = float(price_text)
        except:
            price = price_text
    else:
        price = 0

    # ===== 配送、備註 =====

    # 找價格(@)的位置
    price_index = None
    for j in range(i + 2, len(parts)):
        if parts[j].startswith("@"):
            price_index = j
            break

    # 取得價格後面的所有內容
    if price_index is None:
        remain = parts[i + 2:]
    else:
        remain = parts[price_index + 1:]

    remain_text = " ".join(remain)

    delivery = detect_delivery(remain_text)

    note = remain_text

    for d in DELIVERY_LIST:
        note = note.replace(d, "", 1)

    note = note.strip()

    orders = []

    for d in dates:
        d = format_date(
            parse_date(d)
        )
        orders.append({
            "date": d,
            "customer": customer,
            "product": product,
            "qty": qty,
            "unit": unit,
            "price": price,
            "delivery": delivery,
            "note": note
        })

    return orders if len(orders) > 1 else orders[0]

def parse_multi_customer_order(text):

    text = text.strip()

    if text.startswith("追加"):
        text = text[2:].strip()

    lines = [x.strip() for x in text.splitlines()]

    orders = []

    current_dates = []
    current_customer = ""

    for line in lines:

        if not line:
            continue

        # 日期分隔符統一
        header = (
            line.replace("、", " ")
        .replace("，", " ")
            .replace(",", " ")
        )

        tokens = header.split()

        

        # 日期在前
        if tokens and re.match(f"^{DATE_PATTERN}$", tokens[0]):

            i = 0

            while i < len(tokens):

                if re.match(f"^{DATE_PATTERN}$", tokens[i]):
                    current_dates.append(tokens[i])
                    i += 1
                else:
                    break

            current_customer = " ".join(tokens[i:])

        # 只有第一行才解析日期/客戶
        if tokens:

            # 日期在前
            if re.match(r"\d{1,2}/\d{1,2}$", tokens[0]):

                current_dates = []

                i = 0
                while i < len(tokens) and re.match(r"\d{1,2}/\d{1,2}$", tokens[i]):
                    current_dates.append(tokens[i])
                    i += 1

                current_customer = " ".join(tokens[i:])

                continue

            # 客戶在前（第二個開始有日期）
            elif (
                len(tokens) > 1
                and re.match(f"^{DATE_PATTERN}$", tokens[1])
            ):

                current_customer = tokens[0]
                current_dates = []

                for t in tokens[1:]:
                    if re.match(r"\d{1,2}/\d{1,2}$", t):
                        current_dates.append(t)

                continue

        if line.startswith("@"):
            continue

        if not current_dates:
            continue

        parts = line.split()

        if len(parts) < 2:
            continue

        product = parts[0]

        qty_text = parts[1]

        # ===== 數量 =====
        if qty_text == "全出":
            qty = "全出"
            unit = ""
        else:
            m = re.match(r"(\d+(?:\.\d+)?)(.*)", qty_text)

            if not m:
                continue

            qty = float(m.group(1))

            unit = parse_unit(m.group(2))

        price = 0
        delivery = ""
        note_list = []

        for p in parts[2:]:

            if p.startswith("@"):

                price_text = p[1:]

                try:
                    price = float(price_text)
                except:
                    price = price_text

                continue

            if p in DELIVERY_LIST:
                delivery = p
                continue

            note_list.append(p)

        note = " ".join(note_list)

        for d in current_dates:
            d = format_date(
                parse_date(d)
            )
            orders.append({
                "date": d,
                "customer": current_customer,
                "product": product,
                "qty": qty,
                "unit": unit,
                "price": price,
                "delivery": delivery,
                "note": note
            })

    return orders

def parse_same_product_orders(text):

    text = text.strip()

    if text.startswith("追加"):
        text = text[2:].strip()

    lines = [x.strip() for x in text.splitlines() if x.strip()]

    if len(lines) < 2:
        return []

    # 第一行：客戶 商品
    head = lines[0].split()

    if len(head) < 2:
        return []

    customer = head[0]
    product = " ".join(head[1:])

    orders = []

    for line in lines[1:]:

        parts = line.split()

        if len(parts) < 2:
            continue

        date = parts[0]

        if not re.match(f"^{DATE_PATTERN}$", date):
            continue

        qty = 0
        unit = "件"
        price = 0
        delivery = ""
        note = ""

        # 第二欄：數量+單位
        # 支援全出
        if parts[1] == "全出":
            qty = "全出"
            unit = ""
        else:
            m = re.match(r"(\d+(?:\.\d+)?)(.*)", parts[1])

            if not m:
                continue

            qty = float(m.group(1))

            if m.group(2):
                unit = parse_unit(m.group(2))

        remain = []

        for p in parts[2:]:

            # 單價
            if p.startswith("@"):

                price_text = p[1:]

                try:
                    price = float(price_text)
                except:
                    price = price_text

                continue

            # 配送
            if p in DELIVERY_LIST:
                delivery = p
                continue

            remain.append(p)

        note = " ".join(remain).strip()

        # 如果沒輸入單價，自動抓 Customers 預設
        if price == 0:

            rows = customer_sheet.get_all_values()

            for r in rows[1:]:

                if len(r) < 4:
                    continue

                if r[0] == customer and r[1] == product:

                    try:
                        price = float(r[3])
                    except:
                        price = 0

                    break

        orders.append({

            "date": format_date(
                parse_date(date)
            ),

            "customer": customer,

            "product": product,

            "qty": qty,

            "unit": unit,

            "price": price,

            "delivery": delivery,

            "note": note

        })

    return orders

def parse_customer_products(text):

    text = text.strip()

    if text.startswith("追加"):
        text = text[2:].strip()

    lines = [x.strip() for x in text.splitlines() if x.strip()]

    rows = customer_sheet.get_all_values()

    orders = []

    customer = ""
    product = ""

    for i, line in enumerate(lines):

        # ===== 日期 =====
        if re.match(f"^{DATE_PATTERN}", line):

            if not customer or not product:
                continue

            parts = line.split()

            date = format_date(
                parse_date(parts[0])
            )

            qty = 0
            unit = ""
            price = 0
            delivery = ""
            note = ""

            # 數量
            if len(parts) >= 2:
                m = re.match(r"(\d+(?:\.\d+)?)(.*)", parts[1])
                if m:
                    qty = float(m.group(1))
                    unit = parse_unit(m.group(2))

            # 預設價格
            for r in rows[1:]:
                if r[0] == customer and r[1] == product:
                    try:
                        price = float(r[3])
                    except:
                        pass
                    break

            remain = []

            for p in parts[2:]:

                if p.startswith("@"):

                    price_text = p[1:]

                    try:
                        price = float(price_text)
                    except:
                        price = price_text

                    continue

                if p in DELIVERY_LIST:
                    delivery = p
                    continue

                remain.append(p)

            note = " ".join(remain)

            orders.append({
                "date": date,
                "customer": customer,
                "product": product,
                "qty": qty,
                "unit": unit,
                "price": price,
                "delivery": delivery,
                "note": note
            })

            continue

        # ===== 非日期 =====

        next_is_date = False

        if i + 1 < len(lines):
            next_is_date = bool(
                re.match(f"^{DATE_PATTERN}", lines[i + 1])
            )

        if next_is_date:
            # 商品
            product = line
        else:
            # 客戶
            customer = line
            product = ""
    
    return orders

def parse_reservation(text):

    text = text.strip()

    if text.startswith("預約"):
        text = text[2:].strip()

    lines = [
        x.strip()
        for x in text.splitlines()
        if x.strip()
    ]

    if len(lines) < 2:
        return []

    rows = customer_sheet.get_all_values()

    customer = lines[0]

    delivery = ""
    note = ""

    orders = []

    for line in lines[1:]:

        # ===== 配送 =====
        if line in DELIVERY_LIST:
            delivery = line
            continue

        # ===== 備註 =====
        if line.startswith("!"):
            note = line[1:].strip()
            continue

        parts = line.split()

        if not parts:
            continue

        product = parts[0]

        qty = 0
        unit = ""
        price = ""

        # 每個商品重新開始
        item_delivery = delivery
        item_note = note

        # 預設價格
        for r in rows[1:]:
            if r[0] == customer and r[1] == product:
                price = r[3]
                break

        remain = []

        for p in parts[1:]:

            # ===== 數量 =====
            m = re.match(r"(\d+(?:\.\d+)?)(.*)", p)

            if m and qty == 0:

                qty = float(m.group(1))
                unit = parse_unit(m.group(2))
                continue

            # ===== 單價 =====
            if p.startswith("@"):

                price_text = p[1:]

                if price_text in ("前價", "訂"):
                    price = "前價"

                else:
                    try:
                        v = float(price_text)

                        if v.is_integer():
                            price = str(int(v))
                        else:
                            price = str(v)

                    except:
                        price = price_text

                continue

            # ===== 配送 =====
            if p in DELIVERY_LIST:

                item_delivery = p
                continue

            remain.append(p)

        if remain:

            if item_note:
                item_note += " "

            item_note += " ".join(remain)

        orders.append({

            "date": "未定",

            "customer": customer,

            "product": product,

            "qty": qty,

            "unit": unit,

            "price": price,

            "delivery": item_delivery,

            "note": item_note,

            "status": "預約"

        })

    return orders

def parse_customer_date_blocks(text):

    text = text.strip()

    if text.startswith("追加"):
        text = text[2:].strip()

    rows = customer_sheet.get_all_values()

    orders = []

    customer = ""
    current_date = ""

    lines = [
        x.strip()
        for x in text.splitlines()
        if x.strip()
    ]

    for line in lines:

        # ===== 日期 =====
        if re.match(f"^{DATE_PATTERN}$", line):

            current_date = format_date(
                parse_date(line)
            )
            continue

        # ===== 客戶 =====
        parts = line.split()

        if (
            len(parts) == 1
            and not re.match(f"^{DATE_PATTERN}$", line)
        ):
            customer = line
            current_date = ""
            continue

        if not customer or not current_date:
            continue

        parts = line.split()

        if not parts:
            continue

        product = parts[0]

        qty = 0
        unit = ""
        price = 0
        delivery = ""
        note = ""

        # Customers 預設價格
        for r in rows[1:]:

            if len(r) < 4:
                continue

            if r[0] == customer and r[1] == product:

                try:
                    price = float(r[3])
                except:
                    price = 0

                break

        remain = []

        for p in parts[1:]:

            # 數量
            # 支援全出
            if p == "全出":
                qty = "全出"
                unit = ""
                continue

            m = re.match(r"(\d+(?:\.\d+)?)(.*)", p)

            if m:

                qty = float(m.group(1))

                if m.group(2):
                    unit = parse_unit(m.group(2))

                continue

            # 單價
            if p.startswith("@"):

                price_text = p[1:]

                try:
                    price = float(price_text)
                except:
                    price = price_text

                continue

            # 配送
            if p in DELIVERY_LIST:

                delivery = p

                continue

            remain.append(p)

        note = " ".join(remain).strip()

        orders.append({

            "date": current_date,

            "customer": customer,

            "product": product,

            "qty": qty,

            "unit": unit,

            "price": price,

            "delivery": delivery,

            "note": note

        })

    return orders

def query_order(text, rows):
    text = expand_short_dates(text)
    keyword = text.replace("查詢", "").strip()

    # ===== 查詢預約 =====
    if keyword.startswith("預約"):

        args = keyword.split()

        customer = args[1] if len(args) >= 2 else ""
        product = args[2] if len(args) >= 3 else ""

        result = []

        for r in rows[1:]:

            status = r[11] if len(r) > 11 else ""

            if status != "預約":
                continue

            if customer and customer not in r[3]:
                continue

            if product and product not in r[4]:
                continue

            result.append(r)

        if not result:
            return "❌ 找不到預約"

        lines = []

        for r in result[:50]:

            qty = r[5] if len(r) > 5 else ""
            unit = r[6] if len(r) > 6 else ""
            price = r[7] if len(r) > 7 else ""
            delivery = r[8] if len(r) > 8 else ""
            note = r[9] if len(r) > 9 else ""

            lines.append(
                f"客戶:{r[3]} 商品:{r[4]} "
                f"數量:{qty}{unit} 單價:{price} "
                f"配送:{delivery} 備註:{note}"
            )

        return "\n".join(lines)
    
    # 保留原本單一查詢
    keywords = keyword.split()

    result = []

    for r in rows[1:]:

        # 跳過已刪除訂單
        status = r[11] if len(r) > 11 else ""

        if status == "已刪除":
            continue

        row_text = " ".join(r)

        # ===== 原本功能：只有一個關鍵字 =====
        if len(keywords) == 1:

            matched = False

            if keyword in row_text:
                matched = True

            else:
                try:
                    if parse_date(keyword) == parse_date(r[2]):
                        matched = True
                except:
                    pass

        # ===== 新增功能：多關鍵字 =====
        else:

            matched = True

            for kw in keywords:

                # 日期
                try:
                    if parse_date(kw) == parse_date(r[2]):
                        continue
                except:
                    pass

                # 文字
                if kw not in row_text:
                    matched = False
                    break

        if matched:
            result.append(r)

    if not result:
        return "❌ 找不到訂單"

    lines = []

    for r in result[:50]:
        qty = r[5] if len(r) > 5 else ""
        unit = r[6] if len(r) > 6 else "件"
        price = r[7] if len(r) > 7 else ""
        delivery = r[8] if len(r) > 8 else ""
        note = r[9] if len(r) > 9 else ""

        lines.append(
            f"單號:{r[0]} 日期:{r[2]} 客戶:{r[3]} 商品:{r[4]} "
            f"數量:{qty}{unit} 單價:{price} 配送:{delivery} 備註:{note}"
        )

    return "\n".join(lines)

def delete_order(text, user_id, rows):
    # 展開日期
    text = expand_short_dates(text)

    # 統一日期區間符號
    text = (
        text.replace("－", "-")   # 全形 -
            .replace("–", "-")   # en dash
            .replace("—", "-")   # em dash
            .replace("～", "~")   # 全形 ~
    )

    # ===== 多行格式 =====
    lines = [
        l.strip()
        for l in text.splitlines()
        if l.strip()
    ]

    parts = text.split()
    if len(parts) < 2:
        return "❌ 刪單格式錯誤"

    deleted = 0

    delete_batch = now_tw().strftime("%Y%m%d%H%M%S")

        # ===== 日期區間刪除 =====
    m = re.search(
        rf"({DATE_PATTERN})\s*[~-]\s*({DATE_PATTERN})",
        text
    )

    if m:

        start_date = m.group(1)
        end_date = m.group(2)

        remain = (
            text.replace("刪單", "", 1)
                .replace(m.group(0), "")
                .strip()
        )

        args = remain.split()

        customer = args[0] if len(args) >= 1 else ""
        product = args[1] if len(args) >= 2 else ""

        # 日期 + 客戶 (+ 商品)

        for i in range(len(rows), 1, -1):

            r = rows[i - 1]

            status = r[11] if len(r) > 11 else ""

            if status == "已刪除":
                continue

            try:
                if not any(
                    parse_date(r[2]) == parse_date(d)
                    for d in dates
                ):
                    continue
            except:
                continue

            if customer and r[3] != customer:
                continue

            if product and r[4] != product:
                continue

            sheet.update(
                f"L{i}:O{i}",
                [[
                    "已刪除",
                    now_tw().strftime("%Y-%m-%d %H:%M:%S"),
                    user_id,
                    delete_batch
                ]]
            )

            deleted += 1

        return f"✅ 已刪 {deleted} 筆" if deleted else "❌ 找不到訂單"

    # ===== 單號刪除（支援多個單號）=====
    order_ids = [x for x in parts[1:] if x.isdigit()]

    if order_ids:

        for i in range(len(rows), 1, -1):

            r = rows[i - 1]

            status = r[11] if len(r) > 11 else ""

            if status == "已刪除":
                continue

            if r[0] not in order_ids:
                continue

            sheet.update(
                f"L{i}:O{i}",
                [[
                    "已刪除",
                    now_tw().strftime("%Y-%m-%d %H:%M:%S"),
                    user_id,
                    delete_batch
                ]]
            )

            deleted += 1

        return f"✅ 已刪 {deleted} 筆" if deleted else "❌ 找不到訂單"

    # ===== 多行刪單 =====
    if len(lines) > 1:

        head = lines[0].split()

        if head[0] == "刪單":
            head = head[1:]

        dates = []
        others = []

        for x in head:
            if re.match(r"\d{1,2}/\d{1,2}", x):
                dates.append(x)
            else:
                others.append(x)

        customer = others[0] if others else ""

        # 第二行開始都是商品
        products = lines[1:]

    else:

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

        products = [product] if product else []

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

            try:
                if any(parse_date(r[2]) == parse_date(d) for d in dates):

                    sheet.update(
                        f"L{i}:O{i}",
                        [[
                            "已刪除",
                            now_tw().strftime("%Y-%m-%d %H:%M:%S"),
                            user_id,
                            delete_batch
                        ]]
                    )

                    deleted += 1
            except:
                pass
        return f"✅ 已刪 {deleted} 筆" if deleted else "❌ 找不到訂單"

# 日期 + 客戶

    for i in range(len(rows), 1, -1):

        r = rows[i - 1]

        status = r[11] if len(r) > 11 else ""

        if status == "已刪除":
            continue

        try:
            if not any(parse_date(r[2]) == parse_date(d) for d in dates):
                continue
        except:
            continue

        if customer and r[3] != customer:
            continue

        if products:
            if r[4] not in products:
                continue

        sheet.update(
            f"L{i}:O{i}",
            [[
                "已刪除",
                now_tw().strftime("%Y-%m-%d %H:%M:%S"),
                user_id,
                delete_batch
            ]]
        )

        deleted += 1

    return f"✅ 已刪 {deleted} 筆" if deleted else "❌ 找不到訂單"

def restore_order(text, rows):
    text = expand_short_dates(text)
    if text.strip() in ["復原最後刪除", "還原最後刪除"]:
        return restore_last_delete()
    parts = text.split()

    if len(parts) < 2:
        return "❌ 復原格式錯誤"

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

            try:
                if any(parse_date(r[2]) == parse_date(d) for d in dates):

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
            except:
                pass
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

        try:
            if parse_date(r[2]) != parse_date(target_date):
                continue
        except:
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

def smart_parse(text: str):
    text = text.strip()
    if text.startswith("追加"):
        text = text[2:].strip()
    # 統一符號
    text = (
        text.replace("／", "/")
            .replace("－", "-")
            .replace("～", "~")
            .replace("　", " ")
            .replace("(", "（")
            .replace(")", "）")
    )
    text = expand_short_dates(text)

    result = {
        "customer": "",
        "dates": [],
        "product": "",
        "qty": 0,
        "unit": "件",
        "price": 0,
        "delivery": "",
        "note": text
    }

    tokens = text.split()

    # =====================
    # 1. 日期（全部抓）
    # =====================
    result["dates"] = re.findall(r"\d{1,2}/\d{1,2}", text)

    # =====================
    # 2. 單價 @80
    # =====================
    m = re.search(r"@(\d+(?:\.\d+)?)", text)
    if m:
        result["price"] = float(m.group(1))

    # =====================
    # 3. 數量 + 單位（抓第一個）
    # =====================
    for t in tokens:
        m = re.match(r"(\d+(?:\.\d+)?)([^\d\s]+)?", t)
        if m:
            result["qty"] = float(m.group(1))
            if m.group(2):
                result["unit"] = parse_unit(m.group(2))
            break

    # =====================
    # 4. 配送
    # =====================
    result["delivery"] = detect_delivery(text)

    # =====================
    # 5. 客戶（用排除法）
    # =====================
    for t in tokens:
        if (
            re.match(r"\d{1,2}/\d{1,2}", t)
            or t.startswith("@")
            or t in DELIVERY_LIST
            or re.match(r"\d", t)
        ):
            continue

        # 第一個合理人名當客戶
        if not result["customer"]:
            result["customer"] = t
            continue

    # =====================
    # 6. 商品（剩下的）
    # =====================
    ignore = set(result["dates"] + [result["customer"]])

    product_tokens = []
    for t in tokens:
        if t in ignore:
            continue
        if t.startswith("@"):
            continue
        if re.match(r"\d+(?:\.\d+)?", t):
            continue
        if t in DELIVERY_LIST:
            continue
        product_tokens.append(t)

    result["product"] = " ".join(product_tokens).strip()

    # =====================
    # 7. 備註（扣掉已解析內容）
    # =====================
    note = text
    note = note.replace(result["customer"], "")
    note = note.replace(result["product"], "")
    for d in DELIVERY_LIST:
        note = note.replace(d, "")
    note = re.sub(r"@\d+(?:\.\d+)?", "", note)
    note = re.sub(r"\d+(?:\.\d+)?[^\s]*", "", note)
    note = note.strip()

    result["note"] = note

    return result

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

    rows = sheet.get_all_values()
    customer_rows = customer_sheet.get_all_values()
    photo_rows = photo_sheet.get_all_values()

    for event in body["events"]:

        if event["type"] != "message":
            continue

        msg_type = event["message"]["type"]

        user_id = event["source"]["userId"]
        user_name = get_user_name(user_id)

        # ===== 收到照片 =====
        if msg_type == "image":

            message_id = event["message"]["id"]

            save_photo(
                user_name,
                message_id
            )

            line_bot_api.reply_message(
                event["replyToken"],
                TextSendMessage(text="📷 照片已儲存")
            )

            continue

        # ===== 以下才處理文字 =====
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

        # ===== 追加下單 =====
        if text.startswith("追加"):

            text = text[2:].strip()

        # 統一全形符號
        text = (
            text.replace("／", "/")
                .replace("－", "-")
                .replace("～", "~")
                .replace("　", " ")
                .replace("(", "（")
                .replace(")", "）")
        )
        text = expand_short_dates(text)
    # ============================

        user_id = event["source"]["userId"]
        user_name = get_user_name(user_id)
        text = text.replace("周", "週")
        if (
            (
                "下週" in text
                and "之後每週" in text
            )
            or
            (
                "下個月起" in text
                and "每個月初送" in text
            )
        ):

            reply = create_schedule_order(text)

            line_bot_api.reply_message(
                event["replyToken"],
                TextSendMessage(text=reply)
            )

            continue

        # 改單 / 排程類：不要拆 block
        if text.startswith((
            "改單",
            "刪單",
            "復原",
            "預約",
            "排單"
        )):
            commands = [text]

        elif (
            "下週" in text
            and "之後每週" in text
        ):
            commands = [text]

        elif (
            "每週" in text
            and "到貨" in text
        ):
            commands = [text]

        if parse_customer_products(text):
            commands = [text]      # 不切 block

        else:
            commands = [text]

        results = []

        for cmd in commands:
            if cmd.startswith("追加"):
                cmd = cmd[2:].strip()
            cmd = cmd.replace("周", "週")
            if cmd.startswith("查詢"):

                results.append(
                    query_order(cmd, rows)
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
                    rows,
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
                        user_name,
                        rows
                    )   
                )

            elif cmd.startswith("預約"):

                orders = parse_reservation(cmd)

                if not orders:

                    results.append("❌ 預約格式錯誤")

                else:

                    count = save_orders_batch(
                        orders,
                        user_name
                    )

                    results.append(
                        f"✅ 已新增 {count} 筆預約"
                    )


            elif cmd.startswith("排單"):

                results.append(
                    schedule_order(
                        cmd,
                        rows,
                        user_name
                    )
                )

            elif cmd.startswith("改單"):

                results.append(
                    edit_order(cmd, rows)
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
                    restore_order(cmd, rows)
                )

            elif cmd.startswith("還原"):

                results.append(
                    restore_order(cmd, rows)
                )

            elif (
                ("下週" in cmd and "之後每週" in cmd)
                or
                ("下個月起" in cmd and "每個月初送" in cmd)
            ):
                results.append(
                    create_schedule_order(cmd)
                )
            
            elif (
                ("每週" in cmd and "到貨" in cmd)
                or
                ("下個月起" in cmd and "每個月初送" in cmd)
            ):

                results.append(
                    create_schedule_order(cmd)
                )

            else:

                count = 0

                lines = []

                for x in cmd.splitlines():

                    x = x.strip()

                    if not x:
                        continue

                    # 整行都是 LINE Tag 時直接忽略
                    if re.fullmatch(r"(?:@\S+\s*)+", x):
                        continue

                    lines.append(x)

                # ===== 多行訂單 =====

                if len(lines) > 1:

                    first = lines[0].split()

    # 新格式：客戶 商品
                    # 新格式
                    # 新格式
                    if (
                        not is_header_line(lines[0])
                        and not re.match(r"\d{1,2}/\d{1,2}$", lines[0])
                    ):

                        # 客戶→日期→商品 (新格式)
                        orders = parse_customer_date_blocks(cmd)

                        # 客戶→商品→日期
                        if not orders:
                            orders = parse_customer_products(cmd)

                        # 客戶+商品固定，下面多日期
                        if not orders:
                            orders = parse_same_product_orders(cmd)

                        if orders:
                            count = save_orders_batch(
                                orders,
                                user_name
                            )

    # 舊格式：日期 客戶
                    elif is_header_line(lines[0]):

                        orders = parse_multi_customer_order(cmd)

                        if orders:
                            count = save_orders_batch(
                                orders,
                                user_name
                            )

    # 每行都是完整訂單
                    else:

                        orders = []

                        for line in lines:

                            data = parse_order_line(line)

                            if isinstance(data, list):
                                orders.extend(data)
                            elif data:
                                orders.append(data)

                        if orders:
                            count = save_orders_batch(
                                orders,
                                user_name
                            )

# ===== 單行 =====
                else:

                    orders = []

                    data = parse_order_line(lines[0])

                    if isinstance(data, list):
                        orders.extend(data)
                    elif data:
                        orders.append(data)

                    if orders:
                        count = save_orders_batch(
                            orders,
                            user_name
                        )
                

                # ===== 每行都是完整訂單 =====
                                

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


from fastapi import Response

@app.api_route("/", methods=["GET", "HEAD"])
def home():
    if Response:
        return {"status": "running"}