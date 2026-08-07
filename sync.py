# -*- coding: utf-8 -*-
"""
sync.py - Синхронизация отчётов "Отгрузки Астана" из почты в дашборд

Логика:
1. Подключается к почте (IMAP) по настройкам из config.json
2. Ищет письмо по теме (email_search.subject_contains)
3. Скачивает XLSX-вложение в downloads/
4. Парсит вложение и baseline-файл, считает суммы по категориям
5. Забирает текущие данные с дашборда (GET api_url)
6. Обновляет запись за текущий месяц (current_period) и заливает обратно (POST api_url)
7. Запоминает UID обработанного письма в state/last_uid.txt, чтобы не грузить повторно

ВАЖНО: этот файл восстановлен по логам и config.json после случайной перезаписи
оригинального sync.py содержимым restore_july.py. Логика парсинга колонок
XLSX (какая колонка = категория, какая = сумма) восстановлена по эвристике
названий заголовков и может потребовать правки под реальную структуру отчёта.
Перед постановкой в Task Scheduler — обязательно прогнать вручную и сверить
вывод (Category vocabulary size, Baseline total, Current total) с прошлыми
логами (sync.log), чтобы убедиться, что числа совпадают с ожидаемыми.
"""

import json
import os
import imaplib
import email
import calendar
from datetime import datetime, timezone, timedelta
from email.header import decode_header

import requests
from openpyxl import load_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
DOWNLOADS_DIR = os.path.join(BASE_DIR, "downloads")
STATE_DIR = os.path.join(BASE_DIR, "state")
LAST_UID_FILE = os.path.join(STATE_DIR, "last_uid.txt")

os.makedirs(DOWNLOADS_DIR, exist_ok=True)
os.makedirs(STATE_DIR, exist_ok=True)


def load_config():
    with open(CONFIG_FILE, "r", encoding="utf-8-sig") as f:
        cfg = json.load(f)

    # Поддержка секретов через переменные окружения (для запуска в GitHub
    # Actions или на сервере, где config.json не должен содержать реальные
    # пароли/токены). Если переменная задана — она перекрывает значение
    # из config.json. Локально на ноутбуке эти переменные обычно не заданы,
    # и используются значения прямо из config.json как раньше.
    env_imap_password = os.environ.get("IMAP_PASSWORD")
    if env_imap_password:
        cfg["imap"]["password"] = env_imap_password

    env_upload_token = os.environ.get("UPLOAD_TOKEN")
    if env_upload_token:
        cfg["dashboard"]["upload_token"] = env_upload_token

    return cfg


def decode_mime_words(s):
    if not s:
        return ""
    decoded = decode_header(s)
    return "".join(
        (t.decode(enc or "utf-8") if isinstance(t, bytes) else t)
        for t, enc in decoded
    )


def get_last_uid():
    if os.path.exists(LAST_UID_FILE):
        with open(LAST_UID_FILE, "r", encoding="utf-8") as f:
            content = f.read().strip()
            return content if content else None
    return None


def save_last_uid(uid):
    with open(LAST_UID_FILE, "w", encoding="utf-8") as f:
        f.write(str(uid))


def find_and_download_email(cfg):
    imap_cfg = cfg["imap"]
    search_cfg = cfg["email_search"]

    mail = imaplib.IMAP4_SSL(imap_cfg["host"], imap_cfg["port"])
    mail.login(imap_cfg["username"], imap_cfg["password"])
    mail.select(imap_cfg.get("mailbox", "INBOX"))

    # ВАЖНО: критерий поиска в IMAP SEARCH ограничен ASCII, а тема письма на
    # кириллице ("Отгрузки Астана") ломает протокол при прямой передаче.
    # Поэтому берём ALL/UNSEEN (это ASCII-safe), а фильтрацию по теме/отправителю
    # делаем уже в Python после декодирования заголовков.
    search_key = "UNSEEN" if search_cfg.get("only_unread") else "ALL"
    status, data = mail.search(None, search_key)
    if status != "OK" or not data or not data[0]:
        print("No emails found.")
        mail.logout()
        return None

    uids = data[0].split()
    last_uid = get_last_uid()

    subject_filter = (search_cfg.get("subject_contains") or "").lower()
    sender_filter = (search_cfg.get("sender_contains") or "").lower()

    target_uid = None
    target_subject = None

    # Идём от новых писем к старым, ищем первое подходящее по теме/отправителю
    for uid in reversed(uids):
        status, hdr_data = mail.fetch(uid, "(BODY.PEEK[HEADER.FIELDS (SUBJECT FROM)])")
        if status != "OK" or not hdr_data or not hdr_data[0]:
            continue
        raw_header = hdr_data[0][1]
        header_msg = email.message_from_bytes(raw_header)
        subject = decode_mime_words(header_msg.get("Subject", ""))
        from_ = decode_mime_words(header_msg.get("From", ""))

        if subject_filter and subject_filter not in subject.lower():
            continue
        if sender_filter and sender_filter not in from_.lower():
            continue

        target_uid = uid
        target_subject = subject
        break

    if target_uid is None:
        print("No matching emails found.")
        mail.logout()
        return None

    if last_uid is not None and target_uid.decode() == last_uid:
        print("Latest matching email already processed (UID {}). Nothing to do.".format(last_uid))
        mail.logout()
        return None

    print("Found email: {}".format(target_subject))

    status, msg_data = mail.fetch(target_uid, "(RFC822)")
    if status != "OK":
        print("Failed to fetch email UID {}".format(target_uid))
        mail.logout()
        return None

    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)

    saved_path = None
    for part in msg.walk():
        if part.get_content_disposition() == "attachment":
            filename = part.get_filename()
            if filename:
                filename = decode_mime_words(filename)
                if filename.lower().endswith(".xlsx"):
                    filepath = os.path.join(DOWNLOADS_DIR, filename)
                    with open(filepath, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    print("Saved attachment to {}".format(filepath))
                    saved_path = filepath

    mail.logout()

    if saved_path:
        save_last_uid(target_uid.decode())

    return saved_path


def parse_report_tree(filepath):
    """
    Парсит иерархический отчёт 1С, используя реальные уровни группировки
    Excel (outline_level), а не визуальные отступы.

    Структура (подтверждена диагностикой файла):
        level 0: "Астана Подразделение" (департамент, итоговая сумма)
        level 1: менеджер/отдел
        level 2: РАЙОН (для менеджеров с "МПП" в названии) ИЛИ сразу КЛИЕНТ
                 (для внутренних отделов: Ремонт Астана, Розница Астана,
                 Розница МСЦ, Сотрудники Астана — у них нет уровня "район")
        level 3: клиент (для веток с районом) ИЛИ категория (для веток без района)
        далее вглубь: категория/подкатегория/... /товар (строка с Артикулом)

    Возвращает (grand_total, clients, products):
        grand_total - число, сумма по всему департаменту
        clients     - dict {(manager, client): сумма}
        products    - dict {(manager, client): {(category, product): сумма}}
    """
    # Отделы, где нет промежуточного уровня "район" между менеджером и клиентом
    NO_DISTRICT_MANAGERS = {
        "Ремонт Астана",
        "Розница Астана",
        "Розница МСЦ",
        "Сотрудники Астана",
    }

    wb = load_workbook(filepath, data_only=True)
    ws = wb.active

    header_row_idx = None
    col_sum = None
    max_scan_rows = 20
    for r in range(1, max_scan_rows + 1):
        row_values = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        row_lower = [str(v).strip().lower() if v else "" for v in row_values]
        sum_idx = next((i for i, v in enumerate(row_lower) if "сумма" in v), None)
        if sum_idx is not None:
            header_row_idx = r
            col_sum = sum_idx + 1
            break

    if header_row_idx is None:
        raise ValueError(
            "Could not find header row with 'Сумма' column in {} "
            "(scanned first {} rows).".format(filepath, max_scan_rows)
        )

    col_article = None
    for r in range(header_row_idx, header_row_idx + 3):
        row_values = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        row_lower = [str(v).strip().lower() if v else "" for v in row_values]
        idx = next((i for i, v in enumerate(row_lower) if "артикул" in v), None)
        if idx is not None:
            col_article = idx + 1
            break

    grand_total = None
    clients = {}
    products = {}
    path = {}
    current_manager = None
    current_client = None

    for r in range(header_row_idx + 1, ws.max_row + 1):
        level = ws.row_dimensions[r].outline_level if r in ws.row_dimensions else 0
        name = ws.cell(row=r, column=1).value
        if name is None:
            continue
        name = str(name).strip()

        sum_val = ws.cell(row=r, column=col_sum).value
        article = ws.cell(row=r, column=col_article).value if col_article else None

        path[level] = name
        for lv in list(path.keys()):
            if lv > level:
                del path[lv]

        if level == 0:
            if grand_total is None and sum_val is not None:
                try:
                    grand_total = float(sum_val)
                except (TypeError, ValueError):
                    pass
            continue

        if level == 1:
            current_manager = name
            current_client = None
            continue

        if level == 2:
            if current_manager in NO_DISTRICT_MANAGERS:
                current_client = name
                if sum_val is not None:
                    try:
                        clients[(current_manager, current_client)] = float(sum_val)
                    except (TypeError, ValueError):
                        pass
            continue

        if level == 3:
            if current_manager not in NO_DISTRICT_MANAGERS:
                current_client = name
                if sum_val is not None:
                    try:
                        clients[(current_manager, current_client)] = float(sum_val)
                    except (TypeError, ValueError):
                        pass
            continue

        # Уровни глубже клиента: категории/подкатегории/товары
        if article and current_manager and current_client and sum_val is not None:
            category = path.get(level - 1, "")
            try:
                amount = float(sum_val)
            except (TypeError, ValueError):
                continue
            key = (current_manager, current_client)
            products.setdefault(key, {})
            pkey = (category, name)
            products[key][pkey] = products[key].get(pkey, 0.0) + amount

    if grand_total is None:
        raise ValueError("Could not find top-level total row in {}".format(filepath))

    return grand_total, clients, products


def build_client_records(baseline_clients, current_clients):
    """Сопоставляет клиентов между baseline (2025) и текущим годом (2026)."""
    keys = set(baseline_clients) | set(current_clients)
    records = []
    for manager, client in keys:
        y2025 = baseline_clients.get((manager, client), 0.0)
        y2026 = current_clients.get((manager, client), 0.0)

        if y2025 == 0 and y2026 > 0:
            status, delta_pct = "new", None
        elif y2025 > 0 and y2026 == 0:
            status, delta_pct = "lost", -100.0
        else:
            delta_pct = (y2026 - y2025) / abs(y2025) * 100 if y2025 != 0 else 0.0
            if abs(delta_pct) <= 15:
                status = "flat"
            elif delta_pct > 0:
                status = "up"
            else:
                status = "down"

        records.append({
            "manager": manager,
            "client": client,
            "y2025": round(y2025, 2),
            "y2026": round(y2026, 2),
            "status": status,
            "deltaPct": round(delta_pct, 1) if delta_pct is not None else None,
        })
    return records


def build_products_dict(baseline_products, current_products):
    """Строит products{"Менеджер||Клиент": [{category, product, y2025, y2026}]}."""
    keys = set(baseline_products) | set(current_products)
    result = {}
    for manager, client in keys:
        dict_key = "{}||{}".format(manager, client)
        b = baseline_products.get((manager, client), {})
        c = current_products.get((manager, client), {})
        pkeys = set(b) | set(c)
        items = []
        for category, product in pkeys:
            items.append({
                "category": category,
                "product": product,
                "y2025": round(b.get((category, product), 0.0), 2),
                "y2026": round(c.get((category, product), 0.0), 2),
            })
        result[dict_key] = items
    return result


def main():
    cfg = load_config()

    new_file = find_and_download_email(cfg)
    if not new_file:
        print("Done. No new data to sync.")
        return

    baseline_file = os.path.join(BASE_DIR, cfg["current_period"]["baseline_file"])
    print("Parsing baseline file: {}".format(baseline_file))
    baseline_total, baseline_clients, baseline_products = parse_report_tree(baseline_file)

    print("Parsing new file: {}".format(new_file))
    current_total, current_clients, current_products = parse_report_tree(new_file)

    print("Baseline total: {:.2f} | Current total: {:.2f}".format(baseline_total, current_total))
    print("Baseline clients: {} | Current clients: {}".format(len(baseline_clients), len(current_clients)))

    clients = build_client_records(baseline_clients, current_clients)
    products = build_products_dict(baseline_products, current_products)

    new_count = sum(1 for c in clients if c["status"] == "new")
    lost_count = sum(1 for c in clients if c["status"] == "lost")
    print("Clients matched: {} total | {} new | {} lost".format(len(clients), new_count, lost_count))

    dash_cfg = cfg["dashboard"]
    api_url = dash_cfg["api_url"]
    upload_token = dash_cfg["upload_token"]

    print("Fetching current dashboard data from {} ...".format(api_url))
    resp = requests.get(api_url, timeout=30)
    resp.raise_for_status()
    try:
        current = resp.json()
    except ValueError:
        print(
            "Dashboard did not return valid JSON. Status={}, Content-Type={}, first 300 chars:\n{}".format(
                resp.status_code, resp.headers.get("Content-Type"), resp.text[:300]
            )
        )
        raise

    months = current.get("months") or []

    # Текущая дата по времени Астаны (UTC+5) — нужна для прогноза "к концу
    # месяца" на дашборде: сколько дней уже прошло и сколько всего дней в месяце.
    astana_now = datetime.now(timezone.utc) + timedelta(hours=5)
    as_of_day = astana_now.day
    days_in_month = calendar.monthrange(astana_now.year, astana_now.month)[1]

    period = cfg["current_period"]
    period_label = period["month_label"].rstrip(",").strip()
    new_entry = {
        "id": period["month_id"],
        "label": period_label,
        "baseYear": 2025,
        "curYear": 2026,
        "hasCompare": True,
        "periodLabel": period_label,
        "asOfDay": as_of_day,
        "daysInMonth": days_in_month,
        "clients": clients,
        "products": products,
    }

    found = False
    for i, m in enumerate(months):
        if m.get("id") == period["month_id"]:
            months[i] = new_entry
            found = True
            break
    if not found:
        months.append(new_entry)

    print("Uploading updated data ({} months total)...".format(len(months)))
    resp = requests.post(
        api_url,
        json={"token": upload_token, "months": months},
        timeout=60,
    )
    resp.raise_for_status()
    result = resp.json()
    print("Upload result:", result)
    print("Done. Refresh {} to check.".format(cfg["dashboard"]["api_url"].replace("/api/data", "")))


if __name__ == "__main__":
    main()
