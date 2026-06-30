from datetime import datetime
from email import message_from_bytes
import os
import pandas as pd
from bs4 import BeautifulSoup
from imapclient import IMAPClient
from env import USER, PASSWORD

HOST = "imap.gmail.com"
USER = USER
PASSWORD = PASSWORD

REPORT_MONTH = "01.05.2026"


def get_month_dates(date_str):
    dt = datetime.strptime(date_str, "%d.%m.%Y")
    since = dt.replace(day=1)

    if since.month == 12:
        before = since.replace(year=since.year + 1, month=1)
    else:
        before = since.replace(month=since.month + 1)

    return (
        since.strftime("%d-%b-%Y"),
        before.strftime("%d-%b-%Y"),
    )


def get_html(msg):
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() != "text/html":
                continue
            payload = part.get_payload(decode=True)
            if payload is None:
                continue
            charset = part.get_content_charset() or "utf-8"
            return payload.decode(charset, errors="replace")
    else:
        if msg.get_content_type() == "text/html":
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or "utf-8"
                return payload.decode(charset, errors="replace")
    return ""


def parse_receipt(html):
    soup = BeautifulSoup(html, "html.parser")
    rows = []

    # 1. Извлекаем номер чека универсальным способом
    receipt_no = "Неизвестно"
    for span in soup.find_all(["span", "td"]):
        text = span.get_text()
        if "КАССОВЫЙ ЧЕК №" in text:
            # Проверяем, идет ли номер следом
            next_span = span.find_next(["span", "td"])
            if next_span:
                num_candidate = next_span.get_text(strip=True)
                if num_candidate.isdigit():
                    receipt_no = num_candidate
                    break

    # 2. Универсальный поиск товаров по маркеру единиц измерения (; шт. или ; кг)
    for span in soup.find_all("span"):
        span_text = span.get_text()
        if "; шт" in span_text or "; кг" in span_text:
            try:
                # Очищаем название товара от лишних переносов строк и пробелов
                name = " ".join(span_text.split())

                # Ищем следующую таблицу, содержащую знак умножения 'x' (пропуская таблицы [M])
                table = span.find_next("table")
                while table:
                    if "x" in table.get_text():
                        break
                    table = table.find_next("table")

                if not table:
                    continue

                tds = table.find_all("td")
                if len(tds) >= 2:
                    # Извлекаем количество (текст до знака 'x' в первой ячейке)
                    qty_text = tds[0].get_text(strip=True).split("x")[0].strip()
                    # Извлекаем финальную сумму из второй ячейки
                    total_text = tds[1].get_text(strip=True)

                    rows.append(
                        {
                            "№ чека": receipt_no,
                            "Товар": name,
                            "Количество": float(qty_text.replace(",", ".")),
                            "Сумма": float(total_text.replace(",", ".")),
                        }
                    )
            except Exception:
                continue

    return rows


def load_all_receipts():
    since, before = get_month_dates(REPORT_MONTH)
    all_items = []

    with IMAPClient(HOST, ssl=True) as client:
        client.login(USER, PASSWORD)
        client.select_folder("INBOX")

        uids = client.search(
            [
                "SUBJECT", "лента",
                "SINCE", since,
                "BEFORE", before,
            ],
            charset="UTF-8",
        )

        print(f"Найдено писем с чеками: {len(uids)}")

        for i, uid in enumerate(uids, start=1):
            print(f"Обрабатываю письмо {i}/{len(uids)}")
            data = client.fetch([uid], ["RFC822"])
            raw = data[uid][b"RFC822"]
            msg = message_from_bytes(raw)
            html = get_html(msg)

            if not html:
                continue

            all_items.extend(parse_receipt(html))

    return all_items


def main():
    rows = load_all_receipts()
    df_all = pd.DataFrame(rows)

    if df_all.empty:
        print("Товары не найдены.")
        return

    # Округляем значения в детальном списке
    df_all["Количество"] = df_all["Количество"].round(3)
    df_all["Сумма"] = df_all["Sумма"] = df_all["Сумма"].round(2)

    # Создаем сводную таблицу для первого листа
    df_summary = (
        df_all.groupby("Товар", as_index=False)
        .agg({"Количество": "sum", "Сумма": "sum"})
        .sort_values("Сумма", ascending=False)
    )
    df_summary["Количество"] = df_summary["Количество"].round(3)
    df_summary["Сумма"] = df_summary["Сумма"].round(2)

    # Запись в Excel на два разных листа
    output_file = "Покупки_Лента.xlsx"
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="Итог", index=False)
        df_all.to_excel(writer, sheet_name="Все чеки", index=False)

    print()
    print(f"Обработка завершена успешно!")
    print(f"Всего позиций в детальном списке: {len(df_all)}")
    print(f"Всего уникальных товаров в своднике: {len(df_summary)}")
    print(f"Файл сохранен: {output_file}")


if __name__ == "__main__":
    main()