import imaplib
import email
from email.header import decode_header
from email.utils import parseaddr
import re
from datetime import datetime
import pandas as pd
from bs4 import BeautifulSoup
import os
from dotenv import load_dotenv

# ────────────────────────────────────────────────
#          Загрузка настроек из .env
# ────────────────────────────────────────────────

load_dotenv()  # загружает .env в os.environ

IMAP_SERVER   = os.getenv('IMAP_SERVER')
EMAIL_FOLDER  = os.getenv('EMAIL_FOLDER')
EXCEL_FILE    = os.getenv('EXCEL_FILE')
EMAIL_USER = os.getenv('EMAIL_USER')
EMAIL_KEY = os.getenv('EMAIL_KEY')

# ────────────────────────────────────────────────
#          Вспомогательные функции
# ────────────────────────────────────────────────

def decode_subject(subject):
    if not subject:
        return '(без темы)'
    decoded = ''
    for part, encoding in decode_header(subject):
        try:
            if isinstance(part, bytes):
                decoded += part.decode(encoding or 'utf-8', errors='replace')
            else:
                decoded += str(part)
        except:
            decoded += str(part)
    return decoded.strip()

def get_html_part(msg):
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == 'text/html':
                charset = part.get_content_charset() or 'utf-8'
                return part.get_payload(decode=True).decode(charset, errors='ignore')
    elif msg.get_content_type() == 'text/html':
        charset = msg.get_content_charset() or 'utf-8'
        return msg.get_payload(decode=True).decode(charset, errors='ignore')
    return None

def extract_datetime_from_text(text):
    m = re.search(r'(\d{2}\.\d{2}\.\d{4})\s+(\d{2}:\d{2})', text)
    if m:
        try:
            return datetime.strptime(f"{m.group(1)} {m.group(2)}", "%d.%m.%Y %H:%M")
        except:
            pass
    m2 = re.search(r'(\d{2}\.\d{2}\.\d{4})', text)
    if m2:
        try:
            return datetime.strptime(m2.group(1), "%d.%m.%Y")
        except:
            pass
    return None

def parse_receipt_items(html_text):
    soup = BeautifulSoup(html_text, 'html.parser')
    full_text = soup.get_text(separator='\n', strip=True)

    dt = extract_datetime_from_text(full_text)

    items = []
    table = soup.find('table')
    if table:
        rows = table.find_all('tr')

        for row in rows:
            tds = row.find_all(['td', 'th'])
            if not tds:
                continue
            texts = [td.get_text(strip=True) for td in tds]

            if 'Наименование' in texts:
                try:
                    num = 1
                    index = texts.index(f'{num}.')
                    while index != -1:
                        if '.' not in texts[index]:
                            break

                        name = texts[index + 1]
                        price_str = texts[index + 2].replace(' ', '').replace(',', '.')
                        qty_str  = texts[index + 3].replace(' ', '')
                        sum_str  = texts[index + 4].replace(' ', '').replace(',', '.')

                        price = float(price_str)
                        qty   = int(qty_str)
                        total = float(sum_str)

                        if price is not None and qty is not None:
                            items.append({
                                'name': name,
                                'quantity': qty,
                                'price_per_unit': price,
                                'sum': total
                            })

                        num += 1
                        index = texts.index(f'{num}.')

                except:
                    break
                break

    return items, dt

# ────────────────────────────────────────────────
#                   Основная логика
# ────────────────────────────────────────────────

def main():
    print("Парсер электронных чеков от 1-ofd.ru (Яндекс Почта)\n")

    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER)
        mail.login(EMAIL_USER, EMAIL_KEY)
        mail.select(EMAIL_FOLDER)
        print(f"Подключено → папка {EMAIL_FOLDER}")
    except Exception as e:
        print(f"Ошибка входа: {e}")
        return

    try:
        status, data = mail.search(None, 'FROM "1-ofd.ru"')
        if status != 'OK':
            print("Ошибка поиска")
            mail.logout()
            return
        msg_ids = data[0].split()
        print(f"Найдено писем от 1-ofd.ru: {len(msg_ids)}")
    except Exception as e:
        print(f"Ошибка поиска: {e}")
        mail.logout()
        return

    all_rows = []

    for idx, msg_id in enumerate(msg_ids, 1):
        try:
            _, msg_data = mail.fetch(msg_id, '(RFC822)')
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)

            subject = decode_subject(msg['Subject'])
            sender = parseaddr(msg['From'])[1]

            html = get_html_part(msg)
            if not html:
                print(f"[{idx:3d}] Нет HTML-части → пропуск")
                continue

            items, receipt_dt = parse_receipt_items(html)

            if not items:
                print(f"[{idx:3d}] Не найдены элементы чека -> пропускаем")
                continue

            date_str = receipt_dt.strftime('%d.%m.%Y') if receipt_dt else ''
            time_str = receipt_dt.strftime('%H:%M')    if receipt_dt else ''

            for item in items:
                all_rows.append({
                    'Дата': date_str,
                    'Время': time_str,
                    'Наименование': item['name'],
                    'Количество': item['quantity'],
                    'Цена за шт': item['price_per_unit'],
                    'Сумма': item['sum'],
                    'Название письма': subject,
                    'Письмо от': sender
                })

            mark = '✓' if items else '–'
            print(f"[{idx:3d}/{len(msg_ids)}] {mark}  {subject[:60]}{'...' if len(subject)>60 else ''}")

        except Exception as e:
            print(f"[{idx:3d}] Ошибка: {str(e)[:80]}")

    mail.logout()

    if not all_rows:
        print("\nПодходящих чеков не найдено.")
        return

    df = pd.DataFrame(all_rows)

    if 'Дата' in df.columns and 'Время' in df.columns:
        df['sort_dt'] = pd.to_datetime(df['Дата'] + ' ' + df['Время'], format='%d.%m.%Y %H:%M', errors='coerce')
        df = df.sort_values(['sort_dt']).drop(columns=['sort_dt']).reset_index(drop=True)

    final_columns = ['Наименование', 'Количество', 'Цена за шт', 'Сумма', 'Название письма']
    existing_cols = [c for c in final_columns if c in df.columns]
    df_final = df[existing_cols]

    try:
        df_final.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
        print(f"\nСохранено: {os.path.abspath(EXCEL_FILE)}")
        print(f"Строк с товарами: {len(df_final)}")
        print(f"Уникальных чеков: {df['Название письма'].nunique()}")
    except Exception as e:
        print(f"Ошибка сохранения Excel: {e}")

if __name__ == '__main__':
    main()