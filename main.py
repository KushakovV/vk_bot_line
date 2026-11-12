import os
import vk_api
from vk_api.longpoll import VkLongPoll, VkEventType
from openpyxl import Workbook, load_workbook
from datetime import datetime
import time
import requests
from flask import Flask, request, jsonify
import threading
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

VK_TOKEN = os.environ.get('VK_TOKEN')
MAX_TOKEN = os.environ.get('MAX_TOKEN')
MAX_CHAT_ID = os.environ.get('MAX_CHAT_ID')
MAX_BOT_TOKEN = os.environ.get('MAX_BOT_TOKEN')
YANDEX_TOKEN = os.environ.get('YANDEX_TOKEN')

YANDEX_BASE_FOLDER = "Материалы к сообщениям жителей Костромской области"

TOPICS = [
    "ЖКХ и коммунальные услуги", "Дороги и транспорт", "Здравоохранение",
    "Образование", "Социальная поддержка и льготы", "Благоустройство и экология",
    "Работа и занятость", "Сельское хозяйство", "Строительство и земельные вопросы", "Другое"
]

QUESTIONS = [
    "Введите Ваше ФИО", "Укажите Ваш пол (М/Ж)", "Укажите Ваш возраст (число полных лет)",
    "Укажите Ваш населенный пункт", "Укажите Ваш телефон или почту для связи",
    "Выберите тематику (напишите номер от 1 до 10)\n" + "\n".join([f"{i+1}. {t}" for i, t in enumerate(TOPICS)])
]

VK_LAST_QUESTION = "Напишите Ваш вопрос (можно приложить фото)"
MAX_LAST_QUESTION = "Напишите Ваш вопрос (можно приложить медиафайл)"

THANKS_TEXT = "Спасибо! Сообщение принято."
FINISHED_TEXT = "Если Вы хотите направить еще одно сообщение, напишите мне «Новое сообщение», и я с радостью помогу Вам."

vk_states = {}
vk_data = {}
max_states = {}
max_data = {}

# Для защиты от дубликатов MAX
processed_max_messages = set()

# --- Яндекс Диск ---
def yandex_api_request(method, endpoint, params=None, data=None, files=None):
    url = f"https://cloud-api.yandex.net/v1/disk{endpoint}"
    headers = {"Authorization": f"OAuth {YANDEX_TOKEN}"}
    try:
        if method == 'GET':
            r = requests.get(url, headers=headers, params=params, timeout=30)
        elif method == 'PUT':
            r = requests.put(url, headers=headers, params=params, timeout=30)
        elif method == 'POST':
            r = requests.post(url, headers=headers, params=params, data=data, files=files, timeout=30)
        return r
    except Exception as e:
        raise Exception(f"Ошибка запроса к Яндекс Диску: {e}")

def create_yandex_folder(folder_name):
    if not YANDEX_TOKEN: return False
    folder_path = f"/{YANDEX_BASE_FOLDER}/{folder_name}"
    yandex_api_request('PUT', '/resources', params={'path': folder_path})
    return True

def upload_to_yandex_disk(file_bytes, filename, folder_name):
    if not YANDEX_TOKEN:
        raise Exception("Не задан YANDEX_TOKEN")
    
    folder_path = f"/{YANDEX_BASE_FOLDER}/{folder_name}"
    remote_path = f"{folder_path}/{filename}"
    
    r = yandex_api_request('GET', '/resources/upload', params={'path': remote_path, 'overwrite': 'true'})
    if not r or r.status_code != 200:
        raise Exception(f"Не удалось получить ссылку: {r.text if r else 'no response'}")
    
    upload_url = r.json().get('href')
    upload_resp = requests.put(upload_url, data=file_bytes, timeout=60)
    if upload_resp.status_code not in (200, 201, 202):
        raise Exception(f"Ошибка загрузки файла: {upload_resp.status_code}")
    
    r = yandex_api_request('PUT', '/resources/publish', params={'path': remote_path})
    if not r or r.status_code != 200:
        raise Exception("Не удалось опубликовать файл")
        
    r = yandex_api_request('GET', '/resources', params={'path': remote_path, 'fields': 'public_url'})
    if r and r.status_code == 200:
        return r.json().get('public_url')
    raise Exception("Не удалось получить ссылку на файл")

# --- Скачивание из ВК и MAX ---
def download_vk_attachment(attachment):
    try:
        url = None
        filename = "attachment"
        ext = "dat"
        
        if attachment.get('type') == 'photo':
            sizes = attachment['photo']['sizes']
            url = max(sizes, key=lambda x: x.get('width', 0) * x.get('height', 0))['url']
            ext = "jpg"
            filename = f"photo_{attachment['photo']['id']}.{ext}"
        elif attachment.get('type') == 'doc':
            url = attachment['doc']['url']
            ext = attachment['doc'].get('ext', 'dat')
            filename = f"doc_{attachment['doc']['id']}.{ext}"
        elif attachment.get('type') == 'video':
            print("Видео из ВК пропущено (нет прямой ссылки)")
            return None, None, None
            
        if not url: return None, None, None
        
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        return resp.content, filename, ext
    except Exception as e:
        raise Exception(f"Ошибка скачивания вложения ВК: {e}")

def download_max_attachment(attachment):
    try:
        att_type = attachment.get('type')
        url = attachment.get('url') or attachment.get('payload', {}).get('url')
        if not url:
            raise Exception(f"Нет URL для {att_type}")
        
        filename = attachment.get('filename', f"max_{int(time.time())}_{att_type}.dat")
        ext = filename.split('.')[-1] if '.' in filename else 'dat'
        
        resp = requests.get(url, timeout=30, verify=False)
        resp.raise_for_status()
        return resp.content, filename, ext
    except Exception as e:
        raise Exception(f"Ошибка скачивания вложения MAX: {e}")

# --- Сохранение и отправка ---
def save_to_excel(data):
    filename = 'zayavki.xlsx'
    if not os.path.exists(filename):
        wb = Workbook()
        ws = wb.active
        ws.append(['№', 'Дата', 'Источник', 'ФИО', 'Пол', 'Возраст', 'Населенный пункт', 'Контакт', 'Тематика', 'Вопрос', 'Ссылки на файлы'])
        wb.save(filename)
    wb = load_workbook(filename)
    ws = wb.active
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row_number = ws.max_row
    ws.append([
        row_number, now, data.get('source', 'VK'), data.get('fio', ''), data.get('gender', ''),
        data.get('age', ''), data.get('city', ''), data.get('contact', ''), data.get('topic', ''), 
        data.get('question', ''), data.get('file_links', '')
    ])
    wb.save(filename)
    return row_number

def send_to_max_channel(data, msg_number):
    if not MAX_TOKEN or not MAX_CHAT_ID: return
    files_text = "\n\nПрикрепленные файлы:\n" + data.get('file_links') if data.get('file_links') else ""
    text = (
        f"Новое сообщение ({data.get('source', 'VK')})\n"
        f"№ {msg_number}\n\n"
        f"ФИО: {data.get('fio', '')}\n"
        f"Пол: {data.get('gender', '')}\n"
        f"Возраст: {data.get('age', '')}\n"
        f"Населенный пункт: {data.get('city', '')}\n"
        f"Контакт: {data.get('contact', '')}\n"
        f"Тематика: {data.get('topic', '')}\n\n"
        f"Текст вопроса:\n{data.get('question', '')}"
        f"{files_text}"
    )
    url = "https://botapi.max.ru/messages"
    headers = {"Authorization": MAX_TOKEN, "Content-Type": "application/json"}
    try:
        requests.post(url, headers=headers, params={"chat_id": MAX_CHAT_ID}, json={"text": text}, verify=False, timeout=10)
    except Exception as e:
        print(f"ОШИБКА ОТПРАВКИ В КАНАЛ: {e}")

def send_max_message(user_id, text):
    if not MAX_BOT_TOKEN: return
    url = "https://botapi.max.ru/messages"
    headers = {"Authorization": MAX_BOT_TOKEN, "Content-Type": "application/json"}
    try:
        requests.post(url, headers=headers, params={"user_id": user_id}, json={"text": text}, verify=False, timeout=10)
    except Exception as e:
        print(f"ОШИБКА ОТПРАВКИ В MAX: {e}")

# --- Логика ВК ---
def vk_main():
    if not VK_TOKEN: return
    vk_session = vk_api.VkApi(token=VK_TOKEN)
    vk = vk_session.get_api()
    longpoll = VkLongPoll(vk_session)
    print("ВК бот запущен...")
    for event in longpoll.listen():
        if event.type == VkEventType.MESSAGE_NEW and event.to_me:
            user_id = event.user_id
            text = event.text.strip().lower()
            
            attachments = []
            if event.message_id:
                try:
                    msg_data = vk.messages.getById(message_ids=event.message_id)
                    if msg_data and 'items' in msg_data and len(msg_data['items']) > 0:
                        attachments = msg_data['items'][0].get('attachments', [])
                except Exception as e:
                    print(f"Ошибка ВК: {e}")

            if text == "" and not attachments: 
                continue
            
            if user_id not in vk_states:
                if "новое сообщение" in text:
                    vk_states[user_id] = 0
                    vk_data[user_id] = {'source': 'ВКонтакте', 'file_links': ''}
                    vk.messages.send(user_id=user_id, message="Здравствуйте! Это бот для сбора сообщений.\n" + QUESTIONS[0], random_id=int(time.time()))
                else:
                    vk.messages.send(user_id=user_id, message=FINISHED_TEXT, random_id=int(time.time()))
            else:
                state = vk_states[user_id]
                
                if state == 5:
                    if not text.isdigit() or not (1 <= int(text) <= 10):
                        vk.messages.send(user_id=user_id, message="Введите число от 1 до 10.", random_id=int(time.time()))
                        continue
                    vk_data[user_id]['topic'] = TOPICS[int(text)-1]
                else:
                    keys = ['fio', 'gender', 'age', 'city', 'contact', 'topic', 'question']
                    if state < len(keys):
                        vk_data[user_id][keys[state]] = text
                
                if attachments:
                    folder_name = f"Сообщение_{user_id}_{int(time.time())}"
                    create_yandex_folder(folder_name)
                    links = []
                    for i, att in enumerate(attachments, 1):
                        try:
                            file_bytes, filename, ext = download_vk_attachment(att)
                            if file_bytes:
                                link = upload_to_yandex_disk(file_bytes, filename, folder_name)
                                if link:
                                    links.append(f"Файл {i}: {link}")
                        except Exception as e:
                            print(f"ОШИБКА ВК ФАЙЛА: {e}")
                    
                    if links:
                        vk_data[user_id]['file_links'] = "\n".join(links)
                
                vk_states[user_id] += 1
                
                if vk_states[user_id] > len(QUESTIONS):
                    vk.messages.send(user_id=user_id, message=THANKS_TEXT, random_id=int(time.time()))
                    time.sleep(0.5)
                    vk.messages.send(user_id=user_id, message=FINISHED_TEXT, random_id=int(time.time()) + 1)
                    
                    msg_num = save_to_excel(vk_data[user_id])
                    send_to_max_channel(vk_data[user_id], msg_num)
                    del vk_states[user_id]; del vk_data[user_id]
                elif vk_states[user_id] == len(QUESTIONS):
                    vk.messages.send(user_id=user_id, message=VK_LAST_QUESTION, random_id=int(time.time()))
                else:
                    vk.messages.send(user_id=user_id, message=QUESTIONS[vk_states[user_id]], random_id=int(time.time()))

# --- Flask сервер для MAX ---

app = Flask(__name__)

def process_max_message(user_id, text, attachments):
    key = str(user_id)
    
    if key not in max_states:
        if "новое сообщение" in text:
            max_states[key] = 0
            max_data[key] = {'source': 'MAX', 'file_links': ''}
            send_max_message(user_id, "Добрый день! Это бот для сбора вопросов в рамках проведения прямой линии губернатора Костромской области Сергея Ситникова.\n" + QUESTIONS[0])
        else:
            send_max_message(user_id, FINISHED_TEXT)
    else:
        state = max_states[key]
        
        if state == 5:
            if not text.isdigit() or not (1 <= int(text) <= 10):
                send_max_message(user_id, "Введите число от 1 до 10.")
                return
            max_data[key]['topic'] = TOPICS[int(text)-1]
        else:
            keys = ['fio', 'gender', 'age', 'city', 'contact', 'topic', 'question']
            if state < len(keys):
                max_data[key][keys[state]] = text
        
        if attachments:
            folder_name = f"Сообщение_MAX_{user_id}_{int(time.time())}"
            create_yandex_folder(folder_name)
            links = []
            for i, att in enumerate(attachments, 1):
                try:
                    file_bytes, filename, ext = download_max_attachment(att)
                    if file_bytes:
                        link = upload_to_yandex_disk(file_bytes, filename, folder_name)
                        if link:
                            links.append(f"Файл {i}: {link}")
                except Exception as e:
                    print(f"ОШИБКА MAX ФАЙЛА: {e}")
                    
            if links:
                max_data[key]['file_links'] = "\n".join(links)
        
        max_states[key] += 1
        
        if max_states[key] > len(QUESTIONS):
            send_max_message(user_id, THANKS_TEXT)
            time.sleep(0.5)
            send_max_message(user_id, FINISHED_TEXT)
            
            msg_num = save_to_excel(max_data[key])
            send_to_max_channel(max_data[key], msg_num)
            del max_states[key]; del max_data[key]
        elif max_states[key] == len(QUESTIONS):
            send_max_message(user_id, MAX_LAST_QUESTION)
        else:
            send_max_message(user_id, QUESTIONS[max_states[key]])


@app.route('/webhook', methods=['POST'])
def webhook():
    try:
        update = request.json
        update_type = update.get('update_type')
        
        if update_type == 'bot_started':
            user_id = update.get('user', {}).get('user_id') or update.get('sender', {}).get('user_id')
            if user_id:
                key = str(user_id)
                max_states[key] = 0
                max_data[key] = {'source': 'MAX', 'file_links': ''}
                send_max_message(user_id, "Добрый день! Это бот для сбора вопросов в рамках проведения прямой линии губернатора Костромской области Сергея Ситникова.\n" + QUESTIONS[0])
            return jsonify({"success": True}), 200

        if update_type == 'message_created':
            message = update.get('message', {})
            user_id = message.get('sender', {}).get('user_id')
            text = message.get('body', {}).get('text', '').strip().lower()
            attachments = message.get('body', {}).get('attachments', [])
            
            # ЗАЩИТА ОТ ДУБЛИКАТОВ: получаем ID сообщения
            msg_id = message.get('id') or message.get('message_id')
            
            if msg_id and msg_id in processed_max_messages:
                print(f"Пропущен дубликат сообщения MAX: {msg_id}")
                return jsonify({"success": True}), 200
            
            if msg_id:
                processed_max_messages.add(msg_id)
                # Ограничиваем размер памяти, чтобы не рос бесконечно
                if len(processed_max_messages) > 1000:
                    processed_max_messages.clear()

            if not text and not attachments:
                return jsonify({"success": True}), 200

            threading.Thread(target=process_max_message, args=(user_id, text, attachments)).start()
            
        return jsonify({"success": True}), 200
    except Exception as e:
        print(f"ОШИБКА WEBHOOK: {e}")
        return jsonify({"success": True}), 200

if __name__ == '__main__':
    vk_thread = threading.Thread(target=vk_main)
    vk_thread.daemon = True
    vk_thread.start()
    
    port = int(os.environ.get('PORT', 3000))
    print(f"Flask сервер запущен на порту {port}")
    app.run(host='0.0.0.0', port=port)
