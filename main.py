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
from yndx_disk.clients import DiskClient  # Библиотека для Яндекс Диска

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

VK_TOKEN = os.environ.get('VK_TOKEN')
MAX_TOKEN = os.environ.get('MAX_TOKEN')
MAX_CHAT_ID = os.environ.get('MAX_CHAT_ID')
MAX_BOT_TOKEN = os.environ.get('MAX_BOT_TOKEN')
YANDEX_TOKEN = os.environ.get('YANDEX_TOKEN')

# Путь к папке на Яндекс Диске (создадим при первом обращении)
YANDEX_BASE_FOLDER = "/Материалы к сообщениям жителей Костромской области"

TOPICS = [
    "ЖКХ и коммунальные услуги", "Дороги и транспорт", "Здравоохранение",
    "Образование", "Социальная поддержка и льготы", "Благоустройство и экология",
    "Работа и занятость", "Сельское хозяйство", "Строительство и земельные вопросы", "Другое"
]
QUESTIONS = [
    "Введите ваше ФИО:", "Укажите ваш пол (М/Ж):", "Укажите ваш возраст (число):",
    "Укажите ваш населенный пункт:", "Укажите ваш телефон или почту для связи:",
    "Выберите тематику (напишите номер от 1 до 10):\n" + "\n".join([f"{i+1}. {t}" for i, t in enumerate(TOPICS)]),
    "Напишите ваш вопрос:"
]

vk_states = {}
vk_data = {}
max_states = {}
max_data = {}

# --- Яндекс Диск ---
def upload_to_yandex_disk(file_bytes, filename, folder_name):
    """Загружает файл на Яндекс Диск и возвращает публичную ссылку."""
    if not YANDEX_TOKEN:
        print("Ошибка: не задан YANDEX_TOKEN")
        return None
    
    try:
        client = DiskClient(token=YANDEX_TOKEN)
        folder_path = f"{YANDEX_BASE_FOLDER}/{folder_name}"
        
        # Создаём папку, если её нет
        try:
            client.mkdir(folder_path)
        except Exception:
            pass  # Папка уже существует — это нормально
        
        # Сохраняем файл локально (временная папка на Bothost)
        local_path = f"/tmp/{filename}"
        with open(local_path, "wb") as f:
            f.write(file_bytes)
        
        # Загружаем на Диск
        remote_path = f"{folder_path}/{filename}"
        client.upload_file(local_path, remote_path, overwrite=True)
        
        # Получаем публичную ссылку
        public_url = client.get_public_url(remote_path)
        
        # Удаляем временный файл
        os.remove(local_path)
        
        return public_url
    except Exception as e:
        print(f"Ошибка загрузки на Яндекс Диск: {e}")
        return None

def download_vk_attachment(attachment):
    """Скачивает вложение из ВК по URL и возвращает (bytes, filename, ext)."""
    try:
        url = None
        filename = "attachment"
        ext = "dat"
        
        if attachment['type'] == 'photo':
            # Берём максимальный размер
            sizes = attachment['photo']['sizes']
            max_size = max(sizes, key=lambda x: x.get('width', 0) * x.get('height', 0))
            url = max_size['url']
            ext = "jpg"
            filename = f"photo_{attachment['photo']['id']}"
        elif attachment['type'] == 'doc':
            url = attachment['doc']['url']
            ext = attachment['doc'].get('ext', 'dat')
            filename = f"doc_{attachment['doc']['id']}"
        elif attachment['type'] == 'video':
            # Для видео в ВК прямая ссылка недоступна без авторизации,
            # поэтому просто сохраним информацию
            return None, None, None
        
        if not url:
            return None, None, None
        
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        
        return resp.content, f"{filename}.{ext}", ext
    except Exception as e:
        print(f"Ошибка скачивания вложения ВК: {e}")
        return None, None, None

def download_max_attachment(attachment):
    """Скачивает вложение из MAX по URL."""
    try:
        # MAX присылает вложения с полем url или payload
        url = attachment.get('url') or attachment.get('payload', {}).get('url')
        if not url:
            return None, None, None
        
        filename = attachment.get('filename', f"max_{int(time.time())}")
        ext = filename.split('.')[-1] if '.' in filename else 'dat'
        
        resp = requests.get(url, timeout=30, verify=False)
        resp.raise_for_status()
        
        return resp.content, filename, ext
    except Exception as e:
        print(f"Ошибка скачивания вложения MAX: {e}")
        return None, None, None

def save_to_excel(data):
    filename = 'zayavki.xlsx'
    if not os.path.exists(filename):
        wb = Workbook()
        ws = wb.active
        ws.append(['№', 'Дата', 'Источник', 'ФИО', 'Пол', 'Возраст', 'Населенный пункт', 'Контакт', 'Тематика', 'Вопрос', 'Ссылки на файлы'])
        wb.save(filename)
    wb = load_workbook(filename)
    ws = wb.active
    row_num = ws.max_row
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([
        row_num, now, data.get('source', 'VK'), data.get('fio', ''), data.get('gender', ''),
        data.get('age', ''), data.get('city', ''), data.get('contact', ''), data.get('topic', ''), 
        data.get('question', ''), data.get('file_links', '')
    ])
    wb.save(filename)

def send_to_max_channel(data):
    if not MAX_TOKEN or not MAX_CHAT_ID: return
    
    files_text = ""
    if data.get('file_links'):
        files_text = f"\n\n📎 **Файлы:**\n{data.get('file_links')}"
    
    text = (
        f"📋 **Новое сообщение ({data.get('source', 'VK')})**\n\n"
        f"👤 **ФИО:** {data.get('fio', '')}\n"
        f"⚧ **Пол:** {data.get('gender', '')}\n"
        f"🎂 **Возраст:** {data.get('age', '')}\n"
        f"📍 **Населенный пункт:** {data.get('city', '')}\n"
        f"📞 **Контакт:** {data.get('contact', '')}\n"
        f"📂 **Тематика:** {data.get('topic', '')}\n\n"
        f"💬 **Текст вопроса:**\n{data.get('question', '')}"
        f"{files_text}"
    )
    url = "https://botapi.max.ru/messages"
    headers = {"Authorization": MAX_TOKEN, "Content-Type": "application/json"}
    params = {"chat_id": MAX_CHAT_ID}
    try:
        r = requests.post(url, headers=headers, params=params, json={"text": text}, verify=False, timeout=10)
        if r.status_code != 200:
            print(f"ОШИБКА ОТПРАВКИ В КАНАЛ: {r.status_code} - {r.text}")
    except Exception as e:
        print(f"Ошибка отправки в канал: {e}")

def send_max_message(user_id, text):
    if not MAX_BOT_TOKEN: return
    url = "https://botapi.max.ru/messages"
    headers = {"Authorization": MAX_BOT_TOKEN, "Content-Type": "application/json"}
    params = {"user_id": user_id}
    try:
        r = requests.post(url, headers=headers, params=params, json={"text": text}, verify=False, timeout=10)
        if r.status_code != 200:
            print(f"ОШИБКА ОТПРАВКИ В MAX: {r.status_code} - {r.text}")
    except Exception as e:
        print(f"Ошибка отправки в MAX: {e}")

def vk_main():
    if not VK_TOKEN: return
    vk_session = vk_api.VkApi(token=VK_TOKEN)
    vk = vk_session.get_api()
    longpoll = VkLongPoll(vk_session)
    print("ВК бот запущен...")
    for event in longpoll.listen():
        if event.type == VkEventType.MESSAGE_NEW and event.to_me:
            user_id = event.user_id
            text = event.text.strip()
            
            # Обработка вложений
            attachments = event.attachments
            file_links = []
            
            if attachments:
                for att in attachments:
                    file_bytes, filename, ext = download_vk_attachment(att)
                    if file_bytes:
                        # Создаём уникальное имя папки для сообщения
                        folder_name = f"Сообщение_{user_id}_{int(time.time())}"
                        public_url = upload_to_yandex_disk(file_bytes, filename, folder_name)
                        if public_url:
                            file_links.append(f"- [{filename}]({public_url})")
            
            if text == "" and not file_links: 
                continue
            
            if user_id not in vk_states:
                vk_states[user_id] = 0
                vk_data[user_id] = {'source': 'VK', 'file_links': ''}
                vk.messages.send(user_id=user_id, message="Здравствуйте! Это бот для сбора сообщений.\n" + QUESTIONS[0], random_id=int(time.time()))
            else:
                state = vk_states[user_id]
                if state == 5:
                    if not text.isdigit() or not (1 <= int(text) <= 10):
                        vk.messages.send(user_id=user_id, message="Введите число от 1 до 10.", random_id=int(time.time()))
                        continue
                    vk_data[user_id]['topic'] = TOPICS[int(text)-1]
                else:
                    keys = ['fio', 'gender', 'age', 'city', 'contact', 'topic', 'question']
                    vk_data[user_id][keys[state]] = text
                
                # Если есть файлы, сохраняем ссылки
                if file_links:
                    vk_data[user_id]['file_links'] = "\n".join(file_links)
                
                vk_states[user_id] += 1
                if vk_states[user_id] >= len(QUESTIONS):
                    vk.messages.send(user_id=user_id, message="Спасибо! Сообщение принято.", random_id=int(time.time()))
                    save_to_excel(vk_data[user_id])
                    send_to_max_channel(vk_data[user_id])
                    del vk_states[user_id]; del vk_data[user_id]
                else:
                    vk.messages.send(user_id=user_id, message=QUESTIONS[vk_states[user_id]], random_id=int(time.time()))

app = Flask(__name__)

@app.route('/webhook', methods=['POST'])
def webhook():
    try:
        update = request.json
        update_type = update.get('update_type')
        
        if update_type == 'bot_started':
            user_id = update.get('user', {}).get('user_id') or update.get('sender', {}).get('user_id')
            if user_id:
                key = str(user_id)
                if key in max_states: del max_states[key]
                if key in max_data: del max_data[key]
                    
                max_states[key] = 0
                max_data[key] = {'source': 'MAX', 'file_links': ''}
                send_max_message(user_id, "Добрый день! Это бот для сбора сообщений в рамках проведения прямой линии губернатора Костромской области.\n" + QUESTIONS[0])
            return jsonify({"success": True}), 200

        if update_type == 'message_created':
            message = update.get('message', {})
            user_id = message.get('sender', {}).get('user_id')
            text = message.get('body', {}).get('text', '').strip()
            
            # Обработка вложений MAX
            attachments = message.get('body', {}).get('attachments', [])
            file_links = []
            
            if attachments:
                for att in attachments:
                    file_bytes, filename, ext = download_max_attachment(att)
                    if file_bytes:
                        folder_name = f"Сообщение_MAX_{user_id}_{int(time.time())}"
                        public_url = upload_to_yandex_disk(file_bytes, filename, folder_name)
                        if public_url:
                            file_links.append(f"- [{filename}]({public_url})")
            
            if not text and not file_links:
                return jsonify({"success": True}), 200

            key = str(user_id)

            if key not in max_states:
                max_states[key] = 0
                max_data[key] = {'source': 'MAX', 'file_links': ''}
                send_max_message(user_id, "Добрый день! Это бот для сбора сообщений в рамках проведения прямой линии губернатора Костромской области.\n" + QUESTIONS[0])
            else:
                state = max_states[key]
                if state == 5:
                    if not text.isdigit() or not (1 <= int(text) <= 10):
                        send_max_message(user_id, "Введите число от 1 до 10.")
                        return jsonify({"success": True}), 200
                    max_data[key]['topic'] = TOPICS[int(text)-1]
                else:
                    keys = ['fio', 'gender', 'age', 'city', 'contact', 'topic', 'question']
                    max_data[key][keys[state]] = text
                
                # Если есть файлы, сохраняем ссылки
                if file_links:
                    max_data[key]['file_links'] = "\n".join(file_links)
                
                max_states[key] += 1
                if max_states[key] >= len(QUESTIONS):
                    send_max_message(user_id, "Спасибо! Сообщение принято.")
                    save_to_excel(max_data[key])
                    send_to_max_channel(max_data[key])
                    del max_states[key]; del max_data[key]
                else:
                    send_max_message(user_id, QUESTIONS[max_states[key]])
        return jsonify({"success": True}), 200
    except Exception as e:
        print(f"Ошибка webhook: {e}")
        return jsonify({"success": True}), 200

if __name__ == '__main__':
    vk_thread = threading.Thread(target=vk_main)
    vk_thread.daemon = True
    vk_thread.start()
    
    port = int(os.environ.get('PORT', 3000))
    print(f"Flask сервер запущен на порту {port}")
    app.run(host='0.0.0.0', port=port)
