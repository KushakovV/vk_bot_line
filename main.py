import os
import vk_api
from vk_api.longpoll import VkLongPoll, VkEventType
from openpyxl import Workbook, load_workbook
from datetime import datetime
import time
import requests
from flask import Flask, request, jsonify
import threading

# --- Переменные окружения ---
VK_TOKEN = os.environ.get('VK_TOKEN')
MAX_TOKEN = os.environ.get('MAX_TOKEN')
MAX_CHAT_ID = os.environ.get('MAX_CHAT_ID')
MAX_BOT_TOKEN = os.environ.get('MAX_BOT_TOKEN')

# --- Данные ---
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

# Состояния
vk_states = {}
vk_data = {}
max_states = {}
max_data = {}

# --- Excel ---
def save_to_excel(data):
    filename = 'zayavki.xlsx'
    if not os.path.exists(filename):
        wb = Workbook()
        ws = wb.active
        ws.append(['№', 'Дата', 'Источник', 'ФИО', 'Пол', 'Возраст', 'Населенный пункт', 'Контакт', 'Тематика', 'Вопрос'])
        wb.save(filename)
    wb = load_workbook(filename)
    ws = wb.active
    row_num = ws.max_row
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.append([
        row_num, now, data.get('source', 'VK'), data.get('fio', ''), data.get('gender', ''),
        data.get('age', ''), data.get('city', ''), data.get('contact', ''), data.get('topic', ''), data.get('question', '')
    ])
    wb.save(filename)

# --- Отправка в канал MAX ---
def send_max_message(user_id, text):
    if not MAX_BOT_TOKEN: return
    url = "https://platform-api2.max.ru/messages"
    # Для личных сообщений используем user_id, а не chat_id
    params = {"access_token": MAX_BOT_TOKEN, "user_id": user_id}
    try:
        requests.post(url, params=params, json={"text": text}, verify=False, timeout=10)
    except Exception as e:
        print(f"Ошибка отправки в MAX: {e}")

# --- Отправка ответа пользователю в MAX ---
def send_max_message(user_id, text):
    if not MAX_BOT_TOKEN: return
    url = "https://platform-api2.max.ru/messages"
    params = {"access_token": MAX_BOT_TOKEN, "user_id": user_id}
    try:
        requests.post(url, params=params, json={"text": text}, verify=False, timeout=10)
    except Exception as e:
        print(f"Ошибка отправки в MAX: {e}")

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
            text = event.text.strip()
            if text == "": continue
            if user_id not in vk_states:
                vk_states[user_id] = 0
                vk_data[user_id] = {'source': 'VK'}
                vk.messages.send(user_id=user_id, message="Здравствуйте! Это бот для сбора обращений.\n" + QUESTIONS[0], random_id=int(time.time()))
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
                vk_states[user_id] += 1
                if vk_states[user_id] >= len(QUESTIONS):
                    vk.messages.send(user_id=user_id, message="Спасибо! Обращение принято.", random_id=int(time.time()))
                    save_to_excel(vk_data[user_id])
                    send_to_max_channel(vk_data[user_id])
                    del vk_states[user_id]; del vk_data[user_id]
                else:
                    vk.messages.send(user_id=user_id, message=QUESTIONS[vk_states[user_id]], random_id=int(time.time()))

# --- Flask сервер для MAX ---
app = Flask(__name__)

@app.route('/webhook', methods=['POST'])
def webhook():
    print("!!! ПОЛУЧЕН ЗАПРОС ОТ MAX !!!")
    try:
        update = request.json
        if update.get('update_type') == 'message_created':
            message = update.get('message', {})
            user_id = message.get('sender', {}).get('user_id')
            text = message.get('body', {}).get('text', '').strip()
            
            if not text or not user_id:
                return jsonify({"success": True}), 200

            key = str(user_id)

            if key not in max_states:
                max_states[key] = 0
                max_data[key] = {'source': 'MAX'}
                send_max_message(user_id, "Здравствуйте! Это бот для сбора обращений.\n" + QUESTIONS[0])
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
                
                max_states[key] += 1
                if max_states[key] >= len(QUESTIONS):
                    send_max_message(user_id, "Спасибо! Обращение принято.")
                    save_to_excel(max_data[key])
                    send_to_max_channel(max_data[key])
                    del max_states[key]; del max_data[key]
                else:
                    send_max_message(user_id, QUESTIONS[max_states[key]])
        return jsonify({"success": True}), 200
    except Exception as e:
        print(f"Ошибка webhook: {e}")
        return jsonify({"success": True}), 200

# --- Запуск ---
if __name__ == '__main__':
    # Запускаем ВК в фоновом потоке
    vk_thread = threading.Thread(target=vk_main)
    vk_thread.daemon = True
    vk_thread.start()
    
    # Запускаем Flask сервер (Bothost сам подставит нужный порт)
    port = int(os.environ.get('PORT', 3000))
    print(f"Flask сервер запущен на порту {port}")
    app.run(host='0.0.0.0', port=port)
