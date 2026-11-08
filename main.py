import os
import asyncio
from aiohttp import web
import vk_api
from vk_api.longpoll import VkLongPoll, VkEventType
from openpyxl import Workbook, load_workbook
from datetime import datetime
import requests
import json

# --- Переменные окружения ---
VK_TOKEN = os.environ.get('VK_TOKEN')
MAX_TOKEN = os.environ.get('MAX_TOKEN')
MAX_CHAT_ID = os.environ.get('MAX_CHAT_ID')
MAX_BOT_TOKEN = os.environ.get('MAX_BOT_TOKEN')
PORT = int(os.environ.get('PORT', 3000))

# --- Данные опроса ---
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

# Состояния для ВК (в памяти)
vk_states = {}
vk_data = {}

# Состояния для MAX (в памяти)
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
def send_to_max_channel(data):
    if not MAX_TOKEN or not MAX_CHAT_ID:
        print("Ошибка: не заданы MAX_TOKEN или MAX_CHAT_ID для отправки в канал")
        return
    text = (
        f"📋 **Новая заявка ({data.get('source', 'VK')})**\n\n"
        f"👤 **ФИО:** {data.get('fio', '')}\n"
        f"⚧ **Пол:** {data.get('gender', '')}\n"
        f"🎂 **Возраст:** {data.get('age', '')}\n"
        f"📍 **Населенный пункт:** {data.get('city', '')}\n"
        f"📞 **Контакт:** {data.get('contact', '')}\n"
        f"📂 **Тематика:** {data.get('topic', '')}\n\n"
        f"💬 **Текст вопроса:**\n{data.get('question', '')}"
    )
    url = "https://platform-api2.max.ru/messages"
    params = {"access_token": MAX_TOKEN, "chat_id": MAX_CHAT_ID}
    try:
        resp = requests.post(url, params=params, json={"text": text})
        print(f"Отправка в канал MAX: {resp.status_code}")
    except Exception as e:
        print(f"Ошибка отправки в MAX: {e}")

# --- Логика для ВК ---
def vk_main():
    if not VK_TOKEN:
        print("Нет VK_TOKEN")
        return
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
                vk.messages.send(user_id=user_id, message="Здравствуйте! Это бот для сбора обращений.\n" + QUESTIONS[0], random_id=int(datetime.now().timestamp()*1000))
            else:
                state = vk_states[user_id]
                if state == 5:
                    if not text.isdigit() or not (1 <= int(text) <= 10):
                        vk.messages.send(user_id=user_id, message="Введите число от 1 до 10.", random_id=int(datetime.now().timestamp()*1000))
                        continue
                    vk_data[user_id]['topic'] = TOPICS[int(text)-1]
                else:
                    keys = ['fio', 'gender', 'age', 'city', 'contact', 'topic', 'question']
                    vk_data[user_id][keys[state]] = text
                vk_states[user_id] += 1
                if vk_states[user_id] >= len(QUESTIONS):
                    vk.messages.send(user_id=user_id, message="Спасибо! Обращение принято.", random_id=int(datetime.now().timestamp()*1000))
                    save_to_excel(vk_data[user_id])
                    send_to_max_channel(vk_data[user_id])
                    del vk_states[user_id]; del vk_data[user_id]
                else:
                    vk.messages.send(user_id=user_id, message=QUESTIONS[vk_states[user_id]], random_id=int(datetime.now().timestamp()*1000))

# --- Логика для MAX (Webhook) ---
async def max_webhook(request):
    try:
        update = await request.json()
        if update.get('update_type') == 'message_created':
            message = update.get('message', {})
            chat_id = message.get('recipient', {}).get('chat_id')
            user_id = message.get('sender', {}).get('user_id')
            text = message.get('body', {}).get('text', '').strip()
            
            if not text or not chat_id:
                return web.Response(status=200)

            key = str(chat_id)

            if key not in max_states:
                max_states[key] = 0
                max_data[key] = {'source': 'MAX'}
                await send_max_message(chat_id, "Здравствуйте! Это бот для сбора обращений.\n" + QUESTIONS[0])
            else:
                state = max_states[key]
                if state == 5:
                    if not text.isdigit() or not (1 <= int(text) <= 10):
                        await send_max_message(chat_id, "Введите число от 1 до 10.")
                        return web.Response(status=200)
                    max_data[key]['topic'] = TOPICS[int(text)-1]
                else:
                    keys = ['fio', 'gender', 'age', 'city', 'contact', 'topic', 'question']
                    max_data[key][keys[state]] = text
                
                max_states[key] += 1
                if max_states[key] >= len(QUESTIONS):
                    await send_max_message(chat_id, "Спасибо! Обращение принято.")
                    save_to_excel(max_data[key])
                    send_to_max_channel(max_data[key])
                    del max_states[key]; del max_data[key]
                else:
                    await send_max_message(chat_id, QUESTIONS[max_states[key]])
        return web.Response(status=200)
    except Exception as e:
        print(f"Ошибка webhook: {e}")
        return web.Response(status=200)

async def send_max_message(chat_id, text):
    if not MAX_BOT_TOKEN: return
    url = "https://platform-api2.max.ru/messages"
    params = {"access_token": MAX_BOT_TOKEN, "chat_id": chat_id}
    try:
        async with request.app['session'].post(url, params=params, json={"text": text}) as resp:
            pass
    except Exception as e:
        print(f"Ошибка отправки в MAX (webhook): {e}")

# --- Запуск ---
async def start_vk_thread(app):
    loop = asyncio.get_event_loop()
    loop.run_in_executor(None, vk_main)

async def main():
    app = web.Application()
    app['session'] = requests.Session()
    app.router.add_post('/webhook', max_webhook)
    
    asyncio.create_task(start_vk_thread(app))
    
    runner = web.AppRunner(app)
    await runner.setup()
    site = web.TCPSite(runner, '0.0.0.0', PORT)
    await site.start()
    print(f"Сервер запущен на порту {PORT}, ждем MAX webhook...")
    await asyncio.Event().wait()

if __name__ == '__main__':
    asyncio.run(main())
