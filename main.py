import os
import vk_api
from vk_api.longpoll import VkLongPoll, VkEventType
from openpyxl import Workbook, load_workbook
from datetime import datetime
import time

# Получаем токен из переменных окружения
VK_TOKEN = os.environ.get('VK_TOKEN')

TOPICS = [
    "ЖКХ и коммунальные услуги",
    "Дороги и транспорт",
    "Здравоохранение",
    "Образование",
    "Социальная поддержка и льготы",
    "Благоустройство и экология",
    "Работа и занятость",
    "Сельское хозяйство",
    "Строительство и земельные вопросы",
    "Другое"
]

QUESTIONS = [
    "Введите ваше ФИО:",
    "Укажите ваш пол (М/Ж):",
    "Укажите ваш возраст (число):",
    "Укажите ваш населенный пункт:",
    "Укажите ваш телефон или почту для связи:",
    "Выберите тематику (напишите номер от 1 до 10):\n" + "\n".join([f"{i+1}. {t}" for i, t in enumerate(TOPICS)]),
    "Напишите ваш вопрос:"
]

user_states = {}
user_data = {}

def save_to_excel(data):
    filename = 'zayavki.xlsx'
    if not os.path.exists(filename):
        wb = Workbook()
        ws = wb.active
        ws.append(['№', 'Дата', 'ФИО', 'Пол', 'Возраст', 'Населенный пункт', 'Контакт', 'Тематика', 'Вопрос'])
        wb.save(filename)

    wb = load_workbook(filename)
    ws = wb.active
    row_num = ws.max_row
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    ws.append([
        row_num, now, data.get('fio', ''), data.get('gender', ''), 
        data.get('age', ''), data.get('city', ''), data.get('contact', ''), 
        data.get('topic', ''), data.get('question', '')
    ])
    wb.save(filename)

def main():
    if not VK_TOKEN:
        print("Ошибка: не задан VK_TOKEN")
        return

    vk_session = vk_api.VkApi(token=VK_TOKEN)
    vk = vk_session.get_api()
    longpoll = VkLongPoll(vk_session)

    print("Бот запущен и слушает сообщения...")

    for event in longpoll.listen():
        if event.type == VkEventType.MESSAGE_NEW and event.to_me:
            user_id = event.user_id
            text = event.text.strip()

            # Игнорируем пустые сообщения (например, нажатие кнопки "Начать")
            if text == "":
                if user_id not in user_states:
                    user_states[user_id] = 0
                    user_data[user_id] = {}
                    vk.messages.send(
                        user_id=user_id,
                        message="Здравствуйте! Это бот для сбора обращений.\n" + QUESTIONS[0],
                        random_id=int(time.time())
                    )
                continue

            if user_id not in user_states:
                user_states[user_id] = 0
                user_data[user_id] = {}
                vk.messages.send(
                    user_id=user_id,
                    message="Здравствуйте! Это бот для сбора обращений.\n" + QUESTIONS[0],
                    random_id=int(time.time())
                )
            else:
                state = user_states[user_id]
                
                if state == 5:
                    if not text.isdigit() or int(text) < 1 or int(text) > 10:
                        vk.messages.send(
                            user_id=user_id,
                            message="Пожалуйста, введите число от 1 до 10.",
                            random_id=int(time.time())
                        )
                        continue
                    user_data[user_id]['topic'] = TOPICS[int(text)-1]
                else:
                    keys = ['fio', 'gender', 'age', 'city', 'contact', 'topic', 'question']
                    user_data[user_id][keys[state]] = text

                user_states[user_id] += 1

                if user_states[user_id] >= len(QUESTIONS):
                    vk.messages.send(
                        user_id=user_id,
                        message="Спасибо! Ваше обращение принято.",
                        random_id=int(time.time())
                    )
                    save_to_excel(user_data[user_id])
                    del user_states[user_id]
                    del user_data[user_id]
                else:
                    vk.messages.send(
                        user_id=user_id,
                        message=QUESTIONS[user_states[user_id]],
                        random_id=int(time.time())
                    )

if __name__ == '__main__':
    main()
