import openpyxl
import random
import datetime
import telebot
from telebot import types, TeleBot
from telebot.types import InputFile
from PIL import Image, ImageDraw, ImageFont

bot: TeleBot = telebot.TeleBot('ключ бота')

workbook = openpyxl.open('main.xlsx')

sheep = workbook.worksheets[0]

def random_emoticon():
    emoticon_arr = ['😀', '👏🏻', '😂', '😭', '👍', '😘', '🤝', '🤣', '🤬', '❤️', '😍', '😊', '😁', '☺️', '😔', '🙈', '😄', '😒', '😜', '😎']
    emoticon = random.choice(emoticon_arr)
    return emoticon


def adding_background_schedule(message):
    image = Image.open("img1.png")

    font = ImageFont.truetype("arial.ttf", 42)
    drawer = ImageDraw.Draw(image)
    drawer.text((25, 200), f'{message}', font=font, fill='white')

    image.save('new_img.png')


def show_date_and_weekday():
    weekdays = {
        0: 'Понедельник',
        1: 'Вторник',
        2: 'Среда',
        3: 'Четверг',
        4: 'Пятница',
        5: 'Суббота',
        6: 'Воскресенье'
    }

    current_date = datetime.date.today()
    weekday = current_date.weekday()

    return weekdays[weekday]


def preparing_message(first_name, day_of_the_week):
    schedule_for_today = f'{first_name} cегодня {datetime.date.today().strftime("%d.%m.%y")} {show_date_and_weekday()}\n'

    cells = sheep[day_of_the_week[show_date_and_weekday()]]
    i = 2

    for article, armorer in cells:
        if armorer.value is None and article.value is not None:
            if i == 2:
                schedule_for_today = f'{schedule_for_today}\n {article.value}\n'
        elif article.value is None:
            pass
        else:
            if i == 2:
                schedule_for_today = f'{schedule_for_today}\n {article.value} |[{armorer.value}]|'
                i = 1
            else:
                i = 2

    return schedule_for_today


def n12131(first_name):
    day_of_the_week = {
        'Понедельник': 'G7:H19',
        'Вторник': 'G20:H31',
        'Среда': 'G32:H43',
        'Четверг': 'G44:H55',
        'Пятница': 'G56:H67',
        'Суббота': 'G7 :H19',
        'Воскресенье': 'G7:H91'
    }
    return preparing_message(first_name, day_of_the_week)


def n12132(first_name):
    day_of_the_week = {
        'Понедельник': 'I7:J19',
        'Вторник': 'I20:J31',
        'Среда': 'I32:J43',
        'Четверг': 'I44:J55',
        'Пятница': 'I56:J67',
        'Суббота': 'I80:J91',
        'Воскресенье': 'I7:J91'
    }
    return preparing_message(first_name, day_of_the_week)


@bot.message_handler(commands=['start'])
def bot_start(message):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    btn_1 = types.KeyboardButton('И12131')
    btn_2 = types.KeyboardButton('И12132')
    btn_3 = types.KeyboardButton('нет моей группы')
    btn_4 = types.KeyboardButton('порекомендуй что нибудь')
    markup.add(btn_1, btn_2, btn_3, btn_4)
    bot.send_message(message.chat.id, f'Выбери свою группу'.format(message.from_user), reply_markup=markup)


@bot.message_handler(content_types=['text'])
def bott(message):
    first_name = message.from_user.first_name
    message_text =message.text
    match message_text:
        case 'И12131':
            adding_background_schedule(n12131(first_name))
            photo = InputFile("new_img.png")
            bot.send_photo(chat_id=message.chat.id, photo=photo, caption=random_emoticon())
        case 'И12132':
            adding_background_schedule(n12132(first_name))
            photo = InputFile("new_img.png")
            bot.send_photo(chat_id=message.chat.id, photo=photo, caption=random_emoticon())
        case 'нет моей группы':
            bot.reply_to(message, text=f'Напиши ему <a href="https://t.me/Xbasan">Хасан</a> он добавит', parse_mode='html')
        case 'порекомендуй что нибудь':
            bot.reply_to(message, text='<a href="https://jut.su/oneepiece/film-1.html"><em>рекомендую</em></a>', parse_mode='html')
        case _:
            bot.reply_to(message, text=f'что')


if __name__ == '__main__':
    print('-|-')
    bot.polling(none_stop=True, interval=0)
