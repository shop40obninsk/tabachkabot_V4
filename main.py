from openpyxl import load_workbook
import telebot
from telebot import types
from os import walk

token="2001889625:AAEZrW1nw5GpI9lRshJX5tYiL-zkDX1g7nI"
#token="1986086924:AAFNbyaH3lHwpIu9H_a_LmOEuqlrIrdKU8M"

delivery=["Москва","Обнинск","РАНХиГС","РЭУ","РНИМУ","Бауманка","ПМГМУ","РГУ им. Косыгина","РУДН"]

base_path="catalog/"
electronic=base_path+"Электронные сигареты/"
giga=base_path+"Жижи/"

table_path="main.xlsx"
table="main"

dan_id=1588645954
#dan_id=808525546
doctors_id=808525546 #степан id

electronic_s_key="Электронные сигареты"
giga_key="Жижа"


def append_in_xlsx(path,table_name,row):
    try:
        wb = load_workbook(path)
        ws = wb[str(table_name)]
        ws.append(row)
        wb.save(path)
        wb.close()
        return True
    except:
        return False

def get_description(path):
    f=open(str(path)+'main.txt',"r",encoding="utf-8")
    array=f.readlines()
    Name=str(array[0]).replace("\n",'')
    Description=str(array[1]).replace("\n",'').replace("&",'\n')
    Taste=str(array[2]).replace("\n","").split(":")
    Price=str(array[3])
    Tastes=''
    for i in Taste:
        Tastes+=i+'  '
    return [Name+"\n"+Description+"\n"+"Цена: "+Price+" ₽", Name,Description,Taste,Price,path+"main.png"]

def get_dir(path):
    f = []
    for (dirpath, dirnames, filenames) in walk(path):
        f.extend(dirnames)
        f.extend(filenames)
        break
    return f

def Keyboard_Generator(buttons):
    keyboard = types.ReplyKeyboardMarkup(resize_keyboard=True)
    for i in buttons:
        button = types.KeyboardButton(text=i)
        keyboard.add(button)
    return keyboard

def Inline_Keyboard_Generator(buttons):
    keyboard = types.InlineKeyboardMarkup(row_width=2)
    for i in buttons:
        button = types.InlineKeyboardButton(i[0],callback_data=i[1])
        keyboard.add(button)
    return keyboard

def command_worker(message,chat_id):
    global electronic
    if message=="/Электронные сигареты":
        mass=[]
        for i in get_dir(electronic):
            mass.append("/" +f"{electronic_s_key} "+str(i))
        mass.append("На главную")
        bot.send_message(chat_id, "Выбери производителя👇", reply_markup=Keyboard_Generator(mass))

    if message=="/Жижи":
        mass=[]
        for i in get_dir(giga):
            mass.append("/"+f"{giga_key} "+str(i))
        mass.append("На главную")
        bot.send_message(chat_id, "Выбери производителя👇", reply_markup=Keyboard_Generator(mass))

    if str(message).replace("/",'').replace(f"{electronic_s_key} ","") in get_dir(electronic) and electronic_s_key in str(message):
        print(1)
        bot.send_message(chat_id, "Выбери товар",reply_markup=Keyboard_Generator(["Назад к производителям эллектронных сигарет"]))
        manufacturer=str(message).replace('/','').replace(f"{electronic_s_key} ","")
        for i in get_dir(electronic+str(message).replace("/",'').replace(f"{electronic_s_key} ","")+"/"):
            path=electronic + str(message).replace("/", '').replace(f"{electronic_s_key} ","") + "/"+str(i)+"/"
            print(path)
            i=get_description(path)
            mass = []
            for t in i[3]:
                mass.append([f"{t}", f"E:{i[1]}:{manufacturer}:{t}"])
            keyboard = Inline_Keyboard_Generator(mass)
            bot.send_photo(chat_id, photo=open(i[5], 'rb'), caption=i[0], reply_markup=keyboard)

    if str(message).replace("/", '').replace(f"{giga_key} ", "") in get_dir(giga) and giga_key in str(message):
        print(2)
        bot.send_message(chat_id, "Выбери товар",
                         reply_markup=Keyboard_Generator(["Назад к производителям Жиж"]))
        manufacturer = str(message).replace('/', '').replace(f"{giga_key} ", "")
        for i in get_dir(giga + str(message).replace("/", '').replace(f"{giga_key} ", "") + "/"):
            path = giga + str(message).replace("/", '').replace(f"{giga_key} ", "") + "/" + str(i) + "/"
            print(path)
            i = get_description(path)
            mass = []
            for t in i[3]:
                mass.append([f"{t}", f"G:{i[1]}:{manufacturer}:{t}"])
            keyboard = Inline_Keyboard_Generator(mass)
            bot.send_photo(chat_id, photo=open(i[5], 'rb'), caption=i[0], reply_markup=keyboard)


bot =telebot.TeleBot(token)

@bot.message_handler(commands=['start'])
def welcome(message):
    bot.send_message(message.chat.id, f"Добро пожаловать, {message.from_user.first_name} ! 🤍"
                                      f"\nТебя приветствует онлайн магазин табачной продукции "
                                      f"Дымовой!🔥\n"
                                      f"Наши соц.сети, где регулярно проходят акции:"
                                      f"\nInstagram: https://instagram.com/dymovoy.dv?utm_medium=copy_link \n"
                                      f"Делай заказ прямо сейчас, в удобном боте! 👇",
                     reply_markup=Keyboard_Generator(["/Электронные сигареты","/Жижи","/Поделиться"]))


@bot.message_handler(content_types=['text'])
def get_message(message):
    chat_id=message.chat.id
    text=message.text
    print(text,"h",message.chat.id, message.from_user.first_name)
    if "/" in text:
        if text=="/admin_get_table" and (chat_id==dan_id or chat_id==doctors_id):
            db=open(table_path,"rb")
            bot.send_document(chat_id,db)

        command_worker(text,chat_id)
        print("/")

    if text=="Назад к производителям эллектронных сигарет":
        mass = []
        for i in get_dir(electronic):
            mass.append("/" +f"{electronic_s_key} "+str(i))
        mass.append("На главную")
        bot.send_message(chat_id, "Выбери производителя👇", reply_markup=Keyboard_Generator(mass))

    if text=="Назад к производителям Жиж":
        mass = []
        for i in get_dir(giga):
            mass.append("/" + f"{giga_key} "+str(i))
        mass.append("На главную")
        bot.send_message(chat_id, "Выбери производителя👇", reply_markup=Keyboard_Generator(mass))

    if text=="На главную":
        bot.send_message(message.chat.id, "Выбери товар",
                         reply_markup=Keyboard_Generator(["/Электронные сигареты","/Жижи","/Поделиться"]))

    if text=="/Поделиться":
        bot.send_photo(message.chat.id,open("TgQR.png", 'rb'))
        bot.send_message(message.chat.id, "Telegram: https://t.me/dymovoy_bot\n"
                                          "Instagram: https://instagram.com/dymovoy.dv?utm_medium=copy_link",
                         reply_markup=Keyboard_Generator(["/Электронные сигареты", "/Жижи","/Поделиться"]))

@bot.callback_query_handler(func=lambda call: True)
def Callback_inline(call):
    if call.message:
        chat_id = call.message.chat.id
        data=str(call.data)
        print(data,"i")
        if data.split(":")[0]=="E":
            mass=[]
            for i in delivery:
                mass.append([i,"D:"+str(data)+":"+i])
            bot.send_message(chat_id,"Выберите доствку:",reply_markup=Inline_Keyboard_Generator(mass))

        if data.split(":")[0]=="G":
            mass=[]
            for i in delivery:
                mass.append([i,"D:"+str(data)+":"+i])
            bot.send_message(chat_id,"Выберите доствку:",reply_markup=Inline_Keyboard_Generator(mass))

        if data.split(":")[0] == "D":
            item=data.split(":")[1]
            name=data.split(":")[2]
            manufacturer=data.split(":")[3]
            taste=data.split(":")[4]
            place = data.split(":")[5]
            bot.edit_message_text(chat_id=call.message.chat.id,
                                  message_id=call.message.message_id,
                                  text=f"Спасибо. Мы приняли ваш заказ:\n"
                                       f"✅ Производитель: {manufacturer}\n"
                                       f"✅ Название: {name}\n"
                                       f"✅ Вкус: {taste}\n"
                                       f"✅ Доставка: {place}\n",
                                  reply_markup=None)
            bot.send_message(chat_id,
                             "Скоро мы с вами свяжемся! ⚡",
                             reply_markup=Keyboard_Generator(["/Электронные сигареты","/Жижи","/Поделиться"]))

            #-----------------------------------------------export------------------------------
            username=call.message.chat.username
            first_name=call.message.chat.first_name
            url="https://t.me/"+username
            mass=(item,manufacturer,name,taste,place,username,first_name,chat_id,url)
            try:
                append_in_xlsx(table_path,table,mass)
                bot.send_message(dan_id,f"Новый заказ:\n{item}\n{manufacturer}\n{name}\n{taste}\n{place}\n{username}\n{first_name}\n{chat_id}\n{url}")
            except:
                bot.send_message(dan_id,f"WRITE ERRR\nНовый заказ:\n{item}\n{manufacturer}\n{name}\n{taste}\n{place}\n{username}\n{first_name}\n{chat_id}\n{url}")



bot.polling(none_stop=True)