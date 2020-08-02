import vk_api
from vk_api.bot_longpoll import VkBotLongPoll, VkBotEventType
from vk_api.upload import VkUpload as Vkupload
from random import randint
import os.path as os
import re

from Functions import old_format

LOG = 2000000001
INFORMATION_ABOUT_BOT = '''
Приветствуем, бот находиться в стадии разработки, но вы можете воспользоваться командами:
Позвать администратора
'''

class VkSession:
    def __init__(self, database, vkgroup=None, username=None, password=None, token=None, users_filename=None, administrators_filename=None):
        self.__random_id = self.__random_id_generator(-9223372036854775800)
        self.database = database
        self.users = {}
        if users_filename != None:
            self.users_filename = users_filename
            if not os.exists(users_filename):
                open(users_filename, 'w').close()
            with open(users_filename, 'r') as file:
                for line in file:
                    line = line.replace('\r', '').replace('\n', '')
                    words = line.split()
                    email = words[0]
                    self.users[email] = []
                    for account in words[1::]:
                        self.users[email].append(account)
        else:
            self.users_filename = 'users.txt'
            open('users.txt', 'w').close()
        self.administrators = []
        if administrators_filename != None:
            self.administrators_filename = administrators_filename
            if not os.exists(administrators_filename):
                open(administrators_filename, 'w').close()
            with open(administrators_filename, 'r') as file:
                for line in file:
                    line = line.replace('\r', '').replace('\n', '')
                    words = line.split()
                    for word in words[::]:
                        self.administrators.append(word)
        else:
            self.administrators_filename = 'administrators.txt'
            open('administrators.txt', 'w').close()
        try:
            if username == None and password == None:
                if token == None:
                    raise Exception('Custom Error: unexpected authorization')
                else:
                    self.vk = vk_api.VkApi(token=token)
            else:
                self.vk = vk_api.VkApi(login=username, password=password)
        except:
            raise Exception('Custom Error: authorization failed')
        self.upload = Vkupload(self.vk)
        if vkgroup == None:
            raise Exception('Custom Error: Failed access into longpoll')
        else:
            self.longpoll = VkBotLongPoll(self.vk, vkgroup)

    def __users_get(self, vk_address):
        attachment = {}
        attachment['user_ids'] = vk_address.replace('https://vk.com/', '')
        answer = self.vk.method('users.get', attachment)
        if len(answer):
            return answer[0]['id']
        else:
            raise Exception('Custom Error: non exist user')

    def __random_id_generator(self, id):
        while id < 9223372036854775800:
            id += 1
            yield id

    def __messages_send(self, **data):
        attachment = {}
        for key in data:
            attachment[key] = data[key]
        try:
            random_id = next(self.__random_id)
            attachment['random_id'] = random_id
        except:
            raise Exception('Custom Error: Random ids was ended, session needs in reload')
        print('Answer: ', self.vk.method('messages.send', attachment))

    def __online_log(self, *elements):
        text = ' '
        text = text.join(map(str, elements))
        self.__messages_send(message=text, peer_id=LOG)

    def __delete_user(self, message, from_id):
        user = message[20::].strip().split()
        if not(re.match(r'https://vk\.com/.+', user)):
            self.__messages_send(message='Неправильно введенные данные: "'+user+'"', peer_id=from_id)
        id = self.__users_get(user)
        self.delete_user(id)
        self.__online_log('Удален пользователь: user - "'+str(id)+'"')
    
    def __delete_administrator(self, message, from_id):
        user = message[22::].strip().split()
        if not(re.match(r'https://vk\.com/.+', user)):
            self.__messages_send(message='Неправильно введенные данные: "'+user+'"', peer_id=from_id)
        id = self.__users_get(user)
        self.delete_administrator(id)
        self.__online_log('Удален администратор: administrator - "'+str(id)+'"')
        
    def __add_new_user(self, message, from_id):
        new_user, email = message[21::].strip().split()
        if re.match(r'.+@.*\..*', new_user) and re.match(r'https://vk\.com/.+', email):
            new_user, email = email, new_user
        if not(re.match(r'.+@.*\..*', email) and re.match(r'https://vk\.com/.+', new_user)):
            self.__messages_send(message='Неправильно введенные данные: "'+email+'", "'+new_user+'"', peer_id=from_id)
        id = self.__users_get(new_user)
        self.new_user(email,  id)
        self.__online_log('Добавлен пользователь: email - "'+email+'", user - "'+str(id)+'"')

    def __add_new_administrator(self, message, from_id):
        new_administrator = message[23::].strip()
        if not(re.match(r'https://vk\.com/.+', new_administrator)):
            self.__messages_send(message='Неправильно введенные данные: "'+new_administrator+'"', peer_id=from_id)
        id = self.__users_get(new_administrator)
        self.new_administrator(id)
        self.__online_log('Добавлен администратор: user - "'+str(id)+'"')
        
    def __send_report(self, inn, peer_id):
        print('send, file')
        try:
            filename, recomendation = old_format(inn, self.database)
        except Exception as e:
            print(e)
            self.__messages_send(peer_id=peer_id, message='ИНН по Вашему запросу не найден, проверьте правильность введенных данных или напишите Администратору в Whatsapp +7 (926) 862-02-09')
            return
        file = open(filename, 'rb')
        answer = self.upload.document_message(file, title='spravka_yourspec_inn_'+str(inn), tags=None, peer_id=peer_id)
        type_doc = answer['type']
        owner_id = answer['doc']['owner_id']
        media_id = answer['doc']['id']
        doc_url = '{type}{owner_id}_{media_id}'.format(type=type_doc, owner_id=owner_id, media_id=media_id)
        self.__messages_send(peer_id=peer_id, message=recomendation, attachment=doc_url)
        file.close()

    def __call_administrator(self, id):
        self.__online_log('Пользователь https://vk.com/id' + str(id), 'желает поговорить')

    def __show_help(self, peer_id):
        self.__messages_send(message=INFORMATION_ABOUT_BOT, peer_id=peer_id)

    def __user_action(self, email, data):
        print('___________________user_action________________')
        message = data['object']['message']['text']
        from_id = data['object']['message']['from_id']
        print(message.lower())
        if 'данные по инн' in message.lower():
            inn = int(message[13::])
            self.__send_report(inn, from_id)
        elif 'позвать администратора' in message.lower():
            self.__call_administrator(from_id)
            self.__messages_send(message='Ожидайте', peer_id=from_id)
        elif 'помощь' in message.lower():
            self.__show_help(from_id)

    def __administrator_action(self, data):
        print('________________administrator_action___________')
        message = data['object']['message']['text']
        from_id = data['object']['message']['from_id']
        if 'добавить пользователя' in message.lower():
            self.__add_new_user(message, from_id)
        elif 'удалить пользователя' in message.lower():
            self.__delete_user(message, from_id)
        else:
            self.__user_action('god@god.ru', data)
        
    def __god_action(self, data):
        print('__________________god_action____________________')
        message = data['object']['message']['text']
        from_id = data['object']['message']['from_id']
        if 'добавить администратора' in message.lower():
            self.__add_new_administrator(message, from_id)
        elif 'удалить администратора' in message.lower():
            self.__delete_administrator(message, from_id)
        elif 'test' in message.lower():
            self.__online_log('Test passed')
        else:
            self.__administrator_action(data)

    def __new_message(self, data):
        print('___________________new_message__________________')
        print('Message: ', data)
        peer_id = data['object']['message']['peer_id']
        from_id = data['object']['message']['from_id']
        if peer_id == 2000000001:
            self.__god_action(data)
        elif str(from_id) in self.administrators:
            self.__administrator_action(data)
        else:
            for email in self.users:
                if str(from_id) in self.users[email]:
                    self.__user_action(email, data)
                    break

    def delete_user(self, user):
        user = str(user)
        finded = False
        for email in self.users:
            if user in self.users[email]:
                self.users[email].remove(user)
                finded = True
        with open(self.users_filename, 'r') as file:
            lines = [line.replace(' '+user+' ', ' ') for line in file]
        with open(self.users_filename, 'w') as file:
            file.writelines(lines)
        if not finded:
            return

    def delete_administrator(self, user):
        user = str(user)
        if user in self.administrators:
            self.administrators.remove(user)
            with open(self.administrators_filename, 'r') as file:
                lines = [line.replace(' '+user+' ', ' ') for line in file]
            with open(self.administrators_filename, 'w') as file:
                file.writelines(lines)
        else:
            return      

    def new_user(self, email, user):
        email = str(email)
        user = str(user)
        if email in self.users:
            if user in self.users[email]:
                return
            self.users[email].append(user)
            with open(self.users_filename, 'r') as file:
                lines = [line.rstrip() + ' ' + user + ' \n' if email in line else line for line in file]
            with open(self.users_filename, 'w') as file:
                file.writelines(lines)
        else:
            self.users[email] = []
            self.users[email].append(user)
            with open(self.users_filename, 'a') as file:
                file.write(email+' '+user+' \n')
        
    def new_administrator(self, user):
        user = str(user)
        if user in self.administrators:
            return
        else:
            self.administrators.append(user)
            with open(self.administrators_filename, 'a') as file:
                file.write(user+' ')

    def update(self):
        try:
            event = next(self.longpoll.listen())
        except KeyboardInterrupt:
            raise KeyboardInterrupt()
        except:
            raise Exception('Custom Error: cann`t take event')
        data = event.raw
        if (data['type'] == 'message_new'):
            self.__new_message(data)
