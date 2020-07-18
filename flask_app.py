from tokens import VK_API_TOKEN, VK_GROUP_ID

import vk_api
from vk_api.bot_longpoll import VkBotLongPoll, VkBotEventType
from random import randint


def write_msg(user_id, message):
    vk.method('messages.send', {'peer_id': user_id, 'message': message, 'random_id': randint(0, 100000)})

# Авторизуемся как сообщество
vk = vk_api.VkApi(token=VK_API_TOKEN)

# Работа с сообщениями
longpoll = VkBotLongPoll(vk, VK_GROUP_ID)
print('Бот запущен')
# Основной цикл
while True:
    event = next(longpoll.listen())
    
    print('__________________')
    message = event.raw
    print(message['type'])
    print()
    try:
        if (message['type'] == 'message_new'):
            peer_id = message['object']['message']['from_id']
            print('Сообщение:', event.raw)
            write_msg(peer_id, 'Приветик')
    except Exception as e:
        print('Error:', e)