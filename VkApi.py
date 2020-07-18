import vk_api
from vk_api.bot_longpoll import VkBotLongPoll, VkBotEventType
from random import randint

class VkSession:
    def __init__(self, vkgroup=None, username=None, password=None, token=None):
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
        if vkgroup == None:
            raise Exception('Custom Error: Failed access into longpoll')
        else:
            self.longpoll = VkBotLongPoll(self.vk, vkgroup)
        
    def __new_message(self, data):
        message = 'Здравствуйте'
        peer_id = data['object']['message']['from_id']
        try:
            random_id = next(self.__random_id_gen())
        except:
            raise Exception('Custom Error: Random ids was ended, session needs to be upload')
        self.vk.method('messages.send', {'peer_id': peer_id, 'random_id': random_id, 'message': message})
    
    def __random_id_gen(self):
        id = -9223372036854775807
        while id < 9223372036854775808:
            id += 1
            yield id

    def update(self):
        event = next(self.longpoll.listen())
        data = event.raw
        if (data['type'] == 'message_new'):
            self.__new_message(data)

