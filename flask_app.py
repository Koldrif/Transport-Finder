from tokens import VK_API_TOKEN, VK_GROUP_ID
import vk_api
from vk_api.bot_longpoll import VkBotLongPoll, VkBotEventType
from random import randint
from VkApi import VkSession

vk_session = VkSession(token=VK_API_TOKEN, vkgroup=VK_GROUP_ID)

print('Бот запущен')

while True:
    vk_session.update()