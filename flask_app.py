from tokens import VK_API_TOKEN, VK_GROUP_ID
from VkApi import VkSession

vk_session = VkSession(token=VK_API_TOKEN, vkgroup=VK_GROUP_ID)
print('Бот запущен')

while True:
    vk_session.update()