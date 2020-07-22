from tokens import VK_API_TOKEN, VK_GROUP_ID
from VkApi import VkSession
from DataBase.DataBase import DataBase as Database


def main():
    database = None #Database(host='localhost', user='root', password='6786')
    vk_session = VkSession(database, token=VK_API_TOKEN, vkgroup=VK_GROUP_ID, users_filename='users.txt', administrators_filename='administrators.txt')
    print('Бот запущен')
    while True:
        try:
            vk_session.update()
        except KeyboardInterrupt:
            raise SystemExit
        except Exception as e:
            print(e)

if __name__ == '__main__':
    main()