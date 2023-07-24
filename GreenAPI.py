import requests
import json

from colorama import Fore
from whatsapp_api_client_python import API
from Support_function import Config


class GreenAPI:
    """
    Класс для работы с подключением к GreenAPI
    """
    def __init__(self):

        self.__status_green_api = "Не авторизован"

        data_id_api = Config.get_data(Config, title='GreenAPI')

        self.__id_instance = data_id_api[0][1]
        self.__api_token = data_id_api[1][1]

        self.__error = None

        self.__green_api = API.GreenApi(self.__id_instance, self.__api_token)

    def __str__(self):
        return f"""
        Текущие данные
        Id_instance: {self.__id_instance}
        Api_token: {self.__api_token}"""

    def check_authorization(self, id_instance=None, api_token=None) -> bool:
        """
        Проверка введенных данных на наличие авторизации на сайте GreenAPI
        """

        if id_instance is None and api_token is None:
            id_instance = self.__id_instance
            api_token = self.__api_token

        url = f"https://api.green-api.com/waInstance{id_instance}/getStateInstance/{api_token}"

        payload = {}
        headers = {}

        response = requests.request("GET", url, headers=headers, data=payload)

        if response.status_code == 401:
            self.__error = '\nПользователь не найден\n'
            print('\nПользователь не найден\n')
            return False

        if response.status_code == 403:
            self.__error = '\nКод 403\n'
            print('\nКод 403\n')
            return False

        if response.status_code == 200:
            result = json.loads(response.content)

            if result['stateInstance'] == 'notAuthorized':
                self.__error = '\nПользователь найден но не авторизован\n'
                print('\nПользователь найден но не авторизован\n')
                return False

            if result['stateInstance'] == 'starting':
                self.__error = '\nПользователь в статусе starting\nПопробуйте позже или поменяйте Instance\n'
                print('\nПользователь в статусе starting\nПопробуйте позже или поменяйте Instance\n')
                return False

        print(f'{Fore.GREEN}\nАвторизация пройдена\n')

        self.__status_green_api = "Авторизован"
        return True

    def write_green_api_data(self, id_instance, api_token) -> bool:
        """
        Ввод данных для сайта GreenAPI и их последующая запись в конфиг
        """

        self.__id_instance = id_instance
        self.__api_token = api_token

        self.__green_api = API.GreenApi(self.__id_instance, self.__api_token)

        data_dict = {
            'id_instance': self.__id_instance,
            'api_token_instance': self.__api_token
        }

        Config.write_to_config(Config, title='GreenAPI', data_dict=data_dict)

        return True

    @property
    def id_instance(self):
        """
        Возврат id_instance
        """
        return self.__id_instance

    @property
    def api_token(self):
        """
        Возврат Api_token
        """
        return self.__api_token

    @property
    def green_api(self):
        """
        Возврат объекта green_api
        """
        return self.__green_api

    @property
    def status_green_api(self):
        return self.__status_green_api

    @property
    def error(self):
        return self.__error


gr_api = GreenAPI()
