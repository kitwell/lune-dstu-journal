# import requests
# from bs4 import BeautifulSoup
#
#
# def parse_track_duration(artist, track):
#     # Формируем URL для запроса к музыкальному сервису
#     url = f'https://www.example.com/search?q={artist} {track}'
#
#     try:
#         # Отправляем GET-запрос к сервису и получаем HTML-код страницы
#         response = requests.get(url)
#         response.raise_for_status()  # Проверяем статус код ответа
#
#         # Используем BeautifulSoup для парсинга HTML
#         soup = BeautifulSoup(response.text, 'html.parser')
#
#         # Находим элемент с информацией о треке, предположим, что это тег <span> с классом "duration"
#         duration_tag = soup.find('span', class_='duration')
#
#         # Получаем текст из тега, предположим, что это строка в формате "ММ:СС"
#         duration_str = duration_tag.text.strip()
#
#         # Разбиваем строку на минуты и секунды, конвертируем в секунды
#         minutes, seconds = map(int, duration_str.split(':'))
#         duration_seconds = minutes * 60 + seconds
#
#         return duration_seconds
#
#     except requests.RequestException as e:
#         print(f"Ошибка при выполнении запроса: {e}")
#         return None
#
#
# # Пример использования
# artist = "Бузова Ольга"
# track = "Мало половин"
# duration = parse_track_duration(artist, track)
# if duration is not None:
#     print(f"Длительность трека '{track}' исполнителя '{artist}': {duration} секунд")
# else:
#     print("Не удалось получить длительность трека.")

a = [1, 2, 2, 2, 2, 3, 4, 4, 4, 5, 5]
a.reverse()
print(len(a) - a.index(4) - 1)
