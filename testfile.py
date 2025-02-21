# import PySimpleGUI as sg
#
# def open_new_window():
#     layout = [[sg.Button('Close', key='-CLOSE-')]]
#     window = sg.Window('Window', layout)
#
#     while True:
#         event, values = window.read()
#         if event == sg.WINDOW_CLOSED or event == '-CLOSE-':
#             window.close()
#             break
#
#
# previous_window = None
#
# while True:
#     if previous_window is not None:
#         previous_window.close()
#
#     open_new_window()
#     previous_window = sg.Window('Window')

# import PySimpleGUI as sg
#
# layout = [
#     [sg.Text('Counter: 0', key='-COUNTER-')],
#     [sg.Button('Increment')]
# ]
#
# window = sg.Window('Refresh Example', layout)
#
# counter = 0
#
# while True:
#     event, values = window.read()
#     if event == sg.WINDOW_CLOSED:
#         break
#     elif event == 'Increment':
#         counter += 1
#         window['-COUNTER-'].update(f'Counter: {counter}')
#         # window.refresh()  # Принудительно обновляем окно
#
# window.close()


# import PySimpleGUI as sg
# import os
# from fontTools.ttLib import TTFont
# import shutil
# import platform
#
# def install_font(font_path):
#     # Определяем операционную систему
#     os_type = platform.system()
#
#     if os_type == "Windows":
#         # Путь к директории шрифтов в Windows
#         fonts_dir = os.path.join(os.environ['WINDIR'], 'Fonts')
#     elif os_type == "Darwin":
#         # Путь к директории шрифтов в macOS
#         fonts_dir = os.path.expanduser('~/Library/Fonts')
#     elif os_type == "Linux":
#         # Путь к директории шрифтов в Linux
#         fonts_dir = os.path.expanduser('~/.local/share/fonts')
#     else:
#         print("Неподдерживаемая операционная система")
#         return
#
#     # Копируем шрифт в директорию шрифтов
#     try:
#         shutil.copy(font_path, fonts_dir)
#         print(f"Шрифт успешно установлен в {fonts_dir}")
#     except Exception as e:
#         print(f"Ошибка при установке шрифта: {e}")
#
# # Определяем путь к шрифту относительно текущего скрипта
# font_path = os.path.join(os.path.dirname(__file__), 'fonts', 'beer-money12.ttf')
# # font = TTFont(font_path)
# # font.save(font_path)
# # Пример использования
# install_font(font_path)
#
#
# # Определение макета
# layout = [
#     [sg.Text('Это текст с кастомным шрифтом', font=("zhizn", 20))],
#     [sg.Button('Выход')]
# ]
#
# # Создание окна
# window = sg.Window('Пример кастомного шрифта', layout)
#
# while True:
#     event, values = window.read()
#     if event == sg.WIN_CLOSED or event == 'Выход':
#         break
#
# # Закрытие окна
# window.close()

print(eval("list('dsd')"))
