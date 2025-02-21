import PySimpleGUI as sg

colors = (
    ('#8B0000', '#8B4500', '#8B6914', '#008B45', '#104E8B', '#00008B', '#68228B', '#000000', '#7D7D7D'),
    ('#CD0000', '#CD6600', '#CD9B1D', '#00CD00', '#1874CD', '#0000CD', '#9932CC', '#141414', '#919191'),
    ('#EE0000', '#EE7600', '#DAA520', '#00EE00', '#1C86EE', '#0000EE', '#B23AEE', '#1A1A1A', '#A6A6A6'),
    ('#FF0000', '#FF7F24', '#EEB422', '#00FF00', '#1E90FF', '#0000FF', '#BF3EFF', '#262626', '#BABABA'),
    ('#FF0000', '#FF7F00', '#FFD700', '#7FFF00', '#1E90FF', '#2A52BE', '#D15FEE', '#2E2E2E', '#CFCFCF'),
    ('#EE2C2C', '#FF8C00', '#EEEE00', '#C0FF3E', '#00B2EE', '#4169E1', '#E066FF', '#383838', '#E3E3E3'),
    ('#FF3030', '#EE9A00', '#FFFF00', '#ADFF2F', '#00CDCD', '#1F75FE', '#EE7AE9', '#454545', '#F7F7F7'),
    ('#EE3B3B', '#FFA500', '#EEE685', '#BCEE68', '#00EEEE', '#008CF0', '#FF83FA', '#616161', '#FFFFFF'),
    ('#FF4040', '#FFA500', '#FFF68F', '#CAFF70', '#00FFFF', '#1E90FF', '#EEAEEE', '#6B6B6B')
)

layout = []
for row in colors:
    layout_row = []
    for color in row:
        layout_row.append(sg.Radio(text='     ', font=('Any', 17), key=color, background_color=color,
                                   group_id=0, circle_color=color, enable_events=True))
    layout.append(layout_row)
layout[-1].append(sg.Radio(text='     ', font=('Any', 17), key='#012F2F', background_color='#012F2F', group_id=0,
                           enable_events=True, circle_color='#012F2F'))
layout.append([sg.Input(default_text='#012F2F', size=8, text_color='#012F2F', tooltip='Цвет по умолчанию справа снизу',
                        key='-HEX_INPUT-', font='* 16 bold'),
               sg.Button('Ок', key='-HEX_OK-'), sg.Button('Отменить', key='-HEX_CANCEL-')])

window = sg.Window(title='', layout=layout, no_titlebar=True, element_padding=((3, 3), (3, 3)), grab_anywhere=True,
                   use_default_focus=False, background_color='#EEDFCC', finalize=True, margins=(5, 5), border_depth=4,
                   keep_on_top=True, modal=True, element_justification='center')

for key in list(window.key_dict.keys())[:-3]:
    print(key)
    window[key].widget.configure(indicatoron=False)

while True:
    event, values = window.read()
    if event == '-HEX_CANCEL-':
        break
    # if event == '-HEX_OK-':
    if event.startswith('#'):
        window['-HEX_INPUT-'].update(event)
        window['-HEX_INPUT-'].update(text_color=event)
window.close()
