import sys
import time
import win32com.client
import PySimpleGUI as sg

from update_dir.update import call_updater, check_version

FINANCE_LABEL_FOR_FOLDER = 'Папка с файлами БИТ.Финанс'
FINANCE_LABEL_FOR_FILE = 'Файл БИТ.Финанс'

BUILD_LABEL_FOR_FOLDER = 'Папка с файлами БИТ.Строительство'
BUILD_LABEL_FOR_FILE = 'Файл БИТ.Строительство'
KEYS = ['fin_ipt', 'build_ipt', 'save_folder', 'is_single_file']
sg.LOOK_AND_FEEL_TABLE['SamoletTheme'] = {
                                        'BACKGROUND': '#007bfb',
                                        'TEXT': '#FFFFFF',
                                        'INPUT': '#FFFFFF',
                                        'TEXT_INPUT': '#000000',
                                        'SCROLL': '#FFFFFF',
                                        'BUTTON': ('#FFFFFF', '#007bfb'),
                                        'PROGRESS': ('#354d73', '#FFFFFF'),
                                        'BORDER': 1, 'SLIDER_DEPTH': 0,
                                        'PROGRESS_DEPTH': 0, }
def start():
    sg.theme('SamoletTheme')
    UPD_FRAME = [sg.Column([[sg.Button('Проверка', key='check_upd'), sg.Text('Нет обновлений', key='not_upd_txt'),
                  sg.Push(),
                  sg.pin(sg.Text('Доступно обновление', justification='center', visible=False, key='upd_txt',
                                 background_color='#007bfb', font='bold')),
                  sg.Push(),
                  sg.pin(sg.Button('Обновить', key='upd_btn', visible=False))],
                 ], size=(420, 50))]
    PANEL = [
        sg.Column([
            [sg.Checkbox('Один файл', background_color='#007bfb', key='is_single_file', enable_events=True, default=False)],
            [sg.Text(FINANCE_LABEL_FOR_FOLDER, visible=True, key='fin_label', background_color='#007bfb', font='bold')],
            [sg.pin(sg.Column([[sg.Input(key='fin_ipt'), sg.FolderBrowse(button_text='Выбрать', key='fin_browse')]],
                              key='--FIN_COL--')),
             sg.pin(sg.Column([[sg.Input(key='fin_ipt_fl'),
                                sg.FileBrowse(button_color='#007bfb', button_text='Выбрать', key='fin_browse_fl'
                                              )]], key='--FIN_COL_FL--', visible=False))],
            [sg.Text(BUILD_LABEL_FOR_FOLDER, key='build_label', font='bold')],
            [sg.pin(sg.Column([[sg.Input(key='build_ipt'), sg.FolderBrowse(button_text='Выбрать', key='build_browse')]], key='--BUILD_COL--')),
             sg.pin(sg.Column([[sg.Input(key='build_ipt_fl'), sg.FileBrowse(button_text='Выбрать', key='build_browse_fl')]],
                              key='--BUILD_COL_FL--', visible=False))],
            [sg.Text('Папка сохранения результатов', font='bold')],
            [sg.Column([[sg.Input(key='save_folder'), sg.FolderBrowse(button_text='Выбрать')]])],
            [sg.OK(button_text='Далее'), sg.Cancel(button_text='Выход')]

        ], key='-FILE_PANEL-', visible=True, size=(420, 300), background_color='#007bfb')]
    layout = [
            [sg.Frame(layout=[UPD_FRAME], title='Обновление', key='--UPD_FRAME--')],
            [sg.Frame(layout=[PANEL], title='Выбор файлов')]]
    yeet = sg.Window('Сверка данных файлов', layout=layout)
    check, upd_check = False, True
    while True:
        event, values = yeet.read(100)
        if check:
            upd_check = check_version()
            check = False
        if event in ('Выход', sg.WIN_CLOSED):
            sys.exit()
        if event == 'check_upd':
            check = True
        if not upd_check:
            yeet['not_upd_txt'].Update(visible=False)
            yeet['upd_txt'].Update(visible=True)
            yeet['upd_btn'].Update(visible=True)
        if event == 'upd_btn':
            yeet.close()
            call_updater('pocket')
        if event == 'is_single_file':
            if values['is_single_file']:
                yeet['fin_label'].Update(FINANCE_LABEL_FOR_FILE)
                yeet['build_label'].Update(BUILD_LABEL_FOR_FILE)
            else:
                yeet['fin_label'].Update(FINANCE_LABEL_FOR_FOLDER)
                yeet['build_label'].Update(BUILD_LABEL_FOR_FOLDER)
            yeet['--BUILD_COL--'].Update(visible=not values['is_single_file'])
            yeet['--BUILD_COL_FL--'].Update(visible=values['is_single_file'])
            yeet['--FIN_COL--'].Update(visible=not values['is_single_file'])
            yeet['--FIN_COL_FL--'].Update(visible=values['is_single_file'])
            yeet.refresh()
            yeet['-FILE_PANEL-'].contents_changed()
        elif event == 'Далее':
            break

    yeet.close()
    check_values = check_user_values(raw_data=values)
    if check_values:
        edit_values = dict()
        for key in KEYS:
            for k, v in values.items():
                if key in k and v != '':
                    edit_values[key] = v
        return edit_values
    else:
        check_input_error = input_error_panel()
        if check_input_error:
            return start()

def check_user_values(raw_data):
    if raw_data['is_single_file']:
        if (raw_data['build_ipt_fl'] == ''  or raw_data['fin_ipt_fl'] == '') \
            or raw_data['build_ipt_fl'] == raw_data['fin_ipt_fl']:
            return False
    else:
        if (raw_data['build_ipt'] == ''  or raw_data['fin_ipt'] == '') \
            or raw_data['build_ipt'] == raw_data['fin_ipt']:
            return False
    if raw_data['save_folder'] == '':
        return False
    return True

def input_error_panel():
    event = sg.popup('Ошибка ввода', 'При вводе данных возникла ошибка.\nВы хотите повторить ввод данных?',
                     background_color='#007bfb', button_color=('white', '#007bfb'),
                     title='Ошибка', custom_text=('Да', 'Нет'))
    if event == 'Да':
        return True
    else:
        sys.exit()

def end(path):
    event = sg.popup('Сверка завершена\nОткрыть обработанный файл?', background_color='#007bfb',
                     button_color=('white', '#007bfb'),
                     title='Завершение работы', custom_text=('Да', 'Нет'))
    if event == 'Да':
        Excel = win32com.client.Dispatch("Excel.Application")
        Excel.Visible = True
        Excel.Workbooks.Open(Filename=path)
        time.sleep(5)
        del Excel
    else:
        sys.exit()

def error():
    sg.popup_auto_close('При выполнении сверки возникла непредвиденная ошибка\nПодробности можно посмотреть в лог файле',
                                background_color='#007bfb',button_color=('white', '#007bfb'),
                                title='Выход с исключением', auto_close_duration = 15)
    sys.exit()
