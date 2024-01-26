

from base64 import b64decode
import os
import sys
import tkinter.font as font
import winreg
from configparser import ConfigParser, RawConfigParser
from ctypes import windll
from tkinter import Menu, Tk, ttk
from tkinter.filedialog import askdirectory
from tkinter.messagebox import showinfo, showwarning
from win32api import SetFileAttributes
from win32con import FILE_ATTRIBUTE_NORMAL, FILE_ATTRIBUTE_HIDDEN
from LT_Dic import language_label, language_list, root_dic
from icon import icon

WIN_SHELL_KEY = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                               r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
DESKTOP_PATH = winreg.QueryValueEx(WIN_SHELL_KEY, "Desktop")[0]
NK_PATH = 'C:/ProgramData/WinFolder-Nicknamer'
NK_BACKUP_DATA = os.path.join(NK_PATH, r'backup_data.ini')
NK_CONFIG = os.path.join(NK_PATH, r'config.ini')
NK_ICON_PATH = os.path.join(NK_PATH, r'Nickname.ico')
NK_OPTION = 'nickname'
NK_ORI_OPTION = 'original'
SHELL_CLASS_INFO = '.ShellClassInfo'
LOCAL_REC_NAME = 'localizedresourcename'
VIEW_STATE = 'ViewState'
MODE = 'Mode'
VID = 'Vid'
FOLDER_TYPE = 'FolderType'
GENERIC = 'Generic'


def get_language_num(language: str):
    lang_dic = {'English': 0, 'Chinese': 1}
    return lang_dic[language]


def load_icon():
    nk_icon = open(NK_ICON_PATH, 'wb')
    nk_icon.write(b64decode(icon))
    nk_icon.close()


def changeNickname(path, nickname):
    dsk_config = os.path.join(path, r'desktop.ini')
    if not os.path.exists(dsk_config):
        f = open(dsk_config, 'w+', encoding='utf-16')
        f.close()
    desktop_config.read(dsk_config, encoding='utf-16')
    if not desktop_config.has_section(SHELL_CLASS_INFO):
        desktop_config.add_section(SHELL_CLASS_INFO)
    desktop_config.set(SHELL_CLASS_INFO, LOCAL_REC_NAME, nickname)
    if not desktop_config.has_section(VIEW_STATE):
        desktop_config.add_section(VIEW_STATE)
    desktop_config.set(VIEW_STATE, MODE, '')
    desktop_config.set(VIEW_STATE, VID, '')
    desktop_config.set(VIEW_STATE, FOLDER_TYPE, GENERIC)
    SetFileAttributes(dsk_config, FILE_ATTRIBUTE_NORMAL)
    desktop_config.write(open(dsk_config, 'w+', encoding='utf-16'))
    SetFileAttributes(dsk_config, FILE_ATTRIBUTE_HIDDEN)


def getDesktopData(path):
    if os.path.exists(path):
        dsk_config = os.path.join(os.path.normpath(path), 'desktop.ini')
        desktop_config.read(dsk_config, encoding='utf-16')
        if desktop_config.has_option(SHELL_CLASS_INFO, LOCAL_REC_NAME):
            return desktop_config.get(SHELL_CLASS_INFO, LOCAL_REC_NAME)
    return ''


def getAllConfigs():
    backup_config.read(NK_BACKUP_DATA)
    data = []
    for path in backup_config.sections():
        data.append([path, backup_config.get(path, NK_OPTION)])
    return data


def mk_ui():

    def run():
        path = targetNameBox.get()
        nickname = nickNameEntry.get()
        if not path or not nickname:
            showwarning(root_dic['warning'][lang_num],
                        root_dic['fill_in'][lang_num])
            return
        original_data = getDesktopData(path)
        if not backup_config.has_section(path):
            backup_config.add_section(path)
        if not backup_config.has_option(path, NK_OPTION) and original_data:
            backup_config.set(path, NK_ORI_OPTION, original_data)
        backup_config.set(path, NK_OPTION, nickname)
        backup_config.write(open(NK_BACKUP_DATA, 'w'))
        changeNickname(path, nickname)
        showinfo(root_dic['success'][lang_num],
                 root_dic['success_info'][lang_num])

    def exit_program():
        root.destroy()
        sys.exit()

    def refresh_value(path_dir=''):
        if not path_dir:
            path_dir = targetNameBox.get()
        else:
            if current_target_values:
                current_target_values.insert(0, path_dir)
            else:
                current_target_values.append(path_dir)
            targetNameBox['value'] = current_target_values
            targetNameBox.current(0)
        nickNameEntry.delete(0, 'end')
        nickNameEntry.insert(0, getDesktopData(path_dir))

    def set_path():
        path_dir = askdirectory()
        if path_dir and path_dir != targetNameBox.get():
            refresh_value(path_dir)

    def revert():
        path = targetNameBox.get()
        if not backup_config.has_section(path):
            showwarning(title=root_dic['warning'][lang_num],
                        message=root_dic['warning_info'][lang_num])
            return
        if backup_config.has_option(path, NK_ORI_OPTION):
            changeNickname(path, backup_config.get(path, NK_ORI_OPTION))
        else:
            dsk_config = os.path.join(path, r'desktop.ini')
            SetFileAttributes(dsk_config, FILE_ATTRIBUTE_NORMAL)
            os.remove(dsk_config)
        backup_config.remove_section(path)
        backup_config.write(open(NK_BACKUP_DATA, 'w'))
        showinfo(title=root_dic['success'][lang_num],
                 message=root_dic['success_rvt_info'][lang_num])

    def set_language(lang_num):
        config.set('GENERAL', 'language', language_list[lang_num])
        config.write(open(NK_CONFIG, 'w'))
        root.destroy()
        mk_ui()

    root = Tk()
    root.geometry('550x125')
    root.title('Windows Folder Nicknamer')
    root.protocol('WM_DELETE_WINDOW', exit_program)
    root.resizable(True, False)
    root.iconbitmap(NK_ICON_PATH)
    tkfont = font.nametofont("TkDefaultFont")
    tkfont.config(family='Microsoft YaHei UI')
    root.option_add("*Font", tkfont)

    language = config.get('GENERAL', 'language')
    lang_num = get_language_num(language)

    targetNameLabel = ttk.Label(root, text=root_dic['target'][lang_num])
    targetNameLabel.grid(row=0, column=0, padx=5, pady=5, sticky='NSE')
    current_target_values = []
    targetNameBox = ttk.Combobox(
        root, values=current_target_values, state='readonly')
    targetNameBox.grid(row=0, column=1, padx=5, pady=6, sticky='NSEW')
    chooseBtn = ttk.Button(
        root, text=root_dic['choose'][lang_num], command=set_path)
    chooseBtn.grid(row=0, column=2, padx=5, pady=5,
                   ipadx=5, sticky='NSE')
    chooseBtn.bind('<<ComboboxSelected>>',
                   lambda event: refresh_value())
    # targetNameEntry.insert(0, getData())
    NickNameLabel = ttk.Label(root, text=root_dic['nickname'][lang_num])
    NickNameLabel.grid(row=1, column=0, padx=5, pady=5, sticky='NSE')
    nickNameEntry = ttk.Entry(root)
    nickNameEntry.grid(row=1, column=1, columnspan=3, padx=5, pady=3,
                       sticky='NSEW')

    rvtBtn = ttk.Button(
        root, text=root_dic['revert'][lang_num], command=revert)
    rvtBtn.grid(row=2, column=0, padx=5, ipadx=10,
                pady=5, sticky='NSE')
    exeBtn = ttk.Button(root, text=root_dic['execute'][lang_num], command=run)
    exeBtn.grid(row=2, column=1, columnspan=3,
                pady=5, padx=5, sticky='NSEW')
    for value in getAllConfigs():
        refresh_value(value[0])
    root.grid_columnconfigure(1, weight=1, minsize=300)

    main_menu = Menu(root)
    language_menu = Menu(main_menu, tearoff=False)
    for i in range(len(language_list)):
        language_menu.add_checkbutton(offvalue=lang_num == get_language_num(language_list[i]),
                                      label=language_label[i], command=lambda i=i: set_language(i))
    main_menu.add_cascade(
        label=root_dic['language'][lang_num], menu=language_menu)
    root.config(menu=main_menu)

    root.bind('<Return>', run)
    root.mainloop()


def init_config():
    config.read(NK_CONFIG)
    if not config.has_section('GENERAL'):
        config.add_section('GENERAL')
    if not config.has_option('GENERAL', 'language'):
        if hex(windll.kernel32.GetSystemDefaultUILanguage()) == '0x804':
            config.set('GENERAL', 'language', 'Chinese')
        else:
            config.set('GENERAL', 'language', 'English')
        config.write(open(NK_CONFIG, 'w'))


def main():
    global backup_config, desktop_config, config
    backup_config = RawConfigParser()
    desktop_config = RawConfigParser()
    config = ConfigParser()

    if not os.path.exists(NK_PATH):
        os.mkdir(NK_PATH)
    if not os.path.exists(NK_BACKUP_DATA):
        backup_config.write(open(NK_BACKUP_DATA, 'w'))
    if not os.path.exists(NK_CONFIG):
        config.write(open(NK_CONFIG, 'w'))
    if not os.path.exists(NK_ICON_PATH):
        load_icon()
    backup_config.read(NK_BACKUP_DATA)
    init_config()
    mk_ui()


if __name__ == '__main__':
    main()
