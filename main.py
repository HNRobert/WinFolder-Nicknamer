

import os
from tkinter import Tk, ttk
from tkinter.filedialog import askdirectory
from tkinter.messagebox import showwarning, showinfo
import tkinter.font as font
from configparser import RawConfigParser
from win32con import FILE_ATTRIBUTE_HIDDEN, FILE_ATTRIBUTE_NORMAL
from win32api import SetFileAttributes

NK_PATH = 'C:/ProgramData/WinFolder-Nicknamer'
NK_BACKUP_DATA = os.path.join(NK_PATH, r'backup_data.ini')
NK_OPTION = 'nickname'
NK_ORI_OPTION = 'original'
DESKTOP_SECTION = '.ShellClassInfo'
DESKTOP_OPTION = 'localizedresourcename'


def changeNickname(path, nickname):
    dsk_config = os.path.join(path, r'desktop.ini')
    SetFileAttributes(dsk_config, FILE_ATTRIBUTE_NORMAL)
    if not os.path.exists(dsk_config):
        f = open(dsk_config, 'w', encoding='utf-16')
        f.close()
    desktop_config.read(dsk_config, encoding='utf-16')
    if not desktop_config.has_section(DESKTOP_SECTION):
        desktop_config.add_section(DESKTOP_SECTION)
    desktop_config.set(DESKTOP_SECTION, DESKTOP_OPTION, nickname)
    desktop_config.write(open(dsk_config, 'w', encoding='utf-16'))
    SetFileAttributes(dsk_config, FILE_ATTRIBUTE_HIDDEN)


def getDesktopData(path):
    if os.path.exists(path):
        dsk_config = os.path.join(path, r'desktop.ini')
        desktop_config.read(dsk_config, encoding='utf-16')
        if desktop_config.has_option(DESKTOP_SECTION, DESKTOP_OPTION):
            return desktop_config.get(DESKTOP_SECTION, DESKTOP_OPTION)
    return ''


def getAllConfigs():
    config.read(NK_BACKUP_DATA)
    data = []
    for path in config.sections():
        data.append([path, config.get(path, NK_OPTION)])
    return data


def mk_ui():

    def run():
        path = targetNameBox.get()
        nickname = nickNameEntry.get()
        if not path or not nickname:
            showwarning('Warning', 'Please fill in all the fields.')
        original_data = getDesktopData(path)
        if not config.has_section(path):
            config.add_section(path)
        config.set(path, NK_OPTION, nickname)
        if not config.has_option(path, NK_ORI_OPTION) and original_data:
            config.set(path, NK_ORI_OPTION, original_data)
            config.write(open(NK_BACKUP_DATA, 'w'))
        changeNickname(path, nickname)
        showinfo('Nicknamer', 'Nickname is Successfully Made!')

    def exit_program():
        root.quit()
        root.destroy()

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
        if not config.has_option(path_dir, NK_ORI_OPTION):
            rvtBtn.config(state='disabled')
        else:
            rvtBtn.config(state='normal')

    def set_path():
        path_dir = askdirectory()
        if path_dir:
            refresh_value(path_dir)

    def revert(path):
        if not config.has_section(path):
            showwarning(title='Warning', message="Previous data wasn't found")
        prev_data = config.get(path, NK_ORI_OPTION)
        if prev_data:
            changeNickname(path, prev_data)
        else:
            showwarning(title='Warning', message="Previous data wasn't found")

    root = Tk()
    root.geometry('550x106')
    root.title('Windows Folder Nicknamer')
    root.protocol('WM_DELETE_WINDOW', exit_program)
    root.resizable(True, False)
    tkfont = font.nametofont("TkDefaultFont")
    tkfont.config(family='Microsoft YaHei UI')
    root.option_add("*Font", tkfont)

    targetNameLabel = ttk.Label(root, text='Target Folder :')
    targetNameLabel.grid(row=0, column=0, padx=5, pady=5, sticky='NSE')
    current_target_values = []
    targetNameBox = ttk.Combobox(
        root, values=current_target_values, state='readonly')
    targetNameBox.grid(row=0, column=1, padx=5, pady=5, sticky='NSEW')
    chooseBtn = ttk.Button(root, text='Choose Folder', command=set_path)
    chooseBtn.grid(row=0, column=2, padx=5, pady=5,
                   ipadx=5, ipady=1, sticky='NSE')
    chooseBtn.bind('<<ComboboxSelected>>',
                   lambda event: refresh_value())
    # targetNameEntry.insert(0, getData())
    NickNameLabel = ttk.Label(root, text='Nickname :')
    NickNameLabel.grid(row=1, column=0, padx=5, pady=5, sticky='NSE')
    nickNameEntry = ttk.Entry(root)
    nickNameEntry.grid(row=1, column=1, columnspan=3, padx=5,
                       sticky='NSEW')
    

    rvtBtn = ttk.Button(root, text='Revert', command=revert)
    rvtBtn.grid(row=2, column=0, padx=5, ipadx=10,
                pady=5, sticky='NSE')
    exeBtn = ttk.Button(root, text='Execute', command=run)
    exeBtn.grid(row=2, column=1, columnspan=3,
                pady=5, padx=5, sticky='NSEW')
    for value in getAllConfigs():
        refresh_value(value[0])
    # root.grid_columnconfigure(0, weight=1, minsize=60)
    root.grid_columnconfigure(1, weight=1, minsize=300)
    # root.grid_columnconfigure(2, weight=1, minsize=70)
    root.bind('<Return>', run)
    root.mainloop()


def main():
    global config, desktop_config
    config = RawConfigParser()
    desktop_config = RawConfigParser()
    config.read(NK_BACKUP_DATA)
    if not os.path.exists(NK_PATH):
        os.mkdir(NK_PATH)
    if not os.path.exists(NK_BACKUP_DATA):
        config.write(open(NK_BACKUP_DATA, 'w'))
    mk_ui()


if __name__ == '__main__':
    main()
