

import os
from tkinter import Tk, ttk
from tkinter.filedialog import askdirectory
from configparser import ConfigParser


WFN_PATH = 'C:/ProgramData/WinFolder-Nicknamer'
WFN_CONFIG = os.path.join(WFN_PATH, r'config.ini')
DESKTOP_SECTION = '.ShellClassInfo'
DESKTOP_OPTION = 'localizedresourcename'


def execute_as(path, nickname):
    dsk_config = os.path.join(path, r'desktop.ini')
    if not os.path.exists(dsk_config):
        f = open(dsk_config, 'w', encoding='utf-16')
        f.write('\n')
        f.close()
    desktop_config.read(dsk_config, encoding='utf-16')
    if not desktop_config.has_section(nickname):
        desktop_config.add_section(nickname)
    sections = desktop_config.sections()

    for section in sections:
        print(section)
        print('--')
        print(desktop_config.options(section))


def getData():
    pass


def mk_ui():

    def run():
        path = targetNameBox.get()
        nickname = nickNameEntry.get()
        if not path or not nickname:
            return
        execute_as(path, nickname)

    def exit_program():
        root.quit()
        root.destroy()

    def insert_dir(pre=''):
        if pre:
            targetNameBox.delete(0, 'end')
            targetNameBox.insert(0, pre)
            return
        dir = askdirectory()
        if dir:
            targetNameBox.insert(0, dir)

    root = Tk()
    root.geometry('600x105')
    root.title('FrzVoid')
    root.protocol('WM_DELETE_WINDOW', exit_program)
    root.resizable(True, False)

    targetNameLabel = ttk.Label(root, text='Target-name:')
    targetNameLabel.grid(row=0, column=0, padx=20, pady=5, sticky='NSEW')
    targetNameBox = ttk.Combobox(root, state='readonly')
    targetNameBox.grid(row=0, column=1, padx=10, pady=5, sticky='NSEW')
    chooseBtn = ttk.Button(root, text='Choose Folder', command=insert_dir)
    chooseBtn.grid(row=0, column=2, padx=10, pady=5, sticky='NSEW')

    # targetNameEntry.insert(0, getData())
    NickNameLabel = ttk.Label(root, text='Nickname:')
    NickNameLabel.grid(row=1, column=0, padx=20, pady=5, sticky='NSEW')
    nickNameEntry = ttk.Entry(root)
    nickNameEntry.grid(row=1, column=1, columnspan=3, padx=10, pady=2,
                       sticky='NSEW')

    rvtBtn = ttk.Button(root, text='Exit', command=exit_program)
    rvtBtn.grid(row=2, column=0, padx=10,
                pady=5, sticky='NSEW')
    exeBtn = ttk.Button(root, text='Execute', command=run)
    exeBtn.grid(row=2, column=1, columnspan=3,
                ipadx=25, pady=5, padx=10, sticky='NSEW')
    root.grid_columnconfigure(0, weight=1, minsize=150)
    root.grid_columnconfigure(1, weight=1, minsize=300)
    root.grid_columnconfigure(2, weight=1, minsize=100)
    root.bind('<Return>', run)
    root.mainloop()


def main():
    global config, desktop_config
    config = ConfigParser()
    desktop_config = ConfigParser()
    config.read('config.ini')
    mk_ui()


if __name__ == '__main__':
    main()
