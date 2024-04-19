from googletrans import Translator
import pandas as pd
import os
import numpy as np
from datetime import datetime
from tkinter import *
from threading import *
from tkinter import filedialog
from tkinter.messagebox import showerror, showwarning, showinfo

class MainWindows:
    def __init__(self, master):
        self.master = master
        master.title("Переводчик")
        master.geometry("350x200")
        master.resizable(0,0)

        self.label = Label(master, text="Выбрать папку с файлом Excel")
        self.label.pack()

        self.greet_button = Button(master, text="Открыть проводник", command=self.path_select)
        self.greet_button.pack(ipadx=50, ipady=10)

        self.label_path = Label(master, text="Путь до файла", bg="#FFFFFF")
        self.label_path.pack(ipadx=100, ipady=10, pady=5)

        self.close_button = Button(master, text="Сделать перевод", command=self.threading)
        self.close_button.pack(ipadx=50, ipady=10, pady=5)
    
    def path_select(self) -> None:
        '''Выбор файла для перевода'''
        excel_path = filedialog.askopenfilenames(filetypes = (('xlsx files','*.xlsx'),))
        self.label_path["text"] = excel_path
    
    def threading(self) -> None:
        '''Запуск перевода в отдельном потоке'''
        thread=Thread(target=self.make_translation)
        thread.start()
    
    def show_win_error(self, error_message:str) -> None:
        '''Вызывается новое окно сообщений с иноформацией об ошибках'''
        showerror(title="Ошибка", message= error_message)
    
    def show_win_info(self) -> None:
        '''Вызывается новое окно сообщений с иноформацией об успешном завершении'''
        showinfo(title="Информация", message="Файл успешно создан!")
        
    def read_excel(self, excel_file:str) -> list:
        '''Чтение файла .xlsx с текстом на английском языке и возвращение
           списка текста для перевода.
        '''
        try:
            df = pd.read_excel(excel_file)
            rows = df.values.tolist()
            words = list()
            words = [row for row in rows]
            return words
        except Exception as e:
            raise Exception(e)
    '''
    def get_file_location(self, file_name:str) -> str:
        try:
            file_location = os.path.join(os.path.dirname(__file__), file_name)
            return file_location
        except Exception as e:
            raise Exception(e)
    ''' 
    def traslate(self, source: list) -> list:
        '''Перевод текста с английского на русский.
           Возвращает список из списка [исходный текст, перевод]
        '''
        translator = Translator()
        res = list()
        traslate_str = str()
        for lines in source:
            for line in lines:
                row = []
                if str(line) != 'nan':
                    try:
                        row.append(line)
                        trans_res = translator.translate(line, dest='ru')
                        row.append(trans_res.text)
                        res.append(row)
                    except Exception as e:
                        raise Exception(e)    
        return res
    
    def make_translation(self) -> None:
        try:
            file_location = self.label_path["text"] # путь до файла с исходным текстом должен быть выбран до начала перевода
            res = self.read_excel(file_location) # чтение файла .xlsx с исходным текстом на английском языке
            translation_text = self.traslate(res) # перевод текста на русский язык
            df = pd.DataFrame(translation_text, columns =['Eng', 'Rus']) # создание структуры DataFrame
            now = datetime.now()
            excel_file_name = os.path.basename(file_location).split('/')[-1] # получение имя файла .xlsx с текстом на английском языке
            excel_file_name_path = os.path.dirname(os.path.abspath(file_location)) # получение пути до файла .xlsx с текстом на английском языке
            df.to_excel(f'{excel_file_name_path}\{excel_file_name}_translation.xlsx', sheet_name='Translation result') # создание нового .xlsx файла с иходным текстом и переводом
            self.show_win_info()
        except Exception as e:
            self.show_win_error(e)
  
if __name__ == "__main__":
    root = Tk()
    main_window = MainWindows(root)
    root.mainloop()