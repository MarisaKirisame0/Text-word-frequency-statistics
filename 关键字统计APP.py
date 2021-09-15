import os
import ctypes
import jieba
import tkinter
import tkinter.ttk as ttk
import tkinter.filedialog
import tkinter.messagebox
from PIL import Image,ImageTk

ctypes.windll.shcore.SetProcessDpiAwareness(1)
ScaleFactor = ctypes.windll.shcore.GetScaleFactorForDevice(0)

class WordDetect():
    def __init__(self):
        self.passage=''
        self.passage_words=[]
        self.important_words_list=[]
        self.abandon_words_list=[]
        self.result_items={}

    def _get_the_file_text(self,address):
        try:
            f = open(address, 'r', encoding="utf-8")
        except IOError as s:
            return None
        try:
            buffer = f.read()
        except:
            return None
        f.close()
        return buffer

    def _get_words_list(self,address):
        txt = self._get_the_file_text(address)
        if txt == None:
            return None
        else:
            txt_list = txt.split('\n')
            return txt_list


    def _text_word_get(self):
        txt = self.passage
        for ch in '!"#$%&()*+,-./;:<=>?@[\\]^‘_{|}~，。“”':
            txt = txt.replace(ch, " ")
        txt = txt.replace('\n', ' ')
        self.passage_words = jieba.lcut(txt)

    def _word_frequency_cal(self):
        counts = {}
        if len(self.important_words_list):
            for word in self.passage_words:
                if len(word) == 1:
                    continue
                elif word in self.important_words_list:
                    counts[word] = counts.get(word, 0) + 1
            if len(self.abandon_words_list):
                for word in self.abandon_words_list:
                    del counts[word]
        else:
            for word in self.passage_words:
                if len(word) == 1:
                    continue
                else:
                    counts[word] = counts.get(word, 0) + 1
            if len(self.abandon_words_list):
                for word in self.abandon_words_list:
                    if word in counts.keys():
                        del counts[word]
        items = list(counts.items())
        items.sort(key=lambda x: x[1], reverse=True)
        self.result_items = items

class App():
    def __init__(self):
        self._root_window=tkinter.Tk()
        self._detector=WordDetect()
        self._str_label =tkinter.StringVar()
        self._str_label2 = tkinter.StringVar()
        self._str_label3 = tkinter.StringVar()
        self._str_label.set('选择文件以开始')
        self._str_label2.set('')
        self._str_label3.set('')

    def _get_address(self):
        default_dir = r"./"
        file_path = tkinter.filedialog.askopenfilename(title=u'选择文件', initialdir=(os.path.expanduser(default_dir)))
        return file_path

    def button_call_back_function1(self):
        add = self._get_address()
        x = self._detector._get_the_file_text(add)
        if x==None:
            tkinter.messagebox.showinfo(title='WARNING', message='待检测文本读取失败')
            self._str_label.set('选择文件以开始')
        else:
            self._detector.passage = x
            self._str_label.set('成功获取文章')
            self._detector._text_word_get()

    def button_call_back_function2(self):
        add = self._get_address()
        x = self._detector._get_words_list(add)
        if x == None:
            tkinter.messagebox.showinfo(title='WARNING', message='关键字词文本读取失败')
            self._str_label2.set('')
        else:
            self._detector.important_words_list = x
            self._str_label2.set('成功读取关键字词文件')

    def button_call_back_function3(self):
        add = self._get_address()
        x = self._detector._get_words_list(add)
        if x == None:
            tkinter.messagebox.showinfo(title='WARNING', message='排除字词文本读取失败')
            self._str_label3.set('')
        else:
            self._detector.abandon_words_list = x
            self._str_label3.set('成功排除关键字词文件')

    def button_call_back_function4(self):
        if len(self._detector.passage_words):
            self._str_label.set('文章关键字统计')
            self._detector._word_frequency_cal()
            newtop = tkinter.Tk()
            newtop.geometry('600x400')
            newtop.tk.call('tk', 'scaling', ScaleFactor / 50)
            newtop.title('统计结果')
            tree = ttk.Treeview(newtop)
            tree.pack()
            label = ttk.Label(newtop, text='scrolling Mouse to see results', font='Arial -32')
            label.pack(side='bottom')
            tree["columns"] = ("关键字词", "出现次数")
            tree.column("关键字词", width=200)
            tree.column("出现次数", width=200)
            tree.heading("关键字词", text="关键字词-words")
            tree.heading("出现次数", text="出现次数-times")
            for i in range(len(self._detector.result_items)):
                j = len(self._detector.result_items)-i-1
                word, count = self._detector.result_items[i]
                tree.insert("", 0, text=str(j), values=(str(word), str(count)))
            newtop.mainloop()
        else:
            tkinter.messagebox.showinfo(title='WARNING', message='未选择待检测文本 或 原始文本字词分割失败')

    def WindowInit(self):
        top=self._root_window
        top.geometry('600x600')
        top.tk.call('tk', 'scaling', ScaleFactor/50)
        top.title('文章关键字词统计')
        label = ttk.Label(top, textvariable=self._str_label, font='Arial -32')
        label.pack(expand=1)
        label = ttk.Label(top, textvariable=self._str_label2, font='Arial -32')
        label.pack(expand=1)
        label = ttk.Label(top, textvariable=self._str_label3, font='Arial -32')
        label.pack(expand=1)
        panel = ttk.Frame(top)
        button1 = ttk.Button(panel, text='选择待检测文本', command=self.button_call_back_function1)
        button1.pack(side='top')
        button2 = ttk.Button(panel, text='选择关键词列表文本文件', command=self.button_call_back_function2)
        button2.pack(side='top')
        button3 = ttk.Button(panel, text='选择排除词列表文本文件', command=self.button_call_back_function3)
        button3.pack(side='top')
        button4 = ttk.Button(panel, text='开始检测', command=self.button_call_back_function4)
        button4.pack(side='bottom')
        panel.pack(side='bottom')
        canvas = tkinter.Canvas(top, width=600, height=600, bg='black')
        canvas.pack()
        image = Image.open('./bg.jpg')
        im = ImageTk.PhotoImage(image)
        canvas.create_image(10, 10, anchor=tkinter.NW, image=im)


    def ShowWindow(self):
        self._root_window.mainloop()



if __name__=='__main__':
    app=App()
    app.WindowInit()
    app.ShowWindow()
