import tkinter as tk

class TextShower(object):
    def __init__(self, requestMessage):
        self.root = tk.Tk()
        self.root.title('Mediprobe Excel Validator')
        self.string = ''
        self.frame = tk.Frame(self.root)
        self.frame.pack()
        self.centerText()
        self.root.bind('<Return>', self.showText)
        self.acceptInput(requestMessage)

    def centerText(self):
        window_width = 350
        window_height = 50
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x_cordinate = int((screen_width/2) - (window_width/2))
        y_cordinate = int((screen_height/2) - (window_height/2))
        self.root.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))

    def showText(self, whatever):
        self.gettext()

    def acceptInput(self, requestMessage):
        r = self.frame
        k = tk.Label(r,text=requestMessage)
        k.pack(side='left')
        v = tk.StringVar(self.root, value='input.xlsx')
        self.e = tk.Entry(r,textvariable=v)
        self.e.pack(side='left')
        self.e.focus_set()
        b = tk.Button(r,text='OK',command=self.gettext)
        b.pack(side='right')

    def gettext(self):
        self.string = self.e.get()
        self.root.destroy()

    def getString(self):
        return self.string

    def waitForInput(self):
        self.root.mainloop()
