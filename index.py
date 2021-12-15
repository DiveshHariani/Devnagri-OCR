from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from PIL import ImageTk,Image
from pdf2image import convert_from_path
import xlsxwriter


class MainPage:
    def __init__(self,window):
        self.window = window
        self.frame = Frame(window, height=640, width=740, bg="white")
        self.frame.pack()
        self.add_labels()
        self.add_buttons()
        self.add_loader()
        #self.convert()



    def add_labels(self):
        self.brand = Label(self.frame, text="CodeKnights",padx=10, font=('Segoe UI Black bold', 25), bg='white')
        self.brand.place(x=30,y=30)

    def add_buttons(self):
        # upload_icon = PhotoImage(file = r'D:\coding\DevnagiriOCR\upload_icon.png')
        self.upload_icon=ImageTk.PhotoImage(Image.open("upload.png"))
        self.upload_label=Label(self.frame,text="Upload the file")
        self.upload_label.place(x=240,y=205)
        self.upload_label.config(font=('Seoge UI Black bold',18))

        
        #self.upload_label.place(x=150,y=200)
        self.upload_pdf = Button(self.frame,image=self.upload_icon,command=self.browse,text='Upload the pdf file',padx=10,bg='white')
        self.upload_pdf.place(x=420,y=200)
        self.convert_button=Button(self.frame,command=self.save,text='Convert',padx=10,font=('Seoge UI Black',10),bg='white')
        self.convert_button.place(x=310,y=255)
        self.excel_sheet = Button(self.frame,text='Converted Excel Sheet',padx=10,font=('Seoge UI Black bold',15),bg='white')
        self.excel_sheet.place(x=235,y=340)

    def add_loader(self):
        self.progress = ttk.Progressbar(self.frame, orient=HORIZONTAL,length=250, mode='determinate')
        self.progress.place(x=225,y=305)
       

    #def convert(self):
        #pass

    def load(self):
        import time
        self.progress['value'] = 20
        self.window.update_idletasks()
        time.sleep(0.5)

        self.progress['value'] = 40
        self.window.update_idletasks()
        time.sleep(0.5)

        self.progress['value'] = 60
        self.window.update_idletasks()
        time.sleep(0.5)

        self.progress['value'] = 80
        self.window.update_idletasks()
        time.sleep(0.5)

        self.progress['value'] = 100
        self.window.update_idletasks()
        time.sleep(0.5)
        self.progress['value'] = 100

    def browse(self):
        self.filename = filedialog.askopenfilename(initialdir='/', title='Select PDF')
        print(type(self.filename), self.filename)
        self.images = convert_from_path(self.filename, poppler_path=r"D:\popler\Release-20.11.0\poppler-20.11.0\bin")
        print(type(self.images), len(self.images))
        self.imageslist=[]
        for ind, image in enumerate(self.images):
            print(image.save(f"output{ind}.jpg", "JPEG"))
            imagename="output"+str(ind)+".jpg"
            self.imageslist.append(imagename)

        print(self.imageslist)

    def save(self):
        self.load()
        files = [('Excel Files', '*.xlsx')] 
        file = filedialog.asksaveasfilename(filetypes = files, defaultextension = '.xlsx')
        print(file)
        workbook = xlsxwriter.Workbook(file)
        for image in self.imageslist:
             worksheet = workbook.add_worksheet()
        #worksheet = workbook.add_worksheet()
        bold = workbook.add_format({'bold': True})
        row=0
        worksheet.write(row, 0, 'महारा',bold)
        workbook.close()

if __name__=='__main__':
    window = Tk()
    window.geometry('740x640')
    MainPage(window)
    window.mainloop()
