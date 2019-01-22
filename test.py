from tkinter import *
from tkinter import ttk
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showerror, showwarning, showinfo
from autoTravLog import *


class CDCFrame(Frame):

    def __init__(self):
        Frame.__init__(self)
        self.master.title("Traitement des fichiers des depenses")
        self.master.rowconfigure(5, weight=1)
        self.master.columnconfigure(5, weight=1)
        self.configure(background='#ffffff', width=400, height=200, relief=SUNKEN, bd = 2)
        self.master.resizable(width=False, height=False)
        self.fname = ""
        self.fname1 = ""
        self.foodPrice =  dict()
        with open('param.json', encoding='utf-8') as json_data:
            self.foodPrice = json.load(json_data)
       
        self.grid(sticky=W+E+N+S)
       
        self.label = Label(self, text='Fichier des depenses ',borderwidth=1, width= 19, bg='#ffffff', relief='flat')
        self.label.grid(row = 3)
        self.text = Text(self, height=1, width=50, relief=RAISED)
        self.text.grid(row=3, column=1, sticky=W)

        self.button = Button(self, text="Fichier 1", command=self.load_file,   width=20, relief=RAISED)
        self.button.grid(row=3, column=2, sticky=W)

        self.label1 = Label(self, text='Fichier a completer ',borderwidth=1, width= 19, bg='#ffffff', relief='flat')
        self.label1.grid(row = 4)
        self.text2 = Text(self, height=1, width=50, relief=RAISED)
        self.text2.grid(row=4, column=1, sticky=W)

        self.button2 = Button(self, text="Fichier 2", command=self.load_file1, width=20,  relief=RAISED)
        self.button2.grid(row=4, column=2, sticky=W)

        self.buttonTr = Button(self, text="Traiter", command=self.process_files, width=20, bg = '#39cccc', fg="white", relief=RAISED)
        self.buttonTr.grid(row=13, column=2, sticky=W)

        self.label3 = Label(self, text='Déjeuner',borderwidth=1, width= 15, bg='#ffffff', relief='flat', anchor=NW)
        self.label3.grid(row = 8)
        self.text3 = Text(self, height=1, width=5, relief=RAISED)
        self.text3.grid(row=8, column=1, sticky=W)
        self.text3.insert(END,self.foodPrice["Dejeuner"])

        self.label4 = Label(self, text='Dîner',borderwidth=1, width= 15, bg='#ffffff', relief='flat', anchor=NW)
        self.label4.grid(row = 9)
        self.text4 = Text(self, height=1, width=5, relief=RAISED)
        self.text4.grid(row=9, column=1, sticky=W)
        self.text4.insert(END,self.foodPrice["Diner"])


        self.label5 = Label(self, text='Souper',borderwidth=1, width= 15, bg='#ffffff', relief='flat', anchor=NW)
        self.label5.grid(row = 10)
        self.text5 = Text(self, height=1, width=5, relief=RAISED)
        self.text5.grid(row=10, column=1, sticky=W)
        self.text5.insert(END,self.foodPrice["Souper"])

        self.label6 = Label(self, text='Mileage V',borderwidth=1, width= 15, bg='#ffffff', relief='flat', anchor=NW)
        self.label6.grid(row = 11)
        self.text6 = Text(self, height=1, width=5, relief=RAISED)
        self.text6.grid(row=11, column=1, sticky=W)
        self.text6.insert(END,self.foodPrice["V"])

        self.label7 = Label(self, text='Mileage CV',borderwidth=1, width= 15, bg='#ffffff', relief='flat', anchor=NW)
        self.label7.grid(row = 12)
        self.text7 = Text(self, height=1, width=5, relief=RAISED)
        self.text7.grid(row=12, column=1, sticky=W)
        self.text7.insert(END,self.foodPrice["CV"])

        self.label8 = Label(self, text='Province',borderwidth=1, width= 15, bg='#ffffff', relief='flat', anchor=NW)
        self.label8.grid(row = 13)
        self.text8 = Text(self, height=1, width=20, relief=RAISED)
        self.text8.grid(row=13, column=1, sticky=W)
        self.text8.insert(END,self.foodPrice["Prov"])

    def load_file(self):
        try :
            self.fname = askopenfilename(filetypes=(("Excel files", "*.xlsx"),
                                            ("All files", "*.*") ))
            self.text.insert(END, self.fname)

        except:
            self.fname = ""
            showerror("Open Source File", "Failed to read file\n'%s'" % self.fname)
            return

    def load_file1(self):
        try :
            self.fname1 = askopenfilename(filetypes=(("Excel files", "*.xlsx"),
                                            ("All files", "*.*") ))
            self.text2.insert(END, self.fname1)

        except:
            self.fname1 = ""
            showerror("Open Source File", "Failed to read file\n'%s'" % self.fname1)
            return


    def close_win(self):
        self.destroy()

    def process_files(self):

        self.foodPrice["Dejeuner"] = self.text3.get(1.0,'end-1c')
        self.foodPrice["Diner"] = self.text4.get(1.0,'end-1c')
        self.foodPrice["Souper"] = self.text5.get(1.0,'end-1c')
        self.foodPrice["V"] = self.text6.get(1.0,'end-1c')
        self.foodPrice["CV"] = self.text7.get(1.0,'end-1c')
        self.foodPrice["Prov"] = self.text8.get(1.0,'end-1c')

        with open('param.json', 'w') as outfile:
            json.dump(self.foodPrice, outfile)
        
        if self.fname and self.fname1:

            sc = ExpenseLog()
            sc.foodData(self.foodPrice["Dejeuner"],self.foodPrice["Diner"],self.foodPrice["Souper"],self.foodPrice["V"],self.foodPrice["CV"], self.foodPrice["Prov"])
            sc.processExpLog(self.fname, self.fname1)
            
            showinfo('Status', 'Nombre de lignes inserées : ' + str(sc.lineNumber) + '\n' + 'Nombre de personnes : ' + str(sc.personNumber))

        else:

            showwarning("Fichier", "Veuillez charger les deux fichiers")




#if __name__ == "__main__":
CDCFrame().mainloop()
