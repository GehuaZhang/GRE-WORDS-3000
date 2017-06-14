import xlrd
import os
import Tkinter as tk
import tkMessageBox
import ttk
#from Tkinter import messagebox
#from tkinter import ttk

class ReadData:
    def __init__(self):
        self.book1 = xlrd.open_workbook("/Users/gehuazhang/Desktop/book1.xlsx", encoding_override="utf-8")
        self.sheet1 = self.book1.sheet_by_index(0)

    def read_sheet(self):
        return self.sheet1


class Application:
    def __init__(self, master):

        self.sheet = ReadData().read_sheet()
        self.notknow_words_position = []

        self.column_Posistion = 0
        self.master = master

        #Set Window Position
        self.width = 400
        self.height = 200
        self.screen_width = self.master.winfo_screenwidth()
        self.screen_height = self.master.winfo_screenheight()
        self.master.geometry("+{}+{}".format(int(self.screen_width/2-self.width/2),self.height)) #Only set position, no size
        #self.master.geometry("{}x{}+{}+{}".format(self.width,self.height,int(self.screen_width/2-self.width/2),self.height))

        #Set Frame
        self.frame_top = tk.Frame(master)
        self.frame_top.grid()
        self.paned_words = tk.PanedWindow(orient=tk.VERTICAL, width=self.width,height=self.height)

        #Set Label
        self.tk_label = tk.Label(self.frame_top, text="Pick List").grid(row=0,column=self.column_Posistion)

        #Set Option Menu
        self.options = []
        for i in range(1, 32):
            self.options.append(str(i))
        self.tk_menu_var = tk.StringVar()
        self.tk_menu_var.set("List")
        self.tk_menu = tk.OptionMenu(self.frame_top, self.tk_menu_var, *self.options)
        self.tk_menu.grid(row=0, column=self.column_Posistion+1)

        #Set Radio Button
        self.mode = [("Review","1"),("Learn","2")]
        self.tk_radio_var1 = tk.StringVar()
        self.tk_radio_var1.set("2")
        for txt, val in self.mode:
            self.tk_radio = tk.Radiobutton(self.frame_top, variable=self.tk_radio_var1, text=txt, value=val)
            self.tk_radio.grid(row=0, column=int(val)+self.column_Posistion+1)

        #Set Button
        self.tk_button = tk.Button(self.frame_top, text='OK', command=self.bt_click)
        self.tk_button.grid(row=0,column=self.column_Posistion+4)


    def bt_click(self):
        self.list_num = self.tk_menu_var.get()
        self.study_mode_num = self.tk_radio_var1.get()

        if self.list_num == "List":
            tkMessageBox.showinfo('Warning', "Choose a list")
            return

        self.position = (int(self.list_num) - 1) * 100
        self.words_interface(self.study_mode_num)


    def words_interface(self, study_mode_num):
        #Set words interface

        self.study_mode_dict = {"1": "Review", "2": "Learn"}
        if study_mode_num == "2":
            self.frame_top.destroy()
            self.paned_words.pack(fill=tk.BOTH, expand=1)
            self.Info = tk.Label(self.paned_words, text="List{} {}".format(int(self.position/100+1), self.study_mode_dict[study_mode_num]),anchor=tk.NW)
            self.paned_words.add(self.Info)
            self.learn_mode()

        if study_mode_num == "1":
            self.review_mode()

    def review_mode(self):
        if not os.path.isfile("unknow_{}.txt".format(str(self.list_num))):
            tkMessageBox.showinfo('Warning', "No words to review in this list.")
            return
        self.frame_top.destroy()

        with open("unknow_{}.txt".format(str(self.list_num)), 'r') as self.text_read:
            self.review_list=self.text_read.read().split("$$$$")

        self.paned_review = tk.PanedWindow()
        self.paned_review.pack(fill=tk.BOTH, expand=1)

        self.end_review_button = tk.Button(self.paned_review, text="OK", command=self.end_study)
        self.end_review_button.pack(side='bottom', fill='y')


        self.review_words_ttk = ttk.Treeview(self.paned_review, columns=["Words", "Meaning"], show='headings')
        self.review_words_ttk.pack(side='left')
        self.review_words_ttk.column('Words', width=100, anchor='w')
        self.review_words_ttk.column('Meaning', width=400, anchor='w')
        self.review_words_ttk.heading('Words', text='Words')
        self.review_words_ttk.heading('Meaning', text='Meaning')

        for x in self.review_list:
            self.review_words_ttk.insert('','end',values=(x.split(",")))

        self.review_words__scroll = ttk.Scrollbar(self.paned_review, orient="vertical")
        self.review_words__scroll.config(command=self.review_words_ttk.yview)
        self.review_words__scroll.pack(side='right', fill='y')
        self.review_words_ttk.configure(yscrollcommand=self.review_words__scroll.set)

    def learn_mode(self):
        self.counter = 0
        self.show_words(self.counter)

    def show_words(self, counter):
        if counter != 0:
            self.English.destroy()
            self.paned_button.destroy()

        if counter == 100:
            self.show_result()
            return

        self.eng_value = self.sheet.cell(self.position+counter,0).value
        self.chin_value = self.sheet.cell(self.position+counter,1).value

        self.English=tk.Label(self.paned_words, text=self.eng_value)
        self.English.config(font=("Courier",20))
        self.paned_words.add(self.English)

        self.paned_button = tk.PanedWindow(orient=tk.VERTICAL)
        self.paned_button.pack(fill=tk.Y)

        self.know_button = tk.Button(self.paned_button, text="Know", command=lambda: self.know_words(counter))
        self.paned_button.add(self.know_button)

        self.notknow_button = tk.Button(self.paned_button, text="Don't Know", command=lambda: self.notknow_words(counter))
        self.paned_button.add(self.notknow_button)

        counter=counter+1

    def know_words(self, counter):
        self.show_chinese(counter)

    #Click Not Know in first window
    def notknow_words(self, counter):
        self.show_chinese(counter)
        self.notcorrect_button.config(state=tk.DISABLED)
        self.notknow_words_position.append(counter-1)


    #Click know in first window but click not correct in second window
    def notcorrect_words(self, counter):
        self.notknow_words_position.append(counter-1)
        self.show_words(counter)


    def show_chinese(self, counter):
        self.English.config(text=self.eng_value+"\n\n"+self.chin_value)

        self.know_button.destroy()
        self.notknow_button.destroy()

        self.next_button = tk.Button(self.paned_button, text="Next", command=lambda :self.show_words(counter))
        self.paned_button.add(self.next_button)

        self.notcorrect_button = tk.Button(self.paned_button, text="Not Correct", command=lambda :self.notcorrect_words(counter))
        self.paned_button.add(self.notcorrect_button)


    def show_result(self):
        self.paned_words.destroy()
        self.paned_button.destroy()

        self.paned_result = tk.PanedWindow()
        self.paned_result.pack(fill=tk.BOTH,expand=1)


        self.end_button = tk.Button(self.paned_result, text="OK", command=self.end_study)
        self.end_button.pack(side='bottom',fill='y')

        self.unknow_words_ttk=ttk.Treeview(self.paned_result,columns=["Words","Meaning"],show = 'headings')
        self.unknow_words_ttk.pack(side='left')
        self.unknow_words_ttk.column('Words', width=100, anchor='w')
        self.unknow_words_ttk.column('Meaning', width=400, anchor='w')
        self.unknow_words_ttk.heading('Words', text='Words')
        self.unknow_words_ttk.heading('Meaning', text='Meaning')

        self.unknow_words_ttk_scroll = ttk.Scrollbar(self.paned_result, orient="vertical")
        self.unknow_words_ttk_scroll.config(command=self.unknow_words_ttk.yview)
        self.unknow_words_ttk_scroll.pack(side='right', fill='y')
        self.unknow_words_ttk.configure(yscrollcommand=self.unknow_words_ttk_scroll.set)

        self.unknow_list = [[],[]]

        for x in self.notknow_words_position:
            self.unknow_words_ttk.insert('','end',values=(self.sheet.cell(self.position+x,0).value,self.sheet.cell(self.position+x,1).value.replace('\n','\t')))
            self.unknow_list[0].append(self.sheet.cell(self.position+x,0).value)
            self.unknow_list[1].append(self.sheet.cell(self.position+x,1).value)



        self.write_excel(self.unknow_list)

    def write_excel(self, unknow_list):
        self.text_write = open("unknow_{}.txt".format(str(self.list_num)),'w')
        
        for i in range(len(unknow_list)):
            self.chin_character = unknow_list[1][i].encode('utf-8').replace('\n','\t')
            self.text_write.writelines(unknow_list[0][i].encode('utf-8')+", "+self.chin_character+"$$$$" )

        self.text_write.close()

    def end_study(self):
        self.master.destroy()




root = tk.Tk()
root.resizable(width=False, height=False)
root.title("3000")
app = Application(root)
root.mainloop()



