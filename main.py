# A Tkinter GUI program which converts word to pdf and pdf to word documents

# import the required modules
import win32com.client
import tkinter as tk
import os
import time
import tkinter.ttk as ttk
import threading
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

# set the fileFormat types of both formats
wdFormatPDF = 17
wdFormatWord = 16

# get the current working directory
cwd = os.getcwd()


class Main(Frame):

    # we create a Window.
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master = master
        self.grid()

        # call the functioons to display the inital GUI
        self.draw_title()
        self.draw_body()

        self.filepath = ''
        # self.isFileSelected = False

    # this draws the top heder of the app
    def draw_title(self):

        self.title = Label(self, width=40, height=1, bg='black',
                           text="Simple Converter", font=('Times', 20, 'bold'), fg='#45B8AC')
        self.title.grid(row=0, column=0, ipady=5)

    # this draws the body of the main window
    def draw_body(self):

        # make the whole body layout
        self.body = Frame(self, width=650, height=310, bg='#c0c0c0')
        self.body.grid(row=1, column=0)
        self.body.grid_propagate(False)

        # create the first button of word to pdf
        self.w_p = Button(self.body, width=170, height=170,
                          relief=RIDGE, image=img_one, cursor='hand2',
                          command=lambda: self.convert_page(word2pdf=True))

        # create the second button of pdf to word
        self.p_w = Button(self.body, width=170, height=170,
                          relief=tk.RIDGE, image=img_two, cursor='hand2',
                          command=lambda: self.convert_page(word2pdf=False))

        # place them on the windows screen
        self.w_p.grid(row=0, column=1, pady=(50, 55), padx=(30, 0))
        self.p_w.grid(row=0, column=2, pady=(50, 55), padx=(160, 0))

    def next_page(self):
        # destroy the widgets of the current body
        for w in self.body.winfo_children():
            w.destroy()
        # now destry the body
        self.body.destroy()

        # create the body layout again
        self.new_page = Frame(self, width=650, height=310, bg='#c0c0c0')
        self.new_page.grid(row=1, column=0)
        self.new_page.grid_propagate(False)

        # create the top label layout
        self.top_label = Frame(self.new_page, width=650,
                               height=310, bg='#c0c0c0')
        self.top_label.grid(row=0, column=0)
        self.top_label.grid_propagate(False)

    def convert_page(self, word2pdf=False):
        self.next_page()

        if word2pdf:
            my_label = 'Convert Word To PDF'
            select_label = 'Select a Word file'
            def convert_command(): return self.start(toPDF=True)
            def select_command(): return self.choose_file(selectPDF=False)
        else:
            my_label = 'Convert PDF To Word'
            select_label = 'Select a PDF file'
            def convert_command(): return self.start(toPDF=False)
            def select_command(): return self.choose_file(selectPDF=True)

        # this displays the top Label headline
        self.label1 = Label(self.top_label, text=my_label, bg='#ff6347', fg='black', font=(
            'Helvectica', 14, 'bold'), width=20)
        self.label1.grid(row=0, column=1, padx=5, columnspan=3, pady=30)

        # this displays the second label below the headline
        self.label2 = Label(self.top_label, text=select_label,
                            bg='#c0c0c0', fg='black', font=('Helvectica', 12))
        self.label2.grid(row=1, column=2, padx=5)

        # this displays the uplod icon on which click event is perfromed to attach the file
        self.upload_button = Button(self.top_label, image=add_img,
                                    bg='#c0c0c0', relief=RAISED, cursor='hand2',
                                    command=select_command)
        self.upload_button.grid(row=1, column=4, padx=5)

        # button to make user convert or go back to home
        # the button to convert the doc
        self.convert_it = Button(self.top_label, text='Convert',
                                 command=convert_command, cursor='hand2')
        self.convert_it.grid(row=3, column=0, padx=20, pady=100)

        # the button to go back to home
        self.go_back = Button(self.top_label, text='Home',
                              command=self.go_home, cursor='hand2')
        self.go_back.grid(row=3,  column=5, padx=100, pady=100)

        # this shows the path of the file uploaded
        self.file_label = Label(self.top_label, width=10, height=3,
                                bg='#c0c0c0', fg='Black',
                                font=('Times', 10))
        self.file_label.grid(row=1, column=1, padx=0, pady=0)

        # when no file is selected convert button ccan't be selected and is greyed out
        if self.filepath:
            self.convert_it.config(state=NORMAL)
        else:
            self.convert_it.config(state=DISABLED)

    def go_home(self):
        # when user comes home, destroy the new converter page and
        # call the function which draws home body page and making path to be null
        self.new_page.destroy()
        self.filepath = ''
        self.draw_body()

    # function which lets to chose between pdf or word based on which button is clicked
    def choose_file(self, selectPDF=True):
        # if user chooses pdf file
        if selectPDF:
            filetypes = (("PDF", "*.pdf"),)
        else:
            filetypes = (("Word", "*.doc"), ("Word", "*.docx"),)

        # Displays a dialog box from which the user can select a file.
        path = filedialog.askopenfilename(initialdir=cwd, filetypes=filetypes)
        if path:
            self.filepath = path
            self.convert_it.config(state=NORMAL)
            self.file_label['text'] = "File Uploaded!"

    def start(self, toPDF=False):
        # when converting starts make both buttons disabled
        self.go_back.config(state=DISABLED)
        self.convert_it.config(state=DISABLED)

        if toPDF:
            my_thread_function = self.make_pdf
        else:
            my_thread_function = self.make_word
        # create a thread and pass it which ever task is chosen to convert the document
        x = threading.Thread(target=my_thread_function)
        x.start()
        self.check_thread(x)

    def check_thread(self, thread):
        # check if thread is still performing or finished
        if thread.is_alive():
            # if it hasnt finished call it again after 100ms to finish it
            self.after(100, lambda: self.check_thread(thread))
        else:
            # if its executed, make the buttons normal again
            self.convert_it.config(state=tk.NORMAL)
            self.go_back.config(state=tk.NORMAL)

    # method which converted the file to pdf format
    def make_pdf(self):
        # Create a SaveAs dialog and return a file object opened in write-only mode.
        # this is done to save in the directory user chooses
        path = filedialog.asksaveasfilename(initialdir=cwd,
                                            filetypes=(("PDF", "*.pdf"),))
        if path:
            # create anew word object
            word = win32com.client.Dispatch("Word.Application")
            word.visible = False  # it shouldnt be visible

            # Normalize the specified path
            # using os.path.normpath() method
            file = os.path.normpath(self.filepath)
            path = os.path.normpath(path)  # where file will be saved
            document = word.Documents.Open(file)  # open the file to write
            document.SaveAs2(path, wdFormatPDF)  # save as pdf file

            document.Close()
            word.Quit()

            messagebox.showinfo("showinfo", "Kindly check Your new converted File!")

    # method which converted the file to pdf format
    def make_word(self):
        path = filedialog.asksaveasfilename(initialdir=cwd,
                                            filetypes=(("Word", "*.doc"), ("Word", "*.docx"),))
        if path:
            word = win32com.client.Dispatch("Word.Application")
            word.visible = False  # it shouldnt be visible

            # Normalize the specified path
            # using os.path.normpath() method
            file = os.path.normpath(self.filepath)
            path = os.path.normpath(path)  # where file will be saved
            document = word.Documents.Open(file)  # open the file to write
            document.SaveAs2(path, wdFormatWord)  # save as pdf file

            document.Close()
            word.Quit()

            messagebox.showinfo("showinfo", "Kindly check Your new converted File!")


if __name__ == "__main__":

    # the layout dimensions and properties
    root = Tk()
    root.geometry('600x350')
    root.title("File Converter")
    root.resizable(0, 0)

    # the images used in program
    img_one = PhotoImage(file='assets/wordToPdf.png')
    img_two = PhotoImage(file='assets/pdfToWord.png')
    add_img = PhotoImage(file='assets/add.png')

    # tell Python to run the Tkinter event loop on this class
    main = Main(master=root)
    main.mainloop()
