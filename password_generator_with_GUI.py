import random
import string


def generate_password():
    characters = string.ascii_letters + string.digits + string.punctuation
    password_len = int(box.get())
    password = random.sample(characters,password_len)
    enter.delete(0, 'end')
    enter.insert(0,password)
    
def copy_to_clipboard():
    password = enter.get()
    root.clipboard_clear()
    root.clipboard_append(password)
    root.update()






import tkinter
from tkinter import *
root = Tk()
root.title("PASSWORD GENERATOR")
root.config (bg="#0779f2")
length = Label(root,text="NUMBER OF CHARACTERS: ",font=("Helvetica",10,"bold"), bg="#0779f2", fg="#1c1c1b")
length.grid(padx= 10,pady=5)
box = Spinbox(root,from_=0,to_=15,width=5,font="Helvetica")
box.grid()
generate = Button(root,text="GENERATE", width=10, height=1, font=("Helvetica",10,"bold"),command = lambda:generate_password(), fg="#fcfeff",bg="#1c1c1b")
generate.grid(pady=5)
enter = Entry(root, width=30,)
enter.grid(padx =10,pady=5)
copy = Button(root,text="COPY", width=10, height=1, font=("Helvetica",10,"bold"),command = lambda:copy_to_clipboard(), fg="#fcfeff",bg="#1c1c1b")
copy.grid(pady=5)
root.mainloop()
