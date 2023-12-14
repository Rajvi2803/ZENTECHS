i = ""
def display(val):
    global i
    i+=val
    output.config(text=i)

def clear():
    global i
    i = ""
    output.config(text=i)

def solve():
    global i
    y =""
    if i !="":
        try:
            y= eval(i)
        except:
            y = "ERROR"
            i = ""
    output.config(text=y)


import tkinter
from tkinter import *
root = Tk()
root.title("CALCULATOR")
root.geometry("400x600")
root.resizable(False,False)
root.config (bg="#17161b")

output= Label(root,width=25,height=2,text="",font=("Helvetica",20))
output.pack()
Button(root,text="AC", width=3, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#f58936",command=lambda: clear()).place(x=10,y=100)
Button(root,text=".", width=3, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#f58936",command=lambda: display(".")).place(x=115,y=100)
Button(root,text="%", width=3, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#f58936",command=lambda: display("%")).place(x=220,y=100)
Button(root,text="+", width=3, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#1f69f2",command=lambda: display("+")).place(x=325,y=100)

Button(root,text="7", width=3, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#1c1c1b",command=lambda: display("7")).place(x=10,y=200)
Button(root,text="8", width=3, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#1c1c1b",command=lambda: display("8")).place(x=115,y=200)
Button(root,text="9", width=3, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#1c1c1b",command=lambda: display("9")).place(x=220,y=200)
Button(root,text="-", width=3, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#1f69f2",command=lambda: display("-")).place(x=325,y=200)

Button(root,text="4", width=3, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#1c1c1b",command=lambda: display("4")).place(x=10,y=300)
Button(root,text="5", width=3, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#1c1c1b",command=lambda: display("5")).place(x=115,y=300)
Button(root,text="6", width=3, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#1c1c1b",command=lambda: display("6")).place(x=220,y=300)
Button(root,text="*", width=3, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#1f69f2",command=lambda: display("*")).place(x=325,y=300)

Button(root,text="1", width=3, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#1c1c1b",command=lambda: display("1")).place(x=10,y=400)
Button(root,text="2", width=3, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#1c1c1b",command=lambda: display("2")).place(x=115,y=400)
Button(root,text="3", width=3, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#1c1c1b",command=lambda: display("3")).place(x=220,y=400)
Button(root,text="/", width=3, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#1f69f2",command=lambda: display("/")).place(x=325,y=400)

Button(root,text="0", width=10, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#1c1c1b",command=lambda: display("0")).place(x=10,y=500)
Button(root,text="=", width=10, height=1, font=("Helvetica",20,"bold"), fg="#fff",bd=1,bg="#1f69f2",command=lambda: solve()).place(x=210,y=500)



root.mainloop()

