import tkinter as tk
from tkinter import messagebox, filedialog, StringVar, Radiobutton, Label, Button, Entry, Text, DISABLED, NORMAL, END, Toplevel
import os
import pandas as pd
from PIL import ImageTk
import code_email
import time

root = tk.Tk()

class Email:
    def __init__(self, root):
        self.root = root
        self.root.title("My Email Sender")
        self.root.geometry("1000x550+200+50")
        self.root.resizable(False, False)
        self.root.config(bg="skyblue")

        # Icons
        self.Email_icon = ImageTk.PhotoImage(file="email01.png")
        self.Setting_icon = ImageTk.PhotoImage(file="setting.png")

        # Radio buttons
        self.var_choice = StringVar()
        Radiobutton(root, text="Single", value="single", command=self.check_single_OR_bulk,
                    activebackground="skyblue", variable=self.var_choice, font=("times new roman", 30, "bold"),
                    bg="skyblue", fg="black").place(x=50, y=120)
        Radiobutton(root, text="Multiple", value="multiple", command=self.check_single_OR_bulk,
                    variable=self.var_choice, activebackground="skyblue",
                    font=("times new roman", 30, "bold"), bg="skyblue", fg="black").place(x=200, y=120)
        self.var_choice.set("single")

        # Title
        Label(self.root, text="Bulk Email Sender Panel", font=("Goudy Old Style", 48, "bold"), bg="blue",
              fg="white").place(x=0, y=0, relwidth=1)
        Label(self.root, text="Use Excel File For Sending Bulk Email At Once", font=("Calibri (body)", 14),
              bg="yellow", fg="black").place(x=0, y=80, relwidth=1)

        # Buttons
        Button(self.root, image=self.Email_icon, bg="blue", bd='0', command=self.send_email, height=65, width=100,
               activebackground="blue").place(x=1, y=4)
        Button(self.root, image=self.Setting_icon, bg="blue", bd='0', command=self.setting_window, height=65, width=100,
               activebackground="blue").place(x=890, y=4)

        # Labels
        Label(self.root, text="To (Email Address)", font=("times new roman", 18, "bold"), bg="skyblue",
              fg="black").place(x=50, y=200)
        Label(self.root, text="SUBJECT", font=("times new roman", 18, "bold"), bg="skyblue",
              fg="black").place(x=50, y=250)
        Label(self.root, text="MESSAGE", font=("times new roman", 18, "bold"), bg="skyblue",
              fg="black").place(x=50, y=300)

        # Status Labels
        self.Total = Label(self.root, font=("times new roman", 18, "bold"), bg="skyblue", fg="black")
        self.Total.place(x=50, y=500)
        self.Sent = Label(self.root, font=("times new roman", 18, "bold"), bg="skyblue", fg="darkgreen")
        self.Sent.place(x=350, y=500)
        self.Left = Label(self.root, font=("times new roman", 18, "bold"), bg="skyblue", fg="orange")
        self.Left.place(x=450, y=500)
        self.Failed = Label(self.root, font=("times new roman", 18, "bold"), bg="skyblue", fg="red")
        self.Failed.place(x=550, y=500)

        # Entry fields
        self.to_entry = Entry(self.root, font=("times new roman", 18), bg="lightgrey")
        self.to_entry.place(x=280, y=200, width=350, height=30)
        self.sub_entry = Entry(self.root, font=("times new roman", 18), bg="lightgrey")
        self.sub_entry.place(x=280, y=250, width=450, height=30)
        self.message_entry = Text(self.root, font=("times new roman", 18), bg="lightgrey")
        self.message_entry.place(x=280, y=300, width=700, height=190)

        # Action buttons
        Button(root, activebackground="skyblue", command=self.send_email, text="SEND",
               font=("times new roman", 20, "bold"), bg="black", fg="white").place(x=700, y=500, width=130, height=30)
        Button(root, activebackground="skyblue", command=self.clear1, text="CLEAR",
               font=("times new roman", 20, "bold"), bg="#ffcccb", fg="black").place(x=850, y=500, width=130, height=30)
        self.btn3 = Button(root, activebackground="skyblue", text="BROWSE", font=("times new roman", 20, "bold"),
                           bg="lightblue", command=self.Browse_button, cursor="hand2", state=DISABLED, fg="black")
        self.btn3.place(x=650, y=200, width=150, height=30)

        self.check_file_exist()

    def Browse_button(self):
        op = filedialog.askopenfile(initialdir='/', title="Select Excel File for Emails",
                                    filetypes=(("All Files", "*.*"), ("Excel Files", ".xlsx")))
        if op:
            data = pd.read_excel(op.name)
            if 'Email' in data.columns:
                self.EMAIL = [i for i in data['Email'] if pd.notnull(i)]
                if self.EMAIL:
                    self.to_entry.config(state=NORMAL)
                    self.to_entry.delete(0, END)
                    self.to_entry.insert(0, str(op.name.split("/")[-1]))
                    self.to_entry.config(state='readonly')
                    self.Total.config(text="Total: " + str(len(self.EMAIL)))
                    self.Sent.config(text="Sent: ")
                    self.Left.config(text="Left: ")
                    self.Failed.config(text="Failed: ")
            else:
                messagebox.showinfo("Error", "Please Select A File Which Has Emails", parent=self.root)

    def send_email(self):
        if not self.to_entry.get() or not self.sub_entry.get() or len(self.message_entry.get('1.0', END)) == 1:
            messagebox.showerror("ERROR", "All fields are required", parent=self.root)
        else:
            if self.var_choice.get() == "single":
                status = code_email.Email_send_function(self.to_entry.get(), self.sub_entry.get(),
                                                        self.message_entry.get('1.0', END), self.uname, self.pasw)
                if status == "s":
                    messagebox.showinfo("SUCCESS", "Email Has Been Sent", parent=self.root)
                else:
                    messagebox.showerror("Failed", "Email Not Sent", parent=self.root)
            else:
                self.failed = []
                self.s_count = 0
                self.f_count = 0
                for email in self.EMAIL:
                    status = code_email.Email_send_function(email, self.sub_entry.get(),
                                                            self.message_entry.get('1.0', END), self.uname, self.pasw).replace(u"\u2019", "'")
                    if status == "s":
                        self.s_count += 1
                    else:
                        self.f_count += 1
                    self.status_bar()
                    time.sleep(1)
                messagebox.showinfo("Success", "Email Has Been Sent, Please Check Status....", parent=self.root)

    def clear1(self):
        self.to_entry.config(state=NORMAL)
        self.to_entry.delete(0, END)
        self.sub_entry.delete(0, END)
        self.message_entry.delete('1.0', END)
        self.var_choice.set("single")
        self.btn3.config(state=DISABLED)
        self.Total.config(text="")
        self.Sent.config(text="")
        self.Left.config(text="")
        self.Failed.config(text="")

    def status_bar(self):
        self.Total.config(text="Status " + str(len(self.EMAIL)) + ":-")
        self.Sent.config(text="Sent: " + str(self.s_count))
        self.Left.config(text="Left: " + str(len(self.EMAIL) - (self.f_count + self.s_count)))
        self.Failed.config(text="Failed: " + str(self.f_count))
        self.Total.update()
        self.Sent.update()
        self.Left.update()
        self.Failed.update()

    def check_single_OR_bulk(self):
        if self.var_choice.get() == "single":
            messagebox.showinfo("single", "Set to Single", parent=self.root)
            self.btn3.config(state=DISABLED)
            self.to_entry.config(state=NORMAL)
            self.to_entry.delete(0, END)
            self.clear1()
        else:
            messagebox.showinfo("multiple", "Set to Bulk", parent=self.root)
            self.btn3.config(state=NORMAL)
            self.to_entry.delete(0, END)
            self.to_entry.config(state='readonly')

    def setting_clear(self):
        self.uname_entry.delete(0, END)
        self.pasw_entry.delete(0, END)

    def setting_window(self):
        self.check_file_exist()
        self.root2 = Toplevel()
        self.root2.title("Setting")
        self.root2.resizable(False, False)
        self.root2.geometry("700x450+350+90")
        self.root2.focus_force()
        self.root2.grab_set()
        self.root2.config(bg="lightgrey")
        Label(self.root2, text="Bulk Email Sender", padx=10, compound=LEFT,
              font=("Goudy Old Style", 48, "bold"), bg="black", fg="white").place(x=0, y=0, relwidth=1)
        Label(self.root2, text="Enter your valid Email Id and Password", font=("Calibri (body)", 14),
              bg="yellow", fg="black").place(x=0, y=80, relwidth=1)
        Label(self.root2, text="Email Address", font=("times new roman", 18, "bold"), bg="lightgrey",
              fg="black").place(x=50, y=150)
        Label(self.root2, text="Password", font=("times new roman", 18, "bold"), bg="lightgrey",
              fg="black").place(x=50, y=200)
        self.uname_entry = Entry(self.root2, font=("times new roman", 18), bg="lightyellow")
        self.uname_entry.place(x=250, y=150, width=330, height=30)
        self.pasw_entry = Entry(self.root2, font=("times new roman", 18), bg="lightyellow", show="*")
        self.pasw_entry.place(x=250, y=200, width=330, height=30)

        Button(self.root2, activebackground="skyblue", text="SAVE", font=("times new roman", 20, "bold"),
               bg="black", fg="white", command=self.save_setting).place(x=250, y=250, width=130, height=30)
        Button(self.root2, activebackground="skyblue", text="CLEAR", font=("times new roman", 20, "bold"),
               bg="#ffcccb", command=self.setting_clear, fg="black").place(x=400, y=250, width=130, height=30)

        self.uname_entry.insert(0, self.uname)
        self.pasw_entry.insert(0, self.pasw)

    def check_file_exist(self):
        if not os.path.exists("important.txt"):
            with open('important.txt', 'w') as f:
                f.write(",")
        with open('important.txt', 'r') as f2:
            self.credentials = [line.strip().split(",") for line in f2]
        self.uname = self.credentials[0][0]
        self.pasw = self.credentials[0][1]

    def save_setting(self):
        if not self.uname_entry.get() or not self.pasw_entry.get():
            messagebox.showinfo("ERROR", "All fields are required", parent=self.root2)
        else:
            with open('important.txt', 'w') as f:
                f.write(self.uname_entry.get() + "," + self.pasw_entry.get())
            messagebox.showinfo("Success", "Email and password are saved successfully", parent=self.root2)
            self.check_file_exist()

obj = Email(root)
root.mainloop()