import tkinter as tk
from tkinter import messagebox
import openpyxl

class ContactManager:
    def add(self):
        iden = self.iden1.get()
        name = self.name1.get()
        phone = self.phone1.get()
        email = self.email1.get()
        if iden and name and phone and email:
            contact_info = f"ID no: {iden}   \nName: {name}   \nPhone: {phone}   \nEmail: {email}"
            self.contacts_listbox.insert(tk.END, contact_info)
            self.iden1.delete(0, tk.END)
            self.name1.delete(0, tk.END)
            self.phone1.delete(0, tk.END)
            self.email1.delete(0, tk.END)
        else:
            messagebox.showwarning("Warning", "Please enter all the information.")

    def delete(self):
        selected_index = self.contacts_listbox.curselection()
        if selected_index:
            self.contacts_listbox.delete(selected_index)
        else:
            messagebox.showwarning("Warning", "Please select a contact to delete.")

    def update(self):
        selected_index = self.contacts_listbox.curselection()
        if selected_index:
            iden = self.iden1.get()
            name = self.name1.get()
            phone = self.phone1.get()
            email = self.email1.get()
            if iden and name and phone and email:
                contact_info = f"ID no: {iden}   \nName: {name}   \nPhone: {phone}   \nEmail: {email}"
                self.contacts_listbox.delete(selected_index)
                self.contacts_listbox.insert(tk.END, contact_info)
                self.iden1.delete(0, tk.END)
                self.name1.delete(0, tk.END)
                self.phone1.delete(0, tk.END)
                self.email1.delete(0, tk.END)
            else:
                messagebox.showwarning("Warning", "Please enter all the information.")
        else:
            messagebox.showwarning("Warning", "Please select a contact to update.")

    def search(self):
        search_term = self.search1.get().lower()

        for index in range(self.contacts_listbox.size()):
            contact_info = self.contacts_listbox.get(index).lower()
            if search_term in contact_info:
                self.contacts_listbox.selection_clear(0, tk.END)
                self.contacts_listbox.selection_set(index)
                self.contacts_listbox.see(index)
                break
        else:
            messagebox.showinfo("Contact Not Found", "No matching contact found.")

    def view(self):
        selected_index = self.contacts_listbox.curselection()
        if selected_index:
            contact_info = self.contacts_listbox.get(selected_index)
            messagebox.showinfo("Contact Details", contact_info)
        else:
            messagebox.showwarning("Warning", "Please select a contact to view.")

     
    def export_to_excel(self):
       wb = openpyxl.Workbook()
       sheet = wb.active

       for index in range(self.contacts_listbox.size()):
         contact_info = self.contacts_listbox.get(index)
         iden, name, phone, email = contact_info.split('\n')
         sheet.append([iden.split(': ')[1], name.split(': ')[1], phone.split(': ')[1], email.split(': ')[1]])

       excel_file = "contacts.xlsx"
       wb.save(excel_file)
       messagebox.showinfo("Export Successful", f"Contact data exported to {excel_file}")

    def create(self):
        self.root = tk.Tk()
        self.root.title("Contact Manager")
        self.root.geometry("400x600")
        self.root.config (bg="#05f5e9")

        self.iden_frame = tk.Frame(self.root)  
        self.iden_frame.pack(pady=10)

        self.iden_label = tk.Label(self.iden_frame, text="ID no:",bg="#05f5e9")
        self.iden_label.pack(side=tk.LEFT)

        self.iden1 = tk.Entry(self.iden_frame,width = 20)
        self.iden1.pack(side=tk.LEFT)
        
        self.name_frame = tk.Frame(self.root)
        self.name_frame.pack(pady=10)

        self.name_label = tk.Label(self.name_frame, text="Name:",bg="#05f5e9")
        self.name_label.pack(side=tk.LEFT)

        self.name1 = tk.Entry(self.name_frame,width = 20)
        self.name1.pack(side=tk.LEFT)


        self.phone_frame = tk.Frame(self.root)
        self.phone_frame.pack(pady=10)

        self.phone_label = tk.Label(self.phone_frame, text="Phone:",bg="#05f5e9")
        self.phone_label.pack(side=tk.LEFT)

        self.phone1 = tk.Entry(self.phone_frame,width = 20)
        self.phone1.pack(side=tk.LEFT)

        self.email_frame = tk.Frame(self.root)
        self.email_frame.pack(pady=10)

        self.email_label = tk.Label(self.email_frame, text="Email:",bg="#05f5e9")
        self.email_label.pack(side=tk.LEFT)

        self.email1 = tk.Entry(self.email_frame, width = 20)
        self.email1.pack(side=tk.LEFT)


        self.add = tk.Button(self.root, text="Add Contact", command=self.add,width = 25)
        self.add.pack(pady = 10)

        self.delete = tk.Button(self.root, text="Delete Contact", command=self.delete,width = 25)
        self.delete.pack(pady = 10)

        self.update = tk.Button(self.root, text="Update Contact", command=self.update,width = 25)
        self.update.pack(pady = 10)

        self.search_frame = tk.Frame(self.root)
        self.search_frame.pack(pady=10)

        self.search_label = tk.Label(self.search_frame, text="Search: ",bg="#05f5e9")
        self.search_label.pack(side=tk.LEFT)

        self.search1 = tk.Entry(self.search_frame, width = 25)
        self.search1.pack(side=tk.LEFT)

        self.search_button = tk.Button(self.root, text="Search Contact", command=self.search,width = 25)
        self.search_button.pack(pady = 10)

        self.view = tk.Button(self.root, text="View Contact", command=self.view,width = 25)
        self.view.pack(pady = 10)

        self.export = tk.Button(self.root, text="Export to Excel", command=self.export_to_excel,width = 25)
        self.export.pack(pady = 10)

        self.contacts_listbox = tk.Listbox(self.root, height = 60, width = 60)
        self.contacts_listbox.pack(pady = 10)

        self.root.mainloop()


contact_manager = ContactManager()
contact_manager.create()
