
import customtkinter
import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
import pandas as pd


class BusinessTrackingApp(customtkinter.CTk):


    def __init__(self):
        super().__init__()

        self.title("Vize Takip Sistemi")
        self.geometry("1920×1080")
        #style = ttk.Style()
        """style.configure("Treeview", background="lightblue", foreground="black")
        style.configure("Treeview.Heading", background="lightblue", foreground="black")"""


        # Create a DataFrame to hold the business data
        self.df = pd.DataFrame(columns=["İsim", "Soy İsim", "E-Mail", "Telefon Numarası", "Vize Başvuru Tarihi", "Ülke"])

        # Create labels and entry fields for data input
        self.label_name = customtkinter.CTkLabel(self, text="İsim:")
        self.label_name.grid(row=0, column=0, padx=10, pady=5, sticky="w")
        self.entry_name = customtkinter.CTkEntry(self)
        self.entry_name.grid(row=0, column=1, padx=10, pady=5, sticky="we")

        self.label_surname = customtkinter.CTkLabel(self, text="Soy İsim:")
        self.label_surname.grid(row=1, column=0, padx=10, pady=5, sticky="w")
        self.entry_surname = customtkinter.CTkEntry(self)
        self.entry_surname.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

        self.label_email = customtkinter.CTkLabel(self, text="E-Mail:")
        self.label_email.grid(row=2, column=0, padx=10, pady=5, sticky="w")
        self.entry_email = customtkinter.CTkEntry(self)
        self.entry_email.grid(row=2, column=1, padx=10, pady=5, sticky="ew")

        self.label_phone = customtkinter.CTkLabel(self, text="Telefon Numarası:")
        self.label_phone.grid(row=3, column=0, padx=10, pady=5, sticky="w")
        self.entry_phone = customtkinter.CTkEntry(self)
        self.entry_phone.grid(row=3, column=1, padx=10, pady=5, sticky="ew")

        self.label_visa_date = customtkinter.CTkLabel(self, text="Vize Başvuru Tarihi:")
        self.label_visa_date.grid(row=4, column=0, padx=10, pady=5, sticky="w")
        self.entry_visa_date = customtkinter.CTkEntry(self)
        self.entry_visa_date.grid(row=4, column=1, padx=10, pady=5, sticky="ew")

        self.label_which_country = customtkinter.CTkLabel(self, text="Ülke:")
        self.label_which_country.grid(row=5, column=0, padx=10, pady=5, sticky="w")
        self.entry_which_country = customtkinter.CTkEntry(self)
        self.entry_which_country.grid(row=5, column=1, padx=10, pady=5, sticky="ew")

        # Create buttons for adding, updating, and deleting data
        self.button_add = customtkinter.CTkButton(self, text="Ekle", command=self.add_data)
        self.button_add.grid(row=6, column=0, padx=10, pady=5, sticky="nsew")

        self.button_update = customtkinter.CTkButton(self, text="Güncelle", command=self.update_data)
        self.button_update.grid(row=6, column=1, padx=10, pady=5, sticky="nsew")

        self.button_delete = customtkinter.CTkButton(self, text="Sil", command=self.delete_data)
        self.button_delete.grid(row=6, column=2, padx=10, pady=5, sticky="nsew")



        # Create a Treeview widget to display the business data

        self.treeview = ttk.Treeview(self, columns=["İsim", "Soy İsim", "E-Mail", "Telefon Numarası", "Vize Başvuru Tarihi", "Ülke"], show="headings", height=20)

        self.treeview.heading("İsim", text="İsim")
        self.treeview.heading("Soy İsim", text="Soy İsim")
        self.treeview.heading("E-Mail", text="E-Mail")
        self.treeview.heading("Telefon Numarası", text="Telefon Numarası")
        self.treeview.heading("Vize Başvuru Tarihi", text="Vize Başvuru Tarihi")
        self.treeview.heading("Ülke", text="Ülke")
        self.treeview.grid(row=7, column=0, columnspan=3, padx=10, pady=5)
        y_scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.treeview.yview)
        self.treeview.configure(yscrollcommand=y_scrollbar.set)
        y_scrollbar.grid(row=7, column=10, sticky='ns')
    def add_data(self):
        name = self.entry_name.get()
        surname = self.entry_surname.get()
        email = self.entry_email.get()
        phone = self.entry_phone.get()
        visa_date = self.entry_visa_date.get()
        country = self.entry_which_country.get()

        if name and surname and email and phone and visa_date and country:
            data = {"İsim": [name], "Soy İsim": [surname], "E-Mail": [email], "Telefon Numarası": [phone], "Vize Başvuru Tarihi": [visa_date], "Ülke": [country]}
            df_new = pd.DataFrame(data)
            self.df = pd.concat([self.df, df_new], ignore_index=True)

            # Update the Treeview to show the new data
            self.update_treeview()

            messagebox.showinfo("Success", "Data added successfully.")
            self.clear_entries()
        else:
            messagebox.showerror("Error", "Please enter all fields.")

    def update_data(self):
        selected_item = self.treeview.selection()
        if selected_item:
            name = self.entry_name.get()
            surname = self.entry_surname.get()
            email = self.entry_email.get()
            phone = self.entry_phone.get()
            visa_date = self.entry_visa_date.get()
            country = self.entry_which_country.get()

            if name or surname or email or phone or visa_date or country:
                index = self.treeview.index(selected_item)

                # Get the existing values from the DataFrame for the selected record
                existing_name = self.df.at[index, "İsim"]
                existing_surname = self.df.at[index, "Soy İsim"]
                existing_email = self.df.at[index, "E-Mail"]
                existing_phone = self.df.at[index, "Telefon Numarası"]
                existing_visa_date = self.df.at[index, "Vize Başvuru Tarihi"]
                existing_country = self.df.at[index, "Ülke"]

                # Update only the fields that have non-empty values from the entry widgets
                self.df.at[index, "İsim"] = name if name else existing_name
                self.df.at[index, "Soy İsim"] = surname if surname else existing_surname
                self.df.at[index, "E-Mail"] = email if email else existing_email
                self.df.at[index, "Telefon Numarası"] = phone if phone else existing_phone
                self.df.at[index, "Vize Başvuru Tarihi"] = visa_date if visa_date else existing_visa_date
                self.df.at[index, "Ülke"] = country if country else existing_country

                # Update the Treeview to reflect the changes
                self.update_treeview()

                messagebox.showinfo("Success", "Data updated successfully.")
                self.clear_entries()
            else:
                messagebox.showerror("Error", "Please enter at least one field to update.")
        else:
            messagebox.showerror("Error", "Please select a record to update.")

    def delete_data(self):
        selected_item = self.treeview.selection()
        if selected_item:
            index = self.treeview.index(selected_item)
            self.df.drop(index, inplace=True)

            # Update the Treeview to remove the deleted data
            self.update_treeview()

            messagebox.showinfo("Success", "Data deleted successfully.")
            self.clear_entries()
        else:
            messagebox.showerror("Error", "Please select a record to delete.")

    def update_treeview(self):
        # Clear the current contents of the Treeview
        for item in self.treeview.get_children():
            self.treeview.delete(item)

        # Populate the Treeview with the updated DataFrame
        for index, row in self.df.iterrows():
            path = 'Vİze2023.xlsx'
            workbook = openpyxl.load_workbook(path)
            sheet = workbook.active
            row_values = [row["İsim"], row["Soy İsim"], row["E-Mail"], row["Telefon Numarası"], row["Vize Başvuru Tarihi"], row["Ülke"]]
            sheet.append(row_values)
            workbook.save(path)
            self.treeview.insert("", "end", values=(row["İsim"], row["Soy İsim"], row["E-Mail"], row["Telefon Numarası"], row["Vize Başvuru Tarihi"], row["Ülke"]))

    """def load_data(self):
        path = 'Vİze2023.xlsx'
        workbook = openpyxl.load_workbook(path)
        sheet = workbook.active
        list_values = list(sheet.values)
        for col_name in list_values[0]:
            self.treeview.heading(col_name, text=col_name)"""
    def clear_entries(self):
        self.entry_name.delete(0, tk.END)
        self.entry_surname.delete(0, tk.END)
        self.entry_email.delete(0, tk.END)
        self.entry_phone.delete(0, tk.END)
        self.entry_visa_date.delete(0, tk.END)
        self.entry_which_country.delete(0, tk.END)

    def destroy(self):
        super().destroy()

def open_business_tracking_app():

    app = BusinessTrackingApp()
    app.title("Vize Takip Sistemi")
    app.geometry("1920×1080")
    root.destroy()
    app.mainloop()


if __name__ == "__main__":
    customtkinter.set_appearance_mode("dark")  # Modes: "System" (standard), "Dark", "Light"
    customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"
    root = customtkinter.CTkToplevel()
    root.title("Anasayfa")
    root.geometry("1080x720")


    button1 = customtkinter.CTkButton(root, text="Müşteri Ekle", command=open_business_tracking_app)

    button1.pack(pady=10)



    root.mainloop()



