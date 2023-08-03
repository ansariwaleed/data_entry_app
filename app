import openpyxl as xl
import customtkinter
from tkinter import *
from tkinter import messagebox
from datetime import date

app = customtkinter.CTk()
app.geometry("600x600")
app.title("SONI ENTERPRISES")

font1 = ('Arial', 20, 'bold')
font2 = ('Arial', 15, 'bold')

# Create a lookup table for place rates
place_rates = {
    "Faizabad": 5.0,
    "Delhi": 7.0,
    "Mumbai": 8.5,
    # Add more place-rate mappings as needed
}

def submit():
    file = xl.load_workbook('mydata.xlsx')
    sheet = file.active

    product_value = product_entry.get()
    category_value = category_entry.get()
    docket_value = docket_entry.get()
    place_value = place_entry.get()
    weight_value = weight_entry.get()
    rate_value = rate_entry.get()
    payment_value = "Yes" if payment_checkbox.get() else "No"

    if not rate_value:
        rate_value = calculate_rate(place_value, weight_value)

    row = sheet.max_row + 1

    # Save the current date in the first column (A)
    sheet.cell(row=row, column=1, value=date.today().strftime("%Y-%m-%d"))

    # Save the other information in subsequent columns
    sheet.cell(row=row, column=2, value=product_value)
    sheet.cell(row=row, column=3, value=category_value)
    sheet.cell(row=row, column=4, value=docket_value)
    sheet.cell(row=row, column=5, value=place_value)
    sheet.cell(row=row, column=6, value=weight_value)
    sheet.cell(row=row, column=7, value=rate_value)
    sheet.cell(row=row, column=8, value=payment_value)

    file.save('mydata.xlsx')
    messagebox.showinfo(title="Success", message="Data has been saved")

    # Clear all the data fields
    clear()

def clear():
    product_entry.delete(0, END)
    category_entry.delete(0, END)
    docket_entry.delete(0, END)
    place_entry.delete(0, END)
    weight_entry.delete(0, END)
    rate_entry.delete(0, END)
    payment_checkbox.deselect()

def calculate_rate(place, weight):
    # Look up the rate based on the place in the place_rates dictionary
    rate = place_rates.get(place, 0.0)
    return rate * float(weight)

product_label = customtkinter.CTkLabel(app, text="Product", font=font1, width=10)
product_label.place(x=20, y=30)

category_label = customtkinter.CTkLabel(app, text="Category", font=font1, width=10)
category_label.place(x=20, y=90)

docket_label = customtkinter.CTkLabel(app, text="Docket", font=font1, width=10)
docket_label.place(x=20, y=150)

place_label = customtkinter.CTkLabel(app, text="Place", font=font1, width=10)
place_label.place(x=20, y=210)

weight_label = customtkinter.CTkLabel(app, text="Weight", font=font1, width=10)
weight_label.place(x=20, y=270)

rate_label = customtkinter.CTkLabel(app, text="Rate", font=font1, width=10)
rate_label.place(x=20, y=330)

payment_label = customtkinter.CTkLabel(app, text="Payment", font=font1, width=10)
payment_label.place(x=20, y=390)

product = StringVar()
category = StringVar()
docket = StringVar()
place = StringVar()
weight = StringVar()
rate = StringVar()
payment_checkbox = BooleanVar()

product_entry = customtkinter.CTkEntry(app, textvariable=product, width=150)
product_entry.place(x=153, y=33)

category_entry = customtkinter.CTkEntry(app, textvariable=category, width=150)
category_entry.place(x=153, y=95)

docket_entry = customtkinter.CTkEntry(app, textvariable=docket, width=150)
docket_entry.place(x=153, y=155)

place_entry = customtkinter.CTkEntry(app, textvariable=place, width=150)
place_entry.place(x=153, y=215)

weight_entry = customtkinter.CTkEntry(app, textvariable=weight, width=150)
weight_entry.place(x=153, y=275)

rate_entry = customtkinter.CTkEntry(app, textvariable=rate, width=150)
rate_entry.place(x=153, y=335)

payment_checkbox = customtkinter.CTkCheckBox(app, text="Bank Account", variable=payment_checkbox, font=font2)
payment_checkbox.place(x=170, y=395)

calculate_button = customtkinter.CTkButton(app, command=lambda: rate.set(calculate_rate(place.get(), weight.get())), text="Calculate")
calculate_button.place(x=50, y=450)

clear_button = customtkinter.CTkButton(app, command=clear, text="Clear")
clear_button.place(x=200, y=500)

submit_button = customtkinter.CTkButton(app, command=submit, text="Submit")
submit_button.place(x=50, y=500)

app.mainloop()
