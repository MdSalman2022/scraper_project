import csv
from tkinter import *
from tkinter import ttk
from PIL import Image, ImageTk
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import Chrome, ChromeOptions
import pandas as pd
import os
import webbrowser
import platform

root = Tk()
root.title("Web Scraper")
root.geometry("500x600")


def find():
    url = entry_url.get()
    category = entry_category.get().strip()
    pages = entry_pages.get().strip()
    save_filename = entry_filename.get()

    if not url or not category or not pages.isdigit() or not save_filename:
        status_label.config(text="Please enter URL, category, number of pages, and filename.", fg="red")
        return

    pages = int(pages)
    browser_options = ChromeOptions()
    browser_options.headless = True

    try:
        driver = Chrome(options=browser_options)
        driver.get(url)

        category_link = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.LINK_TEXT, category))
        )
        category_link.click()

        data = []
        current_page = 1

        while current_page <= pages:
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".product_pod"))
            )
            
            books = driver.find_elements(By.CSS_SELECTOR, ".product_pod")
            for book in books:
                title = book.find_element(By.CSS_SELECTOR, "h3 > a")
                price = book.find_element(By.CSS_SELECTOR, ".price_color")
                stock = book.find_element(By.CSS_SELECTOR, ".instock.availability")
                book_item = {
                    'Title': title.get_attribute("title"),
                    'Price': price.text,
                    'Stock': stock.text.strip()
                }
                data.append(book_item)

            try:
                next_button = driver.find_element(By.CSS_SELECTOR, ".next > a")
                next_button.click()
                current_page += 1
            except:
                break

        driver.quit()

        file_directory = "./files" 
        filename = f"{save_filename}.xlsx"
        save_path = os.path.join(file_directory, filename)
        
        index = 1
        while os.path.exists(save_path):
            filename = f"{save_filename}({index}).xlsx"
            save_path = os.path.join(file_directory, filename)
            index += 1

        df = pd.DataFrame(data)
        df.to_excel(save_path, index=False)

        status_label.config(text=f"Data saved to {filename}", fg="green")

        file_label.config(text="Here is the scrapped data :")
        open_button.config(state=NORMAL)
    except Exception as e:
        status_label.config(text=f"Error: {e}", fg="red")


def open_file_explorer():
    file_directory = "files"
    filename = f"{entry_filename.get()}.xlsx"
    file_path = os.path.join(file_directory, filename)
    if os.path.exists(file_path):
        os.startfile(file_directory)

        if platform.system() == "Windows":
            os.system(f"powershell.exe -command \"Select-Object -Path '{file_path}' | Out-Null\"")
        elif platform.system() == "Linux":
            os.system(f"xdg-open '{file_directory}'")
        elif platform.system() == "Darwin":
            os.system(f"open '{file_directory}'")
    else:
        status_label.config(text="File not found.", fg="red")



# UI Elements
welcome_frame = Frame(root)
welcome_frame.pack(pady=20)

logo_path = os.path.join(os.path.dirname(__file__), "./logo.png")
logo_image = Image.open(logo_path) 
logo_image = logo_image.resize((100, 100))

logo_photo = ImageTk.PhotoImage(logo_image)

logo_label = Label(welcome_frame, image=logo_photo)
logo_label.pack()

welcome_message = Label(welcome_frame, text="Welcome to Web Scraper", font=('arial', 16))
welcome_message.pack()

input_frame = Frame(root)
input_frame.pack(pady=20)
style = ttk.Style()
style.configure('TEntry', font=('Arial', 12), padding=5, relief='solid')

label_url = Label(input_frame, text="Enter Correct URL:", font=('arial', 14))
label_url.grid(row=0, column=0, padx=10, pady=10, sticky=E)

entry_url = ttk.Entry(input_frame, width=40, style='TEntry')
entry_url.grid(row=0, column=1, padx=10, pady=10)

label_category = Label(input_frame, text="Enter Category:", font=('arial', 14))
label_category.grid(row=1, column=0, padx=10, pady=10, sticky=E)

entry_category = ttk.Entry(input_frame, width=40, style='TEntry')
entry_category.grid(row=1, column=1, padx=10, pady=10)

label_pages = Label(input_frame, text="Number of Pages:", font=('arial', 14))
label_pages.grid(row=2, column=0, padx=10, pady=10, sticky=E)

entry_pages = ttk.Entry(input_frame, width=40, style='TEntry')
entry_pages.grid(row=2, column=1, padx=10, pady=10)

label_filename = Label(input_frame, text="Save filename:", font=('arial', 14))
label_filename.grid(row=3, column=0, padx=10, pady=10, sticky=E)

entry_filename = ttk.Entry(input_frame, width=40, style='TEntry')
entry_filename.grid(row=3, column=1, padx=10, pady=10)

button = Button(input_frame, text="Scrape Data", command=find, font=('arial', 14), bg='#4CAF50', fg='white', relief='raised')
button.grid(row=4, columnspan=2, pady=20)

status_label = Label(root, text="", font=('arial', 12))
status_label.pack()

file_frame = Frame(root)
file_frame.pack(pady=10)

file_label = Label(file_frame, text="", font=('arial', 12))
file_label.pack(side=LEFT, padx=(10, 20))

open_button = Button(file_frame, 
                     text="Open File Location", 
                     command=open_file_explorer, 
                     font=('arial', 12), 
                     state=DISABLED, 
                     bg='#007bff', 
                     fg='white',   
                     relief='raised', 
                     padx=10,      
                     pady=5,       
                     borderwidth=2, 
                     cursor='hand2')

open_button.pack(side=LEFT)

root.mainloop()
