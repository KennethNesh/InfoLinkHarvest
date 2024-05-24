import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from docx import Document
import os

def search_and_extract_links(name, specific_search):
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--ignore-certificate-errors')
    with webdriver.Chrome(service=Service(quiet=True), options=chrome_options) as driver:
        driver.get("https://www.google.com")
        search_box = driver.find_element(By.NAME, "q")
        search_term = f"{name} {specific_search}"
        search_box.send_keys(search_term)
        search_box.send_keys(Keys.RETURN)
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, '.tF2Cxc a')))
        except Exception as e:
            print("Error waiting for search results:", str(e))
        search_results = driver.find_elements(By.CSS_SELECTOR, '.tF2Cxc a')
        links = [result.get_attribute('href') for result in search_results]
    return links

def save_links_to_doc(links, filename):
    doc = Document()
    for link in links:
        doc.add_paragraph(link)
    doc.save(filename)

def search_and_save():
    name = name_entry.get()
    specific_search = specific_entry.get()
    folder_path = filedialog.askdirectory()
    if folder_path:
        try:
            links = search_and_extract_links(name, specific_search)
            filename = f"{name}_{specific_search}_search_results.docx"
            filepath = os.path.join(folder_path, filename)
            save_links_to_doc(links, filepath)
            messagebox.showinfo("Success", f"Search results saved to {filepath}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    else:
        messagebox.showinfo("Info", "No folder selected")

root = tk.Tk()
root.geometry("500x300")
root.title("InfoLinkHarvest")
root.configure(bg="#F0F0F0")

name_label = tk.Label(root, text="Enter a name to search:", font=("Arial", 16), bg="#F0F0F0")
name_label.pack(pady=20)
name_entry = tk.Entry(root, width=30, font=("Arial", 14), bg="#E0E0E0")
name_entry.pack()

specific_label = tk.Label(root, text="Enter specific search term:", font=("Arial", 16), bg="#F0F0F0")
specific_label.pack(pady=20)
specific_entry = tk.Entry(root, width=30, font=("Arial", 14), bg="#E0E0E0")
specific_entry.pack()

search_button = tk.Button(root, text="Search and Save", command=search_and_save, font=("Arial", 18), bg="#007BFF", fg="white")
search_button.pack(pady=20)

root.mainloop()