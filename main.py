from bs4 import BeautifulSoup
import requests
import pandas as pd
from pathlib import Path
import win32com.client as win32

URL = "https://books.toscrape.com"
EXCEL_FILE = "book.xlsx"

def get_data(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.text, "lxml")
    books = soup.find_all("article", class_="product_pod")

    data=[]

    for book in books:
        item={}
        item["Title"] = book.find("img", class_="thumbnail").attrs["alt"]
        item["Price"] = book.find("p", class_="price_color").text[1:]
        data.append(item)
    return data

def export_data(data):
    df = pd.DataFrame(data)
    df.to_excel(EXCEL_FILE)
    df.to_csv("books.csv")

def create_email():
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = "mpatel@cortado.ventures"
    mail.CC = "nharding@cortado.ventures"
    mail.Subject = f"Cortado Ventures Interview Take-Home Project"
    mail.HTMLBody = f""" 
                    Hi Mensi!<br><br>
                    Here is the data of books you requested!<br><br>
                    Best Regards, <br>
                    <b>Joshua Yustana</b>
                    """
    attachment_path = str(Path.cwd() / EXCEL_FILE)
    mail.Attachments.Add(Source=attachment_path)
    mail.Display()

if __name__ == '__main__':
    data = get_data(URL)
    export_data(data)
    create_email()
    print('Done')