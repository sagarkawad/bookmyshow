from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
import nltk
import tkinter
import customtkinter
from nltk.sentiment import SentimentIntensityAnalyzer


def backend(user_link):
    browser = webdriver.Chrome()
    browser.get(f"{user_link}")

    elements = browser.find_elements(By.CLASS_NAME, "sc-r6zm4d-14")

    wb = openpyxl.load_workbook("sheets/movie-data.xlsx")
    sid = SentimentIntensityAnalyzer()
    print(wb.sheetnames)

    sheet = wb["Sheet1"]

    # variables
    sum = 0

    for i, element in enumerate(elements):
        text = element.text
        score = sid.polarity_scores(text)
        compound_score = score["compound"]
        percentage = (compound_score + 1) * 50
        sum = sum + percentage
        print(
            f"this is the review  {i + 1}  with a score of {percentage}% - {text}  \n"
        )
        sheet[f"a{i+1}"] = text
        sheet[f"b{i+1}"] = percentage

    avg = sum / len(elements)
    print(avg)
    sheet.append(["Final Conclusion", avg])

    browser.quit()
    # extract 38th letter from the link_string
    m_name = ""
    for i, e in enumerate(user_link):
        if i >= 38 and e != "/":
            m_name = m_name + e
        elif i > 38 and e == "/":
            break
    print("m-name : ", m_name)
    wb.save(f"sheets/{m_name}.xlsx")
    # https://in.bookmyshow.com/pune/movies/the-little-mermaid/ET00058086/user-reviews


# ui-functions
def run():
    try:
        ip = link.get()
        print("ran : ", ip)
        backend(ip)
        print("It worked!")
    except:
        print("error!")


# system-settings
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("blue")

# our app frame
app = customtkinter.CTk()
app.geometry("720x480")
app.title("review sentiment reader!")

# adding ui elements
title = customtkinter.CTkLabel(app, text="movie reviews link here")
title.pack(padx=10, pady=10)


# link input
movie_var = tkinter.StringVar()
link = customtkinter.CTkEntry(app, width=350, height=30, textvariable=movie_var)
link.pack()


# Run Button
download = customtkinter.CTkButton(app, text="Run", command=run)
download.pack(padx=10, pady=10)

# run app
app.mainloop()
