from datetime import datetime
from tkinter import *
from tkinter import messagebox
import time
import datetime
from selenium import webdriver
import pandas as pd
import pathlib

# The line below will automatically fetch the current working path
PARENT_PATH = str(pathlib.Path(__file__).parent.resolve())

# The Line below will create the GUI for Login and other tasks
w = Tk()
w.geometry('350x500')
w.title('MMOB-EXTRACTER')
w.resizable(0, 0)

# Making gradient frame
j = 0
r = 10
for i in range(100):
    c = str(222222 + r)
    Frame(w, width=10, height=500, bg="#" + c).place(x=j, y=0)
    j = j + 10
    r = r + 1

Frame(w, width=250, height=400, bg='white').place(x=50, y=50)

l1 = Label(w, text='Username', bg='white')
l = ('Consolas', 13)
l1.config(font=l)
l1.place(x=80, y=200)

# e1 entry for username entry
e1 = Entry(w, width=20, border=0)
l = ('Consolas', 13)
e1.config(font=l)
e1.place(x=80, y=230)

# e2 entry for password entry
e2 = Entry(w, width=20, border=0, show='*')
e2.config(font=l)
e2.place(x=80, y=310)

l2 = Label(w, text='Password', bg='white')
l = ('Consolas', 13)
l2.config(font=l)
l2.place(x=80, y=280)

###lineframe on entry

Frame(w, width=180, height=2, bg='#141414').place(x=80, y=332)
Frame(w, width=180, height=2, bg='#141414').place(x=80, y=252)

from PIL import ImageTk, Image

imagea = Image.open("log.PNG")
imageb = ImageTk.PhotoImage(imagea)

label1 = Label(image=imageb,
               border=0,

               justify=CENTER)

label1.place(x=115, y=50)


# Command Funtion is to recognize the Login and password and to perform further simulations
def cmd():
    if e1.get() == '0' and e2.get() == '0':
        def cmd_for_robot():
            current_date_and_time = datetime.datetime.now()
            current_date_and_time_string = str(current_date_and_time)
            current_date_and_time_string = current_date_and_time_string.replace(':', '')
            current_date_and_time_string = current_date_and_time_string.replace('-', '')
            current_date_and_time_string = current_date_and_time_string.replace('.', '')
            current_date_and_time_string = current_date_and_time_string.replace(' ', '')
            current_date_and_time_string = current_date_and_time_string.replace('_', '')

            extension = ".xls"

            file_name = 'MMOB_EXTRACT' + current_date_and_time_string[0:18] + extension

            a = int(e3.get())
            b = int(e4.get())
            if a >= int(0) and b >= int(0):
                url = 'http://ola.salt.ch'
                chrome_options = webdriver.ChromeOptions()
                chrome_options.add_argument('headless')
                chrome_options.add_argument('window-size=1920x1080')
                chrome_options.add_argument("disable-gpu")
                driver = webdriver.Chrome(PARENT_PATH + '/chromedriver', options=chrome_options)

                driver.get(url)
                f = open(file_name, "a+")
                f.write(
                    'Name' + '\t' + 'Date Of Birth' + '\t' + 'Address' + '\t' + 'Email' + '\t' + 'Color' + '\t\t' + 'Number' + '\t' + 'Subscription Type' + '\n')
                f.close()

                # The function below will automatically check the record line by line and will store in file
                def check_record(phonenumer):
                    driver.find_element_by_name('msisdn').clear()
                    driver.find_element_by_name('msisdn').send_keys(phonenumer)
                    driver.find_element_by_id("_search").click()
                    try:
                        time.sleep(1)
                        a = driver.find_element_by_xpath('//*[@id="top"]/div[2]/div/div/div/div/div[1]/div[4]').text
                        time.sleep(1)

                        backColor = driver.find_element_by_xpath("//*[@id=\"tab-content\"]/div/div[1]/div[2]/span")
                        b = backColor.value_of_css_property("background-color")
                        time.sleep(1)
                        new_var = driver.find_element_by_xpath('//*[@id=\"tab-content\"]/div/div[2]/div[2]').text

                        temp = a.split("\n")
                        time.sleep(1)
                        new_var = new_var.replace('\n', '')

                        num = str(phonenumer)
                        f = open(file_name, "a+")
                        f.write(temp[1] + '\t' + temp[3] + '\t' + temp[5] + temp[6] + '\t' + temp[
                            8] + '\t' + b + '\t\t' + num + '\t' + new_var + '\n')
                        f.close()
                        time.sleep(1)
                        driver.find_element_by_xpath(
                            '//*[@id="top"]/div[2]/div/div/div/div/div[1]/div[6]/div/a').click()
                        time.sleep(1)
                        driver.find_element_by_id("existing-residential-button").click()
                        return True
                    except:
                        f = open('No_Record' + file_name, "a+")
                        num = str(phonenumer)
                        f.write(num)
                        f.write('\n')
                        f.close()
                        return

                driver.find_element_by_name('username').send_keys('jsakthik')
                driver.find_element_by_name('password').send_keys('Inno-tech2022')
                driver.find_element_by_xpath("//*[@id='grp-content-container']/div/form/div/ul/li/input").click()
                driver.find_element_by_class_name('home-link').click()
                iframe = driver.find_element_by_id('external_tools_iframe')
                driver.switch_to.frame(iframe)
                driver.find_element_by_id("existing-customer").click()
                driver.find_element_by_id("existing-residential-button").click()

                df = pd.read_excel('homosapian.xlsx')
                df = df[a:a+b]
                df = df['telephone']
                count = 1
                start = a
                for i in df:
                    print(count,"/",start+b)
                    count = count + 1
                    check_record(i)

                messagebox.showinfo("Robot Has Been Completed Search SUCCESSFULLY", "         W E L C O M E        ")

        j = 0
        r = 10
        for i in range(100):
            c = str(222222 + r)
            Frame(w, width=10, height=500, bg="#" + c).place(x=j, y=0)
            j = j + 10
            r = r + 1

        Frame(w, width=250, height=400, bg='white').place(x=50, y=50)

        l1 = Label(w, text='Start', bg='white')
        l = ('Consolas', 13)
        l1.config(font=l)
        l1.place(x=80, y=200)

        # e3 entry for Start robot
        e3 = Entry(w, width=20, border=0)
        l = ('Consolas', 13)
        e3.config(font=l)
        e3.place(x=80, y=230)

        # e4 entry for Total entry
        e4 = Entry(w, width=20, border=0)
        e4.config(font=l)
        e4.place(x=80, y=310)

        l2 = Label(w, text='Total', bg='white')
        l = ('Consolas', 13)
        l2.config(font=l)
        l2.place(x=80, y=280)

        ###lineframe on entry

        Frame(w, width=180, height=2, bg='#141414').place(x=80, y=332)
        Frame(w, width=180, height=2, bg='#141414').place(x=80, y=252)

        def bttn(x, y, text, ecolor, lcolor):
            def on_entera(e):
                myButton1['background'] = ecolor  # ffcc66
                myButton1['foreground'] = lcolor  # 000d33

            def on_leavea(e):
                myButton1['background'] = lcolor
                myButton1['foreground'] = ecolor

            myButton1 = Button(w, text=text,
                               width=20,
                               height=2,
                               fg=ecolor,
                               border=0,
                               bg=lcolor,
                               activeforeground=lcolor,
                               activebackground=ecolor,
                               command=cmd_for_robot)

            myButton1.bind("<Enter>", on_entera)
            myButton1.bind("<Leave>", on_leavea)

            myButton1.place(x=x, y=y)

        bttn(100, 375, 'Start', 'white', '#994422')
        # q.mainloop()

    else:
        messagebox.showwarning("LOGIN FAILED", "        PLEASE TRY AGAIN        ")


# Button_with hover effect
def bttn(x, y, text, ecolor, lcolor):
    def on_entera(e):
        myButton1['background'] = ecolor  # ffcc66
        myButton1['foreground'] = lcolor  # 000d33

    def on_leavea(e):
        myButton1['background'] = lcolor
        myButton1['foreground'] = ecolor

    myButton1 = Button(w, text=text,
                       width=20,
                       height=2,
                       fg=ecolor,
                       border=0,
                       bg=lcolor,
                       activeforeground=lcolor,
                       activebackground=ecolor,
                       command=cmd)

    myButton1.bind("<Enter>", on_entera)
    myButton1.bind("<Leave>", on_leavea)

    myButton1.place(x=x, y=y)


bttn(100, 375, 'L O G I N', 'white', '#994422')

w.mainloop()
