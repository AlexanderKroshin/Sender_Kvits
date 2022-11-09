import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from tkinter import *
from tkinter import filedialog
from tkinter.ttk import Progressbar
from tkinter import messagebox
import configparser
from datetime import date, datetime
import os


config = configparser.ConfigParser()
config.read("settings.ini")


class IncorrectDateError(Exception):
    pass


class KvitNotFoundError(Exception):
    pass


def select_dir():
    """Функция выбора папки с квитанциями"""
    file_dir = filedialog.askdirectory(initialdir="kvits/")
    ent_dir.delete(0, END)
    ent_dir.insert(0, file_dir)


def test_dir():
    """Функция проверки квитанций"""
    try:
        lst = open(ent_dir.get() + "/kvits_for_sending.txt")
        log = open(ent_dir.get() + "/log.txt", "w")
        kvits_in_dir = os.listdir(ent_dir.get())

        date_from_file = datetime.strptime(lst.readline().split()[1], "%Y-%m-%d").date()

        if date_from_file != date.today():
            raise IncorrectDateError

        data = lst.readlines()
        kvits_in_lst = [row.split()[0] for row in data]

        for kvit in kvits_in_lst:
            if kvit not in kvits_in_dir:
                log.write(f"Квитанция {kvit} не найдена")
                raise KvitNotFoundError

    except FileNotFoundError:
        log.write("Файл списка квитанций не найден")
        messagebox.showerror("Ошибка", "Файл списка квитанций не найден")
    except KvitNotFoundError:
        log.write("Найденны не все квитанции")
        messagebox.showerror("Ошибка", "Найденны не все квитанции")
    except IncorrectDateError:
        log.write("Не корректная дата выгрузки!")
        messagebox.showerror("Ошибка", "Не корректная дата выгрузки")
    except:
        log.write("Непредвиденная ошибка")
        messagebox.showerror("Ошибка", "Непредвиденная ошибка")
    else:
        percent = 100/len(data)
        for row in data:
            kvit, mail = row.split()
            send_kvit(f"{ent_dir.get()}/{kvit}", mail)
            log.write(f"Квитанция {ent_dir.get()}/{kvit} отправленна на почту: {mail}\n")
            bar['value'] += percent
            act_form.update()
    finally:
        lst.close()
        log.close()


def send_kvit(filename: str, to: str):
    """Функция отправки письма с вложением"""
    user = config["Sender_kvits"]["mail_from"]
    passwd = config["Sender_kvits"]["password"]
    server = config["Sender_kvits"]["smtp_server"]
    port = int(config["Sender_kvits"]["smtp_port"])
    subject = config["Sender_kvits"]["subject"]
    print(user, passwd, subject, server)

    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = user
    msg['To'] = to

    try:
        pdf = MIMEApplication(open(filename, 'rb').read())
        pdf.add_header('Content-Disposition', 'attachment; filename="%s"' % filename)
        msg.attach(pdf)

        smtp = smtplib.SMTP(server, port)
        smtp.starttls()
        smtp.ehlo()
        smtp.login(user, passwd)
        smtp.sendmail(user, to, msg.as_string())

    except FileNotFoundError:
        print(f"Квитанция {filename} не найдена")
    except smtplib.SMTPException as err:
        print('Ошибка отправки письма')
        raise err
    except:
        print('Непредвиденная ошибка отправки письма')
    finally:
        smtp.quit()


window = Tk()
window.title("Программа рассылки квитанций по списку из файла")
window.geometry('580x75')

frm_form = Frame(relief=RIDGE, borderwidth=3)
frm_form.pack()

lbl_dir = Label(master=frm_form, text="Выберите каталог с квитанциями: ")
ent_dir = Entry(master=frm_form, width=50)
btn_dir = Button(master=frm_form, text="Выбрать", command=select_dir)

lbl_dir.grid(row=0, column=0, sticky="e")
ent_dir.grid(row=0, column=1)
btn_dir.grid(row=0, column=2)

act_form = Frame(relief=RIDGE, borderwidth=3)
act_form.pack()

bar = Progressbar(master=act_form, length=490)
bar.grid(row=0, column=0)

btn_tst = Button(master=act_form, text="Отправить", command=test_dir)
btn_tst.grid(row=0, column=1)


window.mainloop()
