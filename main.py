from tkinter import *
from tkinter.ttk import Progressbar
from tkinter import filedialog
from tkinter import messagebox
from datetime import datetime, date
import time
import os
import logging
import configparser
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from platform import python_version
import imaplib
from email import encoders


class KvitNotFoundError(Exception):
    pass


class Sender:
    def __init__(self):
        self.root = Tk()
        self.root.title("Программа рассылки квитанций по списку из файла")
        self.root.geometry('580x75')
        self.file_dir = "kvits/"
        self.container = Frame(relief=RIDGE, borderwidth=3)
        self.container.pack()
        self.lbl_dir = Label(master=self.container, text="Выберите каталог с квитанциями: ")
        self.ent_dir = Entry(master=self.container, width=50)
        self.btn_dir = Button(master=self.container, text="Выбрать", command=self.select_dir)
        self.lbl_dir.grid(row=0, column=0, sticky="e")
        self.ent_dir.grid(row=0, column=1)
        self.btn_dir.grid(row=0, column=2)
        self.container_act = Frame(relief=RIDGE, borderwidth=3)
        self.container_act.pack()
        self.bar = Progressbar(master=self.container_act, length=490)
        self.bar.grid(row=0, column=0)
        self.btn_snd = Button(master=self.container_act, text="Отправить", command=self.send_files)
        self.btn_snd.grid(row=0, column=1)
        self.config = configparser.ConfigParser()
        self.config.read("settings.ini")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.mail = None
        self.imap = None

    def on_closing(self):
        if messagebox.askokcancel("Выход", "Выйти из программы?"):
            self.root.destroy()

    def select_dir(self):
        """Функция выбора папки с квитанциями"""
        self.file_dir = filedialog.askdirectory(initialdir=self.file_dir)
        self.ent_dir.delete(0, END)
        self.ent_dir.insert(0, self.file_dir)
        self.bar['value'] = 0
        self.container_act.update()

    def send_files(self):
        """Функция проверяет наличие всех файлов для отправки и в случае успеха, отправляет их на e_mail"""
        status = True
        available_files = True
        file_list = None
        files_in_list = None
        self.bar['value'] = 0
        self.container_act.update()

        smtp_server = self.config["sender_files"]["smtp_server"]
        smtp_port = int(self.config["sender_files"]["smtp_port"])
        imap_server = self.config["sender_files"]["imap_server"]
        imap_port = int(self.config["sender_files"]["imap_port"])
        user = self.config["sender_files"]["user"]
        password = self.config["sender_files"]["password"]

        logging.basicConfig(level=logging.DEBUG, filename="log_file.log",
                            format="%(asctime)s %(levelname)s:%(message)s")

        try:
            file_list = open(self.ent_dir.get() + "/files_for_send.txt")
            date_from_file = datetime.strptime(file_list.readline().split()[1], "%Y-%m-%d").date()
            if date_from_file != date.today():
                status = False
                logging.error("Не корректная дата выгрузки!")
                messagebox.showerror("Ошибка", "Не корректная дата выгрузки")

            if status:
                files_in_list = [row.split()[0] for row in file_list.readlines()]
                for file in files_in_list:
                    if file not in os.listdir(self.ent_dir.get()):
                        available_files = False
                        logging.error(f"Файл {file} не найден")

                if not available_files:
                    raise KvitNotFoundError

        except FileNotFoundError as e:
            print(e)
            logging.error("Не найден список файлов для отправки")
            messagebox.showerror("Ошибка", "Не найден список файлов для отправки")
        except KvitNotFoundError as e:
            print(e)
            logging.error("Найдены не все квитанции")
            messagebox.showerror("Ошибка", "Найдены не все квитанции")
        except EXCEPTION as e:
            print(e)
            logging.error(e)
            messagebox.showerror(e)
        else:
            try:
                self.mail = smtplib.SMTP_SSL(smtp_server, smtp_port)
                self.mail.login(user, password)
                self.imap = imaplib.IMAP4_SSL(imap_server, imap_port)
                self.imap.login(user, password)
            except smtplib.SMTPException as e:
                logging.error(f"Не корректные настройки почты{e}")
                messagebox.showerror("Не корректные настройки почты")
            except EXCEPTION as e:
                logging.error(f"Не корректные настройки почты{e}")
                messagebox.showerror("Не корректные настройки почты")
            else:
                percent = 100 / len(files_in_list)
                file_list.seek(0)
                file_list.readline()
                for line in file_list.readlines():
                    file, mail = line.split()
                    try:
                        self.send_file(f"{self.ent_dir.get()}/{file}", mail)
                    except smtplib.SMTPException as e:
                        print(e)
                        logging.error(e)
                    except EXCEPTION as e:
                        print(e)
                        logging.error(e)
                    finally:
                        self.bar['value'] += percent
                        self.container_act.update()
            finally:
                if self.imap:
                    self.imap.logout()
                if self.mail:
                    self.mail.quit()
        finally:
            if file_list:
                file_list.close()

    def send_file(self, filename: str, recipient: str):
        """Функция отправки письма с вложением"""
        msg = MIMEMultipart('alternative')
        msg['Subject'] = self.config["sender_files"]["subject"]
        msg['From'] = self.config["sender_files"]["sender"]
        msg['To'] = recipient
        msg['Reply-To'] = self.config["sender_files"]["sender"]
        msg['Return-Path'] = self.config["sender_files"]["sender"]
        msg['X-Mailer'] = 'Python/' + (python_version())

        text = self.config["sender_files"]["text"]
        html = '<html><head></head><body><p>' + text + '</p></body></html>'
        basename = os.path.basename(filename)
        filesize = os.path.getsize(filename)

        part_text = MIMEText(text, 'plain')
        part_html = MIMEText(html, 'html')
        part_file = MIMEBase('application', 'octet-stream; name="{}"'.format(basename))
        part_file.set_payload(open(filename, "rb").read())
        part_file.add_header('Content-Description', basename)
        part_file.add_header('Content-Disposition', 'attachment; filename="{}"; size={}'.format(basename, filesize))
        encoders.encode_base64(part_file)

        msg.attach(part_text)
        msg.attach(part_html)
        msg.attach(part_file)

        try:
            self.mail.sendmail(self.config["sender_files"]["sender"], recipient, msg.as_string())
            print(f"Квитанция {filename} отправленна на почту: {recipient}")
            logging.info(f"Квитанция {filename} отправленна на почту: {recipient}")
            seen = '\\SEEN'
            directory = '"&BB4EQgQ,BEAEMAQyBDsENQQ9BD0ESwQ1-"'
        except smtplib.SMTPException as e:
            print(e)
            logging.error(f"Квитаниция не доставлена, проверьте адрес электроной почты {e}")
            seen = '\\UNSEEN'
            directory = 'INBOX'
            msg['Subject'] = "Ваше сообщение не доставлено!"

        self.imap.append(directory, seen, imaplib.Time2Internaldate(time.time()),
                         msg.as_string().encode('utf8'))

    def run(self):
        self.root.mainloop()


def main():
    application = Sender()
    application.run()


if __name__ == "__main__":
    main()
