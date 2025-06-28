import os.path
import datetime
import pickle

import tkinter as tk
from tkinter import simpledialog
import cv2
from PIL import Image, ImageTk
import face_recognition
import openpyxl

import util
from test import test


class App:
    def __init__(self):
        self.admin_password = "admin123"  # admin password

        self.main_window = tk.Tk()
        self.main_window.title("Attendance System")
        self.main_window.attributes('-fullscreen', True)
        self.main_window.bind("<Escape>", self.request_exit_full_screen)

        self.login_button_main_window = util.get_button(self.main_window, 'login', 'green', self.login)
        self.login_button_main_window.place(x=750, y=200)

        self.logout_button_main_window = util.get_button(self.main_window, 'logout', 'red', self.logout)
        self.logout_button_main_window.place(x=750, y=300)

        self.register_new_user_button_main_window = util.get_button(self.main_window, 'register new user', 'gray',
                                                                    self.request_admin_password, fg='black')
        self.register_new_user_button_main_window.place(x=750, y=400)

        self.webcam_label = util.get_img_label(self.main_window)
        self.webcam_label.place(x=10, y=0, width=700, height=500)

        self.add_webcam(self.webcam_label)

        self.db_dir = './db'
        if not os.path.exists(self.db_dir):
            os.mkdir(self.db_dir)

        self.log_path = './Attendance Log.xlsx'
        self.initialize_log()

    def initialize_log(self):
        if not os.path.exists(self.log_path):
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Attendance Log"
            sheet.append(["Name", "Timestamp", "Action"])
            workbook.save(self.log_path)

    def add_webcam(self, label):
        if 'cap' not in self.__dict__:
            self.cap = cv2.VideoCapture(0)

        self._label = label
        self.process_webcam()

    def process_webcam(self):
        ret, frame = self.cap.read()

        self.most_recent_capture_arr = frame
        img_ = cv2.cvtColor(self.most_recent_capture_arr, cv2.COLOR_BGR2RGB)
        self.most_recent_capture_pil = Image.fromarray(img_)
        imgtk = ImageTk.PhotoImage(image=self.most_recent_capture_pil)
        self._label.imgtk = imgtk
        self._label.configure(image=imgtk)

        self._label.after(20, self.process_webcam)

    def login(self):
        label = test(
            image=self.most_recent_capture_arr,
            model_dir='/home/slman/face-attendance-system/Silent-Face-Anti-Spoofing/resources/anti_spoof_models',
            device_id=0
        )

        if label == 1:
            name = util.recognize(self.most_recent_capture_arr, self.db_dir)

            if name in ['unknown_person', 'no_persons_found']:
                util.msg_box('Ups...', 'Unknown user. Please register new user or try again.')
            else:
                util.msg_box('Welcome back !', 'Welcome, {}.'.format(name))
                self.log_action(name, "in")

        else:
            util.msg_box('Spoofing Alert!', 'You are fake!')

    def logout(self):
        label = test(
            image=self.most_recent_capture_arr,
            model_dir='/home/slman/face-attendance-system/Silent-Face-Anti-Spoofing/resources/anti_spoof_models',
            device_id=0
        )

        if label == 1:
            name = util.recognize(self.most_recent_capture_arr, self.db_dir)

            if name in ['unknown_person', 'no_persons_found']:
                util.msg_box(' ', 'Unknown user. Please register new user or try again.')
            else:
                util.msg_box(' ', 'Goodbye, {}.'.format(name))
                self.log_action(name, "out")

        else:
            util.msg_box('Spoofing Alert!', 'You are fake!')

    def log_action(self, name, action):
        workbook = openpyxl.load_workbook(self.log_path)
        sheet = workbook.active
        sheet.append([name, datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'), action])
        workbook.save(self.log_path)

    def request_admin_password(self):
        password = simpledialog.askstring("Admin Password", "Enter admin password:", show='*')
        if password == self.admin_password:
            self.register_new_user()
        else:
            util.msg_box("Error", "Incorrect admin password!")

    def request_exit_full_screen(self, event=None):
        password = simpledialog.askstring("Admin Password", "Enter admin password to exit the Attendance System:", show='*')
        if password == self.admin_password:
            self.exit_full_screen()
        else:
            util.msg_box("Error", "Incorrect admin password!")

    def exit_full_screen(self):
        self.main_window.attributes('-fullscreen', False)

    def register_new_user(self):
        self.register_new_user_window = tk.Toplevel(self.main_window)
        self.register_new_user_window.geometry("1200x520+370+120")

        self.accept_button_register_new_user_window = util.get_button(self.register_new_user_window, 'Accept', 'green',
                                                                      self.accept_register_new_user)
        self.accept_button_register_new_user_window.place(x=750, y=300)

        self.try_again_button_register_new_user_window = util.get_button(self.register_new_user_window, 'Try again',
                                                                         'red', self.try_again_register_new_user)
        self.try_again_button_register_new_user_window.place(x=750, y=400)

        self.capture_label = util.get_img_label(self.register_new_user_window)
        self.capture_label.place(x=10, y=0, width=700, height=500)

        self.add_img_to_label(self.capture_label)

        self.entry_text_register_new_user = util.get_entry_text(self.register_new_user_window)
        self.entry_text_register_new_user.place(x=750, y=150)

        self.text_label_register_new_user = util.get_text_label(self.register_new_user_window, 'Please, \ninput username:')
        self.text_label_register_new_user.place(x=750, y=70)

    def try_again_register_new_user(self):
        self.register_new_user_window.destroy()

    def add_img_to_label(self, label):
        imgtk = ImageTk.PhotoImage(image=self.most_recent_capture_pil)
        label.imgtk = imgtk
        label.configure(image=imgtk)

        self.register_new_user_capture = self.most_recent_capture_arr.copy()

    def start(self):
        self.main_window.mainloop()

    def accept_register_new_user(self):
        name = self.entry_text_register_new_user.get(1.0, "end-1c")

        embeddings = face_recognition.face_encodings(self.register_new_user_capture)[0]

        file = open(os.path.join(self.db_dir, '{}.pickle'.format(name)), 'wb')
        pickle.dump(embeddings, file)

        util.msg_box('Success!', 'User was registered successfully!')

        self.register_new_user_window.destroy()


if __name__ == "__main__":
    app = App()
    app.start()