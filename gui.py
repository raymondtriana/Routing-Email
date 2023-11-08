# helloworld.py
import pathlib
import tkinter as tk
import tkinter.ttk as ttk
import pygubu

PROJECT_PATH = pathlib.Path(__file__).parent
PROJECT_UI = PROJECT_PATH / "gui.ui"

splitter_label: any
splitter_path: any
splitter_button: any

email_scrolledFrame: any
email_button: any


class HelloworldApp:
    def __init__(self, master=None):
        # 1: Create a builder and setup resources path (if you have images)
        self.builder = builder = pygubu.Builder()
        builder.add_resource_path(PROJECT_PATH)

        # 2: Load a ui file
        builder.add_from_file(PROJECT_UI)

        # 3: Create the mainwindow
        self.mainwindow = builder.get_object('mainwindow', master)

        builder.import_variables(self, ['pdf_file_path'])

        # 4: Connect callbacks
        builder.connect_callbacks(self)

        # 5: Cache pointers to elements
        splitter_label = builder.get_object('splitter_label', master)
        splitter_path = builder.get_object('splitter_path', master)
        splitter_button = builder.get_object('splitter_button', master)

        email_scrolledFrame = builder.get_object('email_scrolledFrame', master)
        email_button = builder.get_object('email_button', master)

        label_text = splitter_label.cget('text')
        print(label_text)

        button = tk.Button(email_scrolledFrame, text="Left")
        button.pack(side=tk.LEFT)
        button.bind('<ButtonRelease>', self.on_email_button_release)

    def run(self):
        self.mainwindow.mainloop()

    ##
    # SPLITTER CALLBACKS
    ##
    def on_pdf_path_changed(self, msg: tk):
        print(msg.widget.cget("path"))
        print(self.builder.tkvariables['pdf_file_path'].get())
        pass

    ##
    # EMAIL CALLBACKS
    ##
    def on_email_button_release(self, msg: tk):
        print(msg.widget.cget("text"))
        print("EMAIL RELEASED")


if __name__ == '__main__':
    app = HelloworldApp()
    app.run()


def email():
    print("EMAIL")
