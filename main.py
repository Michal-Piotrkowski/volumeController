from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
import tkinter as tk
from tkinter import *
from PIL import Image, ImageTk
from VolumeManager import VolumeManager


#####################
#   CONSOLE APP
####################
class App():
    def __init__(self):
        super(App, self).__init__()
        self.width = None
        self.height = None
        self.sessions = None
        self.root = tk.Tk()

    def create_App(self):
        ico = Image.open('volume.png')
        photo = ImageTk.PhotoImage(ico)
        bg = PhotoImage(file="bg.png")  # adding photo that will become the background
        self.root.wm_iconphoto(False, photo)
        self.root.title("Volume Master - management for volumes")
        self.width = self.root.winfo_screenwidth()
        self.height = self.root.winfo_screenheight()
        self.root.geometry(f'{int(self.width * 0.65)}x{int(self.height * 0.65)}')
        label1 = Label(self.root, image=bg)
        label1.place(x=0, y=0)
        self.root.resizable(1, 1)
        self.sessions = AudioUtilities.GetAllSessions()
        for session in self.sessions:
            if session.Process:
                num_input = Scale(self.root, from_=0, to=10, sliderrelief='flat', orient="horizontal",
                                  highlightthickness=0, background='black', fg='grey', troughcolor='#73B5FA',
                                  activebackground='#1065BF')
                num_input.pack(side=LEFT)
                Button(self.root, text=f'{session.Process.name()}',
                       command=lambda: self.show_values(int(num_input.get()))).pack(side=LEFT)
        self.root.mainloop()

    def show_values(self, num):
        volume_manager = VolumeManager()
        volume_manager.changeVolume(num)

if __name__ == "__main__":
    app = App()
    app.create_App()
