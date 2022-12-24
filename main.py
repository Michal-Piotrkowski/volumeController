from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
import tkinter as tk
from tkinter import messagebox
from tkinter import *
from PIL import Image, ImageTk

#####################
#   CONSOLE APP
####################
class App():
    def __init__(self):
        super(App, self).__init__()
        self.root = tk.Tk()

    def create_App(self):
        ico = Image.open('volume.png')
        photo = ImageTk.PhotoImage(ico)
        self.root.wm_iconphoto(False, photo)
        self.root.title("Volume Master - management for volumes")
        self.width = self.root.winfo_screenwidth()
        self.height = self.root.winfo_screenheight()
        self.root.geometry(f'{self.width // 2}x{self.height // 2}')
        self.root.resizable(0, 0)
        self.sessions = AudioUtilities.GetAllSessions()
        for session in self.sessions:
            if session.Process:
                num_input = Scale(self.root, from_=0, to=10, orient=VERTICAL)
                num_input.pack(side=BOTTOM)
                Button(self.root, text=f'{session.Process.name()}', command=lambda: self.show_values(int(num_input.get()))).pack(side=BOTTOM)
        self.root.mainloop()
    def show_values(self, num):
        volumeManager = VolumeManager()
        volumeManager.changeVolume(num)

####################
# VOLUME MANAGEMENT
###################
class VolumeManager():
    def __init__(self):
        super(VolumeManager, self).__init__()
        self.devices = AudioUtilities.GetSpeakers()
        self.interface = self.devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        self.volume = cast(self.interface, POINTER(IAudioEndpointVolume))

    def changeVolume(self, level):
        print(level)
        self.sessions = AudioUtilities.GetAllSessions()
        for session in self.sessions:
            if session.Process:
                self.volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                self.volume.SetMasterVolume(level/10, None)

if __name__ == "__main__":
    app = App()
    app.create_App()