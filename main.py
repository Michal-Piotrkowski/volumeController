from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
import tkinter as tk
from tkinter import messagebox
from tkinter import *

#####################
#   CONSOLE APP
####################
class App():
    def __init__(self):
        super(App, self).__init__()
        self.root = tk.Tk()

    def create_App(self):
        self.root.title("Volume Master - management for volumes")
        self.width = self.root.winfo_screenwidth()
        self.height = self.root.winfo_screenheight()
        self.root.geometry(f'{self.width // 2}x{self.height // 2}')
        self.root.resizable(0, 0)
        num_input = Scale(self.root, from_=100, to=0)
        num_input.pack()
        Button(self.root, text='set', command=self.show_values(num_input.get())).pack()
        self.root.mainloop()
    def show_values(self, num):
        messagebox.showinfo('Message', 'You clicked the Submit button!')
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
        self.sessions = AudioUtilities.GetAllSessions()
        for session in self.sessions:
            if session.Process:
                print(session.Process.name())
            self.volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            if session.Process and session.Process.name() == "chrome.exe":
                self.volume.SetMasterVolume(level, None)

if __name__ == "__main__":
    app = App()
    app.create_App()