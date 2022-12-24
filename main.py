from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
from tkinter import *

#####################
#   CONSOLE APP
####################
class App():
    def __init__(self):
        super(App, self).__init__()
        self.root = Tk()

    def create_App(self):
        volumeManager = VolumeManager()
        volumeManager.changeVolume()
        self.root.title("Volume Master - management for volumes")
        self.width = self.root.winfo_screenwidth()
        self.height = self.root.winfo_screenheight()
        self.root.geometry(f'{self.width // 2}x{self.height // 2}')
        self.root.resizable(0, 0)
        self.root.mainloop()

####################
# VOLUME MANAGEMENT
###################
class VolumeManager():
    def __init__(self):
        super(VolumeManager, self).__init__()
        self.devices = AudioUtilities.GetSpeakers()
        self.interface = self.devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        self.volume = cast(self.interface, POINTER(IAudioEndpointVolume))

    def changeVolume(self):
        print("X")
        self.sessions = AudioUtilities.GetAllSessions()
        for session in self.sessions:
            self.volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            if session.Process and session.Process.name() == "chrome.exe":
                self.volume.SetMasterVolume(1, None)

if __name__ == "__main__":
    app = App()
    app.create_App()