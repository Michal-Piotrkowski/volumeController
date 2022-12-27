from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, ISimpleAudioVolume
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL

####################
# VOLUME MANAGEMENT
###################
class VolumeManager():
    def __init__(self):
        super(VolumeManager, self).__init__()
        self.sessions = None
        self.devices = AudioUtilities.GetSpeakers()
        self.interface = self.devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        self.volume = cast(self.interface, POINTER(IAudioEndpointVolume))

    def changeVolume(self, level):
        print(level)
        self.sessions = AudioUtilities.GetAllSessions()
        for session in self.sessions:
            if session.Process:
                self.volume = session._ctl.QueryInterface(ISimpleAudioVolume)
                self.volume.SetMasterVolume(level / 10, None)