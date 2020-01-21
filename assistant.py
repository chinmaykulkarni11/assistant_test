import speech_recognition as sr
import os
import wx
import win32com.client as wc
from urllib.request import urlopen
import webbrowser
import winshell
import time
import ctypes
import random
import subprocess

speak = wc.Dispatch("SAPI.SpVoice")


class CSfx(wx.Frame):
    #GUI
    def __init__(self,parent):
        wx.Frame.__init__(self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.Size(410,165), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        bSizer1 = wx.BoxSizer(wx.VERTICAL)

        self.m_staticText1 = wx.StaticText(self, wx.ID_ANY, u"CSfx", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText1.Wrap(-1)
        bSizer1.Add(self.m_staticText1, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_textCtrl1 = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer1.Add( self.m_textCtrl1, 0, wx.ALL|wx.EXPAND, 5 )

        self.m_button1 = wx.Button(self, wx.ID_ANY, u"Work on it", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer1.Add(self.m_button1, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_button2 = wx.Button(self, wx.ID_ANY, u"Voice command", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer1.Add(self.m_button2, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.SetSizer(bSizer1)
        self.Layout()

        self.Centre(wx.BOTH)
        speak.Speak("hey mate, i am at your service")

        # Connect Events
        self.m_button1.Bind(wx.EVT_BUTTON, self.txtin)
        self.m_button2.Bind(wx.EVT_BUTTON, self.rec)
        

    def __del__(self):
        pass


    #take text command
    def txtin( self, event ):
        txt = str(self.m_textCtrl1.GetValue())
        CSfx.work(str(txt))


    #command processor
    def work( term ):
        trm = term
        try:
            #fb
            if trm == "fb":
                speak.speak("opening facebook")
                webbrowser.open("www.fb.com")
            #basic fb
            elif trm == "mb":
                speak.speak("opening basic facebook")
                webbrowser.open("www.mbasic.faceook.com")
            #youtube
            elif trm == "yt":
                speak.speak("opening youtube")
                webbrowser.open("www.youtube.com")
            #empty recycle bin or use ccleaner
            elif trm == "clean":
                try:
                    winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
                    speak.speak("trash emptied buddy!")
                except:
                    speak.speak("there's no trash yet, try cleaning")
                    os.startfile("C:\\Program Files\\CCleaner\\CCleaner64.exe")
            #lock device
            elif trm == "lock":
                try:
                    speak.speak("locking your device!")
                    ctypes.windll.user32.LockWorkStation()
                except Exception as e:
                   print(str(e))
            #any google search       
            else:
                speak.speak("searching it on google")
                webbrowser.open("www.google.com/search?q="+trm)
        except Exception as e:
            print(str(e))


    #recognize audio
    def rec(self, event):
        r = sr.Recognizer()
        with sr.Microphone() as src:
            print("Say something!")
            atxt = r.listen(src)

        #google audio recogniotion
        try:
            speak.speak("you said " + r.recognize_google(atxt))
            print("you said " + r.recognize_google(atxt))
            txt = atxt
            CSfx.work(str(txt))
        except sr.UnknownValueError:
            speak,speak("Pardon me")
        except sr.RequestError as e:
            print("Sphinx error; {0}".format(e))

        #recognize audio command with sphinx
        try:
            print("you said " + r.recognize_sphinx(atxt))
            speak.speak("you said " + r.recognize_sphinx(atxt))
            txt = atxt
            CSfx.work(str(txt))
        except sr.UnknownValueError:
            speak,speak("Pardon me")
        except sr.RequestError as e:
            print("Sphinx error; {0}".format(e))

        
# Trigger GUI
app = wx.App(False) 
frame = CSfx(None) 
frame.Show(True) 
#start the applications 
app.MainLoop()
