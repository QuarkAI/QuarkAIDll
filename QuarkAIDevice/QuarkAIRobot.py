import sys
import subprocess
import time
import threading
import math
import cv2
import zmq
import base64
import socket
import time
import pickle
import struct
import zlib
import numpy as np
import tkinter as tk
from tkinter import messagebox
from PIL import Image
from PIL import ImageTk
from datetime import datetime
import win32com.client as Com
ControlCDKey=""
ConnectionRecord= 0
DistanceRecord=0
TurnAngleRecord=0
HorizontalAngleRecord=0
VerticalAngleRecord=0
VideoIP = 'video.quarkai.com'
VideoBufferSize=3072
VideoPort = 6337
TCPRecord=""
ControlCDKeyRecord=0
DistanceRecord=0
TurnAngleRecord=0
HorizontalAngleRecord=0
VerticalAngleRecord=0
QuarkAIRobotControl = Com.Dispatch("QuarkAI.dll")
def VideoUpdateThread():
    global LabelVideo
    global frame
    global QuarkAIRobotVideo
    while True:
        if ControlCDKeyRecord==1:
            VideoSocket=socket.socket(socket.AF_INET,socket.SOCK_STREAM)
            VideoSocket.connect((VideoIP, VideoPort))
            data = b""
            payload_size = struct.calcsize(">L")
            while ControlCDKeyRecord==1:
                while len(data) < payload_size:
                    data += VideoSocket.recv(VideoBufferSize)
                    
                packed_msg_size = data[:payload_size]
                data = data[payload_size:]
                msg_size = struct.unpack(">L", packed_msg_size)[0]
                while len(data) < msg_size:
                    data += VideoSocket.recv(VideoBufferSize)                    
                
                frame_data = data[:msg_size]
                data = data[msg_size:]
                frame=pickle.loads(frame_data, fix_imports=True, encoding="bytes")
                frame = cv2.imdecode(frame, cv2.IMREAD_COLOR)
                image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                image = Image.fromarray(image)
                image = ImageTk.PhotoImage(image)
                LabelVideo.configure(image=image)
                LabelVideo.image = image
            
            VideoSocket.close()
            video =tk.PhotoImage(file = 'video.png')
            LabelVideo.configure(image=video)
            LabelVideo.image = video
        
        time.sleep(0.5)


def sendReceive(message):
    global TCPRecord
    global QuarkAIRobotControl
    reply = QuarkAIRobotControl.OperateDevice(message)
    if not reply:        
        TCPRecord ="CDKeyRecord:0|DistanceRecord:0|TurnAngleRecord:0|HorizontalAngleRecord:0|VerticalAngleRecord:0"    
    else:
        TCPRecord = reply
        
    Decode()
    LabelDistanceRecord.config(text='Front Distance: '+str(DistanceRecord)+' mm')
    LabelTurnAngleRecord.config(text='Wheel Angle: '+str(TurnAngleRecord)+' degree')
    LabelHorizontalAngleRecord.config(text='Camera Horizontal Angle: '+str(HorizontalAngleRecord)+' degree')
    LabelVerticalAngleRecord.config(text='Camera Vertical Angle: '+str(VerticalAngleRecord)+' degree')

def KeyPressQ(event):
    sendReceive("LeftBlue|0")

def KeyPressE(event):
    sendReceive("RightBlue|0")

def KeyPress1(event):
    sendReceive("Icon|0")

def KeyPress2(event):
    sendReceive("Shapes|0")

def KeyPress3(event):
    sendReceive("Buttons|0")

def KeyPress4(event):
    sendReceive("Noface|0")

def KeyPress5(event):
    sendReceive("Neutral|0")

def KeyPress6(event):
    sendReceive("Smile|0")

def KeyPress7(event):
    sendReceive("Unhappy|0")

def KeyPress8(event):
    sendReceive("Silence|0")

def KeyPress9(event):
    sendReceive("OK|0")

def KeyPress0(event):
    sendReceive("Laugh|0")

def KeyPressW(event):
    sendReceive("MoveRecord|1")

def KeyReleaseW(event):
    sendReceive("MoveRecord|0")

def KeyPressS(event):
    sendReceive("MoveRecord|-1")

def KeyReleaseS(event):
    sendReceive("MoveRecord|0")

def KeyPressA(event):
    sendReceive("TurnRecord|1")

def KeyReleaseA(event):
    sendReceive("TurnRecord|0")

def KeyPressD(event):
    sendReceive("TurnRecord|-1")

def KeyReleaseD(event):
    sendReceive("TurnRecord|0")

def KeyPressLeft(event):
    sendReceive("HorizontalRecord|1")

def KeyReleaseLeft(event):    
    sendReceive("HorizontalRecord|0")
    
def KeyPressRight(event):
    sendReceive("HorizontalRecord|-1")

def KeyReleaseRight(event):
    sendReceive("HorizontalRecord|0")
    

def KeyPressUp(event):
    sendReceive("VerticalRecord|1")

def KeyReleaseUp(event):
    sendReceive("VerticalRecord|0")

def KeyPressDown(event):
    sendReceive("VerticalRecord|-1")

def KeyReleaseDown(event):
    sendReceive("VerticalRecord|0")

def KeyPressR(event):
    sendReceive("MotorReset|0")

def KeyPressSpace(event):
    Snapshot()

def Snapshot():
     global frame
     now = datetime.now()
     date_time = now.strftime("%m%d%Y%H%M%S")
     cv2.imwrite(date_time+'.png',frame)
     print('Iamge written to: '+date_time+'.png')

def Connect():
    global CDKey
    global ControlCDKeyRecord
    global EntryCDKey
    global QuarkAIRobotControl
    if ControlCDKeyRecord==0:
        ControlCDKey=EntryCDKey.get()
        print()
        Result = QuarkAIRobotControl.BindDevice(ControlCDKey)
        sendReceive("Check|0")
        Decode()
        time.sleep(0.5)
        if ControlCDKeyRecord==1:
            BindKeyboard()
            EntryCDKey.config(state='disabled')
            BtnConnect.config(bg='#990000')
            BtnConnect.config(text='Disconnect')
        else:
            CDKey=""
            MyResult = QuarkAIRobotControl.UnBindDevice()
            MyResult = QuarkAIRobotControl.Message("Connection failed, please check the CD-Key and the network connection.", "QuarkAI", 0)            
    
    elif ControlCDKeyRecord==1:        
        sendReceive("Exit|0")
        ControlCDKeyRecord=0
        MyResult = QuarkAIRobotControl.UnBindDevice()
        UnbindKeyboard()
        EntryCDKey.config(state='normal')
        BtnConnect.config(bg='#212121')
        BtnConnect.config(text='Connect')        

def Decode():
    global TCPRecord
    global ControlCDKeyRecord
    global DistanceRecord
    global TurnAngleRecord
    global HorizontalAngleRecord
    global VerticalAngleRecord
    RecordMessage = TCPRecord.split('|', 5)
    temp = RecordMessage[0].split(':', 2)
    ControlCDKeyRecord=int(temp[1])
    temp = RecordMessage[1].split(':', 2)
    DistanceRecord=int(temp[1])
    temp = RecordMessage[2].split(':', 2)
    TurnAngleRecord=int(temp[1])
    temp = RecordMessage[3].split(':', 2)
    HorizontalAngleRecord=int(temp[1])
    temp = RecordMessage[4].split(':', 2)
    VerticalAngleRecord=int(temp[1])

def CheckingThread():    
    while True:
        if ControlCDKeyRecord==1:
            time.sleep(0.5)
            if ControlCDKeyRecord==1:
                sendReceive("Check|0")
            
        time.sleep(0.5) 

def BindKeyboard():
    root.bind('<KeyPress-q>', KeyPressQ) 
    root.bind('<KeyPress-e>', KeyPressE)
    root.bind('<KeyPress-1>', KeyPress1)
    root.bind('<KeyPress-2>', KeyPress2)
    root.bind('<KeyPress-3>', KeyPress3)
    root.bind('<KeyPress-4>', KeyPress4)
    root.bind('<KeyPress-5>', KeyPress5)
    root.bind('<KeyPress-6>', KeyPress6)
    root.bind('<KeyPress-7>', KeyPress7)
    root.bind('<KeyPress-8>', KeyPress8)
    root.bind('<KeyPress-9>', KeyPress9)
    root.bind('<KeyPress-0>', KeyPress0)
    root.bind('<KeyPress-w>', KeyPressW)
    root.bind('<KeyPress-s>', KeyPressS)
    root.bind('<KeyPress-a>', KeyPressA)
    root.bind('<KeyPress-d>', KeyPressD)
    root.bind('<KeyPress-Left>', KeyPressLeft)
    root.bind('<KeyPress-Right>', KeyPressRight)
    root.bind('<KeyPress-Up>', KeyPressUp)
    root.bind('<KeyPress-Down>', KeyPressDown)
    root.bind('<KeyPress-space>', KeyPressSpace)
    root.bind('<KeyPress-r>', KeyPressR)
    root.bind('<KeyRelease-w>', KeyReleaseW)
    root.bind('<KeyRelease-s>', KeyReleaseS)
    root.bind('<KeyRelease-a>', KeyReleaseA)
    root.bind('<KeyRelease-d>', KeyReleaseD)
    root.bind('<KeyRelease-Left>', KeyReleaseLeft)
    root.bind('<KeyRelease-Right>', KeyReleaseRight)
    root.bind('<KeyRelease-Up>', KeyReleaseUp)
    root.bind('<KeyRelease-Down>', KeyReleaseDown)

def UnbindKeyboard():
    root.unbind('<KeyPress-q>') 
    root.unbind('<KeyPress-e>')
    root.unbind('<KeyPress-1>')
    root.unbind('<KeyPress-2>')
    root.unbind('<KeyPress-3>')
    root.unbind('<KeyPress-4>')
    root.unbind('<KeyPress-5>')
    root.unbind('<KeyPress-6>')
    root.unbind('<KeyPress-7>')
    root.unbind('<KeyPress-8>')
    root.unbind('<KeyPress-9>')
    root.unbind('<KeyPress-0>')
    root.unbind('<KeyPress-w>')
    root.unbind('<KeyPress-s>')
    root.unbind('<KeyPress-a>')
    root.unbind('<KeyPress-d>')
    root.unbind('<KeyPress-Left>')
    root.unbind('<KeyPress-Right>')
    root.unbind('<KeyPress-Up>')
    root.unbind('<KeyPress-Down>')
    root.unbind('<KeyPress-r>')
    root.unbind('<KeyRelease-w>')
    root.unbind('<KeyRelease-s>')
    root.unbind('<KeyRelease-a>')
    root.unbind('<KeyRelease-d>')
    root.unbind('<KeyRelease-Left>')
    root.unbind('<KeyRelease-Right>')
    root.unbind('<KeyRelease-Up>')
    root.unbind('<KeyRelease-Down>')
    root.unbind('<KeyRelease-space>')

color_bg='#000000'
color_text='#FFFFFF'
color_btn='#212121'
color_line='#01579B'
color_can='#212121'
color_oval='#2196F3'
target_color='#FF6D00'
root = tk.Tk()
root.title('QuarkAI Robot')
root.wm_iconbitmap('ico64.ico')
root.geometry('1000x600')
root.config(bg=color_bg)
logo =tk.PhotoImage(file = 'logo.png')
LabelLogo=tk.Label(root,image = logo,bg=color_bg)
LabelLogo.place(x=0,y=0)
video =tk.PhotoImage(file = 'video.png')
LabelVideo=tk.Label(root,image = video,bg=color_bg)
LabelVideo.place(x=40,y=50)  
Label1=tk.Label(root,width=30,text='Move-Forward: W',fg='#212121',bg='#90a4ae')
Label1.place(x=760,y=140)
Label2=tk.Label(root,width=30,text='Move-Backward: S',fg='#212121',bg='#90a4ae')
Label2.place(x=760,y=160)
Label3=tk.Label(root,width=30,text='Turn-Left: A',fg='#212121',bg='#90a4ae')
Label3.place(x=760,y=180)
Label4=tk.Label(root,width=30,text='Turn-Right: D',fg='#212121',bg='#90a4ae')
Label4.place(x=760,y=200)
Label5=tk.Label(root,width=30,text='Look-Up: Up', fg='#212121',bg='#90a4ae')
Label5.place(x=760,y=220)
Label6=tk.Label(root,width=30,text='Look-Down: Down',fg='#212121',bg='#90a4ae')
Label6.place(x=760,y=240)
Label7=tk.Label(root,width=30,text='Look-Left: Left',fg='#212121',bg='#90a4ae')
Label7.place(x=760,y=260)
Label8=tk.Label(root,width=30,text='Look-Right: Right',fg='#212121',bg='#90a4ae')
Label8.place(x=760,y=280)
Label9=tk.Label(root,width=30,text='Left-LED: Q',fg='#212121',bg='#90a4ae')
Label9.place(x=760,y=300)
Label10=tk.Label(root,width=30,text='Right-LED: E',fg='#212121',bg='#90a4ae')
Label10.place(x=760,y=320)
Label9=tk.Label(root,width=30,text='Reset: R',fg='#212121',bg='#90a4ae')
Label9.place(x=760,y=340)
Label10=tk.Label(root,width=30,text='Snapshot: Space',fg='#212121',bg='#90a4ae')
Label10.place(x=760,y=360)
Label11=tk.Label(root,width=30,text='Lauging-Man: 1',fg='#212121',bg='#90a4ae')
Label11.place(x=760,y=380)
Label12=tk.Label(root,width=30,text='Hello-World: 2',fg='#212121',bg='#90a4ae')
Label12.place(x=760,y=400)
Label13=tk.Label(root,width=30,text='Star-Show: 3',fg='#212121',bg='#90a4ae')
Label13.place(x=760,y=420)
Label14=tk.Label(root,width=30,text='No-Face: 4',fg='#212121',bg='#90a4ae')
Label14.place(x=760,y=440)
Label15=tk.Label(root,width=30,text='Neutral-Face: 5',fg='#212121',bg='#90a4ae')
Label15.place(x=760,y=460)
Label16=tk.Label(root,width=30,text='Smiling-Face: 6',fg='#212121',bg='#90a4ae')
Label16.place(x=760,y=480)
Label17=tk.Label(root,width=30,text='Unhappy-Face: 7',fg='#212121',bg='#90a4ae')
Label17.place(x=760,y=500)
Label18=tk.Label(root,width=30,text='Silence-Face: 8',fg='#212121',bg='#90a4ae')
Label18.place(x=760,y=520)
Label19=tk.Label(root,width=30,text='OK-Face: 9',fg='#212121',bg='#90a4ae')
Label19.place(x=760,y=540)
Label20=tk.Label(root,width=30,text='Laughing-Face: 0',fg='#212121',bg='#90a4ae')
Label20.place(x=760,y=560)
LabelCDKey=tk.Label(root,width=12,text='Control CDKey:',fg=color_text,bg='#000000')
LabelCDKey.place(x=640,y=15)
EntryCDKey = tk.Entry(root,show=None,width=24,bg="#37474F",fg='#eceff1')
EntryCDKey.place(x=740,y=15)   
BtnConnect= tk.Button(root, width=10, height=1, text='Connect',fg=color_text,bg=color_btn,command=Connect,relief='ridge')
BtnConnect.place(x=900,y=10)
LabelDistanceRecord=tk.Label(root,width=30,text='Front Spacing Distance: ',fg=color_text,bg='#000000')
LabelDistanceRecord.place(x=760,y=40)
LabelTurnAngleRecord=tk.Label(root,width=30,text='Front Wheel Angle: ',fg=color_text,bg='#000000')
LabelTurnAngleRecord.place(x=760,y=65)
LabelHorizontalAngleRecord=tk.Label(root,width=30,text='Camera Horizontal Angle: ',fg=color_text,bg='#000000')
LabelHorizontalAngleRecord.place(x=760,y=90)
LabelVerticalAngleRecord=tk.Label(root,width=30,text='Camera Vertical Angle: ',fg=color_text,bg='#000000')
LabelVerticalAngleRecord.place(x=760,y=115)
ThreadChecking=threading.Thread(target=CheckingThread)
ThreadChecking.setDaemon(True)
ThreadChecking.start()
ThreadVideoUpdate=threading.Thread(target=VideoUpdateThread)
ThreadVideoUpdate.setDaemon(True)
ThreadVideoUpdate.start()
root.mainloop()
