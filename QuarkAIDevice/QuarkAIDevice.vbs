'Step 1: Connect the QuarkAI device to the computer using a USB cable.
'Step 2: Input setup information to the device using function SetDevice. For example. SetDevice("SSID=WIFINAME;Password:WIFIPASSWORD;CDKEY=YOURQUARKAICDKEY"). After this, you can plug off the USB cable, and the QuarkAI device will run automatically.
'Step 3: Now you can control that QuarkAI device with any computer has internet connection. The first step is bind the device using function BindDevice("A_CDKey")
'Step 4: Use function OperateDevice to control the QuarkAI device. For example, you can use the following function the control the QuarkAI robot:
'Camera up: OperateDevice(HorizontalRecord|1)
'Camera down: OperateDevice(HorizontalRecord|-1)
'Camera stop: OperateDevice(HorizontalRecord|0)
'Camera left: OperateDevice(VerticalRecord|1)
'Camera right: OperateDevice(VerticalRecord|-1)
'Camera stop: OperateDevice(VerticalRecord|0)
'Step 5: When your control is down, remember to use function UnBindDevice().
'The following script shows how to control a QuarkAI device, if you want to have a try, please ask Dr. Xu for the CDKey.
MyCDKey = "AAAAAAAAAAAAAAAA"
Set MyQuarkAI = CreateObject("QuarkAI.Dll")
MyResult = MyQuarkAI.BindDevice(MyCDKey)
MyResult = MyQuarkAI.OperateDevice("HorizontalRecord|1")
MyResult = MyQuarkAI.Delay(500)
MyResult = MyQuarkAI.OperateDevice("HorizontalRecord|0")
MyResult = MyQuarkAI.Delay(500)
MyResult = MyQuarkAI.OperateDevice("HorizontalRecord|-1")
MyResult = MyQuarkAI.Delay(500)
MyResult = MyQuarkAI.OperateDevice("HorizontalRecord|0")
MyResult = MyQuarkAI.Delay(500)
MyResult = MyQuarkAI.OperateDevice("VerticalRecord|1")
MyResult = MyQuarkAI.Delay(500)
MyResult = MyQuarkAI.OperateDevice("VerticalRecord|0")
MyResult = MyQuarkAI.Delay(500)
MyResult = MyQuarkAI.OperateDevice("VerticalRecord|-1")
MyResult = MyQuarkAI.Delay(500)
MyResult = MyQuarkAI.OperateDevice("VerticalRecord|0")
MyResult = MyQuarkAI.Delay(500)
MyResult = MyQuarkAI.UnBindDevice()
Set MyQuarkAI = Nothing
