QuarkAI device control functions can help you control a QuarkAI as long as you have a computer that has internet access.
First of you, you need to connect the device to a computer using a USB cable.
Then you can use SetDevice function to set up device name, wifi name, wifi password, QuarkAI CDKey, etc.
After this you can plug off the USB cable. 
Now the device is ready to be controlled by any computer on-line.
To control the device, you first need to bind the device through the CDKey that you set up, using function BindDevice.
Then you can use OperateDeivce to control it.
Finally, when you no longer need to use the device, just call the function UnBindDevice.

For example, now you have QuarkAI lamp:
You firstly, connect it to a computer through a USB cable and run the VB script:
Set MyQuarkAI = CreateObject("QuarkAI.dll")
MyResult = MyQuarkAI.SetDevice("SSID=Wifiname;Password=WifiPassword;DeviceName=MyLamp;CDKey=0123456789ABCDEF")
Set MyQuarkAI = Nothing

Then you just plug off the USB cable and run the following VB script to control this lamp:
Set MyQuarkAI = CreateObject("QuarkAI.dll")
MyResult = MyQuarkAI.BindDevice("0123456789ABCDEF")
MyResult = MyQuarkAI.OperateDevice("LightOn")
MyResult = MyQuarkAI.Delay(360000)
MyResult = MyQuarkAI.OperateDevice("LightOff")
MyResult = MyQuarkAI.UnBindDevice()
Set MyQuarkAI = Nothing

Dr. Xu also prepare a robot for you to test the QuarkAI device control functions.
QuarkAIRobot.exe is the client for that robot, and you need to ask Dr. Xu for the CDKey before using it.
QuarkAIRobot.py is the source code of QuarkAIRobot.exe.
QuarkAIDevice.vbs is a small script to control the camera of the robot, which also needs the CDKey.
