		Dim MyWScript
        Set MyWScript = CreateObject("WScript.Shell")
        Dim MyQuarkAI
        Set MyQuarkAI = CreateObject("QuarkAI.dll")
        Dim MyResult
        Dim MyString
        MyString = "In this lesson, we will learn how to control 3-D game windows using binding mode. Now let us open the test program. This process may take a minite, please wait."
        MyResult = MyQuarkAI.Message(MyString, "Lesson 4", 0)
        Dim MyIndex
        Dim MyMainWindowHandle
        MyMainWindowHandle = 0
        MyMainWindowHandle = MyQuarkAI.FindWindow("", "QuarkAITestWindows")
        If MyMainWindowHandle = 0 Then
            MyResult = MyQuarkAI.SetCurrentPath("QuarkAIDllTestWindowWithBinding")
            MyResult = MyWScript.run("QuarkAIDllTestWindowWithBinding.exe")
            MyResult = MyQuarkAI.SetCurrentPath("..")
            For MyIndex = 1 To 100
                MyMainWindowHandle = MyQuarkAI.FindWindow("", "QuarkAITestWindows")
                If MyMainWindowHandle > 0 Then
                    MyIndex = 100
                End If
                MyResult = MyQuarkAI.Delay(1000)
            Next
        End If
		Set MyWScript = Nothing
        If MyMainWindowHandle > 0 Then
			MyString = "Now you can try to play with the test game window by pressing the keys: w, s, a, d, left, up, right, down. And also click the rest button."
			MyResult = MyQuarkAI.Message(MyString, "Lesson 4", 0)
			MyString = "If you are ready, we can begin this lesson by moving the game window to the top-left of the desktop."
			MyResult = MyQuarkAI.Message(MyString, "Lesson 4", 0)
            MyResult = MyQuarkAI.Delay(1000)
            MyResult = MyQuarkAI.MoveWindow(MyMainWindowHandle, 0, 0)
            MyResult = MyQuarkAI.Delay(1000)
            MyResult = MyQuarkAI.SetWindowState(MyMainWindowHandle, 1)
            Dim MyFrameHandle
            MyFrameHandle = MyQuarkAI.FindWindowEx(MyMainWindowHandle, "", "")
            Dim MyClientHandle
            MyClientHandle = MyQuarkAI.FindWindowEx(MyFrameHandle, "", "QuarkAITest")
            MyString = "The handle Of the client region is " + MyQuarkAI.IntegerToString(MyClientHandle) + ". Now let us bind this window and control it."
            MyResult = MyQuarkAI.Message(MyString, "Lesson 4", 0)
            Dim MyQuarkAIID
            MyQuarkAIID = MyQuarkAI.Login("0123456789ABCDEF")
            MyString = "To bind a window , you need to login With a CDKey first. And you QuarkAI ID will be obtained after login, which is " + MyQuarkAI.IntegerToString(MyQuarkAIID) + " for the present CDKey."
            MyResult = MyQuarkAI.Message(MyString, "Lesson 4", 0)
            MyString = "Function BindWindow can bind the AI to the window to control it. The graphic mode can be normal (front-end), gdi (back-end for ordinary windows), dx or dx2 (back-end For 3D-game windows). The mouse And keyboard mode can be normal (front-end), windows (back-end For ordinary windows) And dx (back-end For game windows). Let us try the back-end mode, which won't influence your front-end keyboard and mouse operations."
            MyResult = MyQuarkAI.Message(MyString, "Lesson 4", 0)
            MyResult = MyQuarkAI.BindWindow(MyClientHandle, "dx2", "windows", "windows", 0)
            If MyResult = 0 Then
                MyString = "Failed to bind the window, please give access to QuarkAI.dll from anti-virus software."
                MyResult = MyQuarkAI.Message(MyString, "Lesson 4", 0)
            Else
                MyString = "Now let us start to control the window, using functiong KeyPressChar, KeyDownChar, and KeyUpChar."
                MyResult = MyQuarkAI.Message(MyString, "Lesson 4", 0)
                MyResult = MyQuarkAI.KeyPressChar("R")
                MyResult = MyQuarkAI.Delay(100)
                MyResult = MyQuarkAI.KeyDownChar("S")
                MyResult = MyQuarkAI.Delay(1000)
                MyResult = MyQuarkAI.KeyUpChar("S")
                MyResult = MyQuarkAI.Delay(100)
                MyResult = MyQuarkAI.KeyDownChar("W")
                MyResult = MyQuarkAI.Delay(1000)
                MyResult = MyQuarkAI.KeyUpChar("W")
                MyResult = MyQuarkAI.Delay(100)
                MyResult = MyQuarkAI.KeyDownChar("S")
                MyResult = MyQuarkAI.Delay(1000)
                MyResult = MyQuarkAI.KeyUpChar("S")
                MyResult = MyQuarkAI.Delay(100)
                MyResult = MyQuarkAI.KeyDownChar("A")
                MyResult = MyQuarkAI.Delay(1000)
                MyResult = MyQuarkAI.KeyUpChar("A")
                MyResult = MyQuarkAI.Delay(100)
                MyResult = MyQuarkAI.KeyDownChar("D")
                MyResult = MyQuarkAI.Delay(1000)
                MyResult = MyQuarkAI.KeyUpChar("D")
                MyResult = MyQuarkAI.Delay(100)
                Dim MyClientX1, MyClientY1, MyClientX2, MyClientY2
                MyResult = MyQuarkAI.GetClientRect(MyClientHandle, MyClientX1, MyClientY1, MyClientX2, MyClientY2)
                Dim ColorX, ColorY, MyColor
                MyColor = MyQuarkAI.GetColor(417, 445)
                MyString = "GetColor gives you the color of a pixel. The color at (425, 475) is " + MyColor
                MyResult = MyQuarkAI.Message(MyString, "Lesson 4", 0)
                MyString = "Now let us try to find the ResetButton and the click it, using function FindPic."
                MyResult = MyQuarkAI.Message(MyString, "Lesson 4", 0)
                Dim PictureX, PictureY
                MyResult = MyQuarkAI.SetCurrentPath("QuarkAIDllTestWindowWithBinding")
                MyResult = MyQuarkAI.FindPic(MyClientX1, MyClientY1, MyClientX2, MyClientX2, "ResetButton.bmp", "000000", 0.9, 0, PictureX, PictureY)
                If PictureX > 0 Then
                    MyResult = MyQuarkAI.MoveTo(PictureX, PictureY)
                    MyResult = MyQuarkAI.Delay(100)
                    MyResult = MyQuarkAI.LeftClick()
                End If
                MyString = "Let us save the snap shot and store it to QuarkAICapture.bmp."
                MyResult = MyQuarkAI.Message(MyString, "Lesson 4", 0)
                MyResult = MyQuarkAI.Capture(MyClientX1, MyClientY1, MyClientX2, MyClientX2, "QuarkAICapture.bmp")
                MyString = "Remember to to unbind the window by callling function UnBindWindow."
                MyResult = MyQuarkAI.Message(MyString, "Lesson 4", 0)
                MyResult = MyQuarkAI.UnBindWindow()
                MyString = "Let us open the file QuarkAICapture.bmp by MSPaint."
                MyResult = MyQuarkAI.Message(MyString, "Lesson 4", 0)
                MyResult = MyQuarkAI.Execute("mspaint QuarkAICapture.bmp")
                MyResult = MyQuarkAI.SetCurrentPath("..")
            End If
        Else
            MyString = "Sorry I cannot run the QuarkAIDllTestWindowWithBinding.exe, which should be placed in the folder QuarkAIDllTestWindowWithBinding. Please check it."
            MyResult = MyQuarkAI.Message(MyString, "Lesson 4", 0)
        End If
		MyString = "This is the end of Lesson 4, thank you."
        MyResult = MyQuarkAI.Message(MyString, "Lesson 4", 0)
        Set MyQuarkAI = Nothing