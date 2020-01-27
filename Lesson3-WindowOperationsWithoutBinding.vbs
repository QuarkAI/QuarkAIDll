        Dim MyWScript
        Set MyWScript = CreateObject("WScript.Shell")
        Dim MyQuarkAI
        Set MyQuarkAI = CreateObject("QuarkAI.dll")
        Dim MyResult
        Dim MyString
        MyString = "In this lesson, we will learn window operations without binding. Now Let us open the test program. This process may take a minite, please wait."
        MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
        Dim MyIndex
        Dim MyMainWindowHandle
        MyMainWindowHandle = 0
        MyMainWindowHandle = MyQuarkAI.FindWindow("", "QuarkAITestWindowWithoutBinding")
        If MyMainWindowHandle = 0 Then
            MyResult = MyQuarkAI.SetCurrentPath("QuarkAIDllTestWindowWithoutBinding")
            MyResult = MyWScript.run("QuarkAITestWindowWithoutBinding.exe")
            MyResult = MyQuarkAI.SetCurrentPath("..")
            For MyIndex = 1 To 100
                MyMainWindowHandle = MyQuarkAI.FindWindow("", "QuarkAITestWindowWithoutBinding")
                If MyMainWindowHandle > 0 Then
                    MyIndex = 100
                End If
                MyResult = MyQuarkAI.Delay(1000)
            Next			
        End If
		Set MyWScript = Nothing
        MyMainWindowHandle = MyQuarkAI.FindWindow("", "QuarkAITestWindowWithoutBinding")
        MyString = "The test window contains a label, two inputboxes and a button. By clicking the button, the text in the first inputboxes will be copied to the label. Now you can play with it. If you are ready for the lesson, press OK."
        MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
        MyString = "Function FindWindow can find the hand of the window by the title Or the class name. It finds the handle of the main window is " + MyQuarkAI.IntegerToString(MyMainWindowHandle) + "."
        MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
        Dim MyWindowX1, MyWindowY1, MyWindowX2, MyWindowY2
        MyResult = MyQuarkAI.GetWindowRect(MyMainWindowHandle, MyWindowX1, MyWindowY1, MyWindowX2, MyWindowY2)
        MyString = "Funtion GetWindowRect gives you the coordinates Of the left-top point And the right-botton point of the window, which is : (" + MyQuarkAI.IntegerToString(MyWindowX1) + ", " + MyQuarkAI.IntegerToString(MyWindowY1) + "), (" + MyQuarkAI.IntegerToString(MyWindowX2) + ", " + MyQuarkAI.IntegerToString(MyWindowY2) + ")."
        MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
        Dim MyClientX1, MyClientY1, MyClientX2, MyClientY2
        MyResult = MyQuarkAI.GetClientRect(MyMainWindowHandle, MyClientX1, MyClientY1, MyClientX2, MyClientY2)
        MyString = "The client region is the main interface region of the window, whose area is slightly smaller than the window region. Function GetClientRect gives you the coordinates of the left-top point and the right-botton point of the client region, which is: (" + MyQuarkAI.IntegerToString(MyClientX1) + ", " + MyQuarkAI.IntegerToString(MyClientY1) + "), (" + MyQuarkAI.IntegerToString(MyClientX2) + ", " + MyQuarkAI.IntegerToString(MyClientY2) + ")."
        MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
        MyString = "Function MoveWindow can move the window to wherever you want. And funtion SetWindowState can set the state of the window. Now let us move the window to the point (0, 0) of the desktop and then set it to top."
        MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
        MyResult = MyQuarkAI.MoveWindow(MyMainWindowHandle, 0, 0)
        MyResult = MyQuarkAI.SetWindowState(MyMainWindowHandle, 1)
        MyString = "Now let us hide the window."
        MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
        MyResult = MyQuarkAI.SetWindowState(MyMainWindowHandle, 6)
        MyString = "Show the window."
        MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
        MyResult = MyQuarkAI.SetWindowState(MyMainWindowHandle, 7)
        MyString = "Flash the window."
        MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
        For MyIndex = 1 To 10
            MyResult = MyQuarkAI.SetWindowState(MyMainWindowHandle, 14)
            MyResult = MyQuarkAI.Delay(200)
        Next
		MyString = "Set the window to the top."
		MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
        MyResult = MyQuarkAI.SetWindowState(MyMainWindowHandle, 1)
        Dim MyButtonHandle
        MyButtonHandle = MyQuarkAI.FindWindowEx(MyMainWindowHandle, "", "Click me")
        MyString = "Buttons, inputboxes and other components are just sub windows of the main window, whose handles can be found by function FindWindowEx. For example, the handle of the Click me button is " + MyQuarkAI.IntegerToString(MyButtonHandle) + "."
        MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
        MyString = "The coordinate origin of the button is just the left-top point of the button. Now let us click the button at coordinate (5, 5) through back-end by the function SendLeftClick."
        MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
        MyResult = MyQuarkAI.SendLeftClick(MyButtonHandle, 5, 5)
        MyString = "Sometimes you have to find the window through the class name. And there exists multiple sub windows with the sample class name. For example, the main window contains two inputboxes. Then you have to use the function EnumWindow to find all of them first. Then try to locate which one is the one you want. Now let us try to find the handle of the correct inputbox."
        MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
        Dim MyAllInputboxHandles
        MyAllInputboxHandles = MyQuarkAI.EnumWindow(MyMainWindowHandle, "", "WindowsForms10.EDIT", 2)
        Dim MyInputBoxHandleString
		Dim MyInputBoxHandle
        MyInputBoxHandleString = MyQuarkAI.GetTargetString(MyAllInputboxHandles, ",", 1)
        MyInputBoxHandle = MyQuarkAI.StringToInteger(MyInputBoxHandleString)
        MyString = "It is found that the handle of the first inputbox is " + MyQuarkAI.IntegerToString(MyInputBoxHandle) + "."
        MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
		MyString = "Now let us type the word hello to the inputbox through the SendkeyPressChar function."
        MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
        MyResult = MyQuarkAI.SetWindowState(MyMainWindowHandle, 1)
        MyResult = MyQuarkAI.Delay(500)
        MyResult = MyQuarkAI.SendKeyPressChar(MyInputBoxHandle, "H")
        MyResult = MyQuarkAI.Delay(200)
        MyResult = MyQuarkAI.SendKeyPressChar(MyInputBoxHandle, "E")
        MyResult = MyQuarkAI.Delay(200)
        MyResult = MyQuarkAI.SendKeyPressChar(MyInputBoxHandle, "L")
        MyResult = MyQuarkAI.Delay(200)
        MyResult = MyQuarkAI.SendKeyPressChar(MyInputBoxHandle, "L")
        MyResult = MyQuarkAI.Delay(200)
        MyResult = MyQuarkAI.SendKeyPressChar(MyInputBoxHandle, "O")
        MyResult = MyQuarkAI.Delay(200)
		MyString = "You can also send a string to the sub window, such as the inputbox, through function SendString. Now let us send the string to the inputbox, which is: world."
		MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
        MyResult = MyQuarkAI.SendString(MyInputBoxHandle, " world")
		MyResult = MyQuarkAI.Delay(500)
		MyString = "Not let us click the Click me button"
		MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
		MyResult = MyQuarkAI.SendLeftClick(MyButtonHandle, 5, 5)
        MyString = "This is the end of lession3, thank you."
        MyResult = MyQuarkAI.Message(MyString, "Lesson 3", 0)
        Set MyQuarkAI = Nothing