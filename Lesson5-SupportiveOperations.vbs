        Dim MyQuarkAI
        Set MyQuarkAI = CreateObject("QuarkAI.dll")
        Dim MyResult
        MyResult = MyQuarkAI.Message("First Of all, Let's learn how to use the Message. Please click Yes or No.", "This is the title", 1)
        If MyResult = 6 Then
            MyResult = MyQuarkAI.Message("You Clicked Yes.", "Lesson 5", 0)
        Else
            MyResult = MyQuarkAI.Message("You Clicked No.", "Lesson 5", 0)
        End If
        Dim MyVersion
        MyVersion = MyQuarkAI.Ver()
        MyResult = MyQuarkAI.Message("Using the function Ver, we obtain the current version of QuarkAI is: " + MyVersion, "Lesson 5", 0)
        Dim MyString
        MyString = MyQuarkAI.IntegerToString(123)
        MyResult = MyQuarkAI.Message("Function IntegerToString gives you the string of an integer. The string of the 123 is """ + MyString + """", "Lesson 5", 0)
        Dim MyInteger
        MyInteger = MyQuarkAI.StringToInteger(MyString)
        MyInteger = MyInteger + 1
        MyString = MyQuarkAI.IntegerToString(MyInteger)
        MyResult = MyQuarkAI.Message("Function StringToInteger gives you the integer of a string. Now the number is """ + MyString + """", "Lesson 5", 0)
        MyInteger = MyQuarkAI.CountTotalStrings("1|2|3|4", "|")
        MyString = MyQuarkAI.IntegerToString(MyInteger)
        MyResult = MyQuarkAI.Message("Function CountTotalStrings the total number of the sub-string in the main-string.So there are " + MyString + " ""|"" inside the ""1|2|3|4""", "Lesson 5", 0)
        MyString = MyQuarkAI.GetTargetString("1|2|3|4", "|", 2)
        MyResult = MyQuarkAI.Message("Function GetTargetString gives you the target string of a string array.So there are the 2rd string is " + MyString + ", where the string array is by obtained splitting ""1|2|3|4"" with ""|"".", "Lesson 5", 0)
        MyResult = MyQuarkAI.Message("Now Let's delay 2 seconds using function Delay.", "Lesson 5", 0)
        MyResult = MyQuarkAI.Delay(2000)
        MyResult = MyQuarkAI.Message("2 second delay finished.", "Lesson 5", 0)
        MyResult = MyQuarkAI.ShowText(0, "Write text at position 10,20 of the desktop, where 0 is the handle of the desktop.", 10, 20)
        MyResult = MyQuarkAI.Message("Pay attention to the left-top corner of the screen. Function ShowText writes text on a window or the desktop.", "Lesson 5", 0)
        Dim MyTime
        MyTime = MyQuarkAI.NetWorkTime()
        MyResult = MyQuarkAI.Message("Function NetworkTime says the global time is: " + MyTime, "Lesson 5", 0)
        Dim MyHttpText
        MyHttpText = MyQuarkAI.GetHttpText("http://www.quarkai.com")
        MyResult = MyQuarkAI.Message("Function MyHttpText can get network text within 64 characters. The http text is: " + MyHttpText, "Lesson 5", 0)
        Dim MyPath
        MyPath = MyQuarkAI.GetCurrentPath()
        MyResult = MyQuarkAI.Message("Get the current path using function CurrentPath. The current path is: " + MyPath, "Lesson 5", 0)
        MyResult = MyQuarkAI.CreateFolder("MyFolder")
        MyResult = MyQuarkAI.Message("Create a folder at the current path using function CreateFolder, the name of the folder is: MyFolder", "Lesson 5", 0)
        MyPath = MyQuarkAI.SetCurrentPath("MyFolder")
        MyResult = MyQuarkAI.Message("Using function SetCurrentPath, set the current path to: MyFolder", "Lesson 5", 0)
        MyPath = MyQuarkAI.GetCurrentPath()
        MyResult = MyQuarkAI.Message("Now the current path is: " + MyPath, "Lesson 5", 0)
        MyResult = MyQuarkAI.CreateFile("MyFile.ini")
        MyResult = MyQuarkAI.Message("Create a file through function CreateFile: MyFile.ini.", "Lesson 5", 0)
        MyResult = MyQuarkAI.WriteINI(MyPath + "\MyFile.ini", "MySection", "MyKey", "MyValue")
        MyResult = MyQuarkAI.Message("With function WriteINI, you can write some user data to the file, using the full path.", "Lesson 5", 0)
        Dim MyKey
        MyKey = MyQuarkAI.ReadINI(MyPath + "\MyFile.ini", "MySection", "MyKey")
        MyResult = MyQuarkAI.Message("Read a key from the file using function ReadINI. The key read from the file is: " + MyKey, "Lesson 5", 0)
        MyResult = MyQuarkAI.Execute("Notepad.exe MyFile.ini")
        MyResult = MyQuarkAI.Delay(1000)
        MyResult = MyQuarkAI.Message("The function Execute can run any windows command. For example, we can use it to open Myfile.ini with Notepad.exe.", "Lesson 5", 0)
        MyResult = MyQuarkAI.SetCurrentPath("..")
        MyResult = MyQuarkAI.Message("Remeber to set the current path back. And this is the end of lesson 5.", "Lesson 5", 0)
        Set MyQuarkAI = Nothing