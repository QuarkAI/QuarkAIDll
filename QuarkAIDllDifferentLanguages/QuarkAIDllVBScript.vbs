Dim MyQuarkAI
Set MyQuarkAI = CreateObject("QuarkAI.dll")
Dim MyResult
MyResult = MyQuarkAI.Message("This is the usage of QuarkAI.dll in VBScript.", "QuarkAI VBScript", 0)
Set MyQuarkAI = Nothing