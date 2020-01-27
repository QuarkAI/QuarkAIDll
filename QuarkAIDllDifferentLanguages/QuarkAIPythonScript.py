import comtypes.client as Com
MyQuarkAI = Com.CreateObject("QuarkAI.dll")
MyResult = MyQuarkAI.Message("This is an example to use QuarkAI in Python", "QuarkAI PythonScript", 0)
MyQuarkAI = None
