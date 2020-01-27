All the program are free of malware or virus. Feel free to download and start to build your own AI. If the anti-virus software blocked QuarkAI, please let anti-virus software trust QuarkAI products or temporary close the anti-virus software.
For the first time, run QuarkAIDllInstall.bat to install QuarkAI.
Double click the VBScript files to run it.
You can also open the VBScript with Notepad to look into the source code.
Run the QuarkAIDllTutorialVB.exe to learn the tutorial in just 10 minutes.
Folder QuarkAIDllTestTool contains the test tool for you to use, which is also a VBScript.
Folder QuarkAIDllDifferentLanguages contains examples of using QuarkAI by different languages.
Folder QuarkAIDllTestWindowWithoutBinding contains a test program for illustrating back-end operations without binding.
Folder QuarkAIDllTestWindowWithoutBinding contains a test program for illustrating back-end operations with binding.
At this point, it should be pointed out that binding mode allows you to manipulate most 3-D game windows.
More information can be found on www.quarkai.com.
Now let me show you a simple example of using QuarkAI in VBScript:
        Dim MyQuarkAI
        Set MyQuarkAI = CreateObject("QuarkAI.dll")
        Dim MyQuarkAIID
        MyQuarkAIID = MyQuarkAI.Login("0123456789ABCDEF")
        Dim MyResult
        MyResult= MyQuarkAI.Message("Hello world!","QuarkAI",0)