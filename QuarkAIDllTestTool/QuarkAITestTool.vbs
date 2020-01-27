dim ie
set ie = wscript.createobject("internetexplorer.application","ie_")
dim wnd
dim id
dim MyQuarkAI
set MyQuarkAI = createobject("QuarkAI.dll")
dim number
dim MyResult
dim ws
Set ws = CreateObject("WScript.Shell")
number = 0
ie.menubar=1
ie.addressbar=0
ie.toolbar=0
ie.statusbar=0
ie.width=1024
ie.height=1024
ie.resizable=1 
ie.navigate "about:blank"
ie.visible=1
dim MyResponse
dim MousePointPixelInformation
dim MousePointWindowInformation
dim MousePointPixelCheckBox
dim MousePointWindowCheckBox
dim MouseX
dim MouseY
dim MouseColor
dim MouseCursor
dim PressKeyCode
dim RunningFlag
dim WindowHandle
dim WindowTitle
dim WindowClass
dim WindowX1
dim WindowX2
dim WindowY1
dim WindowY2
dim ClientX1
dim ClientX2
dim ClientY1
dim ClientY2
dim ClientWidth
dim ClientHeight
dim WindowParent
dim ProcessID
dim ProcessInfomation
dim ParentTitle
dim ParentClass
dim ParentX1
dim ParentX2
dim ParentY1
dim ParentY2
dim ChildHandles
dim ChildNumber
dim ChildIndex
dim ChildHandle
dim ChildTitle
dim ChildClass
dim ChildX1
dim ChildX2
dim ChildY1
dim ChildY2
dim MyOutput
dim MouseHandleWithoutBind
dim KeyboardHandleWithoutBind
dim MouseXWithoutBind
dim MouseYWithoutBind
dim MouseMessageWithoutBind
dim KeycodeWithoutBind
dim KeyboardMessageWithoutBind
dim WindowHandleWithBind
dim CDKey
dim GraphicsMode
dim MouseMode
dim KeyboardMode
dim LoginFlag
dim ColorXWithBind
dim ColorYWithBind
dim PictureName
dim ColorToleranceWithBind
dim SimlarityWithBind
dim BindMode
dim MouseXWithBind
dim MouseYWithBind
dim KeycodeWithBind
dim MouseXDesktop
dim MouseYDesktop
dim KeycodeDesktop
dim WindowHandleWithoutBind
dim WindowXWithoutBind
dim WindowYWithoutBind
dim ClientWidthWithoutBind
dim ClientHeightWithoutBind
ColorXWithBind = 0
ColorYWithBind = 0
PictureName = ""
ColorToleranceWithBind = ""
SimlarityWithBind = ""
WindowHandleWithoutBind = 0
WindowXWithoutBind = 0
WindowYWithoutBind = 0
ClientWidthWithoutBind = 0
ClientHeightWithoutBind = 0
MouseHandleWithoutBind = 0
KeyboardHandleWithoutBind = 0
MouseXWithoutBind = 0
MouseYWithoutBind = 0
WindowHandleWithBind = 0
BindMode = 0
MouseXWithBind = 0
MouseYWithBind = 0
KeycodeWithBind = 0
MouseXDesktop = 0
MouseYDesktop = 0
KeycodeDesktop = 0
LoginFlag = false
SetWindowStateSelection = "CloseWindow"
MouseMessageWithoutBind = "MoveTo"
KeycodeWithoutBind = 0
KeyboardMessageWithoutBind = "KeyPress"
CDKey = "0123456789ABCDEF"
GraphicsMode = "normal"
MouseMode = "normal"
KeyboardMode = "normal"
RunningFlag = true
MousePointPixelCheckBox = true
MousePointWindowCheckBox = true
dim BindFlag
dim QuarkAIID
QuarkAIID = 0
BindFlag = False
dim PressQ
dim Press1
dim Press2
sub rewrite
	MyResult = ie.document.close()
	MyResponse = "<html>"
	MyResponse = MyResponse + "<title>QuarkAI Tool</title>"
	MyResponse = MyResponse + "<body overflow: auto>"
	MyResponse = MyResponse + "<div>Supportive functions:<br>"
	MyResponse = MyResponse + "<div><p><input id=Version type=button value=Version>, <input id=BasePath type=button value=BasePath>, <input id=CurrentPath type=button value=CurrentPath>, <input id=MachineCode type=button value=MachineCode>, <input id=Magnifier type=button value=Magnifier>, <input id=SnippingTool type=button value=SnippingTool>, <input id=KeyCodeCheck type=button value=KeyCode>, <input id=MSPaint type=button value=Paint>.</p>"
	MyResponse = MyResponse + "<hr></div>"
	MyResponse = MyResponse + "<div>Desktop operation functions:<br>"
	if MousePointPixelCheckBox then
		MyResponse = MyResponse + "<p><input id=MousePointedPixel type=checkbox checked>Watch mouse pointed pixel information by pressing Q + 1</p>"
	else
		MyResponse = MyResponse + "<p><input id=MousePointedPixel type=checkbox>Watch mouse pointed pixel information by pressing Q + 1</p>"
	end if
	MyResponse = MyResponse + "<p>Mouse: X = <input type=text id=MouseXDesktop size=1 value=" + cstr(MouseXDesktop) + ">, Y = <input type=text id=MouseYDesktop size=1 value=" + cstr(MouseYDesktop) + ">, <input id=MoveToDesktopButton type=button value=MoveTo>, "
	MyResponse = MyResponse + "<input id=LeftClickDesktopButton type=button value=LeftClick>, "
	MyResponse = MyResponse + "<input id=LeftDownDesktopButton type=button value=LeftDown>, "
	MyResponse = MyResponse + "<input id=LeftUpDesktopButton type=button value=LeftUp>, "
	MyResponse = MyResponse + "<input id=RightClickDesktopButton type=button value=RightClick>, "
	MyResponse = MyResponse + "<input id=RightDownDesktopButton type=button value=RightDown>, "
	MyResponse = MyResponse + "<input id=RightUpDesktopButton type=button value=RightUp>.</p>"
	MyResponse = MyResponse + "<p>Keyboard: KeyCode = <input type=text id=KeyCodeDesktop size=1 value=" + cstr(KeyCodeDesktop) + ">, "
	MyResponse = MyResponse + "<input id=KeyPressDesktopButton type=button value=KeyPress>, "
	MyResponse = MyResponse + "<input id=KeyDownDesktopButton type=button value=KeyDown>, "
	MyResponse = MyResponse + "<input id=KeyUpDesktopButton type=button value=KeyUp>.</p>"
	MyResponse = MyResponse + "<hr></div>"
	MyResponse = MyResponse + "<div>Window operation functions without binding:"
	if MousePointWindowCheckBox then
		MyResponse = MyResponse + "<p><input id=MousePointedWindow type=checkbox checked>Watch mouse pointed window information by pressing Q + 2.</p>"
	else
		MyResponse = MyResponse + "<p><input id=MousePointedWindow type=checkbox>Watch mouse pointed window information by pressing Q + 2.</p>"
	end if
	MyResponse = MyResponse + "<p>Window: Handle = <input type=text id=WindowHandleWithoutBind size=4 value=" + cstr(WindowHandleWithoutBind) + ">, X = <input type=text id=WindowXWithoutBind size=1 value=" + cstr(WindowXWithoutBind) + ">, Y = <input type=text id=WindowYWithoutBind size=1 value=" + cstr(MouseYWithoutBind) + ">, <input id=MoveWindowWithoutBind type=button value=MoveWindow>, Width = <input type=text id=ClientWidthWithoutBind size=1 value=" + cstr(ClientWidthWithoutBind) + ">, Height = <input type=text id=ClientHeightWithoutBind size=1 value=" + cstr(ClientHeightWithoutBind) + ">, <input id=SetClientSize type=button value=SetSize>, "
	if SetWindowStateSelection = "CloseWindow" then
	MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"" selected>CloseWindow</option><option value=""ActivateWindow"">ActivateWindow</option><option value=""MinimizeWindow"">MinimizeWindow</option><option value=""MaximizeDeactivate"">MaximizeDeactivate</option><option value=""MaximizeWindow"">MaximizeWindow</option><option value=""RestoreWindow"">RestoreWindow</option><option value=""HideWindow"">HideWindow</option><option value=""ShowWindow"">ShowWindow</option><option value=""EnableTopFixed"">EnableTopFixed</option><option value=""DisableTopFixed"">DisableTopFixed</option><option value=""EnableIsolated"">EnableIsolated</option><option value=""DisableIsolation"">DisableIsolation</option><option value=""RestoreActivate"">RestoreActivate</option><option value=""KillWindow"">KillWindow</option><option value=""BlinkWindow"">BlinkWindow</option><option value=""FocusWindow"">FocusWindow</option><option value=""EnableFakeActivated"">EnableFakeActivated</option><option value=""DisableFakeActivated"">DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"
	elseif SetWindowStateSelection = "ActivateWindow" then
	MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"">CloseWindow</option><option value=""ActivateWindow"" selected>ActivateWindow</option><option value=""MinimizeWindow"">MinimizeWindow</option><option value=""MaximizeDeactivate"">MaximizeDeactivate</option><option value=""MaximizeWindow"">MaximizeWindow</option><option value=""RestoreWindow"">RestoreWindow</option><option value=""HideWindow"">HideWindow</option><option value=""ShowWindow"">ShowWindow</option><option value=""EnableTopFixed"">EnableTopFixed</option><option value=""DisableTopFixed"">DisableTopFixed</option><option value=""EnableIsolated"">EnableIsolated</option><option value=""DisableIsolation"">DisableIsolation</option><option value=""RestoreActivate"">RestoreActivate</option><option value=""KillWindow"">KillWindow</option><option value=""BlinkWindow"">BlinkWindow</option><option value=""FocusWindow"">FocusWindow</option><option value=""EnableFakeActivated"">EnableFakeActivated</option><option value=""DisableFakeActivated"">DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"
	elseif SetWindowStateSelection = "MinimizeWindow" then
		MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"">CloseWindow</option><option value=""ActivateWindow"">ActivateWindow</option><option value=""MinimizeWindow"" selected>MinimizeWindow</option><option value=""MaximizeDeactivate"">MaximizeDeactivate</option><option value=""MaximizeWindow"">MaximizeWindow</option><option value=""RestoreWindow"">RestoreWindow</option><option value=""HideWindow"">HideWindow</option><option value=""ShowWindow"">ShowWindow</option><option value=""EnableTopFixed"">EnableTopFixed</option><option value=""DisableTopFixed"">DisableTopFixed</option><option value=""EnableIsolated"">EnableIsolated</option><option value=""DisableIsolation"">DisableIsolation</option><option value=""RestoreActivate"">RestoreActivate</option><option value=""KillWindow"">KillWindow</option><option value=""BlinkWindow"">BlinkWindow</option><option value=""FocusWindow"">FocusWindow</option><option value=""EnableFakeActivated"">EnableFakeActivated</option><option value=""DisableFakeActivated"">DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"
	elseif SetWindowStateSelection = "MaximizeDeactivate" then
		MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"">CloseWindow</option><option value=""ActivateWindow"">ActivateWindow</option><option value=""MinimizeWindow"">MinimizeWindow</option><option value=""MaximizeDeactivate"" selected>MaximizeDeactivate</option><option value=""MaximizeWindow"">MaximizeWindow</option><option value=""RestoreWindow"">RestoreWindow</option><option value=""HideWindow"">HideWindow</option><option value=""ShowWindow"">ShowWindow</option><option value=""EnableTopFixed"">EnableTopFixed</option><option value=""DisableTopFixed"">DisableTopFixed</option><option value=""EnableIsolated"">EnableIsolated</option><option value=""DisableIsolation"">DisableIsolation</option><option value=""RestoreActivate"">RestoreActivate</option><option value=""KillWindow"">KillWindow</option><option value=""BlinkWindow"">BlinkWindow</option><option value=""FocusWindow"">FocusWindow</option><option value=""EnableFakeActivated"">EnableFakeActivated</option><option value=""DisableFakeActivated"">DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"
	elseif SetWindowStateSelection = "MaximizeWindow" then
		MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"">CloseWindow</option><option value=""ActivateWindow"">ActivateWindow</option><option value=""MinimizeWindow"">MinimizeWindow</option><option value=""MaximizeDeactivate"">MaximizeDeactivate</option><option value=""MaximizeWindow"" selected>MaximizeWindow</option><option value=""RestoreWindow"">RestoreWindow</option><option value=""HideWindow"">HideWindow</option><option value=""ShowWindow"">ShowWindow</option><option value=""EnableTopFixed"">EnableTopFixed</option><option value=""DisableTopFixed"">DisableTopFixed</option><option value=""EnableIsolated"">EnableIsolated</option><option value=""DisableIsolation"">DisableIsolation</option><option value=""RestoreActivate"">RestoreActivate</option><option value=""KillWindow"">KillWindow</option><option value=""BlinkWindow"">BlinkWindow</option><option value=""FocusWindow"">FocusWindow</option><option value=""EnableFakeActivated"">EnableFakeActivated</option><option value=""DisableFakeActivated"">DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"
	elseif SetWindowStateSelection = "RestoreWindow" then
		MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"">CloseWindow</option><option value=""ActivateWindow"">ActivateWindow</option><option value=""MinimizeWindow"">MinimizeWindow</option><option value=""MaximizeDeactivate"">MaximizeDeactivate</option><option value=""MaximizeWindow"">MaximizeWindow</option><option value=""RestoreWindow"" selected>RestoreWindow</option><option value=""HideWindow"">HideWindow</option><option value=""ShowWindow"">ShowWindow</option><option value=""EnableTopFixed"">EnableTopFixed</option><option value=""DisableTopFixed"">DisableTopFixed</option><option value=""EnableIsolated"">EnableIsolated</option><option value=""DisableIsolation"">DisableIsolation</option><option value=""RestoreActivate"">RestoreActivate</option><option value=""KillWindow"">KillWindow</option><option value=""BlinkWindow"">BlinkWindow</option><option value=""FocusWindow"">FocusWindow</option><option value=""EnableFakeActivated"">EnableFakeActivated</option><option value=""DisableFakeActivated"">DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"
	elseif SetWindowStateSelection = "HideWindow" then
		MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"">CloseWindow</option><option value=""ActivateWindow"">ActivateWindow</option><option value=""MinimizeWindow"">MinimizeWindow</option><option value=""MaximizeDeactivate"">MaximizeDeactivate</option><option value=""MaximizeWindow"">MaximizeWindow</option><option value=""RestoreWindow"">RestoreWindow</option><option value=""HideWindow"" selected>HideWindow</option><option value=""ShowWindow"">ShowWindow</option><option value=""EnableTopFixed"">EnableTopFixed</option><option value=""DisableTopFixed"">DisableTopFixed</option><option value=""EnableIsolated"">EnableIsolated</option><option value=""DisableIsolation"">DisableIsolation</option><option value=""RestoreActivate"">RestoreActivate</option><option value=""KillWindow"">KillWindow</option><option value=""BlinkWindow"">BlinkWindow</option><option value=""FocusWindow"">FocusWindow</option><option value=""EnableFakeActivated"">EnableFakeActivated</option><option value=""DisableFakeActivated"">DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"
	elseif SetWindowStateSelection = "ShowWindow" then
		MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"">CloseWindow</option><option value=""ActivateWindow"">ActivateWindow</option><option value=""MinimizeWindow"">MinimizeWindow</option><option value=""MaximizeDeactivate"">MaximizeDeactivate</option><option value=""MaximizeWindow"">MaximizeWindow</option><option value=""RestoreWindow"">RestoreWindow</option><option value=""HideWindow"">HideWindow</option><option value=""ShowWindow"" selected>ShowWindow</option><option value=""EnableTopFixed"">EnableTopFixed</option><option value=""DisableTopFixed"">DisableTopFixed</option><option value=""EnableIsolated"">EnableIsolated</option><option value=""DisableIsolation"">DisableIsolation</option><option value=""RestoreActivate"">RestoreActivate</option><option value=""KillWindow"">KillWindow</option><option value=""BlinkWindow"">BlinkWindow</option><option value=""FocusWindow"">FocusWindow</option><option value=""EnableFakeActivated"">EnableFakeActivated</option><option value=""DisableFakeActivated"">DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"
	elseif SetWindowStateSelection = "EnableTopFixed" then
		MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"">CloseWindow</option><option value=""ActivateWindow"">ActivateWindow</option><option value=""MinimizeWindow"">MinimizeWindow</option><option value=""MaximizeDeactivate"">MaximizeDeactivate</option><option value=""MaximizeWindow"">MaximizeWindow</option><option value=""RestoreWindow"">RestoreWindow</option><option value=""HideWindow"">HideWindow</option><option value=""ShowWindow"">ShowWindow</option><option value=""EnableTopFixed"" selected>EnableTopFixed</option><option value=""DisableTopFixed"">DisableTopFixed</option><option value=""EnableIsolated"">EnableIsolated</option><option value=""DisableIsolation"">DisableIsolation</option><option value=""RestoreActivate"">RestoreActivate</option><option value=""KillWindow"">KillWindow</option><option value=""BlinkWindow"">BlinkWindow</option><option value=""FocusWindow"">FocusWindow</option><option value=""EnableFakeActivated"">EnableFakeActivated</option><option value=""DisableFakeActivated"">DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"
	elseif SetWindowStateSelection = "DisableTopFixed" then
		MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"">CloseWindow</option><option value=""ActivateWindow"">ActivateWindow</option><option value=""MinimizeWindow"">MinimizeWindow</option><option value=""MaximizeDeactivate"">MaximizeDeactivate</option><option value=""MaximizeWindow"">MaximizeWindow</option><option value=""RestoreWindow"">RestoreWindow</option><option value=""HideWindow"">HideWindow</option><option value=""ShowWindow"">ShowWindow</option><option value=""EnableTopFixed"">EnableTopFixed</option><option value=""DisableTopFixed"" selected>DisableTopFixed</option><option value=""EnableIsolated"">EnableIsolated</option><option value=""DisableIsolation"">DisableIsolation</option><option value=""RestoreActivate"">RestoreActivate</option><option value=""KillWindow"">KillWindow</option><option value=""BlinkWindow"">BlinkWindow</option><option value=""FocusWindow"">FocusWindow</option><option value=""EnableFakeActivated"">EnableFakeActivated</option><option value=""DisableFakeActivated"">DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"
	elseif SetWindowStateSelection = "EnableIsolated" then
		MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"">CloseWindow</option><option value=""ActivateWindow"">ActivateWindow</option><option value=""MinimizeWindow"">MinimizeWindow</option><option value=""MaximizeDeactivate"">MaximizeDeactivate</option><option value=""MaximizeWindow"">MaximizeWindow</option><option value=""RestoreWindow"">RestoreWindow</option><option value=""HideWindow"">HideWindow</option><option value=""ShowWindow"">ShowWindow</option><option value=""EnableTopFixed"">EnableTopFixed</option><option value=""DisableTopFixed"">DisableTopFixed</option><option value=""EnableIsolated"" selected>EnableIsolated</option><option value=""DisableIsolated"">DisableIsolated</option><option value=""RestoreActivate"">RestoreActivate</option><option value=""KillWindow"">KillWindow</option><option value=""BlinkWindow"">BlinkWindow</option><option value=""FocusWindow"">FocusWindow</option><option value=""EnableFakeActivated"">EnableFakeActivated</option><option value=""DisableFakeActivated"">DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"
	elseif SetWindowStateSelection = "DisableIsolated" then
		MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"">CloseWindow</option><option value=""ActivateWindow"">ActivateWindow</option><option value=""MinimizeWindow"">MinimizeWindow</option><option value=""MaximizeDeactivate"">MaximizeDeactivate</option><option value=""MaximizeWindow"">MaximizeWindow</option><option value=""RestoreWindow"">RestoreWindow</option><option value=""HideWindow"">HideWindow</option><option value=""ShowWindow"">ShowWindow</option><option value=""EnableTopFixed"">EnableTopFixed</option><option value=""DisableTopFixed"">DisableTopFixed</option><option value=""EnableIsolated"">EnableIsolated</option><option value=""DisableIsolated"" selected>DisableIsolated</option><option value=""RestoreActivate"">RestoreActivate</option><option value=""KillWindow"">KillWindow</option><option value=""BlinkWindow"">BlinkWindow</option><option value=""FocusWindow"">FocusWindow</option><option value=""EnableFakeActivated"">EnableFakeActivated</option><option value=""DisableFakeActivated"">DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"
	elseif SetWindowStateSelection = "RestoreActivate" then
		MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"">CloseWindow</option><option value=""ActivateWindow"">ActivateWindow</option><option value=""MinimizeWindow"">MinimizeWindow</option><option value=""MaximizeDeactivate"">MaximizeDeactivate</option><option value=""MaximizeWindow"">MaximizeWindow</option><option value=""RestoreWindow"">RestoreWindow</option><option value=""HideWindow"">HideWindow</option><option value=""ShowWindow"">ShowWindow</option><option value=""EnableTopFixed"">EnableTopFixed</option><option value=""DisableTopFixed"">DisableTopFixed</option><option value=""EnableIsolated"">EnableIsolated</option><option value=""DisableIsolated"">DisableIsolated</option><option value=""RestoreActivate"" selected>RestoreActivate</option><option value=""KillWindow"">KillWindow</option><option value=""BlinkWindow"">BlinkWindow</option><option value=""FocusWindow"">FocusWindow</option><option value=""EnableFakeActivated"">EnableFakeActivated</option><option value=""DisableFakeActivated"">DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"	
	elseif SetWindowStateSelection = "KillWindow" then
		MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"">CloseWindow</option><option value=""ActivateWindow"">ActivateWindow</option><option value=""MinimizeWindow"">MinimizeWindow</option><option value=""MaximizeDeactivate"">MaximizeDeactivate</option><option value=""MaximizeWindow"">MaximizeWindow</option><option value=""RestoreWindow"">RestoreWindow</option><option value=""HideWindow"">HideWindow</option><option value=""ShowWindow"">ShowWindow</option><option value=""EnableTopFixed"">EnableTopFixed</option><option value=""DisableTopFixed"">DisableTopFixed</option><option value=""EnableIsolated"">EnableIsolated</option><option value=""DisableIsolated"">DisableIsolated</option><option value=""RestoreActivate"">RestoreActivate</option><option value=""KillWindow"" selected>KillWindow</option><option value=""BlinkWindow"">BlinkWindow</option><option value=""FocusWindow"">FocusWindow</option><option value=""EnableFakeActivated"">EnableFakeActivated</option><option value=""DisableFakeActivated"">DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"	
	elseif SetWindowStateSelection = "BlinkWindow" then
		MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"">CloseWindow</option><option value=""ActivateWindow"">ActivateWindow</option><option value=""MinimizeWindow"">MinimizeWindow</option><option value=""MaximizeDeactivate"">MaximizeDeactivate</option><option value=""MaximizeWindow"">MaximizeWindow</option><option value=""RestoreWindow"">RestoreWindow</option><option value=""HideWindow"">HideWindow</option><option value=""ShowWindow"">ShowWindow</option><option value=""EnableTopFixed"">EnableTopFixed</option><option value=""DisableTopFixed"">DisableTopFixed</option><option value=""EnableIsolated"">EnableIsolated</option><option value=""DisableIsolated"">DisableIsolated</option><option value=""RestoreActivate"">RestoreActivate</option><option value=""KillWindow"">KillWindow</option><option value=""BlinkWindow"" selected>BlinkWindow</option><option value=""FocusWindow"">FocusWindow</option><option value=""EnableFakeActivated"">EnableFakeActivated</option><option value=""DisableFakeActivated"">DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"	
	elseif SetWindowStateSelection = "FocusWindow" then
		MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"">CloseWindow</option><option value=""ActivateWindow"">ActivateWindow</option><option value=""MinimizeWindow"">MinimizeWindow</option><option value=""MaximizeDeactivate"">MaximizeDeactivate</option><option value=""MaximizeWindow"">MaximizeWindow</option><option value=""RestoreWindow"">RestoreWindow</option><option value=""HideWindow"">HideWindow</option><option value=""ShowWindow"">ShowWindow</option><option value=""EnableTopFixed"">EnableTopFixed</option><option value=""DisableTopFixed"">DisableTopFixed</option><option value=""EnableIsolated"">EnableIsolated</option><option value=""DisableIsolated"">DisableIsolated</option><option value=""RestoreActivate"">RestoreActivate</option><option value=""KillWindow"">KillWindow</option><option value=""BlinkWindow"">BlinkWindow</option><option value=""FocusWindow"" selected>FocusWindow</option><option value=""EnableFakeActivated"">EnableFakeActivated</option><option value=""DisableFakeActivated"">DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"	
	elseif SetWindowStateSelection = "EnableFakeActivated" then
		MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"">CloseWindow</option><option value=""ActivateWindow"">ActivateWindow</option><option value=""MinimizeWindow"">MinimizeWindow</option><option value=""MaximizeDeactivate"">MaximizeDeactivate</option><option value=""MaximizeWindow"">MaximizeWindow</option><option value=""RestoreWindow"">RestoreWindow</option><option value=""HideWindow"">HideWindow</option><option value=""ShowWindow"">ShowWindow</option><option value=""EnableTopFixed"">EnableTopFixed</option><option value=""DisableTopFixed"">DisableTopFixed</option><option value=""EnableIsolated"">EnableIsolated</option><option value=""DisableIsolated"">DisableIsolated</option><option value=""RestoreActivate"">RestoreActivate</option><option value=""KillWindow"">KillWindow</option><option value=""BlinkWindow"">BlinkWindow</option><option value=""FocusWindow"">FocusWindow</option><option value=""EnableFakeActivated"" selected>EnableFakeActivated</option><option value=""DisableFakeActivated"">DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"	
	elseif SetWindowStateSelection = "DisableFakeActivated" then
		MyResponse = MyResponse + "<select id=SetWindowStateSelection><option value=""CloseWindow"">CloseWindow</option><option value=""ActivateWindow"">ActivateWindow</option><option value=""MinimizeWindow"">MinimizeWindow</option><option value=""MaximizeDeactivate"">MaximizeDeactivate</option><option value=""MaximizeWindow"">MaximizeWindow</option><option value=""RestoreWindow"">RestoreWindow</option><option value=""HideWindow"">HideWindow</option><option value=""ShowWindow"">ShowWindow</option><option value=""EnableTopFixed"">EnableTopFixed</option><option value=""DisableTopFixed"">DisableTopFixed</option><option value=""EnableIsolated"">EnableIsolated</option><option value=""DisableIsolated"">DisableIsolated</option><option value=""RestoreActivate"">RestoreActivate</option><option value=""KillWindow"">KillWindow</option><option value=""BlinkWindow"">BlinkWindow</option><option value=""FocusWindow"">FocusWindow</option><option value=""EnableFakeActivated"">EnableFakeActivated</option><option value=""DisableFakeActivated"" selected>DisableFakeActivated</option></select>, <input id=SetWindowState type=button value=SetState>.</p></div>"	
	end if
	if MouseMessageWithoutBind = "MoveTo" then
		MyResponse = MyResponse + "<p>Mouse: Handle = <input type=text id=MouseHandleWithoutBind size=4 value=" + cstr(MouseHandleWithoutBind) + ">, X = <input type=text id=MouseXWithoutBind size=1 value=" + cstr(MouseXWithoutBind) + ">, Y = <input type=text id=MouseYWithoutBind size=1 value=" + cstr(MouseYWithoutBind) + ">, option = <select id=MouseMessageWithoutBind><option value=""MoveTo"" selected>MoveTo</option><option value=""LeftClick"">LeftClick</option><option value=""LeftDown"">LeftDown</option><option value=""LeftUp"">LeftUp</option><option value=""RightClick"">RightClick</option><option value=""RightDown"">RightDown</option><option value=""RightUp"">RightUp</option></select>, <input id=MouseWithoutBind type=button value=Send>.</p></div>"
	elseif MouseMessageWithoutBind = "LeftClick" then
		MyResponse = MyResponse + "<p>Mouse: Handle = <input type=text id=MouseHandleWithoutBind size=4 value=" + cstr(MouseHandleWithoutBind) + ">, X = <input type=text id=MouseXWithoutBind size=1 value=" + cstr(MouseXWithoutBind) + ">, Y = <input type=text id=MouseYWithoutBind size=1 value=" + cstr(MouseYWithoutBind) + ">, option = <select id=MouseMessageWithoutBind><option value=""MoveTo"">MoveTo</option><option value=""LeftClick"" selected>LeftClick</option><option value=""LeftDown"">LeftDown</option><option value=""LeftUp"">LeftUp</option><option value=""RightClick"">RightClick</option><option value=""RightDown"">RightDown</option><option value=""RightUp"">RightUp</option></select>, <input id=MouseWithoutBind type=button value=Send>.</p></div>"
	elseif MouseMessageWithoutBind = "LeftDown" then
		MyResponse = MyResponse + "<p>Mouse: Handle = <input type=text id=MouseHandleWithoutBind size=4 value=" + cstr(MouseHandleWithoutBind) + ">, X = <input type=text id=MouseXWithoutBind size=1 value=" + cstr(MouseXWithoutBind) + ">, Y = <input type=text id=MouseYWithoutBind size=1 value=" + cstr(MouseYWithoutBind) + ">, option = <select id=MouseMessageWithoutBind><option value=""MoveTo"">MoveTo</option><option value=""LeftClick"">LeftClick</option><option value=""LeftDown"" selected>LeftDown</option><option value=""LeftUp"">LeftUp</option><option value=""RightClick"">RightClick</option><option value=""RightDown"">RightDown</option><option value=""RightUp"">RightUp</option></select>, <input id=MouseWithoutBind type=button value=Send>.</p></div>"
	elseif MouseMessageWithoutBind = "LeftUp" then
		MyResponse = MyResponse + "<p>Mouse: Handle = <input type=text id=MouseHandleWithoutBind size=4 value=" + cstr(MouseHandleWithoutBind) + ">, X = <input type=text id=MouseXWithoutBind size=1 value=" + cstr(MouseXWithoutBind) + ">, Y = <input type=text id=MouseYWithoutBind size=1 value=" + cstr(MouseYWithoutBind) + ">, option = <select id=MouseMessageWithoutBind><option value=""MoveTo"">MoveTo</option><option value=""LeftClick"">LeftClick</option><option value=""LeftDown"">LeftDown</option><option value=""LeftUp"" selected>LeftUp</option><option value=""RightClick"">RightClick</option><option value=""RightDown"">RightDown</option><option value=""RightUp"">RightUp</option></select>, <input id=MouseWithoutBind type=button value=Send>.</p></div>"
	elseif MouseMessageWithoutBind = "RightClick" then
		MyResponse = MyResponse + "<p>Mouse: Handle = <input type=text id=MouseHandleWithoutBind size=4 value=" + cstr(MouseHandleWithoutBind) + ">, X = <input type=text id=MouseXWithoutBind size=1 value=" + cstr(MouseXWithoutBind) + ">, Y = <input type=text id=MouseYWithoutBind size=1 value=" + cstr(MouseYWithoutBind) + ">, option = <select id=MouseMessageWithoutBind><option value=""MoveTo"">MoveTo</option><option value=""LeftClick"">LeftClick</option><option value=""LeftDown"">LeftDown</option><option value=""LeftUp"">LeftUp</option><option value=""RightClick"" selected>RightClick</option><option value=""RightDown"">RightDown</option><option value=""RightUp"">RightUp</option></select>, <input id=MouseWithoutBind type=button value=Send>.</p></div>"
	elseif MouseMessageWithoutBind = "RightDown" then
		MyResponse = MyResponse + "<p>Mouse: Handle = <input type=text id=MouseHandleWithoutBind size=4 value=" + cstr(MouseHandleWithoutBind) + ">, X = <input type=text id=MouseXWithoutBind size=1 value=" + cstr(MouseXWithoutBind) + ">, Y = <input type=text id=MouseYWithoutBind size=1 value=" + cstr(MouseYWithoutBind) + ">, option = <select id=MouseMessageWithoutBind><option value=""MoveTo"">MoveTo</option><option value=""LeftClick"">LeftClick</option><option value=""LeftDown"">LeftDown</option><option value=""LeftUp"">LeftUp</option><option value=""RightClick"">RightClick</option><option value=""RightDown"" selected>RightDown</option><option value=""RightUp"">RightUp</option></select>, <input id=MouseWithoutBind type=button value=Send>.</p></div>"
	elseif MouseMessageWithoutBind = "RightUp" then
		MyResponse = MyResponse + "<p>Mouse: Handle = <input type=text id=MouseHandleWithoutBind size=4 value=" + cstr(MouseHandleWithoutBind) + ">, X = <input type=text id=MouseXWithoutBind size=1 value=" + cstr(MouseXWithoutBind) + ">, Y = <input type=text id=MouseYWithoutBind size=1 value=" + cstr(MouseYWithoutBind) + ">, option = <select id=MouseMessageWithoutBind><option value=""MoveTo"">MoveTo</option><option value=""LeftClick"">LeftClick</option><option value=""LeftDown"">LeftDown</option><option value=""LeftUp"">LeftUp</option><option value=""RightClick"">RightClick</option><option value=""RightDown"">RightDown</option><option value=""RightUp"" selected>RightUp</option></select>, <input id=MouseWithoutBind type=button value=Send>.</p></div>"
	end if
	if KeyboardMessageWithoutBind = "KeyPress" then
		MyResponse = MyResponse + "<p>Keyboard: Handle = <input type=text id=KeyboardHandleWithoutBind size=4 value=" + cstr(KeyboardHandleWithoutBind) + ">, KeyCode = <input type=text id=KeyCodeWithoutBind size=1 value=" + cstr(KeyCodeWithoutBind) + "> option = <select id=KeyboardMessageWithoutBind><option value=""KeyPress"" selected>KeyPress</option><option value=""KeyDown"">KeyDown</option><option value=""KeyUp"">KeyUp</option><option value=""LeftUp"">LeftUp</option></select>, <input id=KeyboardWithoutBind type=button value=Send>.</p></div>"
	elseif KeyboardMessageWithoutBind = "KeyDown" then
		MyResponse = MyResponse + "<p>Keyboard: Handle = <input type=text id=KeyboardHandleWithoutBind size=4 value=" + cstr(KeyboardHandleWithoutBind) + ">, KeyCode = <input type=text id=KeyCodeWithoutBind size=1 value=" + cstr(KeyCodeWithoutBind) + "> option = <select id=KeyboardMessageWithoutBind><option value=""KeyPress"">KeyPress</option><option value=""KeyDown"" selected>KeyDown</option><option value=""KeyUp"">KeyUp</option><option value=""LeftUp"">LeftUp</option></select>, <input id=KeyboardWithoutBind type=button value=Send>.</p></div>"
	elseif KeyboardMessageWithoutBind = "KeyUp" then
		MyResponse = MyResponse + "<p>Keyboard: Handle = <input type=text id=KeyboardHandleWithoutBind size=4 value=" + cstr(KeyboardHandleWithoutBind) + ">, KeyCode = <input type=text id=KeyCodeWithoutBind size=1 value=" + cstr(KeyCodeWithoutBind) + "> option = <select id=KeyboardMessageWithoutBind><option value=""KeyPress"">KeyPress</option><option value=""KeyDown"">KeyDown</option><option value=""KeyUp"">KeyUp</option><option value=""LeftUp"" selected>LeftUp</option></select>, <input id=KeyboardWithoutBind type=button value=Send>.</p></div>"
	end if
	MyResponse = MyResponse + "<hr></div>"
	MyResponse = MyResponse + "<div>Window operation functions with binding:<br>"
	MyResponse = MyResponse + "<p>Login: CDKey = <input type=text id=CDKey size=20 value=" + CDKey + ">, <input id=Login type=button value=Login>, QuarkAIID = " + cstr(QuarkAIID) + ".</p>"
	MyResponse = MyResponse + "<p>Window: Handle = <input type=text id=WindowHandleWithBind size=4 value=" + cstr(WindowHandleWithBind) + ">, "
	if GraphicsMode = "normal" then
		MyResponse = MyResponse + "GraphicsMode = <select id=GraphicsMode><option value=""normal"" selected>normal</option><option value=""gdi"">gdi</option><option value=""dx"">dx</option><option value=""dx2"">dx2</option></select>, "
	elseif GraphicsMode = "gdi" then
		MyResponse = MyResponse + "GraphicsMode = <select id=GraphicsMode><option value=""normal"">normal</option><option value=""gdi"" selected>gdi</option><option value=""dx"">dx</option><option value=""dx2"">dx2</option></select>, "
	elseif GraphicsMode = "dx" then
		MyResponse = MyResponse + "GraphicsMode = <select id=GraphicsMode><option value=""normal"">normal</option><option value=""gdi"">gdi</option><option value=""dx"" selected>dx</option><option value=""dx2"">dx2</option></select>, "
	elseif GraphicsMode = "dx2" then
		MyResponse = MyResponse + "GraphicsMode = <select id=GraphicsMode><option value=""normal"">normal</option><option value=""gdi"">gdi</option><option value=""dx"">dx</option><option value=""dx2"" selected>dx2</option></select>, "
	end if
	if MouseMode = "normal" then
		MyResponse = MyResponse + "MouseMode = <select id=MouseMode><option value=""normal"" selected>normal</option><option value=""windows"">windows</option><option value=""dx"">dx</option></select>, "
	elseif MouseMode = "windows" then
		MyResponse = MyResponse + "MouseMode = <select id=MouseMode><option value=""normal"">normal</option><option value=""windows"" selected>windows</option><option value=""dx"">dx</option></select>, "
	elseif MouseMode = "dx" then
		MyResponse = MyResponse + "MouseMode = <select id=MouseMode><option value=""normal"">normal</option><option value=""windows"">windows</option><option value=""dx"" selected>dx</option></select>, "
	end if
	if KeyboardMode = "normal" then
		MyResponse = MyResponse + "KeyboardMode = <select id=KeyboardMode><option value=""normal""  selected>normal</option><option value=""windows"">windows</option><option value=""dx"">dx</option></select>, "
	elseif KeyboardMode = "windows" then
		MyResponse = MyResponse + "KeyboardMode = <select id=KeyboardMode><option value=""normal"">normal</option><option value=""windows""  selected>windows</option><option value=""dx"">dx</option></select>, "
	elseif KeyboardMode = "dx" then
		MyResponse = MyResponse + "KeyboardMode = <select id=KeyboardMode><option value=""normal"">normal</option><option value=""windows"">windows</option><option value=""dx""  selected>dx</option></select>, "
	end if
	MyResponse = MyResponse + "BindMode = <input type=text id=BindMode size=1 value=" + cstr(BindMode) + ">, "
	if BindFlag then
		MyResponse = MyResponse + "<input id=BindButton type=button value=UnBind>.</p></div>"
	else		
		MyResponse = MyResponse + "<input id=BindButton type=button value=Bind>.</p></div>"
	end if
	MyResponse = MyResponse + "<p>Graphic: X = <input type=text id=ColorXWithBind size=1 value=" + cstr(ColorXWithBind) + ">, Y = <input type=text id=ColorYWithBind size=1 value=" + cstr(ColorYWithBind) + ">, <input id=GetColor type=button value=GetColor>, PictureName = <input type=text id=PictureName size=8 value=" + PictureName + ">, ColorTolerance = <input type=text id=ColorToleranceWithBind size=4 value=" + ColorToleranceWithBind + ">, Simlarity = <input type=text id=SimlarityWithBind size=1 value=" + cstr(SimlarityWithBind) + ">, <input id=FindPictures type=button value=FindPictures>, <input id=Capture type=button value=Capture>.</p>"
	MyResponse = MyResponse + "<p>Mouse: X = <input type=text id=MouseXWithBind size=1 value=" + cstr(MouseXWithBind) + ">, Y = <input type=text id=MouseYWithBind size=1 value=" + cstr(MouseYWithBind) + ">, <input id=MoveToButton type=button value=MoveTo>, "
	MyResponse = MyResponse + "<input id=LeftClickButton type=button value=LeftClick>, "
	MyResponse = MyResponse + "<input id=LeftDownButton type=button value=LeftDown>, "
	MyResponse = MyResponse + "<input id=LeftUpButton type=button value=LeftUp>, "
	MyResponse = MyResponse + "<input id=RightClickButton type=button value=RightClick>, "
	MyResponse = MyResponse + "<input id=RightDownButton type=button value=RightDown>, "
	MyResponse = MyResponse + "<input id=RightUpButton type=button value=RightUp>.</p>"
	MyResponse = MyResponse + "<p>Keyboard: KeyCode = <input type=text id=KeyCodeWithBind size=1 value=" + cstr(KeyCodeWithBind) + ">, "
	MyResponse = MyResponse + "<input id=KeyPressButton type=button value=KeyPress>, "
	MyResponse = MyResponse + "<input id=KeyDownButton type=button value=KeyDown>, "
	MyResponse = MyResponse + "<input id=KeyUpButton type=button value=KeyUp>.</p>"
	MyResponse = MyResponse + "<hr></div>"
	MyResponse = MyResponse + MyOutput
	MyResponse = MyResponse + "</body></html>"
	MyResult = ie.document.write(MyResponse)
	set id = ie.document.all
	id.GetColor.onclick=getref("GetColor")
	id.FindPictures.onclick=getref("FindPictures")
	id.Capture.onclick=getref("Capture")
	id.MoveWindowWithoutBind.onclick=getref("MoveWindowWithoutBind")
	id.SetClientSize.onclick=getref("SetClientSize")
	id.Version.onclick=getref("Version")
	id.BasePath.onclick=getref("BasePath")
	id.CurrentPath.onclick=getref("CurrentPath")
	id.MachineCode.onclick=getref("MachineCode")
	id.Magnifier.onclick=getref("Magnifier")
	id.SnippingTool.onclick=getref("SnippingTool")
	id.KeyCodeCheck.onclick=getref("KeyCodeCheck")
	id.MSPaint.onclick=getref("MSPaint")
	id.MousePointedPixel.onclick=getref("MousePointedPixel")
	id.MousePointedWindow.onclick=getref("MousePointedWindow")
	id.SetWindowState.onclick=getref("SetWindowState")
	id.MouseWithoutBind.onclick=getref("MouseWithoutBind")
	id.KeyboardWithoutBind.onclick=getref("KeyboardWithoutBind")
	id.Login.onclick=getref("Login")
	id.BindButton.onclick=getref("BindButton")
	id.MoveToButton.onclick=getref("MoveToButton")
	id.LeftClickButton.onclick=getref("LeftClickButton")
	id.LeftDownButton.onclick=getref("LeftDownButton")
	id.LeftUpButton.onclick=getref("LeftUpButton")
	id.RightClickButton.onclick=getref("RightClickButton")
	id.RightDownButton.onclick=getref("RightDownButton")
	id.RightUpButton.onclick=getref("RightUpButton")
	id.KeyPressButton.onclick=getref("KeyPressButton")
	id.KeyDownButton.onclick=getref("KeyDownButton")
	id.KeyUpButton.onclick=getref("KeyUpButton")
	id.MoveToDesktopButton.onclick=getref("MoveToDesktopButton")
	id.LeftClickDesktopButton.onclick=getref("LeftClickDesktopButton")
	id.LeftDownDesktopButton.onclick=getref("LeftDownDesktopButton")
	id.LeftUpDesktopButton.onclick=getref("LeftUpDesktopButton")
	id.RightClickDesktopButton.onclick=getref("RightClickDesktopButton")
	id.RightDownDesktopButton.onclick=getref("RightDownDesktopButton")
	id.RightUpDesktopButton.onclick=getref("RightUpDesktopButton")
	id.KeyPressDesktopButton.onclick=getref("KeyPressDesktopButton")
	id.KeyDownDesktopButton.onclick=getref("KeyDownDesktopButton")
	id.KeyUpDesktopButton.onclick=getref("KeyUpDesktopButton")
end sub
rewrite
while RunningFlag
	number = number + 1
	KeyPressQ = MyQuarkAI.GetKeyPress(81)
	KeyPress1 = MyQuarkAI.GetKeyPress(49)
	KeyPress2 = MyQuarkAI.GetKeyPress(50)
	if KeyPressQ = 1 and KeyPress1 = 1 and MousePointPixelCheckBox then
		MyResult = MyQuarkAI.GetMousePosition(MouseX, MouseY)
		MouseColor = MyQuarkAI.GetDesktopColor(MouseX, MouseY)
		MyOutput = "<div><p>Result:</p><p>Mouse Pointed Point: X = " + cstr(MouseX) + ",Y = " + cstr(MouseY) + ", Color = " + MouseColor + "</p></div>"
		rewrite
	elseif KeyPressQ = 1 and KeyPress2 = 1 and MousePointWindowCheckBox then
		WindowHandle = MyQuarkAI.GetMousePointWindow()
		WindowTitle = MyQuarkAI.GetWindowTitle(WindowHandle)
		WindowClass = MyQuarkAI.GetWindowClass(WindowHandle)
		MyResult = MyQuarkAI.GetWindowRect(WindowHandle, WindowX1, WindowY1, WindowX2, WindowY2)
		MyResult = MyQuarkAI.GetClientRect(WindowHandle, ClientX1, ClientY1, ClientX2, ClientY2)
		MyResult = MyQuarkAI.GetClientSize(WindowHandle, ClientWidth, ClientHeight)		
		ProcessID = MyQuarkAI.GetWindowProcessId(WindowHandle)
		ProcessInfomation = MyQuarkAI.GetProcessInfo(ProcessID)
		ParentHandle = MyQuarkAI.GetWindow(WindowHandle, 0)
		ParentTitle = MyQuarkAI.GetWindowTitle(ParentHandle)
		ParentClass = MyQuarkAI.GetWindowClass(ParentHandle)
		MyResult = MyQuarkAI.GetWindowRect(ParentHandle, ParentX1, ParentY1, ParentX2, ParentY2)
		ChildHandles = MyQuarkAI.EnumWindow(WindowHandle, "", "", 0)
		ChildNumber = MyQuarkAI.CountTotalStrings(ChildHandles, ",")
		MyOutput = "<div><p>Result:</p>This Window:</br>Handle = " + cstr(WindowHandle) + ", Title = " + cstr(WindowTitle)+ ", Class = " + cstr(WindowClass) + " Region = (" + cstr(WindowX1) + ", " + cstr(WindowY1) + ") &times (" + cstr(WindowX2) + ", "+ cstr(WindowY2) +").</br>"
		MyOutput = MyOutput + "Client Region = (" + cstr(ClientX1) + ", "+ cstr(ClientY1) + " ) &times (" + cstr(ClientX2) + ", "+ cstr(ClientY2) + "), Width =" + cstr(ClientWidth) + ". Height =" + cstr(ClientHeight) + ".</br>"
		MyOutput = MyOutput + "Process ID = " + cstr(ProcessID) + "</br> Process Information: ProcessName|PrceossPath|CPU|Memory = " + ProcessInfomation + ".</p>"
		MyOutput = MyOutput + "<p>Parent Window: </br>Handle = " + cstr(ParentHandle) + ", Title = " + cstr(ParentTitle)+ ", Class = " + cstr(ParentClass) + ", Region = (" + cstr(ChildX1) + ", "+ cstr(ChildY1)+") &times (" + cstr(ChildX2) + ", "+ cstr(ChildY2)+ ").</p>"
		MyOutput = MyOutput + "<p>Child windows:</br>"
		for ChildIndex = 0 to ChildNumber - 1
			ChildHandle = clng(MyQuarkAI.GetTargetString(ChildHandles, "," , ChildIndex))
			ChildTitle = MyQuarkAI.GetWindowTitle(ChildHandle)
			ChildClass = MyQuarkAI.GetWindowClass(ChildHandle)
			MyResult = MyQuarkAI.GetWindowRect(ChildHandle, ChildX1, ChildY1, ChildX2, ChildY2)
			MyOutput = MyOutput + "Handle = " + cstr(ChildHandle) + ", Title = " + cstr(ChildTitle)+ ", Class = " + cstr(ChildClass) + ", Region = (" + cstr(ChildX1) + ", " + cstr(ChildY1) + ") &times (" + cstr(ChildX2) + ", "+ cstr(ChildY2) +").</br>"
		next
		MyOutput = MyOutput + "</div>"
		rewrite
	end if
	MyResult = wscript.sleep(50)
wend

sub FindPictures
	PictureName = id.PictureName.value
	ColorToleranceWithBind  = id.ColorToleranceWithBind.value
	SimlarityWithBind  = cdbl(id.SimlarityWithBind.value)
	MyResult = MyQuarkAI.FindPicExS(0, 0, 2048, 2048, PictureName, ColorToleranceWithBind, SimlarityWithBind, 0)
	MyOutput = "<div><p>Result:</p><p>FindPicExS(0, 0, 2048, 2048,""" + PictureName+""", " + cstr(ColorToleranceWithBind) + ", " + cstr(SimlarityWithBind) + "0) = """ + MyResult + """.</p></div>"
	rewrite
end sub

sub Capture
	MyFileName = InputBox("Enter the file name:")
	MyFileName = MyFileName + ".bmp"	
	MyResult = MyQuarkAI.Capture(0, 0, 2048, 2048, MyFileName)
	if MyResult = 1 then
        MyOutput = "<div><p>Result:</p><p>Capture(0, 0, 2048, 2048, """ + MyFileName + """) = Success.</p></div>"		
	else
		MyOutput = "<div><p>Result:</p><p>Capture(0, 0, 2048, 2048, """ + MyFileName + """) = Failed.</p></div>"
	end if	
	rewrite
end sub

sub GetColor
	ColorXWithBind  = clng(id.ColorXWithBind.value)
	ColorYWithBind  = clng(id.ColorYWithBind.value)
	MyResult = MyQuarkAI.GetColor(ColorXWithBind, ColorYWithBind)
	MyOutput = "<div><p>Result:</p><p>GetColor(" + cstr(ColorXWithBind) + ", " + cstr(ColorYWithBind) + ") = " + MyResult + ".</p></div>"
	rewrite
end sub

sub MoveWindowWithoutBind
	WindowHandleWithoutBind  = clng(id.WindowHandleWithoutBind.value)
	WindowXWithoutBind  = clng(id.WindowXWithoutBind.value)
	WindowYWithoutBind  = clng(id.WindowYWithoutBind.value)
	MyResult = MyQuarkAI.MoveWindow(WindowHandleWithoutBind,WindowXWithoutBind,WindowYWithoutBind)
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>MoveWindowWithoutBind(" + cstr(WindowHandleWithoutBind) + ", "+ cstr(WindowXWithoutBind) + ", " +cstr(WindowYWithoutBind) + ") = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>MoveWindowWithoutBind(" + cstr(WindowHandleWithoutBind) + ", "+ cstr(WindowXWithoutBind) + ", " +cstr(WindowYWithoutBind) + ") = Failed.</p></div>"
	end if
	rewrite
end sub

sub SetClientSize
	WindowHandleWithoutBind  = clng(id.WindowHandleWithoutBind.value)
	ClientWidthWithoutBind  = clng(id.ClientWidthWithoutBind.value)
	ClientHeightWithoutBind  = clng(id.ClientHeightWithoutBind.value)
	MyResult = MyQuarkAI.SetClientSize(WindowHandleWithoutBind,ClientWidthWithoutBind,ClientHeightWithoutBind)
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>SetClientSize(" + cstr(WindowHandleWithoutBind) + ", "+ cstr(ClientWidthWithoutBind) + ", " +cstr(ClientHeightWithoutBind) + ") = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>SetClientSize(" + cstr(WindowHandleWithoutBind) + ", "+ cstr(ClientidthWithoutBind) + ", " +cstr(ClientHeightWithoutBind) + ") = Failed.</p></div>"
	end if
	rewrite
end sub

sub SetWindowState
	WindowHandleWithoutBind = clng(id.WindowHandleWithoutBind.value)
	SetWindowStateSelection = id.SetWindowStateSelection.value
	Dim MyState
	if SetWindowStateSelection = "CloseWindow" then
		MyState = 0
	elseif SetWindowStateSelection = "ActivateWindow" then
		MyState = 1
	elseif SetWindowStateSelection = "MinimizeWindow" then
		MyState = 2
	elseif SetWindowStateSelection = "MaximizeDeactivate" then
		MyState = 3
	elseif SetWindowStateSelection = "MaximizeWindow" then
		MyState = 4
	elseif SetWindowStateSelection = "RestoreWindow" then
		MyState = 5
	elseif SetWindowStateSelection = "HideWindow" then
		MyState = 6
	elseif SetWindowStateSelection = "ShowWindow" then
		MyState = 7
	elseif SetWindowStateSelection = "EnableTopFixed" then
		MyState = 8
	elseif SetWindowStateSelection = "DisableTopFixed" then
		MyState = 9
	elseif SetWindowStateSelection = "EnableIsolated" then
		MyState = 10
	elseif SetWindowStateSelection = "DisableIsolated" then
		MyState = 11
	elseif SetWindowStateSelection = "RestoreActivate" then
		MyState = 12
	elseif SetWindowStateSelection = "KillWindow" then
		MyState = 13	
	elseif SetWindowStateSelection = "BlinkWindow" then
		MyState = 14	
	elseif SetWindowStateSelection = "FocusWindow" then
		MyState = 15
	elseif SetWindowStateSelection = "EnableFakeActivated" then
		MyState = 16	
	elseif SetWindowStateSelection = "DisableFakeActivated" then
		MyState = 17
	end if	
	MyResult = MyQuarkAI.SetWindowState(WindowHandleWithoutBind, MyState)
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>SetWindowState("+cstr(WindowHandleWithoutBind)+", "+cstr(MyState)+") = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>SetWindowState("+cstr(WindowHandleWithoutBind)+", "+cstr(MyState)+") = Failed.</p></div>"
	end if
	rewrite
end sub	

sub KeyPressDesktopButton
	KeycodeDesktop  = clng(id.KeyCodeDesktop.value)
	MyResult = MyQuarkAI.SimKeyPress(KeycodeDesktop)
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>SimKeyPress(" + cstr(KeycodeDesktop ) + ") = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>SimKeyPress(" + cstr(KeycodeDesktop ) + ") = Failed.</p></div>"
	end if
	rewrite
end sub

sub KeyDownDesktopButton
	KeyCodeDesktop = clng(id.KeyCodeDesktop.value)
	MyResult = MyQuarkAI.SimKeyDown(KeyCodeDesktop)
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>SimKeyDown(" + cstr(KeyCodeDesktop) + ") = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>SimKeyDown(" + cstr(KeyCodeDesktop) + ") = Failed.</p></div>"
	end if
	rewrite
end sub

sub KeyUpDesktopButton
	KeyCodeDesktop = clng(id.KeyCodeDesktop.value)
	MyResult = MyQuarkAI.SimKeyUp(KeyCodeDesktop)
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>SimKeyUp(" + cstr(KeyCodeDesktop) + ") = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>SimKeyUp(" + cstr(KeyCodeDesktop) + ") = Failed.</p></div>"
	end if
	rewrite
end sub

sub MoveToDesktopButton
	MouseXDesktop = clng(id.MouseXDesktop.value)
	MouseYDesktop = clng(id.MouseYDesktop.value)
	MyResult = MyQuarkAI.SimMoveTo(MouseXDesktop, MouseYDesktop)
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>SimMoveTo(" + cstr(MouseXDesktop) + ", " + cstr(MouseYDesktop) + ") = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>SimMoveTo(" + cstr(MouseXDesktop) + ", " + cstr(MouseYDesktop) + ") = Failed.</p></div>"
	end if
	rewrite
end sub

sub LeftClickDesktopButton
	MyResult = MyQuarkAI.SimLeftClick()
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>SimLeftClick() = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>SimLeftClick() = Failed.</p></div>"
	end if
	rewrite
end sub

sub LeftDownDesktopButton
	MyResult = MyQuarkAI.SimLeftDown()
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>SimLeftDown() = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>SimLeftDown() = Failed.</p></div>"
	end if
	rewrite
end sub

sub LeftUpDesktopButton
	MyResult = MyQuarkAI.SimLeftUp()
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>SimLeftUp() = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>SimLeftUp() = Failed.</p></div>"
	end if
	rewrite
end sub

sub RightClickDesktopButton
	MyResult = MyQuarkAI.SimRightClick()
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>SimRightClick() = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>SimRightClick() = Failed.</p></div>"
	end if
	rewrite
end sub

sub RightDownDesktopButton
	MyResult = MyQuarkAI.SimRightDown()
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>SimRightDown() = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>SimRightDown() = Failed.</p></div>"
	end if
	rewrite
end sub

sub RightUpDesktopButton
	MyResult = MyQuarkAI.SimRightUp()
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>SimRightUp() = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>SimRightUp() = Failed.</p></div>"
	end if
	rewrite
end sub

sub KeyPressButton
	KeyCodeWithBind = clng(id.KeyCodeWithBind.value)
	MyResult = MyQuarkAI.KeyPress(KeyCodeWithBind)
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>KeyPress(" + cstr(KeyCodeWithBind) + ") = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>KeyPress(" + cstr(KeyCodeWithBind) + ") = Failed.</p></div>"
	end if
	rewrite
end sub

sub KeyDownButton
	KeyCodeWithBind = clng(id.KeyCodeWithBind.value)
	MyResult = MyQuarkAI.KeyDown(KeyCodeWithBind)
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>KeyDown(" + cstr(KeyCodeWithBind) + ") = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>KeyDown(" + cstr(KeyCodeWithBind) + ") = Failed.</p></div>"
	end if
	rewrite
end sub

sub KeyUpButton
	KeyCodeWithBind = clng(id.KeyCodeWithBind.value)
	MyResult = MyQuarkAI.KeyUp(KeyCodeWithBind)
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>KeyUp(" + cstr(KeyCodeWithBind) + ") = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>KeyUp(" + cstr(KeyCodeWithBind) + ") = Failed.</p></div>"
	end if
	rewrite
end sub

sub MoveToButton
	MouseXWithBind = clng(id.MouseXWithBind.value)
	MouseYWithBind = clng(id.MouseYWithBind.value)
	MyResult = MyQuarkAI.MoveTo(MouseXWithBind, MouseYWithBind)
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>MoveTo(" + cstr(MouseXWithBind) + ", " + cstr(MouseYWithBind) + ") = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>MoveTo(" + cstr(MouseXWithBind) + ", " + cstr(MouseYWithBind) + ") = Failed.</p></div>"
	end if
	rewrite
end sub

sub LeftClickButton
	MyResult = MyQuarkAI.LeftClick()
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>LeftClick() = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>LeftClick() = Failed.</p></div>"
	end if
	rewrite
end sub

sub LeftDownButton
	MyResult = MyQuarkAI.LeftDown()
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>LeftDown() = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>LeftDown() = Failed.</p></div>"
	end if
	rewrite
end sub

sub LeftUpButton
	MyResult = MyQuarkAI.LeftUp()
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>LeftUp() = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>LeftUp() = Failed.</p></div>"
	end if
	rewrite
end sub

sub RightClickButton
	MyResult = MyQuarkAI.RightClick()
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>RightClick() = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>RightClick() = Failed.</p></div>"
	end if
	rewrite
end sub

sub RightDownButton
	MyResult = MyQuarkAI.RightDown()
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>RightDown() = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>RightDown() = Failed.</p></div>"
	end if
	rewrite
end sub

sub RightUpButton
	MyResult = MyQuarkAI.RightUp()
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>RightUp() = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>RightUp() = Failed.</p></div>"
	end if
	rewrite
end sub

sub MousePointedPixel
	if ie.document.all.MousePointedPixel.checked then
		MousePointPixelCheckBox = true
	else
		MousePointPixelCheckBox = false
	end if
end sub

sub MousePointedWindow
	if ie.document.all.MousePointedWindow.checked then
		MousePointWindowCheckBox = true
	else
		MousePointWindowCheckBox = false
	end if
end sub

sub Version
	MyOutput = "<div><p>Result:</p><p>Ver() = " + MyQuarkAI.Ver() + ".</p></div>"
	rewrite
end sub

sub BasePath
	MyOutput = "<div><p>Result:</p><p>GetBasePath() = " + MyQuarkAI.GetBasePath() + ".</p></div>"
	rewrite
end sub

sub CurrentPath
	MyOutput = "<div><p>Result:</p><p>GetCurrentPath() = " + MyQuarkAI.GetCurrentPath() + ".</p></div>"
	rewrite
end sub

sub MachineCode
	MyOutput = "<div><p>Result:</p><p>GetMachineCode() = " + MyQuarkAI.GetMachineCode() + ".</p></div>"
	rewrite
end sub

sub Magnifier
	MyResult = MyQuarkAI.Execute("magnify.exe")	
end sub

sub SnippingTool
	MyResult = MyQuarkAI.Execute("SnippingTool.exe")	
end sub

sub MSPaint
	MyResult = MyQuarkAI.Execute("MSPaint.exe")	
end sub


sub KeyCodeCheck
	MyResult = ws.run("https://keycode.info")
end sub

sub Login
	CDKey = id.CDKey.value
	QuarkAIID = MyQuarkAI.Login(CDKey)
	if QuarkAIID > 0 then
		MyOutput = "<div><p>Result:</p><p>Login() = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>Login() = Failed.</p></div>"
	end if
	rewrite
end sub

sub BindButton
	if BindFlag then
		MyResult = MyQuarkAI.UnBindWindow()		
		if MyResult = 1 then
			BindFlag = False
			MyOutput = "<div><p>Result:</p><p>UnBindWindow() = Success.</p></div>"
		else
			MyOutput = "<div><p>Result:</p><p>UnBindWindow() = Failed.</p></div>"
		end if
		rewrite
	else
		WindowHandleWithBind = clng(id.WindowHandleWithBind.value)
		GraphicsMode = id.GraphicsMode.value
		MouseMode = id.MouseMode.value
		KeyboardMode = id.KeyboardMode.value
		BindMode = clng(id.BindMode.value)
		MyResult = MyQuarkAI.BindWindow(WindowHandleWithBind, GraphicsMode, MouseMode, KeyboardMode, BindMode)
		if MyResult = 1 then
			BindFlag = True
			MyOutput = "<div><p>Result:</p><p>BindWindow(" + cstr(WindowHandleWithBind) + ", " + GraphicsMode + ", " + MouseMode + ", " + KeyboardMode + ", " + cstr(BindMode) + ") = Success.</p></div>"
		else
			MyOutput = "<div><p>Result:</p><p>BindWindow(" + cstr(WindowHandleWithBind) + ", " + GraphicsMode + ", " + MouseMode + ", " + KeyboardMode + ", " + cstr(BindMode) + ") = Failed.</p></div>"
		end if
		rewrite
	end if
end sub

sub MouseWithoutBind
	MouseHandleWithoutBind = clng(id.MouseHandleWithoutBind.value)
	MouseXWithoutBind = cint(id.MouseXWithoutBind.value)
	MouseYWithoutBind = cint(id.MouseYWithoutBind.value)
	MouseMessageWithoutBind = id.MouseMessageWithoutBind.value
	if MouseMessageWithoutBind = "MoveTo" then
		MyResult = MyQuarkAI.SendMoveTo(MouseHandleWithoutBind, MouseXWithoutBind, MouseYWithoutBind)
	elseif MouseMessageWithoutBind = "LeftClick" then
		MyResult = MyQuarkAI.SendLeftClick(MouseHandleWithoutBind, MouseXWithoutBind, MouseYWithoutBind)
	elseif MouseMessageWithoutBind = "LeftDown" then
		MyResult = MyQuarkAI.SendLeftDown(MouseHandleWithoutBind, MouseXWithoutBind, MouseYWithoutBind)
	elseif MouseMessageWithoutBind = "LeftUp" then
		MyResult = MyQuarkAI.SendLeftUp(MouseHandleWithoutBind, MouseXWithoutBind, MouseYWithoutBind)
	elseif MouseMessageWithoutBind = "RightClick" then
		MyResult = MyQuarkAI.SendRightClick(MouseHandleWithoutBind, MouseXWithoutBind, MouseYWithoutBind)
	elseif MouseMessageWithoutBind = "RightDown" then
		MyResult = MyQuarkAI.SendRightDown(MouseHandleWithoutBind, MouseXWithoutBind, MouseYWithoutBind)
	elseif MouseMessageWithoutBind = "RightUp" then
		MyResult = MyQuarkAI.SendRightUp(MouseHandleWithoutBind, MouseXWithoutBind, MouseYWithoutBind)
	end if
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>MouseWithoutBind() = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>MouseWithoutBind() = Failed.</p></div>"
	end if
	rewrite
end sub

sub KeyboardWithoutBind
	KeyboardHandleWithoutBind = clng(id.KeyboardHandleWithoutBind.value)
	KeycodeWithoutBind = cint(id.KeycodeWithoutBind.value)
	KeyboardMessageWithoutBind = id.KeyboardMessageWithoutBind.value
	if KeyboardMessageWithoutBind = "KeyPress" then
		MyResult = MyQuarkAI.SendKeyPress(KeyboardHandleWithoutBind, KeyCodeWithoutBind)
	elseif KeyboardMessageWithoutBind = "KeyDown" then
		MyResult = MyQuarkAI.SendKeyDown(KeyboardHandleWithoutBind, KeyCodeWithoutBind)
	elseif KeyboardMessageWithoutBind = "LeftUp" then
		MyResult = MyQuarkAI.SendKeyUp(KeyboardHandleWithoutBind, KeyCodeWithoutBind)
	end if
	if MyResult = 1 then
		MyOutput = "<div><p>Result:</p><p>KeyboardWithoutBind() = Success.</p></div>"
	else
		MyOutput = "<div><p>Result:</p><p>KeyboardWithoutBind() = Failed.</p></div>"
	end if
	rewrite	
end sub

sub ie_OnQuit
	RunningFlag = false
	set MyQuarkAI = nothing
	ie.quit	
	set ie = nothing
	set ws = nothing
	wscript.quit
	set wscript = nothing	
end sub



