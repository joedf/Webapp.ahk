#NoEnv
#SingleInstance off
#NoTrayIcon
SendMode Input
SetWorkingDir %A_ScriptDir%

;Some parts from Lexikos' Installer.ahk

OnExit,GuiClose

MYAPP_NAME:="Webapp.ahk"
MYAPP_PROTOCOL:="myapp"

HTML_page =
( Ltrim Join
<!DOCTYPE html>
<html>
	<head>
		<style>
			body,html{overflow:hidden;}
			body{font-family:sans-serif;background-color:#dde4ec;}
			#title{font-size:36px;}
			#corner{font-size:10px;position:absolute;top:8px;right:8px;}
			p{font-size:16px;background-color:#efefef;border:solid 1px #666;padding:4px;}
			#footer{text-align:center;}
		</style>
	</head>
	<body>
		<div id="title">Lorem Ipsum</div>
		<div id="corner">Welcome!</div>
		<p>The standard Lorem Ipsum passage, used since the 1500s</p>
		<p>Lorem ipsum dolor sit amet, consectetur adipisicing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.</p>
		<textarea rows="4" cols="70" id="MyTextBox">1234567890-=\ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz!@#$^&*()_+|~</textarea>
		<p id="footer">
			<a href="%MYAPP_PROTOCOL%://msgbox/hello">Click me for a MsgBox</a>&nbsp;-&nbsp;
			<a href="NOPE://msgbox/hello">Click me for nothing</a>&nbsp;-&nbsp;
			<a href="%MYAPP_PROTOCOL%://soundplay/ding">Click me for a ding sound!</a><br>
			<a href="#" onclick="AHK('RunButton4_function')">Link using onlick instead</a>
			<span style="display:block;height:14px;">&nbsp;</span>
			<input type="button" id="MyButton1" value="Show Content in AHK MsgBox">
			<input type="button" id="MyButton2" value="Change Content with AHK">
			<input type="button" id="MyButton3" value="Greetings from AHK">
			<input type="button" id="MyButton4" value="AHK MsgBox using function" onclick="AHK('RunButton4_function')">
		</p>
	</body>
</html>
)

Gui Margin, 0, 0
Gui +LastFound

OnMessage(0x100, "gui_KeyDown", 2)
Gui Add, ActiveX, vwb w640 h480 hwndhwb, Shell.Explorer
wb.silent := true ;Surpress JS Error boxes
Display(wb,HTML_page)
ComObjConnect(wb, wb_events)

DisableClickSounds()

w := wb.Document.parentWindow
w.AHK := Func("JS_AHK")

;Wait for IE to load the page, before we connect the event handlers
while wb.readystate != 4 or wb.busy
	sleep 10

;Use DOM access just like javascript!
MyButton1 := wb.document.getElementById("MyButton1")
MyButton2 := wb.document.getElementById("MyButton2")
MyButton3 := wb.document.getElementById("MyButton3")
ComObjConnect(MyButton1, "MyButton1_") ;connect button events
ComObjConnect(MyButton2, "MyButton2_")
ComObjConnect(MyButton3, "MyButton3_")

Gui Show, w640 h480, %MYAPP_NAME%
return


GuiEscape:
	MsgBox 0x34, Webapp.ahk, Are you sure you want to quit?
	IfMsgBox No
		return
GuiClose:
	Gui Destroy
	FileDelete,%A_Temp%\*.DELETEME.html ;clean tmp file
	ExitApp
return

class wb_events
{
	;for more events and other, see http://msdn.microsoft.com/en-us/library/aa752085
	
	NavigateComplete2(wb) { ;blocked all navigation, we want our own stuff happening
		wb.Stop()
	}
	DownloadComplete(wb, NewURL) {
		wb.Stop()
	}
	DocumentComplete(wb, NewURL) {
		wb.Stop()
	}
	
	BeforeNavigate2(wb, NewURL)
	{
		wb.Stop() ;blocked all navigation, we want our own stuff happening
		;parse the url
		global MYAPP_PROTOCOL
		if (InStr(NewURL,MYAPP_PROTOCOL "://")==1) { ;if url starts with "myapp://"
			what := SubStr(NewURL,Strlen(MYAPP_PROTOCOL)+4) ;get stuff after "myapp://"
			if InStr(what,"msgbox/hello")
				MsgBox Hello world!
			else if InStr(what,"soundplay/ding")
				SoundPlay, %A_WinDir%\Media\ding.wav
		}
		;else do nothing
	}
}

; Our Event Handlers
MyButton1_OnClick() {
	global wb
	MsgBox % wb.Document.getElementById("MyTextBox").Value
}
MyButton2_OnClick() {
	global wb
	FormatTime, TimeString, %A_Now%, dddd MMMM d, yyyy HH:mm:ss
	data := "AHK Version " A_AhkVersion " - " (A_IsUnicode ? "Unicode" : "Ansi") " " (A_PtrSize == 4 ? "32" : "64") "bit`nCurrent time: " TimeString "`nRandom number: " RandNum()
	wb.Document.getElementById("MyTextBox").value := data
}
MyButton3_OnClick() {
	MsgBox Hello world!
}

RunButton4_function() {
	MsgBox Hello from JS_AHK :)
}

Display(wb,html_str) {
	Count:=0
	while % FileExist(f:=A_Temp "\" A_TickCount A_NowUTC "-tmp" Count ".DELETEME.html")
		Count+=1
	FileAppend,%html_str%,%f%
	wb.Navigate("file://" . f)
}

RandNum(min=0,max=65535) {
	Random, x, %min%, %max%
	return x
}

DisableClickSounds() {
	FEATURE_DISABLE_NAVIGATION_SOUNDS := 21
	SET_FEATURE_ON_PROCESS := 0x00000002
    DllCall("urlmon.dll\CoInternetSetFeatureEnabled","uint",FEATURE_DISABLE_NAVIGATION_SOUNDS,"uint",SET_FEATURE_ON_PROCESS,1)
}

_InitUI() {
    local w
    SetWBClientSite()
    ;gosub DefineUI
    wb.Silent := true
    wb.Navigate("about:blank")
    while wb.ReadyState != 4 {
        Sleep 10
        if (A_TickCount-initTime > 2000)
            throw 1
    }
    wb.Document.open()
    wb.Document.write(html)
    wb.Document.close()
    w := wb.Document.parentWindow
    if !w || !w.initOptions
        throw 1
    w.AHK := Func("JS_AHK")
    if (!CurrentType && A_ScriptDir != DefaultPath)
        CurrentName := ""  ; Avoid showing the Reinstall option since we don't know which version it was.
    w.initOptions(CurrentName, CurrentVersion, CurrentType
                , ProductVersion, DefaultPath, DefaultStartMenu
                , DefaultType, A_Is64bitOS = 1)
    if (A_ScriptDir = DefaultPath) {
        w.installdir.disabled := true
        w.installdir_browse.disabled := true
        w.installcompiler.disabled := !DefaultCompiler
        w.installcompilernote.style.display := "block"
        w.ci_nav_install.innerText := "apply"
        w.install_button.innerText := "Apply"
        w.extract.style.display := "None"
        w.opt1.disabled := true
        w.opt1.firstChild.innerText := "Checking for updates..."
    }
    w.installcompiler.checked := DefaultCompiler
    w.enabledragdrop.checked := DefaultDragDrop
    w.separatebuttons.checked := DefaultIsHostApp
    ; w.defaulttoutf8.checked := DefaultToUTF8
    if !A_Is64bitOS
        w.it_x64.style.display := "None"
    if A_OSVersion in WIN_2000,WIN_2003,WIN_XP ; i.e. not WIN_7, WIN_8 or a future OS.
        w.separatebuttons.parentNode.style.display := "none"
    ;w.switchPage("start")
    w.document.body.focus()
    ; Scale UI by screen DPI.  My testing showed that Vista with IE7 or IE9
    ; did not scale by default, but Win8.1 with IE10 did.  The scaling being
    ; done by the control itself = deviceDPI / logicalDPI.
    logicalDPI := w.screen.logicalXDPI, deviceDPI := w.screen.deviceXDPI
    if (A_ScreenDPI != 96)
        w.document.body.style.zoom := A_ScreenDPI/96 * (logicalDPI/deviceDPI)
    ;if (A_ScriptDir = DefaultPath)
    ;    CheckForUpdates()
}

/*  Complex workaround to override "Active scripting" setting
 *  and ensure scripts can run within the WebBrowser control.
 */

global WBClientSite

SetWBClientSite()
{
    interfaces := {
    (Join,
        IOleClientSite: [0,3,1,0,1,0]
        IServiceProvider: [3]
        IInternetSecurityManager: [1,1,3,4,8,7,3,3]
    )}
    unkQI      := RegisterCallback("WBClientSite_QI", "Fast")
    unkAddRef  := RegisterCallback("WBClientSite_AddRef", "Fast")
    unkRelease := RegisterCallback("WBClientSite_Release", "Fast")
    WBClientSite := {_buffers: bufs := {}}, bufn := 0, 
    for name, prms in interfaces
    {
        bufn += 1
        bufs.SetCapacity(bufn, (4 + prms.MaxIndex()) * A_PtrSize)
        buf := bufs.GetAddress(bufn)
        NumPut(unkQI,       buf + 1*A_PtrSize)
        NumPut(unkAddRef,   buf + 2*A_PtrSize)
        NumPut(unkRelease,  buf + 3*A_PtrSize)
        for i, prmc in prms
            NumPut(RegisterCallback("WBClientSite_" name, "Fast", prmc+1, i), buf + (3+i)*A_PtrSize)
        NumPut(buf + A_PtrSize, buf + 0)
        WBClientSite[name] := buf
    }
    global wb
    if pOleObject := ComObjQuery(wb, "{00000112-0000-0000-C000-000000000046}")
    {   ; IOleObject::SetClientSite
        DllCall(NumGet(NumGet(pOleObject+0)+3*A_PtrSize), "ptr"
            , pOleObject, "ptr", WBClientSite.IOleClientSite, "uint")
        ObjRelease(pOleObject)
    }
}

WBClientSite_QI(p, piid, ppvObject)
{
    static IID_IUnknown := "{00000000-0000-0000-C000-000000000046}"
    static IID_IOleClientSite := "{00000118-0000-0000-C000-000000000046}"
    static IID_IServiceProvider := "{6d5140c1-7436-11ce-8034-00aa006009fa}"
    iid := _String4GUID(piid)
    if (iid = IID_IOleClientSite || iid = IID_IUnknown)
    {
        NumPut(WBClientSite.IOleClientSite, ppvObject+0)
        return 0 ; S_OK
    }
    if (iid = IID_IServiceProvider)
    {
        NumPut(WBClientSite.IServiceProvider, ppvObject+0)
        return 0 ; S_OK
    }
    NumPut(0, ppvObject+0)
    return 0x80004002 ; E_NOINTERFACE
}

WBClientSite_AddRef(p)
{
    return 1
}

WBClientSite_Release(p)
{
    return 1
}

WBClientSite_IOleClientSite(p, p1="", p2="", p3="")
{
    if (A_EventInfo = 3) ; GetContainer
    {
        NumPut(0, p1+0) ; *ppContainer := NULL
        return 0x80004002 ; E_NOINTERFACE
    }
    return 0x80004001 ; E_NOTIMPL
}

WBClientSite_IServiceProvider(p, pguidService, piid, ppvObject)
{
    static IID_IUnknown := "{00000000-0000-0000-C000-000000000046}"
    static IID_IInternetSecurityManager := "{79eac9ee-baf9-11ce-8c82-00aa004ba90b}"
    if (_String4GUID(pguidService) = IID_IInternetSecurityManager)
    {
        iid := _String4GUID(piid)
        if (iid = IID_IInternetSecurityManager || iid = IID_IUnknown)
        {
            NumPut(WBClientSite.IInternetSecurityManager, ppvObject+0)
            return 0 ; S_OK
        }
        NumPut(0, ppvObject+0)
        return 0x80004002 ; E_NOINTERFACE
    }
    NumPut(0, ppvObject+0)
    return 0x80004001 ; E_NOTIMPL
}

WBClientSite_IInternetSecurityManager(p, p1="", p2="", p3="", p4="", p5="", p6="", p7="", p8="")
{
    if (A_EventInfo = 5) ; ProcessUrlAction
    {
        if (p2 = 0x1400) ; dwAction = URLACTION_SCRIPT_RUN
        {
            NumPut(0, p3+0)  ; *pPolicy := URLPOLICY_ALLOW
            return 0 ; S_OK
        }
    }
    return 0x800C0011 ; INET_E_DEFAULT_ACTION
}

_String4GUID(pGUID)
{
	VarSetCapacity(String,38*2)
	DllCall("ole32\StringFromGUID2", "ptr", pGUID, "str", String, "int", 39)
	Return	String
}


/*  Fix keyboard shortcuts in WebBrowser control.
 *  References:
 *    http://www.autohotkey.com/community/viewtopic.php?p=186254#p186254
 *    http://msdn.microsoft.com/en-us/library/ms693360
 */

gui_KeyDown(wParam, lParam, nMsg, hWnd) {
    global wb
    if (Chr(wParam) ~= "[A-Z]" || wParam = 0x74) ; Disable Ctrl+O/L/F/N and F5.
        return
    pipa := ComObjQuery(wb, "{00000117-0000-0000-C000-000000000046}")
    VarSetCapacity(kMsg, 48), NumPut(A_GuiY, NumPut(A_GuiX
    , NumPut(A_EventInfo, NumPut(lParam, NumPut(wParam
    , NumPut(nMsg, NumPut(hWnd, kMsg)))), "uint"), "int"), "int")
    Loop 2
    r := DllCall(NumGet(NumGet(1*pipa)+5*A_PtrSize), "ptr", pipa, "ptr", &kMsg)
    ; Loop to work around an odd tabbing issue (it's as if there
    ; is a non-existent element at the end of the tab order).
    until wParam != 9 || wb.Document.activeElement != ""
    ObjRelease(pipa)
    if r = 0 ; S_OK: the message was translated to an accelerator.
        return 0
}


/*  javascript:AHK('Func') --> Func()
 */

JS_AHK(func, prms*) {
    global wb
    ; Stop navigation prior to calling the function, in case it uses Exit.
    wb.Stop(),  %func%(prms*)
}


/*  Utility Functions
 */

getWindow() {
    global wb
    return wb.document.parentWindow
}

ErrorExit(errMsg) {
    global
    if !SilentMode
        MsgBox 16, Webapp.ahk, %errMsg%
    ExitApp 1
}

GetErrorMessage(error_code="") {
    VarSetCapacity(buf, 1024) ; Probably won't exceed 1024 chars.
    if DllCall("FormatMessage", "uint", 0x1200, "ptr", 0, "int", error_code!=""
                ? error_code : A_LastError, "uint", 1024, "str", buf, "uint", 1024, "ptr", 0)
        return buf
}

Quit() {
    ExitApp
}