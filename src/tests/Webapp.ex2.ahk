#NoEnv

class Webapp {
	
	;static Name, Width, Height, Protocol, HtmlAddress, hWebCtrl, iWebCtrl, wWebCtrl
	
	/*
	Name:="Untitled Webapp"
	Width:=640
	Height:=480
	Protocol:="webapp"
	HtmlAddress:="index.html"
	
	hWebCtrl:=0
	iWebCtrl:=0
	wWebCtrl:=0
	*/
	
	__New(Name="Untitled Webapp",Width=640,Height=480,Protocol="webapp",HtmlAddress="index.html",HtmlData="") {
		Gui _Webapp_:New
		Gui _Webapp_:Default
		Gui _Webapp_:Margin, 0, 0
		Gui _Webapp_:+LastFound +Resize
		
		;OnExit,_Webapp_Exit
		OnExit(this.Quit)
		
		static iWebCtrl
		
		;OnMessage(0x100, "gui_KeyDown", 2)
		Gui _Webapp_:Add, ActiveX, viWebCtrl w%Width% h%Height% hwndhWebCtrl, Shell.Explorer
		;SetWBClientSite()
		iWebCtrl.silent := true ;Surpress JS Error boxes
		
		this.Name 			:= Name
		this.Width 			:= Width
		this.Height 		:= Height
		this.Protocol 		:= Protocol
		this.hWebCtrl 		:= hWebCtrl
		this.iWebCtrl 		:= iWebCtrl
		this.HtmlAddress 	:= HtmlAddress
		
		/*
		if !Strlen(HtmlData) {
			if FileExist(HtmlAddress)
				f := HtmlAddress
			else
				f := FileExist(A_ScriptDir . "\index.html")
			iWebCtrl.Navigate("file://" . f)
		} else {
			;this.SetHtml(HtmlData)
			iWebCtrl.Document.open()
			iWebCtrl.Document.write(HtmlData)
			iWebCtrl.Document.close()
		}
		*/
		iWebCtrl.Document.write("test")
		
		;ComObjConnect(iWebCtrl, this.Webapp_Events)
		
		this.wWebCtrl := iWebCtrl.Document.parentWindow
		this.wWebCtrl.AHK := Func("JS_AHK")
		
		while iWebCtrl.readystate != 4 or iWebCtrl.busy
			sleep 10
		
		Gui _Webapp_:Show, w%Width% h%Height%, %Name%
		;return
		
		;/////////// [ LABELS ] ///////////
		
		
		
		/*
		_Webapp_GuiSize:
			GuiControl, _Webapp_:Move, iWebCtrl, W%A_GuiWidth% H%A_GuiHeight%
		return
		_Webapp_GuiEscape:
			MsgBox 0x34, Webapp.ahk, Are you sure you want to quit?
			IfMsgBox No
				return
		_Webapp_Exit:
		_Webapp_GuiClose:
			ExitApp
		return
		*/
	}
	
	ConnectItem(oRef,cEvents) {
		if IsObject(oRef)
			Item := oRef
		else
			Item := this.iWebCtrl.document.getElementById(oRef)
		return ComObjConnect(Item, cEvents)
		
		;Use DOM access just like javascript!
		;MyButton1 := iWebCtrl.document.getElementById("MyButton1")
		;ComObjConnect(MyButton1, "MyButton1_") ;connect button events
	}
	
	SetHtml(data) {
		;Display(this.iWebCtrl,HTML_page)
		this.iWebCtrl.Document.open()
		this.iWebCtrl.Document.write(data)
		this.iWebCtrl.Document.close()
	}
	
	getWindow() {
		return this.iWebCtrl.document.parentWindow
	}

	ErrorExit(errMsg) {
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
		;MsgBox Goodbye!
		Gui _Webapp_:Destroy
		ExitApp
	}
	
	class Webapp_Events
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
			aP := this.Protocol
			if (InStr(NewURL,aP "://")==1) { ;if url starts with "myapp://"
				what := SubStr(NewURL,Strlen(aP)+4) ;get stuff after "myapp://"
				MsgBox What = "%what%"
			}
			;else do nothing
		}
	}
	
	/*  javascript:AHK('Func') --> Func()
	 */
	JS_AHK(func, prms*) {
		wb := this.iWebCtrl
		; Stop navigation prior to calling the function, in case it uses Exit.
		wb.Stop(),  %func%(prms*)
	}
}

/*  Fix keyboard shortcuts in WebBrowser control.
 *  References:
 *    http://www.autohotkey.com/community/viewtopic.php?p=186254#p186254
 *    http://msdn.microsoft.com/en-us/library/ms693360
 */
gui_KeyDown(wParam, lParam, nMsg, hWnd) {
	wb := this.iWebCtrl
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
    global iWebCtrl
    if pOleObject := ComObjQuery(iWebCtrl, "{00000112-0000-0000-C000-000000000046}")
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