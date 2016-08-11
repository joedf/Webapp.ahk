class CWebView extends CWebViewer
{
	__New(gui, options:="", src:="", kwargs:="")
	{
		static WB

		Gui %gui%:+LastFoundExist
		if !WinExist()
			throw "GUI '" . gui . "' does not exist."
		Gui %gui%:Add, ActiveX, %options% vWB HwndhAtlAxWin, Shell.Explorer

		/* Store a weak reference to this instance in the IWebBrowser2 object
		 * via PutProperty(), assign it to the property name: 'CWV_REF'. User
		 * can then retrieve a reference to the instance via Object() and
		 * GetProperty() -> e.g.: instance := Object( WB.GetProperty("CWV_REF") )
		 */
		WB.PutProperty("CWV_REF", &this)
		
		this.__Ptr  := WB, WB := ""
		this.__Host := hAtlAxWin + 0 ;// AtlAxWin

		;// Make sure that scripts are allowed to run
		; this.__Site := new CWebClientSite(this)
		; this._SetClientSite(this.__Site.__Ptr)

		;// Initialize with blank document
		this.SetURL()
		
		;// Load/Write HTML(src) if specified
		if (src != "")
			this[src ~= "s)^\s*?<.*>\s*$" ? "SetHTML" : "SetURL"](src)
	}

	__Delete()
	{
		/* Without this line, AHK crashes on instance object's release (inst := "")
		 * if SetClientSite is implemented (currently commented out, see __New)
		 */
		this._SetClientSite(0)
		
		this.__Ptr.PutProperty("CWV_REF", 0)
		this.__Ptr := ""
	}

	SetURL(url:="about:blank")
	{
		wb := this.__Ptr
		wb.Navigate(url)
		while (wb.ReadyState != 4) && (wb.Document.readyState != "complete")
			Sleep 10
	}

	SetHTML(html)
	{
		doc := this.__Ptr.Document
		doc.open(), doc.write(html), doc.close()
	}

	Window {
		get {
			return new CWebPage("HTMLWindow", this.__Ptr.Document.parentWindow)
		}
		set {
			return value
		}
	}

	Document[selector:=""] {
		get {
			return new CWebPage("HTMLDocument", this.__Ptr.Document)
		}
		set {
			return value
		}
	}

	TranslateAccelerator[keys:="DU", args*] { ;// keys := D|U|DU -> down|up|down&up
		;// WM_KEYDOWN = 0x100 | WM_KEYUP = 0x101
		get {
			static msg := { "D": 0x100, "U": 0x101 }
			if !NumGet(&args + 4*A_PtrSize) ;// get()
				return OnMessage(msg[SubStr(keys, 1, 1)])
			;// set()
			fn := args[1] ? (IsFunc(args[1]) ? args[1] : "_CWV_OnKeyPress") : ""
			for i, key in StrSplit(keys)
				if ObjHasKey(msg, key)
					OnMessage(msg[key], fn)
			return this.TranslateAccelerator[keys]
		}
		set {
			return this.TranslateAccelerator[keys, value] ;// trigger get()
		}
	}

	__Handle {
		get {
			static cw := ["Shell Embedding", "Shell DocObject View", "Internet Explorer_Server"]
			h := this.__Host, VarSetCapacity(WinClass, 256)
			while (WinClass != "Internet Explorer_Server")
			{
				if !( h := DllCall("FindWindowEx", "Ptr", h, "Ptr", 0, "Str", cw[A_Index], "Ptr", 0) )
					break
				DllCall("GetClassName", "Ptr", h, "Ptr", &WinClass, "Int", 256)
			}
			return h
		}
		set {
			return value
		}
	}

	__Gui {
		get {
			return DllCall("GetParent", "Ptr", this.__Host)
		}
		set {
			return value
		}
	}

	_AutoScale()
	{
		window := this.Window.__Ptr, body := this.Document.body.__Ptr
		logicalXDPI := window.screen.logicalXDPI, deviceXDPI  := window.screen.deviceXDPI
		if (A_ScreenDPI > 96) && (logicalXDPI == deviceXDPI)
			body.style.zoom := A_ScreenDPI/96 * (logicalXDPI/deviceXDPI)
	}

	_SetClientSite(site)
	{
		wb := this.__Ptr
		if pOleObject := ComObjQuery(wb, "{00000112-0000-0000-C000-000000000046}")
		{
			;// IOleObject::SetClientSite
			SetClientSite := NumGet(NumGet(pOleObject + 0) + 3*A_PtrSize)
			DllCall(SetClientSite, "Ptr", pOleObject, "Ptr", site, "UInt")
			ObjRelease(pOleObject)
		}
	}
} ; class CWebView

class CWebViewer extends CWebCommon
{
	__New()
	{
		return false
	}

	Version[short:=1] {
		get {
			static L := ComObjCreate("WScript.Shell").RegRead("HKLM\SOFTWARE\Microsoft\Internet Explorer\svcVersion")
			static S := Round(L)
			return short ? S : L ;// S=short version, L=long version
		}
		set {
			throw "This property is read-only`nSpecifically: Version"
		}
	} ; property Version

	BrowserEmulation[args*] { ;// FEATURE_BROWSER_EMULATION -> http://goo.gl/V01Frx
		get {
			static skey := "HKCU\Software\" . ( A_Is64bitOS ? "Wow6432Node\" : "" )
			. "Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION\"
			. [ VarSetCapacity(lpFilename, nSize := 260)
			  , DllCall("GetModuleFileName", "Ptr", 0, "Str", lpFilename, "UInt", nSize)
			  , StrGet(DllCall("Shlwapi\PathFindFileName", "Ptr", &lpFilename)) ][3]

			WshShell := ComObjCreate("WScript.Shell")

			if !NumGet(&args + 4*A_PtrSize) ;// no argument(s), get()
			{
				try val := WshShell.RegRead(skey)
				catch error
					val := 0
				finally
					return val
			}

			;// set()
			if ( (ie_version := this.Version) < 8 )
				throw "FEATURE_BROWSER_EMULATION is only available for IE8 and later."
			value := args[1]

			;// http://goo.gl/V01Frx
			static fbe := { ;// values indicated by negative keys ignore !DOCTYPE directives
			(LTrim Join Q C
				7:  7000,              ; 0x1B58           - IE7 Standards mode.
				8:  8000,  -8:  8888,  ; 0x1F40  | 0x22B8 - IE8 mode, IE8 Standards mode
				9:  9000,  -9:  9999,  ; 0x2328  | 0x270F - IE9 mode, IE9 Standards mode
				10: 10000, -10: 10001, ; 0x02710 | 0x2711 - IE10 Standards mode
				11: 11000, -11: 11001  ; 0x2AF8  | 0x2AF9 - IE11 edge mode
			)}

			if (Abs(value) != "")
				value := Round(value)
			if value
				value := fbe[ (value = "edge") || !ObjHasKey(fbe, value) ? ie_version : value ]

			try value? WshShell.RegWrite(skey, value, "REG_DWORD") : WshShell.RegDelete(skey)
			catch error
				throw error
			finally
				return this.BrowserEmulation
		}
		set {
			return this.BrowserEmulation[value] ;// trigger get()
		}
	} ; property BrowserEmulation
} ; class CWebViewer

class CWebPage ;// namespace + constructor
{
	__New(itype, ptr, args*)
	{
		itype := this.base[InStr(itype, "HTML") ? itype : "HTML" . itype]
		return IsObject(itype) ? new itype(ptr, args*) : 0
	}

	class HTMLWindow extends CWebCommon
	{
		__New(window)
		{
			this.__Ptr := window
		}
	} ; class CWebPage.HTMLWindow

	class HTMLDocument extends CWebPage.HTMLNode
	{
		__New(document)
		{
			this.__Ptr := document
		}

		Load(src)
		{
			window := (doc := this.__Ptr).parentWindow
			window.location.assign(src)
			while (doc.readyState != "complete")
				Sleep 10
		}

		Write(html)
		{
			doc := this.__Ptr
			doc.open(), doc.write(html), doc.close()
		}
	} ; class CWebPage.HTMLDocument

	class HTMLElement extends CWebPage.HTMLNode
	{
		__New(element)
		{
			this.__Ptr := element
		}
	} ; class CWebPage.HTMLElement

	class HTMLNode extends CWebCommon
	{
		Type[t:=""] {
			get {
				static types := { 1: "Element", 9: "Document" }
				type := this.__Ptr.nodeType
				return (t = "String" || t = "Str") ? types[type] : type
			}
			set {
				return value
			}
		}

		Query(selector:="*", all:=false)
		{
			node  := this.__Ptr
			doc   := this.Type == 9 ? node : node.ownerDocument
			query := all? "querySelectorAll" : "querySelector"
			if (doc.documentMode < 8)
				return false
				; throw query . "() is only available on IE8 and later." ;// crashes

			if this._IsMemberOf(node, query)
				r := node[query](selector)
			else
			{
				expr := (tmp_id := node.nodeType != 9)
				? "(function() { if (document." . query . ")"
				  . "{ return document.getElementById('"
				  . ( (tmp_id := node.id == "") ? node.id := "QS_" . &this : node.id )
				  . "')." . query . "('" . selector . "'); }}())"
				: "(function() { if (document." . query . ")"
				  . "{ return document." . query . "('" . selector . "'); }}())"

				r := doc.parentWindow.eval(expr)

				if tmp_id
					node.id := ""
			}
			return this._Wrap(r)
		}

		QueryAll(selector:="*")
		{
			return this.Query(selector, true)
		}
	} ; class CWebPage.HTMLNode

	class HTMLCollection extends CWebCommon
	{
		__New(collection)
		{
			this.__Ptr := collection
		}

		__Get(key:="", args*)
		{
			if (key >= 0 && key < this.__Ptr.length)
				return this._Wrap( this.__Ptr.item(key) )
		}

		_NewEnum()
		{
			return new this.base._Enumerator(this)
		}

		class _Enumerator
		{
			__New(collection)
			{
				this.idx := 0
				this.obj := collection
			}

			Next(ByRef key, ByRef val:="")
			{
				obj := this.obj.__Ptr
				if (r := this.idx < obj.length)
					key := this.obj._Wrap(obj.item(this.idx++))
				return r
			}
		}
	} ; class CWebPage.HTMLCollection
} ; class CWebPage

class CWebCommon extends CWebPrototype
{
	__New()
	{
		return false
	}

	_Wrap(obj)
	{
		if (ComObjType(obj) == "")
			return obj
		static types := { 1: "HTMLElement", 9: "HTMLDocument" }
		if (ComObjType(obj, "name") == "IWebBrowser2")
			return (ref := obj.GetProperty("CWV_REF")) ? Object(ref) : obj
		else if (pwin := ComObjQuery(obj, "{332C4427-26CB-11D0-B483-00C04FD90119}")) ;// IID_IHTMLWindow2
			type := "HTMLWindow", ObjRelease(pwin)
		else
			type := this._IsMemberOf(obj, "nodeType") ? types[obj.nodeType]
			     :  this._IsMemberOf(obj, "length")   ? "HTMLCollection"
			     :  ""
		return (type != "") ? new CWebPage(type, obj) : obj
	}

	_IsMemberOf(obj, name)
	{
		pDisp := ComObjValue(obj)
		GetIDsOfNames := NumGet(NumGet(pDisp + 0), 5*A_PtrSize)
		VarSetCapacity(IID_NULL, 16, 0), DISPID := 0 ;// Make #Warn happy
		r := DllCall(GetIDsOfNames, "Ptr", pDisp, "Ptr", &IID_NULL, "Ptr*", &name, "UInt", 1, "UInt", 1024, "Int*", DISPID)
		return (r == 0) && (DISPID + 1)
	}

	_WBFrom(from:="ahk_class AutoHotkeyGUI", nn:=1, raw:=0)
	{
		wb := IsObject(from) ? this._WBFromObject(from) : this._WBFromWindow(from, nn)
		return raw? wb : this._Wrap(wb)
	}

	_WBFromWindow(WinTitle:="ahk_class AutoHotkeyGUI", nn:=1)
	{
		dhw := A_DetectHiddenWindows
		DetectHiddenWindows On
		static IE_ServerClass := "Internet Explorer_Server"
		WinGetClass WinClass, %WinTitle%
		IE_Server := (WinClass != IE_ServerClass) ? IE_ServerClass . nn : ""
		static msg := DllCall("RegisterWindowMessage", "Str", "WM_HTML_GETOBJECT")
		SendMessage %msg%, 0, 0, %IE_Server%, %WinTitle%
		DetectHiddenWindows %dhw%
		static ERROR := A_AhkVersion < "2" ? "FAIL" : "ERROR"
		if (ErrorLevel == ERROR)
			return
		static IID_IHTMLDocument2 := "{332C4425-26CB-11D0-B483-00C04FD90119}"
		lResult := ErrorLevel
		, VarSetCapacity(GUID, 16, 0)
		, DllCall("ole32\CLSIDFromString", "WStr", IID_IHTMLDocument2, "Ptr", &GUID)
		, DllCall("oleacc\ObjectFromLresult", "Ptr", lResult, "Ptr", &GUID, "Ptr", 0, "Ptr*", pdoc)

		static IID_IWebBrowserApp := "{0002DF05-0000-0000-C000-000000000046}"
		     , SID_SWebBrowserApp := IID_IWebBrowserApp
		pweb := ComObjQuery(pdoc, SID_SWebBrowserApp, IID_IWebBrowserApp)

		ObjRelease(pdoc)

		return ComObject(9, pweb, 1) ;// VT_DISPATCH=9, F_OWNVALUE=1
	}

	_WBFromObject(obj)
	{
		static IID_IWebBrowserApp := "{0002DF05-0000-0000-C000-000000000046}"
		     , SID_SWebBrowserApp := IID_IWebBrowserApp
		pweb := ComObjQuery(obj, SID_SWebBrowserApp, IID_IWebBrowserApp)
		return ComObject(9, pweb, 1)
	}
}

class CWebPrototype
{
	__Get(key:="", args*)
	{
		if !key
			return this.__Ptr

		; print(this.__Class, key, "sep=->")

		if (key != "base" && key != "__Class") && !ObjHasKey(CWebPrototype, key)
		{
			if this._IsMemberOf(this.__Ptr, key)
				return this._Wrap( (this.__Ptr)[key, args*] )
		}
	}

	__Set(key, value, args*)
	{
		; print(this.__Class "->" key, value, "sep= = ")

		if (key != "base" && key != "__Class" && key != "__Ptr")
		{
			if this._IsMemberOf(this.__Ptr, key)
			{
				if IsObject(value)
				&& (ComObjType(value) == "")
				&& (ComObjType(value.__Ptr) != "")
					value := value.__Ptr
				return (this.__Ptr)[key] := value
			}
		}
	}

	__Call(method:="", args*)
	{
		; print(this.__Class, method, args, "sep=->")

		if !ObjHasKey(CWebPrototype, method)
		{
			if this._IsMemberOf(this.__Ptr, method)
			{
				for i, arg in args
				{
					if IsObject(arg)
					&& (ComObjType(arg) == "")
					&& (ComObjType(arg.__Ptr) != "")
						args[i] := arg.__Ptr
				}
				return this._Wrap( (this.__Ptr)[method](args*) )
			}
		}
	}
}

class CWebClientSite
{
	__New(self)
	{
		ObjSetCapacity(this, "__Site", 3*A_PtrSize)
		NumPut(&self
		, NumPut(this.base._vftable("_vft_IInternetSecurityManager", "11348733")
		, NumPut(this.base._vftable("_vft_IServiceProvider", "3")
		, NumPut(this.base._vftable("_vft_IOleClientSite", "031010")
		    , this.__Ptr := ObjGetAddress(this, "__Site") ) ) ) )
	}

	_vftable(name, args)
	{
		if ( ptr := ObjGetAddress(this, name) )
			return ptr
		
		static IUnknown := {
		(Join, Q C
			"QueryInterface": RegisterCallback(CWebClientSite._QueryInterface, "Fast")
			"AddRef":         RegisterCallback(CWebClientSite._AddRef, "Fast")
			"Release":        RegisterCallback(CWebClientSite._Release, "Fast")
		)}

		sizeof_VFTABLE := (3 + StrLen(args)) * A_PtrSize
		ObjSetCapacity(this, name, sizeof_VFTABLE)
		ptr := ObjGetAddress(this, name)

		NumPut(IUnknown.Release, NumPut(IUnknown.AddRef, NumPut(IUnknown.QueryInterface, ptr+0)))

		fn := this[SubStr(name, 6)]
		for i, argc in StrSplit(args)
			NumPut(RegisterCallback(fn, "Fast", argc+1, i), ptr + (3+i-1)*A_PtrSize)

		return ptr
	}

	_QueryInterface(riid, ppvObject)
	{
		static IID_IUnknown := "{00000000-0000-0000-C000-000000000046}"
		     , IID_IOleClientSite := "{00000118-0000-0000-C000-000000000046}"
		     , IID_IServiceProvider := "{6d5140c1-7436-11ce-8034-00aa006009fa}"
		
		iid := CWebClientSite._GUID2String(riid)
		if (iid = IID_IOleClientSite || iid = IID_IUnknown)
		{
			NumPut(this, ppvObject+0) ;// IOleClientSite
			return 0 ;// S_OK
		}
		if (iid = IID_IServiceProvider)
		{
			NumPut(this + A_PtrSize, ppvObject+0) ;// IServiceProvider
			return 0 ;// S_OK
		}
		NumPut(0, ppvObject+0)
		return 0x80004002 ;// E_NOINTERFACE
	}

	_AddRef()
	{
		return 1
	}

	_Release()
	{
		return 1
	}

	IOleClientSite(p1:="", p2:="", p3:="")
	{
		if (A_EventInfo == 3) ;// GetContainer
		{
			NumPut(0, p1+0) ;// *ppContainer := NULL
			return 0x80004002 ;// E_NOINTERFACE
		}
		return 0x80004001 ;// E_NOTIMPL
	}

	IServiceProvider(guidService, riid, ppv) ;// QueryService
	{
		static IID_IUnknown := "{00000000-0000-0000-C000-000000000046}"
		     , IID_IInternetSecurityManager := "{79eac9ee-baf9-11ce-8c82-00aa004ba90b}"
		
		if (CWebClientSite._GUID2String(guidService) = IID_IInternetSecurityManager)
		{
			iid := CWebClientSite._GUID2String(riid)
			if (iid = IID_IInternetSecurityManager || iid = IID_IUnknown)
			{
				NumPut(this + A_PtrSize, ppv+0) ;// IInternetSecurityManager
				return 0 ;// S_OK
			}
			NumPut(0, ppv+0)
			return 0x80004002 ;// E_NOINTERFACE
		}
		NumPut(0, ppv+0)
		return 0x80004001 ;// E_NOTIMPL
	}

	IInternetSecurityManager(p1:="", p2:="", p3:="", p4:="", p5:="", p6:="", p7:="", p8:="")
	{
		if (A_EventInfo == 5) ;// ProcessUrlAction
		{
		 	if (p2 == 0x1400) ;// dwAction = URLACTION_SCRIPT_RUN
		 	{
		 		NumPut(0x00, p3+0) ;// *pPolicy := URLPOLICY_ALLOW = 0x00
		 		return 0 ;// S_OK
		 	}
		}
		return 0x800C0011 ;// INET_E_DEFAULT_ACTION
	}

	_GUID2String(pGUID)
	{
		VarSetCapacity(string, 38*2)
		DllCall("ole32\StringFromGUID2", "Ptr", pGUID, "Str", string, "Int", 39)
		return string
	}
}

_CWV_OnKeyPress(wParam, lParam, nMsg, hWnd)
{
	static IID_IOleInPlaceActiveObject := "{00000117-0000-0000-C000-000000000046}"
	static IID_IDispatch := "{00020400-0000-0000-C000-000000000046}"
	static IID_IHTMLWindow2 := "{332C4427-26CB-11D0-B483-00C04FD90119}"
	static IID_IWebBrowserApp := "{0002DF05-0000-0000-C000-000000000046}"
	     , SID_SWebBrowserApp := IID_IWebBrowserApp

	WinGetClass WinClass, ahk_id %hWnd%
	if (WinClass == "Internet Explorer_Server")
	{
		VarSetCapacity(GUID, 16, 0), pacc := 0 ;// Make #Warn happy
		DllCall("ole32\CLSIDFromString", "WStr", IID_IDispatch, "Ptr", &GUID)
		DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", 0xFFFFFFFC, "Ptr", &GUID, "Ptr*", pacc) ;// OBJID_CLIENT:=0xFFFFFFFC
		pwin := ComObjQuery(pacc, IID_IHTMLWindow2, IID_IHTMLWindow2)
		     , ObjRelease(pacc)
		pweb := ComObjQuery(pwin, SID_SWebBrowserApp, IID_IWebBrowserApp)
		     , ObjRelease(pwin)
		wb := ComObject(9, pweb, 1) ; OR simply wb := CWebView._WBFromWindow("ahk_id " hWnd)
		pIOIPAO := ComObjQuery(wb, IID_IOleInPlaceActiveObject)

		/* http://goo.gl/GX6GNm
		typedef struct tagMSG {
		  HWND   hwnd;
		  UINT   message;
		  WPARAM wParam;
		  LPARAM lParam;
		  DWORD  time;
		  POINT  pt;
		} MSG, *PMSG, *LPMSG;
		*/
		VarSetCapacity(MSG, 12 + 4*A_PtrSize, 0)
		, NumPut(A_GuiY
		, NumPut(A_GuiX
		, NumPut(A_EventInfo
		, NumPut(lParam
		, NumPut(wParam
		, NumPut(nMsg
		, NumPut(hWnd, MSG)))), "UInt"), "Int"), "Int")

		TranslateAccelerator := NumGet(NumGet(pIOIPAO + 0) + 5*A_PtrSize)
		Loop 2
			r := DllCall(TranslateAccelerator, "Ptr", pIOIPAO, "Ptr", &MSG)
		until (wParam != 9 || wb.Document.activeElement != "")
		ObjRelease(pIOIPAO)
		if (r == 0)
			return 0
	}
} ; _CWV_OnKeyPress()