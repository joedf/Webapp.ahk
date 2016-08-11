#NoEnv
#Include CWebView.ahk

class Webapp {

	static g := 0

	__New(Name="Untitled Webapp",Width=640,Height=480,html="null",Protocol="webapp") {
		g+=1
		
		;Gui %g%:New
		Gui %g%:Default
		Gui %g%:Margin, 0, 0
		Gui %g%:+LastFound +Resize
		
		this.Name 		:= Name
		this.Width 		:= Width
		this.Height 	:= Height
		this.Protocol 	:= Protocol
		this.hWebCtrl 	:= hWebCtrl
		this.iWebCtrl 	:= iWebCtrl
		this.Html 		:= Html
		
		OnExit(this.Quit)
		
		GuiOptions:="x0 y0 w" . Width . " h" . Height
		
		c := new CWebView(g, GuiOptions, html)
		c.TranslateAccelerator := ["D", gui_KeyDown]
		;w := c.__Ptr.document.parentWindow
		;w.AHK := Func("JS_AHK")
		
		Gui %g%:Show, w%Width% h%Height%, %Name%
	}
	
	Quit() {
		MsgBox Goodbye!
		ExitApp
	}
}

	/*  javascript:AHK('Func') --> Func()
	 */
	JS_AHK(func, prms*) {
		wb := this.iWebCtrl
		; Stop navigation prior to calling the function, in case it uses Exit.
		wb.Stop(),  %func%(prms*)
	}