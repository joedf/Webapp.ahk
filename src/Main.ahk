#SingleInstance off
#NoTrayIcon
SendMode Input
SetWorkingDir %A_ScriptDir%

#Include Lib\Webapp.ahk
_AppStart:
;<< Header End >>


;Get our HTML DOM object
iWebCtrl := getDOM()


;connect button events
b1 := iWebCtrl.document.getElementById("MyButton1")
b2 := iWebCtrl.document.getElementById("MyButton2")
ComObjConnect(b1,"MyButton1_")
ComObjConnect(b2,"MyButton2_")


; Our Buttom Event Handlers
MyButton1_OnClick() {
	wb := getDOM()
	MsgBox % wb.Document.getElementById("MyTextBox").Value
}
MyButton2_OnClick() {
	wb := getDOM()
	FormatTime, TimeString, %A_Now%, dddd MMMM d, yyyy HH:mm:ss
    Random, x, %min%, %max%
	data := "AHK Version " A_AhkVersion " - " (A_IsUnicode ? "Unicode" : "Ansi") " " (A_PtrSize == 4 ? "32" : "64") "bit`nCurrent time: " TimeString "`nRandom number: " x
	wb.Document.getElementById("MyTextBox").value := data
}


; Our custom protocol's url event handler
app_call(args) {
	MsgBox %args%
}


;Example function to be called from the html/js source
Hello() {
	MsgBox Hello from JS_AHK :)
}