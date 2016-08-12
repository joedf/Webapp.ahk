#SingleInstance off
#NoTrayIcon
SendMode Input

#Include Lib\Webapp.ahk
__Webapp_AppStart:
;<< Header End >>


;Get our HTML DOM object
iWebCtrl := getDOM()


; Our custom protocol's url event handler
app_call(args) {
	MsgBox %args%
}


; Functions to be called from the html/js source
Hello() {
	MsgBox Hello from JS_AHK :)
}
Multiply(a,b) {
	;MsgBox % a " * " b " = " a*b
	return a * b
}
MyButton1() {
	wb := getDOM()
	MsgBox % wb.Document.getElementById("MyTextBox").Value
}
MyButton2() {
	wb := getDOM()
	FormatTime, TimeString, %A_Now%, dddd MMMM d, yyyy HH:mm:ss
    Random, x, %min%, %max%
	data := "AHK Version " A_AhkVersion " - " (A_IsUnicode ? "Unicode" : "Ansi") " " (A_PtrSize == 4 ? "32" : "64") "bit`nCurrent time: " TimeString "`nRandom number: " x
	wb.Document.getElementById("MyTextBox").value := data
}