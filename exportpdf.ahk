; AutoHotkey script for PDF export automation
; This is a basic template. Customize as needed.
Global filename
; Example hotkey: Ctrl+Alt+E to export PDF
;
; Add your PDF export logic here
CoordMode("Mouse", "Window")
filename := [13, 14, 15, 16]
MsgBox(" Start", , " T1")
len := filename.Length
loop len
{

    Sleep(1000)
    ;<open word>
    winActivate "ใบเสร็จรับเงิน - Word"
    Sleep(1000)
    WinActive("ใบเสร็จรับเงิน - Word")
    ;<open word>
    Sleep(1000)
    if (A_Index > 1)
    {
        nextrecord()
        Sleep(1000)
        savepdf(filename[A_Index])
    } else {
        savepdf(filename[A_Index])
    }

}
MsgBox(" complete", , " T1")
ExitApp

^!x:: ;exit hotkey
{
    MsgBox(" Exit app", , " T1")
    ExitApp
}


nextrecord() {
    ;MsgBox(" Next record triggered!" ,, " T1" )
    ;Sleep(1000)
    ; Define an array of keys and key combinations
    ; keysToSend := ["{Alt}", "M", "x", "{Enter}", "!{Tab}"]
    keysToSend := ["{Alt}", "M", "x"]

    ; Loop through the array and send each key
    for key in keysToSend {
        Send(key)
        Sleep(1000) ; Optional: Add a small delay (100ms) between keys
    }

}

savepdf(fn) {
    ;MsgBox(" Save PDF triggered!" ,, " T1" )
    ;Sleep(1000)
    ; Define an array of keys and key combinations for saving PDF
    keysToSend := ["{Alt}", "F", "{A}", "Y4"]
    Sleep(1000)
    ; Loop through the array and send each key
    for key in keysToSend {
        Send(key)
        Sleep(1000) ; Optional: Add a small delay (100ms) between keys
    }
    WinActive("Save As")
    Sleep(200)
    MouseClick "left", 453, 449 ;click choose file type PDF
    Sleep(1000)
    loop 6
    {
        Send("{Down}")
        Sleep(100)
    }
    Sleep(500)
    Send("{Enter}")

    ;Change filename
    Sleep(1500)
    Send("+{Tab}") ;sift+tab to filename
    Sleep(1500)
    fname := 'ใบเสร็จรับเงิน No.' . fn . '.pdf'

    Send(fname)
    Sleep(1500)
    Send("{Enter}")
    Sleep(3000)

    ;click save
    Send "{Alt down}{S}{Alt up}"
    Sleep(1000)

    ;close pdf new one
    Send "{Alt down}{F4}{Alt up}"


}