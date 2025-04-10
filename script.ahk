#Requires AutoHotkey v2.0
#NoTrayIcon
#SingleInstance


WinTitle := "ahk_exe Business Sender Pro V35 PRO.exe"
ImportForm := "ahk_class WindowsForms10.Window.8.app.0.141b42a_r7_ad1"
;ContactFile := "D:\Works\whatsappauto\ctcs.xlsx"
ContactFile := "D:\Works\whatsappauto\import.csv"
PhotoFile := "D:\Works\whatsappauto\photo.jpg"
MessageFile := "D:\Works\whatsappauto\message.txt"

Run '"C:\Program Files (x86)\Marketerpro Enterprise\Business Sender Pro V35 Pro\Business Sender Pro V35 PRO.exe"'

; Wait a moment
Sleep(10000)

if WinExist(WinTitle)
{
    ; active window
    WinActivate(WinTitle)
    Sleep(500)

    ; right click in number list, and then choose "imports from files"
    SetControlDelay -1
    ControlClick "WindowsForms10.SysListView32.app.0.141b42a_r7_ad13", , , "Right"
    Sleep(500)
    Send("{Down}")
    Sleep(500)
    Send("{Enter}")
    Sleep(2000)

    ; click "Browse" button
    ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r7_ad14", WinTitle
    Sleep(1000)

    ; file path in File Open
    Send(ContactFile)
    Send("{Enter}")
    Sleep(5000)

    ;;
    ;;
    ;TreatHeader()
    ;;
    ;;

    ; click import button
    ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r7_ad17", ImportForm
    Sleep(5000)

    ; Attach Files -> menu, photo
    if FileExist(PhotoFile) {
        WinActivate(WinTitle)
        ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r7_ad122", WinTitle
        Send("{Down}")
        Send("{Enter}")
        Sleep(500)
        Send(PhotoFile)
        Send("{Enter}")
        Sleep(1000)
    } else {
        ;MsgBox("Photo file does not exist")
    }

    ; open message content, copy all text
    if FileExist(MessageFile) {
        MyText := FileRead(MessageFile, "UTF-8")
        Sleep(1000)
    } else {
        MsgBox("Message content does not exist")
        ExitApp
    }

    ; paste the content into message box
    EditPaste(MyText, "WindowsForms10.EDIT.app.0.141b42a_r7_ad12", WinTitle)

    ; wait ready
    Sleep(500)
    Loop {
        WinActivate(WinTitle)
        Sleep(500)
        ; Retrieve the button's caption
        Text := ControlGetText("WindowsForms10.STATIC.app.0.141b42a_r7_ad11", WinTitle)

        if (Text == "Ready") {
            break
        }
        Sleep(500)
    }

    ; click "send now" button
    ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r7_ad123", WinTitle
    Sleep(1500)

    ; check if can send
    if ControlExists("Static2", "#32770") {
        MsgBox("Cannot send")
        ExitApp
    } else {
        ;MsgBox("The control does not exist.")
    }

    ; click blinde mode
    ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r7_ad15", WinTitle
    Sleep(500)

    ;;
    ;;

    ;MsgBox "Really send?"
    ;ExitApp
    ;;
    ;;

    ; click ok button
    ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r7_ad14", WinTitle

    ; wait to send all, close
    ;Sleep(10000)
    Loop {
        ; Retrieve the modal's caption
        if ControlExists("Static2", "#32770") {
            Text := ControlGetText("Static2", WinTitle)

            if (Text == "Campaign has been done!") {
                break
            }
        }
        Sleep(1000)
    }
    ProcessClose("Business Sender Pro V35 PRO.exe")
}
else
{
    MsgBox "Business Sender Pro V35 is NOT open."
}


ControlExists(ControlName, WinTitle) {
    try {
        ControlHwnd := ControlGetHwnd(ControlName, WinTitle)
        return true
    } catch {
        return false
    }
}

TreatHeader() {
    ; check "Use first row as header"
    ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r7_ad13", ImportForm
    Sleep(300)
    ; check "Reove duplicates"
    ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r7_ad12", ImportForm
    Sleep(300)

    ; set name field
    ControlClick "WindowsForms10.COMBOBOX.app.0.141b42a_r7_ad12", ImportForm
    Sleep(300)
    Send("{Down}")
    Sleep(300)
    Send("{Enter}")
    ; set number field
    ControlClick "WindowsForms10.COMBOBOX.app.0.141b42a_r7_ad13", ImportForm
    Sleep(300)
    Send("{Down}")
    Sleep(300)
    Send("{Enter}")
}