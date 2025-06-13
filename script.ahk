#Requires AutoHotkey v2.0
;#NoTrayIcon
#SingleInstance


WinTitle := "ahk_exe Business Sender Pro V35 PRO.exe"
ImportForm := "ahk_class WindowsForms10.Window.8.app.0.141b42a_r8_ad1"
ContactFile := "C:\Users\user\Documents\Whatsappauto\ctcs.xlsx"
targetFile := "C:\Users\user\Documents\Whatsappauto\target.xlsx"
;ContactFile := "C:\Users\user\Documents\Whatsappauto\import.csv"
PhotoFile := "C:\Users\user\Documents\Whatsappauto\photo.jpg"
MessageFile := "C:\Users\user\Documents\Whatsappauto\message.txt"
PerPage := 2 ; number of contacts per page

RunWait('icacls "' ContactFile '" /grant user:M', , 'Hide')
FileSetAttrib("-R", ContactFile)

Page := 1
;CreateCSV()
CopyContactsXlsx()
RunProfile()

CopyContactsXlsx()
RunProfile('01121659058')

CopyContactsXlsx()
RunProfile('01121679033')

CopyContactsXlsx()
RunProfile('01121682805')

CopyContactsXlsx()
RunProfile('01121695920')

CopyContactsXlsx()
RunProfile('01121695970')

CopyContactsXlsx()
RunProfile('01121697582')

CopyContactsXlsx()
RunProfile('01163619946')

CopyContactsXlsx()
RunProfile('01163629915')

RunProfile(profile := "") {
    tail := profile == "" ? "" : " " . profile
    Run '"C:\Program Files (x86)\Marketerpro Enterprise\Business Sender Pro V35 Pro\Business Sender Pro V35 PRO.exe"' . tail

    ; Wait a moment
    ;Sleep(10000)
    WinWait(winTitle)
    while !ControlExists("WindowsForms10.SysListView32.app.0.141b42a_r8_ad13", WinTitle) {
        Sleep(500)
    }


    if WinExist(WinTitle)
    {
        ; active window
        WinActivate(WinTitle)
        Sleep(500)

        ; right click in number list, and then choose "imports from files"
        SetControlDelay -1
        ControlClick "WindowsForms10.SysListView32.app.0.141b42a_r8_ad13", , , "Right"
        Sleep(500)
        Send("{Down}")
        Sleep(500)
        Send("{Enter}")
        Sleep(2000)

        ; click "Browse" button
        ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r8_ad14", WinTitle
        Sleep(1000)

        ; file path in File Open
        Send(targetFile)
        Send("{Enter}")
        Sleep(5000)

        TreatHeader()

        ; click import button
        ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r8_ad17", ImportForm
        Sleep(5000)

        ; Attach Files -> menu, photo
        if FileExist(PhotoFile) {
            WinActivate(WinTitle)
            ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r8_ad122", WinTitle
            Send("{Down}")
            Send("{Enter}")
            Sleep(500)
            Send(PhotoFile)
            Send("{Enter}")
            Sleep(1000)
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
        EditPaste(MyText, "WindowsForms10.EDIT.app.0.141b42a_r8_ad12", WinTitle)

        ; wait ready
        Sleep(500)
        Loop {
            WinActivate(WinTitle)
            Sleep(500)
            ; Retrieve the button's caption
            Text := ControlGetText("WindowsForms10.STATIC.app.0.141b42a_r8_ad11", WinTitle)

            if (Text == "Ready") {
                break
            }
            Sleep(500)
        }

        ; click "send now" button
        ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r8_ad123", WinTitle
        Sleep(1500)

        ; check if can send
        if ControlExists("Static2", "#32770") {
            MsgBox("Cannot send")
            ExitApp
        } else {
            ;MsgBox("The control does not exist.")
        }

        ; click blinde mode
        ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r8_ad15", WinTitle
        Sleep(500)

        ;;
        ;;
        ;ProcessClose("Business Sender Pro V35 PRO.exe")
        ;return
        ;ExitApp
        ;;
        ;;

        ; click ok button
        ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r8_ad14", WinTitle

        ; wait to send all, close
        Loop {
            WinActivate(WinTitle)
            ; Retrieve the modal's caption
            if ControlExists("Static2", "ahk_class #32770") {
                Text := ControlGetText("Static2", "ahk_class #32770")

                if (Text == "Campaign has been done!") {
                    DeleteContactsXlsx()
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
    ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r8_ad13", ImportForm
    Sleep(300)
    ; check "Reove duplicates"
    ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r8_ad12", ImportForm
    Sleep(300)

    ; set name field
    ControlClick "WindowsForms10.COMBOBOX.app.0.141b42a_r8_ad12", ImportForm
    Sleep(300)
    Send("{Down}")
    Sleep(300)
    Send("{Enter}")
    ; set number field
    ControlClick "WindowsForms10.COMBOBOX.app.0.141b42a_r8_ad13", ImportForm
    Sleep(300)
    Send("{Down}")
    Sleep(300)
    Send("{Enter}")
}

CreateCSV() {
    startRow := PerPage * (Page - 1) + 2

    ; Read the source file
    fileContent := FileRead(ContactFile, "UTF-8")
    if !fileContent {
        ExitApp
    }

    ; Split content into lines
    lines := StrSplit(fileContent, "`n")

    if lines.Length < startRow {
        ExitApp
    }

    ; Open the target file for writing
    if FileExist(targetFile) {
        FileDelete(targetFile)
    }
    FileAppend(lines[1] "`n", targetFile)  ; Write the header row

    ; Extract rows based on startRow and rowCount
    endRow := startRow + PerPage - 1
    for i, line in lines {
        if i >= startRow && i <= endRow {
            FileAppend(line "`n", targetFile)
        }
    }
}

CopyContactsXlsx() {
    startRow := 2

    if !FileExist(ContactFile) {
        MsgBox("Error: Unable to read the ctcs.xlsx file.")
        ExitApp
    }
    global targetFile := "C:\Users\user\Documents\Whatsappauto\target_" . A_Now . ".xlsx"
    FileCopy(ContactFile, targetFile, true)


    xl := ComObject("Excel.Application")
    xl.Visible := false ; Set true if you want to see the file being modified

    ; Open the workbook
    wb := xl.Workbooks.Open(targetFile)

    try {
        totalRows := xl.ActiveSheet.UsedRange.Rows.Count
        if (totalRows == 1) {
            wb.Close()
            xl.Quit()
            ExitApp
        }
        endRow := startRow + PerPage - 1

        ; Delete rows after selected range
        if (totalRows > endRow)
            xl.ActiveSheet.Rows((endRow + 1) . ":" . totalRows).EntireRow.Delete
    } catch {
        MsgBox("An error occurred while deleting rows.")
        wb.Save()
        wb.Close()
        xl.Quit()
        ExitApp
    }

    ; Save and close the workbook (optional)
    wb.Save()
    wb.Close()
    xl.Quit()
}

DeleteContactsXlsx() {
    startRow := 2
    endRow := startRow + PerPage - 1

    if !FileExist(ContactFile) {
        MsgBox("Error: Unable to read the ctcs.xlsx file.")
        ExitApp
    }

    xl := ComObject("Excel.Application")
    xl.Visible := false ; Set true if you want to see the file being modified

    ; Open the workbook
    wb := xl.Workbooks.Open(ContactFile)
    try {
        totalRows := xl.ActiveSheet.UsedRange.Rows.Count
        xl.ActiveSheet.Rows(startRow . ":" . Min(totalRows, endRow)).EntireRow.Delete
    } catch {
        MsgBox("An error occurred while deleting rows.")
        ExitApp
    }

    ; Save and close the workbook (optional)
    wb.Save()
    wb.Close()
    xl.Quit()
}

DeleteRowsCSV(RowsToDelete := 10) {
    HeaderRow := ""  ; Store the header row

    ; Read the file and remove rows
    output := ""
    try {
        Loop Read, ContactFile {
            if (A_Index = 1) {
                HeaderRow := A_LoopReadLine  ; Save the header
                output .= HeaderRow "`n"  ; Keep the header
                continue
            }
            if (A_Index > 1 && A_Index <= (1 + RowsToDelete)) {
                continue  ; Skip the next 10 rows after the header
            }
            output .= A_LoopReadLine "`n"  ; Keep all other rows
        }

        ; Overwrite the file with the modified content
        FileDelete(ContactFile)
        FileAppend(output, ContactFile)
    } catch {
        MsgBox("Error: The file could not be processed.")
    }
}