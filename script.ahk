; =====================
; CONFIGURATION
; =====================
#Requires AutoHotkey v2.0
;#NoTrayIcon
#SingleInstance

; ---- File Paths and Constants ----
global WinTitle := "ahk_exe Business Sender Pro V35 PRO.exe"
global ImportForm := "ahk_class WindowsForms10.Window.8.app.0.141b42a_r8_ad1"
global ContactFile := "C:\Users\user\Documents\Whatsappauto\ctcs.xlsx"
global PhotoFile := "C:\Users\user\Documents\Whatsappauto\photo.jpg"
global MessageFile := "C:\Users\user\Documents\Whatsappauto\message.txt"
global PerPage := 10 ; number of contacts per page
global targetFile := "" ; Will be set per profile
global Profiles := [
    "601159362833",
    "601159363618",
    "601159361617",
    "601159335834",
    "601159353732",
    "601159339348",
    "601159345446",
    "601159347077",
    "601159337961",
    "601159353311"
]

; =====================
; MAIN EXECUTION
; =====================
Main() {
    ; Grant permissions and remove read-only
    RunWait('icacls "' ContactFile '" /grant user:M', , 'Hide')
    FileSetAttrib("-R", ContactFile)

    ; Process each profile
    for profile in Profiles {
        CopyContactsXlsx(profile)
        RunProfile(profile)
    }
}

; =====================
; PROFILE AUTOMATION
; =====================
/**
 * Automates the process for a given profile:
 *  - Ensures the app window is open
 *  - Imports contacts
 *  - Attaches files and message
 *  - Sends the campaign
 */
RunProfile(profile := "") {
    tail := profile == "" ? "" : " " . profile
    if !WinExist(WinTitle) {
        Run '"C:\Program Files (x86)\Marketerpro Enterprise\Business Sender Pro V35 Pro\Business Sender Pro V35 PRO.exe"' . tail
        WinWait(WinTitle)
        while !ControlExists("WindowsForms10.SysListView32.app.0.141b42a_r8_ad13", WinTitle) {
            Sleep(500)
        }
    }

    if !WinExist(WinTitle) {
        MsgBox "Business Sender Pro V35 is NOT open."
        return
    }

    WinActivate(WinTitle)
    Sleep(500)

    ; Import contacts from file
    SetControlDelay -1
    ControlClick "WindowsForms10.SysListView32.app.0.141b42a_r8_ad13", , , "Right"
    Sleep(500)
    Send("{Down}")
    Sleep(500)
    Send("{Enter}")
    Sleep(12000)

    ; Click "Browse" button
    ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r8_ad14", WinTitle
    Sleep(1000)
    Send(targetFile)
    Send("{Enter}")
    Sleep(15000)

    TreatHeader()

    ; Click import button
    ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r8_ad17", ImportForm
    Sleep(15000)

    ; Attach photo if exists
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

    ; Attach message if exists
    if FileExist(MessageFile) {
        MyText := FileRead(MessageFile, "UTF-8")
        Sleep(1000)
    } else {
        MsgBox("Message content does not exist")
        ExitApp
    }
    EditPaste(MyText, "WindowsForms10.EDIT.app.0.141b42a_r8_ad12", WinTitle)

    ; Wait for ready
    MaxWaitTime := 600
    WaitedTime := 0
    Sleep(500)
    Loop {
        WinActivate(WinTitle)
        Sleep(500)
        Text := ControlGetText("WindowsForms10.STATIC.app.0.141b42a_r8_ad11", WinTitle)
        if (Text == "Ready") {
            break
        }
        Sleep(500)
        WaitedTime += 1
        if (WaitedTime >= MaxWaitTime) {
            ProcessClose("Business Sender Pro V35 PRO.exe")
            return
        }
    }

    ; Click "send now" button
    ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r8_ad123", WinTitle
    Sleep(1500)

    ; Check if can send
    if ControlExists("Static2", "#32770") {
        MsgBox("Cannot send")
        ExitApp
    }

    ; Click blinde mode
    ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r8_ad15", WinTitle
    Sleep(500)
    ; Click ok button
    ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r8_ad14", WinTitle

    ; Wait for campaign completion
    Loop {
        WinActivate(WinTitle)
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

; =====================
; EXCEL UTILITIES
; =====================
/**
 * Copies the contact file for a profile and trims to PerPage rows.
 */
CopyContactsXlsx(profile := "default") {
    startRow := 2
    if !FileExist(ContactFile) {
        MsgBox("Error: Unable to read the ctcs.xlsx file.")
        ExitApp
    }
    global targetFile := "C:\Users\user\Documents\Whatsappauto\" . profile . "_" . A_Now . ".xlsx"
    FileCopy(ContactFile, targetFile, true)
    xl := ComObject("Excel.Application")
    xl.Visible := false
    wb := xl.Workbooks.Open(targetFile)
    try {
        totalRows := xl.ActiveSheet.UsedRange.Rows.Count
        if (totalRows == 1) {
            wb.Close()
            xl.Quit()
            ExitApp
        }
        endRow := startRow + PerPage - 1
        if (totalRows > endRow)
            xl.ActiveSheet.Rows((endRow + 1) . ":" . totalRows).EntireRow.Delete
    } catch {
        MsgBox("An error occurred while deleting rows.")
        wb.Save()
        wb.Close()
        xl.Quit()
        ExitApp
    }
    wb.Save()
    wb.Close()
    xl.Quit()
}

/**
 * Deletes PerPage rows from the contact file after sending.
 */
DeleteContactsXlsx() {
    startRow := 2
    endRow := startRow + PerPage - 1
    if !FileExist(ContactFile) {
        MsgBox("Error: Unable to read the ctcs.xlsx file.")
        ExitApp
    }
    xl := ComObject("Excel.Application")
    xl.Visible := false
    wb := xl.Workbooks.Open(ContactFile)
    try {
        totalRows := xl.ActiveSheet.UsedRange.Rows.Count
        xl.ActiveSheet.Rows(startRow . ":" . Min(totalRows, endRow)).EntireRow.Delete
    } catch {
        MsgBox("An error occurred while deleting rows.")
        ExitApp
    }
    wb.Save()
    wb.Close()
    xl.Quit()
}

; =====================
; UI/CONTROL UTILITIES
; =====================
/**
 * Checks if a control exists in a window.
 */
ControlExists(ControlName, WinTitle) {
    try {
        ControlHwnd := ControlGetHwnd(ControlName, WinTitle)
        return true
    } catch {
        return false
    }
}

/**
 * Sets import options in the import dialog.
 */
TreatHeader() {
    ; click "Use first row as header"
    ControlClick "WindowsForms10.BUTTON.app.0.141b42a_r8_ad13", ImportForm
    Sleep(300)
    ; click "Reove duplicates"
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

; =====================
; SCRIPT ENTRY POINT
; =====================
Main()