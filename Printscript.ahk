#Requires AutoHotkey v2.0

langID := DllCall("GetUserDefaultUILanguage")   ;get used Language of the System / User

;Set user language, if not supported fallback to english
switch langID {
    case 1031: lang := "de"
    case 1033: lang := "en"
    Default: lang := "en"
}

; Messages in the different languages
messages := Map(
    "msg_error", Map("de", "Kein PDF-Dokument geöffnet oder nicht im Fokus!", "en", "No PDF document opened or not in focus!"),
    "title_error", Map("de", "Ups, da ist was schiefgelaufen!", "en", "Oops, there went something wrong!"),
    "msg_success", Map("de", "Dokumente erfolgreich gedruckt!", "en", "documents printed successfully!"),
    "title_success", Map("de", "Das hat prima funktioniert!", "en", "This worked like a charm!")
)

;Default-Print: Print all documents but don't print last page of each document
^PrintScreen:: {
    printDocuments(0, 1)
}    

;Custom-Print with dialog for entering the number of documents and number of pages to not to print
^!PrintScreen:: {
    try{
        winTitle := WinGetTitle("A")
        ;get process name of the active window (e.g. acrobat or acrobat reader)
        proc := WinGetProcessName("A")
        if (!RegExMatch(winTitle, "\.pdf.*Adobe\xA0Acrobat")) {
            throw Error(messages["msg_error"][lang])
        }
        MyInput := Gui(, "Settings")
        MyInput.MarginX := 100
        MyInput.MarginY := 20
        MyInput.AddText("x25 y20", "Anzahl Dokumente:")
        MyInput.AddText(, "Seiten nicht drucken:")
        Pages := MyInput.AddEdit("ys-5 x+30 w50 Number")
        Skip := MyInput.AddEdit("w50 Number")
        MyInput.Add("Button","Default w80", "OK").OnEvent("Click", OK_Click)
        MyInput.Show()

        OK_Click(*) {
            MyInput.Hide()
            ;if Edit empty then 0, else Edit.Value
            pages := (Pages.Value = "") ? 0 : Pages.Value
            skip := (Skip.Value = "") ? 0 : Skip.Value
            
            WinWaitActive("ahk_exe " . proc,,5)
            printDocuments(pages, skip)
        }
    }
    catch Error as fail{
        MsgBox fail.Message, messages["title_error"][lang], "IconX"
    }
}





printDocuments(pdfTotal, skipLastPages) {
    try {
        print_count := 0
        loop {
            winTitle := WinGetTitle("A")
            ;get process name of the active window (e.g. acrobat or acrobat reader)
            proc := WinGetProcessName("A")
            if (!RegExMatch(winTitle, "\.pdf.*Adobe\xA0Acrobat")) {
                if (print_count = 0) {
                    throw Error(messages["msg_error"][lang])
                }
                else {
                    break
                }
            }
            if (!pdfTotal = 0) {
                if (print_count = pdfTotal) {
                    break
                }
            }
            ;empty clipboard
            A_Clipboard := ""
            RegExMatch(winTitle, ".*\.pdf", &docTitle)
            Send("^{End}")
            Sleep 500
            Send("^p")
            WinWaitActive("ahk_class #32770 ahk_exe " . proc)
            dialogID := WinGetID("A")
            Sleep 200
            Send("{Tab 7}")
            Sleep 300
            Send("{Right 2}")
            Sleep 300
            Send ("{Tab}")
            Sleep 500
            ;copy how many pages are in the document
            Send ("^c")
            ;wait until the page-number is in clipboard
            ClipWait(2)
            pagesTotal := A_Clipboard
            Sleep 500
            ;determine how many pages to print
            printPages := PagesToPrint(pagesTotal, skipLastPages)
            Send (printPages)
            Sleep 500
            ;if it is the first document to print, wait for the user to hit Print
            if (print_count = 0) {
                MsgBox "Bitte die Druckeinstellungen überprüfen mit Drucken bestätigen!"
            }
            else {
                Send ("{Enter}")
            }
            print_count++
            ;Wait until print dialog has been closed
            WinWaitClose("ahk_class  #32770 ahk_exe " . proc)
            ;Wait until printing progess has been opened
            WinWaitActive("ahk_class  #32770 ahk_exe " . proc)
            ;Wait until printing progress has been closed
            WinWaitClose("ahk_class  #32770 ahk_exe " . proc)
            Sleep 500
            ;Close the actual document
            Send("^{F4}")
            WinWaitClose(winTitle)
            Sleep 200
            ;wenn Anzahl an Dokumenten gedruckt - break
            ;MsgBox pdfTotal . " " . print_count
            if (pdfTotal > 0 and print_count = pdfTotal) {
                break
            }
            ;Wait until Adobe Acrobat (Reader) has focus
            WinWaitActive("ahk_exe" . proc)
        }
    MsgBox print_count . " " . messages["msg_success"][lang], messages["title_success"][lang], "Iconi"
    }
    catch Error as fail{
        MsgBox fail.Message, messages["title_error"][lang], "IconX"
    }
}

;function to determine how many pages to print
PagesToPrint(total, skip) {
    if (total - skip <= 1){
        return "1"
    }
    else {
        return "1 - " . total - skip
    }
}