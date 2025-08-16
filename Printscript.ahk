#Requires AutoHotkey v2.0

langID := DllCall("GetUserDefaultUILanguage")   ;get used Language of the System / User

;Set user language, if not supported fallback to english
switch langID {
    case 1031: lang := "de"
    case 1033: lang := "en"
    Default: lang := "en"
}

; text in the different languages
translation := Map(
    "msg_error", Map("de", "Kein PDF-Dokument geöffnet oder nicht im Fokus!", "en", "No PDF document opened or not in focus!"),
    "title_error", Map("de", "Ups, da ist was schiefgelaufen!", "en", "Oops, there went something wrong!"),
    "msg_success", Map("de", "Dokumente erfolgreich gedruckt!", "en", "documents printed successfully!"),
    "title_success", Map("de", "Das hat prima funktioniert!", "en", "This worked like a charm!"),
    "msg_check", Map("de", "Bitte die Druckeinstellungen überprüfen und mit `"Drucken`" bestätigen.`nDas Script wird dann ihren Druckauftrag automatisch abarbeiten.", "en", "Please check the print settings and confirm with `"Print`".`nThe script will then process your print job automatically."),
    "title_check", Map("de", "Durckeinstellungen prüfen", "en", "Check Print Settings"),
    "dialog_title", Map("de", "Anzahl Dokumente & wegzulassender Endseiten", "en", "Number of Documents & End Pages to Skip"),
    "dialog_headline", Map("de", "Druckeinstellungen:", "en", "Print Settings"),
    "dialog_lbl_page", Map("de", "Anzahl Dokumente:", "en", "Number of Documents:"),
    "dialog_lbl_skip", Map("de", "Endseiten überspringen:", "en", "Skip Endpages:")  
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
            throw Error(translation["msg_error"][lang])
        }
        MyInput := Gui(, translation["dialog_title"][lang])
        MyInput.AddText("x40 y20", translation["dialog_headline"][lang])
        MyInput.AddText("Section", translation["dialog_lbl_page"][lang])
        MyInput.AddText(, translation["dialog_lbl_skip"][lang])
        Pages := MyInput.AddEdit("ys-5 x+50 w50 Number")
        Skip := MyInput.AddEdit("w50 Number")
        MyInput.AddButton("Default w80 x185 y115", "OK").OnEvent("Click", OK_Click)
        MyInput.Show("w450")
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
        MsgBox fail.Message, translation["title_error"][lang], "IconX"
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
                    throw Error(translation["msg_error"][lang])
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
                MsgBox translation["msg_check"][lang], translation["title_check"][lang], "Iconi"
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
    MsgBox print_count . " " . translation["msg_success"][lang], translation["title_success"][lang], "Iconi"
    }
    catch Error as fail{
        MsgBox fail.Message, translation["title_error"][lang], "IconX"
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