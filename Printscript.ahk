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
    


^PrintScreen:: {
    try {
        print_count := 0
        loop {
            winTitle := WinGetTitle("A")
            if (!RegExMatch(winTitle, "\.pdf.*Adobe\xA0Acrobat")) {
                if (print_count = 0) {
                    throw Error(messages["msg_error"][lang])
                }
                else {
                    break
                }
            }

            Send("^{End}")
            Sleep 500
            Send("^p")
            WinWaitActive("ahk_class  #32770")
            Sleep 200
            Send("{Tab 7}")
            Sleep 300
            Send("{Right 2}")
            Sleep 300
            Send ("{Tab}")
            Sleep 200
            Send ("^c")
            pages := A_Clipboard
            Sleep 200                           
            Send ("1 - " pages - 1)
            Sleep 200
            Send ("{Enter}")
            WinWaitClose("ahk_class #32770")
            print_count++
            Send ("^{F4}")
            WinWaitClose(winTitle)
        }
    MsgBox print_count . " " . messages["msg_success"][lang], messages["title_success"][lang], "Iconi"
    }
    catch Error as fail{
        MsgBox fail.Message, messages["title_error"][lang], "IconX"
    }
}
