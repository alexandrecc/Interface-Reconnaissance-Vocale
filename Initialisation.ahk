#Requires AutoHotkey v2.0

WM_COPYDATA := 0x004A
CMD := Map(
  "SetText",1, "InsertText",2, "SetFile",3, "InsertFile",4, "RequestTemp",5,
  "Response",6, "Error",7, "SetTitle",8, "GetTitle",9, "TitleResponse",10,
  "SetName",11, "GetName",12, "NameResponse",13, "GotoEnd",14,
  "FixFont",15, "CleanUpEnd",16, "SetHtmlFile", 17, "RequestHtmlFile", 18, "HtmlFileResponse", 19,
  "SetDataContext", 20, "GetDataContext", 21, "DataContextResponse", 22
)

global PATH := ".\Textes\"
global reqnb_title := ""
global report_file := ""          ; réponse TempFileResponse (cmd 6)
global html_file := ""            ; réponse HtmlFileResponse (cmd 19)
global reqnb_name := ""           ; optionnel (cmd 13)
global last_radedit_error := ""   ; optionnel (cmd 7)
global data_context_json := ""    ; optionnel (cmd 22)
global reqnb := ""
global patdos := ""
global patnom := ""
global reqnb_norm := ""
global dateexam := ""
global proc := ""
SetTitleMatchMode 2

OnMessage(WM_COPYDATA, CopyDataHandler)


global target := WinExist("RadEdit ahk_exe RadEdit.exe")    ; or WinExist("RadEdit ahk_exe RadEdit.exe")
if !target {
    MsgBox "RadEdit window not found."
    ExitApp
}

SendCopyData(hwnd, command, text := "") {
    text := text . Chr(0)
    buf := Buffer(StrLen(text) * 2, 0)
    StrPut(text, buf, "UTF-16")
    cds := Buffer(3 * A_PtrSize, 0)
    NumPut("UPtr", command, cds, 0)
    NumPut("UInt", buf.Size, cds, A_PtrSize)
    NumPut("Ptr", buf.Ptr, cds, 2 * A_PtrSize)
    return DllCall("user32\SendMessageW", "Ptr", hwnd, "UInt", WM_COPYDATA, "Ptr", A_ScriptHwnd, "Ptr", cds.Ptr, "Ptr")
}


if (A_Args.Length == 8) {
    proc := A_Args[1]
    modal := A_Args[2]
    loc := A_Args[3]
    first := A_Args[4]
    reqnb := A_Args[5]
    patdos := A_Args[6]
    patnom := A_Args[7]
    dateexam := A_Args[8]

    ctx := BuildInitContextJson(proc, modal, loc, first, reqnb, patdos, patnom, dateexam)
    safeProc := SanitizeFileName(proc)

    if (target) {
    skipCorps := false
    if first = 1
        skipCorps := InitRadEdit(target, reqnb, patdos, patnom, modal, loc, safeProc)

    if !skipCorps {
	SendCopyData(target, CMD["GotoEnd"], "")
	titleRtf := (IsSRJ(loc) && (first = 1))
            ? "{\rtf1\ansi\b\ul " proc "\b0\ul0\par\par}"
            : "{\rtf1\ansi\par\par\par\b\ul " proc "\b0\ul0\par\par}"
        SendCopyData(target, CMD["InsertText"], titleRtf)
	if !InsertFileIfExists(PATH "Corps\" modal "\" safeProc ".rtf")
		if !InsertFileIfExists(PATH "Corps\" safeProc ".rtf")
		        InsertFileIfExists(PATH "Corps\" modal ".rtf")
    }

    SendCopyData(target, CMD["FixFont"], "Arial;10")

    safeReqnb := RegExReplace(reqnb, "[^\w-]", "_")
    htmlPath := A_Temp "\RadEdit_" safeReqnb "_" safeProc ".html"

    if FileExist(htmlPath) {
        SendCopyData(target, CMD["SetHtmlFile"], htmlPath)
    } else {
        htmlNew := SelectHtmlTemplate(modal, safeProc)
        if (htmlNew != "") {
            SendCopyData(target, CMD["SetHtmlFile"], htmlNew)
            } else {
	        SendCopyData(target, CMD["SetHtmlFile"], "")
            }
    }

    Sleep 500
    SendCopyData(target, CMD["SetDataContext"], ctx)

    PrepareForDictation()
    EraseTempFile(reqnb)

   }
}

SanitizeFileName(str) {
    invalid := '[<>:"/\\|?*]'
    clean := RegExReplace(str, invalid)
    Return Trim(clean)
}

IsSRJ(loc) {
    ; comparaison tolérante aux espaces et casse
    return RegExMatch(loc ?? "", "i)^\s*Services Radiologiques Joliette\s*$")
}


InitRadEdit(target, reqnb, patdos, patnom, modal, loc, safeProc) {

    oldTitle := GetTitleFromRadEdit(target)
    if (oldTitle != "") {
        safeOld := RegExReplace(oldTitle, "[^\w-]", "_")
	SendCopyData(target, CMD["RequestTemp"], "RadEdit_" safeOld ".rtf")
        global report_file := ""
        t0 := A_TickCount
        while (report_file = "" && (A_TickCount - t0 < 1000))
            Sleep 20

        ; --- FUTUR (à activer quand RadEdit ignorera RequestHtmlFile si HTML non actif
        ;           ET retournera quand même une réponse) ---
        ; oldSafeProc := ""  ; TODO: lire dataContext.safeproc (ancien examen) depuis RadEdit
        ; html_file := "" ;à sauvegarder dans le répertoire html si on veux que les liens fonctionnent
        ; SendCopyData(target, CMD["RequestHtmlFile"], "RadEdit_" safeOld "_" oldSafeProc ".html")
        ; t1 := A_TickCount
        ; while (html_file = "" && (A_TickCount - t1 < 1000))
        ;     Sleep 20

    }

    safeReqnb := RegExReplace(reqnb, "[^\w-]", "_")
    tempPath := A_Temp "\RadEdit_" safeReqnb ".rtf"

    if FileExist(tempPath) {
        ; --- RTF trouvé : on le CHARGE puis on FORCE la barre ---
        SendCopyData(target, CMD["SetFile"], tempPath)          ; remplace tout le contenu
        Sleep 50                                                ; laisse le temps de charger
        SendCopyData(target, CMD["SetTitle"], reqnb)            ; MAJ bandeau gauche
        SendCopyData(target, CMD["SetName"],  patdos " - " StrUpper(patnom))
	
        ; --- Vérifie si le proc est déjà présent dans le RTF ---
        skipCorps := false
        try rtfText := FileRead(tempPath, "CP1252")
        catch as err {
            MsgBox "Erreur de lecture du fichier : " err.Message
            rtfText := ""
        }

	; --- Mini-normalisation du RTF (léger) ---
	plain := DecodeRtfHex(rtfText)

	; 1) préserver les retours de paragraphes/lignes AVANT de stripper les commandes
	plain := RegExReplace(plain, "\\par\b", "`n")
	plain := RegExReplace(plain, "\\line\b", "`n")

	; 2) ensuite seulement, enlever les commandes RTF
	plain := RegExReplace(plain, "\\[a-zA-Z]+(?:-?\d+)?[ ]?", "")
	plain := RegExReplace(plain, "\{\\[^}]+\}", "")
	plain := StrReplace(plain, "{", "")
	plain := StrReplace(plain, "}", "")
	plain := RegExReplace(plain, "[\x00-\x08\x0B-\x1F]", "")

	; 3) proc doit être SEUL sur une ligne (espaces tolérés)
	escProc := RegExReplace(proc, "([\\\.\*\+\?\|\{\}\[\]\(\)\^\$\-])", "\$1")
	found := RegExMatch(plain, "im)^\s*" escProc "\s*$")

        if (found) {
	    SendCopyData(target, CMD["GotoEnd"], "")
            skipCorps := true
        } else {
            skipCorps := false
        }

        return skipCorps
    }
    else {
        ; --- Nouveau rapport : init standard ---
        SendCopyData(target, CMD["SetTitle"], reqnb)
        SendCopyData(target, CMD["SetName"],  patdos " - " StrUpper(patnom))
	SendCopyData(target, CMD["SetText"], "{\rtf1\ansi\deff0 {\fonttbl {\f0\fnil\fcharset0 Arial;}}\f0\fs20}")
	if !IsSRJ(loc) {
        if !InsertFileIfExists(PATH "Avant-titre\" modal "_" safeProc ".rtf")
            if !InsertFileIfExists(PATH "Avant-titre\" modal "_" loc ".rtf")
                InsertFileIfExists(PATH "Avant-titre\Default.rtf")
	}
	return false
    }
}


CopyDataHandler(wParam, lParam, msg, hwnd) {
    static CMD_RESPONSE := 6
    static CMD_ERROR := 7
    static CMD_TITLERESPONSE := 10
    static CMD_NAMERESPONSE := 13
    static CMD_HTMLFILERESPONSE := 19
    static CMD_DATACONTEXTRESPONSE := 22

    global report_file, html_file, reqnb_title, reqnb_name, last_radedit_error, data_context_json

    cmd  := NumGet(lParam, 0, "UPtr")
    size := NumGet(lParam, A_PtrSize, "UInt")
    text := StrGet(NumGet(lParam, 2*A_PtrSize, "Ptr"), size/2, "UTF-16")
    text := Trim(text)

    if (cmd = CMD_RESPONSE) {
        report_file := text
    } else if (cmd = CMD_HTMLFILERESPONSE) {
        html_file := text
    } else if (cmd = CMD_TITLERESPONSE) {
        reqnb_title := text
    } else if (cmd = CMD_NAMERESPONSE) {
        reqnb_name := text
    } else if (cmd = CMD_DATACONTEXTRESPONSE) {
        data_context_json := text
    } else if (cmd = CMD_ERROR) {
        last_radedit_error := text
        ; MsgBox "RadEdit error: " text  ; (optionnel) je le laisse commenté pour éviter les popups
    }
    return true
}

InsertFileIfExists(fpath) {
   if FileExist(fpath) {
      SendCopyData(target, CMD["InsertFile"], fpath)
      return 1
   }
   else {
      return 0
   }
}

PrepareForDictation() {
    target := 0
    maxTries := 10        ; nombre maximum de tentatives
    interval := 200       ; délai entre chaque tentative (ms)

    ; === Boucle de détection de la fenêtre RadEdit ===
    Loop maxTries {
        target := WinExist("RadEdit ahk_exe RadEdit.exe")
        if (target)
            break
        Sleep interval
    }

    if !target {
        MsgBox "Fenêtre RadEdit introuvable après " (maxTries * interval) " ms."
        ExitApp
    }

    ; === Activation de la fenêtre ===
    WinActivate(target)
    if WinWaitActive("RadEdit ahk_exe RadEdit.exe", , 3)
    {
        Send "^!{Right}"
        return
    }

    ; --- Si la première tentative échoue, on temporise et recommence ---
    Sleep 200
    WinActivate(target)
    if !WinWaitActive("RadEdit ahk_exe RadEdit.exe", , 3) {
        MsgBox "RadEdit introuvable ou inactive après double tentative."
        ExitApp
    }

    Send "^!{Right}"
}
GetTitleFromRadEdit(hwnd) {
    global reqnb_title
    reqnb_title := ""   ; reset

    SendCopyData(hwnd, CMD["GetTitle"], "")

    t0 := A_TickCount
    while (reqnb_title = "" && (A_TickCount - t0 < 1000))
        Sleep 20

    return reqnb_title
}

EraseTempFile(reqnb) {
    tempFile := A_Temp "\RadEdit_" reqnb ".rtf"
    try {
        if FileExist(tempFile) {
            FileDelete(tempFile)
        }
    } catch as err {
        MsgBox "Impossible d'effacer le fichier temporaire :`n" tempFile "`n`nErreur : " err.Message, "ATTENTION", 262144
    }
}

DecodeRtfHex(str) {
    pos := 1, out := ""
    while (pos <= StrLen(str)) {
        if RegExMatch(str, "i)\\'([0-9A-F]{2})", &m, pos) {
            ; texte avant l’occurrence
            out .= SubStr(str, pos, m.Pos(0) - pos)
            ; sous-capture 1 = les 2 hex (ex: "C9") → caractère CP1252
            out .= Chr("0x" m[1])
            ; avancer après l’occurrence
            pos := m.Pos(0) + m.Len(0)
        } else {
            out .= SubStr(str, pos)
            break
        }
    }
    return out
}

JsonEscape(s) {
    s := s ?? ""
    s := StrReplace(s, "\", "\\")
    s := StrReplace(s, '"', '\"')
    s := StrReplace(s, "`r", "\r")
    s := StrReplace(s, "`n", "\n")
    s := StrReplace(s, "`t", "\t")
    return s
}

BuildInitContextJson(proc, modal, loc, first, reqnb, patdos, patnom, dateexam) {
    studydate := GetStudyDate(dateexam)
    safeproc := SanitizeFileName(proc)

    ; IMPORTANT: payload "replace" attendu par RadEdit:
    ; {"__mode":"replace","data":{...}}
    return "{"
        . '"__mode":"replace",'
        . '"data":{'
        . '"proc":"'      JsonEscape(proc)      '",'
        . '"modal":"'     JsonEscape(modal)     '",'
        . '"loc":"'       JsonEscape(loc)       '",'
        . '"first":"'     JsonEscape(first)     '",'
        . '"reqnb":"'     JsonEscape(reqnb)     '",'
        . '"patdos":"'    JsonEscape(patdos)    '",'
        . '"patnom":"'    JsonEscape(patnom)    '",'
        . '"safeproc":"'  JsonEscape(safeproc)  '",'
        . '"studydate":"' JsonEscape(studydate) '"'
        . "}}"
}

SelectHtmlTemplate(modal, safeProc) {
    global PATH

    p1 := PATH "HTML\" modal "\" safeProc ".html"
    if FileExist(p1)
        return p1

    p2 := PATH "HTML\" safeProc ".html"
    if FileExist(p2)
        return p2

    ;p3 := PATH "HTML\" modal ".html"
    ;if FileExist(p3)
    ;    return p3

    p4 := PATH "HTML\Default.html"
    if FileExist(p4)
        return p4

    return ""
}

GetStudyDate(dateexam) {
    ; Retourne "JJ/MM/AAAA" ou "" si format non reconnu

    dateexam := Trim(dateexam)

    ; 1) Format ISO: YYYY-MM-DD (ex: 2025-10-22)
    if RegExMatch(dateexam, "^\s*(\d{4})-(\d{2})-(\d{2})", &m) {
        y  := m[1], mo := m[2], d := m[3]
        return Format("{:02}/{:02}/{}", d, mo, y)
    }

    ; 2) Format US: MM/DD/YYYY (ex: 10/22/2025)
    if RegExMatch(dateexam, "^\s*(\d{1,2})/(\d{1,2})/(\d{4})", &m) {
        mo := m[1], d := m[2], y := m[3]
        return Format("{:02}/{:02}/{}", d, mo, y)
    }

    return ""
}
