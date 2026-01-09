#Requires AutoHotkey v2.0

WM_COPYDATA := 0x004A
CMD := Map("SetText",1,"InsertText",2,"SetFile",3,"InsertFile",4,"RequestTemp",5,"SetTitle",8,"GetTitle",9,"SetName",11,"GetName",12,"GotoEnd",14,"FixFont",15)
PATH := ".\Textes\"
global reqnb_title := ""
OnMessage(WM_COPYDATA, CopyDataHandler)


target := WinExist("ahk_exe RadEdit.exe")    ; or WinExist("RadEdit ahk_exe RadEdit.exe")
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
        target := WinExist("ahk_exe RadEdit.exe")
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
    if WinWaitActive("ahk_exe RadEdit.exe", , 3)
    {
        Send "^!{Right}"
        return
    }

    ; --- Si la première tentative échoue, on temporise et recommence ---
    Sleep 200
    WinActivate(target)
    if !WinWaitActive("ahk_exe RadEdit.exe", , 3) {
        MsgBox "RadEdit introuvable ou inactive après double tentative."
        ExitApp
    }

    Send "^!{Right}"
}


if (A_Args.Length == 7) {
    proc := A_Args[1]
    modal := A_Args[2]
    loc := A_Args[3]
    first := A_Args[4]
    reqnb := A_Args[5]
    patdos := A_Args[6]
    patnom := A_Args[7]

    safeProc := SanitizeFileName(proc)

    if (target) {
    skipCorps := false
    if first = 1
        skipCorps := InitRadEdit(target, reqnb, patdos, patnom, modal, loc, safeProc)

    if !skipCorps {
	SendCopyData(target, CMD["GotoEnd"], "")
	titleRtf := IsSRJ(loc)
            ? "{\rtf1\ansi\b\ul " proc "\b0\ul0\par\par}"
            : "{\rtf1\ansi\par\par\par\b\ul " proc "\b0\ul0\par\par}"
        SendCopyData(target, CMD["InsertText"], titleRtf)
	if !InsertFileIfExists(PATH "Corps\" modal "\" safeProc ".rtf")
		if !InsertFileIfExists(PATH "Corps\" safeProc ".rtf")
		        InsertFileIfExists(PATH "Corps\" modal ".rtf")
    }

    SendCopyData(target, CMD["FixFont"], "Arial;10")
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
    ; (facultatif) sauvegarde de l’ancien si présent…
    oldTitle := GetTitleFromRadEdit(target)
    if (oldTitle != "") {
        safeOld := RegExReplace(oldTitle, "[^\w-]", "_")
	SendCopyData(target, CMD["RequestTemp"], "RadEdit_" safeOld ".rtf")
        global report_file := ""
        t0 := A_TickCount
        while (report_file = "" && (A_TickCount - t0 < 1000))
            Sleep 20
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
	plain := RegExReplace(plain, "\\[a-zA-Z]+(?:-?\d+)?[ ]?", "")   
	plain := RegExReplace(plain, "\{\\[^}]+\}", "") 
	plain := StrReplace(plain, "{", "")
	plain := StrReplace(plain, "}", "")
	plain := RegExReplace(plain, "[\x00-\x08\x0B-\x1F]", "")  
	found := InStr(StrLower(plain), StrLower(proc)) > 0

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
    static CMD_TITLERESPONSE := 10
    global reqnb_title

    cmd  := NumGet(lParam, 0, "UPtr")
    size := NumGet(lParam, A_PtrSize, "UInt")
    text := StrGet(NumGet(lParam, 2*A_PtrSize, "Ptr"), size/2, "UTF-16")

    if (cmd = CMD_TITLERESPONSE) {
        reqnb_title := Trim(text)
    }
    return true
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
