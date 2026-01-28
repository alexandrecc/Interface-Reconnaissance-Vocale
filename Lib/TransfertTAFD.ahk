LUNGRADS_MARK := "X"

TransfertTAFD(word, doc, report_file, mode) {

tafdVars := ReadTAFDVars(report_file)

;DebugShowMap("TAFD vars après MergeTAFDWithWebForm", tafdVars)  ; <<< DEBUG

tafdVars := MergeTAFDWithWebForm(tafdVars, downloadsDir)

;DebugShowMap("TAFD vars après MergeTAFDWithWebForm", tafdVars)  ; <<< DEBUG

missing := CheckTAFDVars(tafdVars)

if (missing.Length > 0) {
    msg := "Les variables suivantes sont vides ou manquantes:`n`n"
    for _, name in missing
        msg .= "- " name "`n"

    msg .= "`nVoulez-vous continuer malgré tout ?`n"
        . "Oui = continuer le transfert`n"
        . "Non = annuler (CancelTransfer)"

    resp := MsgBox(msg, "Transfert TAFD - Variables manquantes", "YesNo Icon?")

    if (resp = "No") {
        CancelTransfer()  ; ta fonction existante
        }
}

invalidVals := CheckTAFDValues(tafdVars)

if (invalidVals.Length > 0) {
    msg := "Les variables suivantes ont une valeur non conforme (hors choix possibles):`n`n"
    for _, name in invalidVals {
        curVal := tafdVars.Has(name) ? tafdVars[name] : "<absente>"
        msg .= "- " name " = " curVal "`n"
    }

    msg .= "`nVoulez-vous continuer malgré tout ?`n"
        . "Oui = continuer le transfert`n"
        . "Non = annuler (CancelTransfer)"

    resp := MsgBox(msg, "Transfert TAFD - Valeurs non conformes", "YesNo Icon?")

    if (resp = "No") {
        CancelTransfer()
        }
}

ComputeLungRadsVars(tafdVars)

nbPlus := 0
if tafdVars.Has("Nb_Plus") {
    raw := Trim(tafdVars["Nb_Plus"])
    if RegExMatch(raw, "^\d+$") {
        nbPlus := Integer(raw)
    }
}

if (nbPlus < 0)
    nbPlus := 0
else if (nbPlus > 5)
    nbPlus := 5

templateDir  := A_ScriptDir "\Textes\Corps\CT\TAFD"
templatePath := templateDir "\ModeleTAFD" nbPlus ".rtf"

if !FileExist(templatePath) {
    MsgBox "Gabarit introuvable :`n" templatePath, "Erreur gabarit TAFD"
    CancelTransfer()
}

doc.Content.Delete

try {
    r := doc.Range(0, 0)          ; début du document
    r.InsertFile(templatePath)    ; insère ModeleTAFD#.rtf avec toute la mise en forme
} catch as err {
    MsgBox "Erreur lors de l'insertion du modèle TAFD :`n" err.Message
        , "Erreur gabarit", "Icon!"
    CancelTransfer()
}

FillDocVariablesFromMap(doc, tafdVars)
AddText(mode)

try {
    rngAll := doc.Content
    rngAll.Font.Name := "Arial"
} catch as err {
    ; pas bloquant, on continue quand même
}
;fin transfertTAFD
}


ReadTAFDVars(path) {
    static tafdMap := Map(
        1,  "PLCO",
        2,  "Info_TAFD",
        3,  "Date_Exam",
        4,  "Comparaison",
        5,  "Date_Comparaison",
        6,  "Qualite",
        7,  "Serie",
        8,  "DLP",
        9,  "Nb_Total",
        10, "Nb_Plus",
        11, "Nodule1_Image",
        12, "Nodule1_Lobe",
        13, "Nodule1_Localisation",
        14, "Nodule1_Type",
        15, "N1D",
        16, "N1S",
        17, "Nodule1_Comparaison",
        18, "Nodule2_Image",
        19, "Nodule2_Lobe",
        20, "Nodule2_Localisation",
        21, "Nodule2_Type",
        22, "N2D",
        23, "N2S",
        24, "Nodule2_Comparaison",
        25, "Nodule3_Image",
        26, "Nodule3_Lobe",
        27, "Nodule3_Localisation",
        28, "Nodule3_Type",
        29, "N3D",
        30, "N3S",
        31, "Nodule3_Comparaison",
        32, "Nodule4_Image",
        33, "Nodule4_Lobe",
        34, "Nodule4_Localisation",
        35, "Nodule4_Type",
        36, "N4D",
        37, "N4S",
        38, "Nodule4_Comparaison",
        39, "Nodule5_Image",
        40, "Nodule5_Lobe",
        41, "Nodule5_Localisation",
        42, "Nodule5_Type",
        43, "N5D",
        44, "N5S",
        45, "Nodule5_Comparaison",
        46, "Nodules_Commentaire",
        47, "Emphyseme",
        48, "Pneumopathie",
        49, "Epanchement",
        50, "Calcif_Coronaire",
        51, "Autre_Comm",
        52, "LungRADS",
        53, "Decouverte",
        54, "Decouverte_Liste",
        55, "Decouverte_recom",
        56, "Double_Lecture"
    )

    vars := Map()

    try {
        text := FileRead(path)
    } catch as err {
        MsgBox "Erreur lors de la lecture du fichier TAFD :`n" path "`n`n" err.Message
        return vars
    }

    ; On ne découpe plus par lignes : on balaie tout le texte,
    ; en autorisant les retours à la ligne dans 'val'
    pos := 1
    pattern := "(?s)\b(?<num>\d{2})(?:\\[a-zA-Z]+\d*)*\\\{(?<val>.*?)\\\}"

    while RegExMatch(text, pattern, &m, pos) {
        pos   := m.Pos(0) + m.Len(0)

        num   := Integer(m.num)
        value := m.val

        ; Trim des espaces + retours de ligne aux extrémités
        value := Trim(value, " `t`r`n")

        ; [] -> vide
        if (value = "[]") {
            value := ""
        } else {
            ; 1) décodage RTF (unicode + hex) en premier
            value := DecodeRTFUnicode(value)
            value := DecodeRTFHex(value)

	    ; 2) nettoyage des mots de contrôle simples (\par, \~)
            value := CleanSimpleRTF(value)

            ; 3) ensuite, conversion des mots ("zéro"/"zero"/"un"/...) en chiffres
            value := ConvertWordDigit(value)
        }

        if tafdMap.Has(num) {
            name  := tafdMap[num]
            value := TranslateTAFDValue(name, value)
            vars[name] := value
        }
    }

    return vars
}

CleanSimpleRTF(s) {
    ; --- 1) \par + espaces/retours → saut de ligne "soft" ---
    s := RegExReplace(s, "\\par\s*", Chr(11))

    ; --- 2) \~ → espace normale ---
    s := StrReplace(s, "\~", " ")

    ; --- 3) Enlever les commandes de format simples (gras, italique, souligné...) ---

    ; attention : il faut écrire "\\b" pour matcher le littéral \b
    ; gras
    s := StrReplace(s, "\\b0", "")
    s := StrReplace(s, "\\b",  "")

    ; italique
    s := StrReplace(s, "\\i0", "")
    s := StrReplace(s, "\\i",  "")

    ; souligné
    s := StrReplace(s, "\\ulnone", "")
    s := StrReplace(s, "\\ul0",    "")
    s := StrReplace(s, "\\ul",     "")

    ; couleur (\cf1, \cf2, etc.)
    s := RegExReplace(s, "\\cf\d+", "")

    ; taille police (\fs20, \fs24, etc.)
    s := RegExReplace(s, "\\fs\d+", "")

    ; police (\f0, \f1, ...)
    s := RegExReplace(s, "\\f\d+", "")

    ; alignement, etc. (optionnel, mais ça ne gêne pas de les virer ici)
    s := RegExReplace(s, "\\qc|\\ql|\\qr|\\qj", "")

    return s
}


ConvertWordDigit(val) {
    v := Trim(val)
    if (v = "")
        return v

    low := StrLower(v)

    switch low {
        case "zero", "aucun", "zéro":
            return "0"
        case "un":
            return "1"
        case "deux":
            return "2"
        case "trois":
            return "3"
        case "quatre":
            return "4"
        case "cinq":
            return "5"
        case "six":
            return "6"
        case "sept":
            return "7"
        case "huit":
            return "8"
        case "neuf":
            return "9"
        default:
            return val
    }
}

; --- Vérifie les variables TAFD et retourne une liste de noms manquants ---
CheckTAFDVars(vars) {
    missing := []  ; tableau des noms de variables manquantes

    ; ---- Helper local via fonction séparée ----
    ; (définie plus bas : RequireVar(vars, name, missing))

    ; --- 1) Variables toujours obligatoires ---
    alwaysReq := [
        "PLCO",           ; 1
        "Info_TAFD",      ; 2
        "Date_Exam",      ; 3
        "Comparaison",    ; 4
        "Qualite",        ; 6
        "Serie",          ; 7
        "DLP",            ; 8
        "Nb_Total",       ; 9
        "Nb_Plus",        ; 10
        "Emphyseme",      ; 47
        "Pneumopathie",   ; 48
        "Epanchement",    ; 49
        "Calcif_Coronaire", ; 50
        "LungRADS",       ; 52
        "Decouverte",     ; 53
        "Double_Lecture"  ; 56
    ]

    for name in alwaysReq
        RequireVar(vars, name, missing)

    ; --- 2) Règle: Si 4=oui, 5 doit exister ---
    comp := ""
    if vars.Has("Comparaison")
        comp := StrLower(Trim(vars["Comparaison"]))

    if (comp = "oui")
        RequireVar(vars, "Date_Comparaison", missing)  ; 5

    ; --- 3) Règles dépendantes de Nb_Plus (10) ---
    nbPlus := 0
    if vars.Has("Nb_Plus") && vars["Nb_Plus"] != ""
        nbPlus := Integer(vars["Nb_Plus"])

    ; Si 10=1, 11 à 15 doivent exister
    if (nbPlus >= 1) {
        for name in [
            "Nodule1_Image",        ; 11
            "Nodule1_Lobe",         ; 12
            "Nodule1_Localisation", ; 13
            "Nodule1_Type",         ; 14
            "N1D"                   ; 15
        ]
            RequireVar(vars, name, missing)
    }

    ; Si 10=2, ajouter 18 à 22
    if (nbPlus >= 2) {
        for name in [
            "Nodule2_Image",        ; 18
            "Nodule2_Lobe",         ; 19
            "Nodule2_Localisation", ; 20
            "Nodule2_Type",         ; 21
            "N2D"                   ; 22
        ]
            RequireVar(vars, name, missing)
    }

    ; Si 10=3, ajouter 25 à 29
    if (nbPlus >= 3) {
        for name in [
            "Nodule3_Image",        ; 25
            "Nodule3_Lobe",         ; 26
            "Nodule3_Localisation", ; 27
            "Nodule3_Type",         ; 28
            "N3D"                   ; 29
        ]
            RequireVar(vars, name, missing)
    }

    ; Si 10=4, ajouter 32 à 36
    if (nbPlus >= 4) {
        for name in [
            "Nodule4_Image",        ; 32
            "Nodule4_Lobe",         ; 33
            "Nodule4_Localisation", ; 34
            "Nodule4_Type",         ; 35
            "N4D"                   ; 36
        ]
            RequireVar(vars, name, missing)
    }

    ; Si 10=5 ou plus, ajouter 39 à 43 (en plus des précédents)
    if (nbPlus >= 5) {
        for name in [
            "Nodule5_Image",        ; 39
            "Nodule5_Lobe",         ; 40
            "Nodule5_Localisation", ; 41
            "Nodule5_Type",         ; 42
            "N5D"                   ; 43
        ]
            RequireVar(vars, name, missing)
    }

    ; --- 4) Si 53=oui, 54 et 55 doivent exister ---
    decouv := ""
    if vars.Has("Decouverte")
        decouv := StrLower(Trim(vars["Decouverte"]))

    if (decouv = "oui") {
        RequireVar(vars, "Decouverte_Liste",  missing)  ; 54
        RequireVar(vars, "Decouverte_recom",  missing)  ; 55
    }

    ; --- 5) Règles combinées 4=oui et 10=... -> comparaisons nodulaires ---
    ; 17, 24, 31, 38, 45
    if (comp = "oui") {
        if (nbPlus >= 1)
            RequireVar(vars, "Nodule1_Comparaison", missing)  ; 17
        if (nbPlus >= 2)
            RequireVar(vars, "Nodule2_Comparaison", missing)  ; 24
        if (nbPlus >= 3)
            RequireVar(vars, "Nodule3_Comparaison", missing)  ; 31
        if (nbPlus >= 4)
            RequireVar(vars, "Nodule4_Comparaison", missing)  ; 38
        if (nbPlus >= 5)
            RequireVar(vars, "Nodule5_Comparaison", missing)  ; 45
    }

    return missing
}

; --- Helper: ajoute 'name' à missing si var absente ou vide ---
RequireVar(vars, name, missing) {
    if !vars.Has(name) {
        missing.Push(name)
    } else if (Trim(vars[name]) = "") {
        missing.Push(name)
    }
}

NormalizeCodeChar(v) {
    v := Trim(v)
    if (v = "")
        return ""

    ch := SubStr(v, 1, 1)
    low := StrLower(ch)

    ; Cas spécifiques qu'on veut "corriger"
    if (low = "à")    ; Dragon entend "A" mais écrit "à"
        return "A"

    ; Tu pourrais ajouter d'autres mappings ici si un jour tu observes
    ; d'autres confusions (é -> E, etc).

    return StrUpper(ch)
}

TranslateTAFDValue(name, val) {
    v   := Trim(val)
    if (v = "")
        return ""

    low := StrLower(v)

    ; --- Info_TAFD (2) ---
    if (name = "Info_TAFD") {
        code := NormalizeCodeChar(v)
        switch code {
            case "A": return "TAFD de base"
            case "B": return "TAFD annuelle suivi"
            case "C": return "TAFD intervalle 6 mois"
            case "D": return "TAFD intervalle 3 mois"
            default:  return v
        }
    }

    ; --- Comparaison (4) / Decouverte (53) / Double_Lecture (56) ---
    if (name = "Comparaison" || name = "Decouverte" || name = "Double_Lecture") {
        if (low = "oui" || v = "8")
            return "Oui"
        if (low = "non")
            return "Non"
        return CapFirst(v)
    }


    ; --- Date_Exam (3), Date_Comparaison (5) ---
    if (name = "Date_Exam" || name = "Date_Comparaison") {
        ; on laisse tel quel, juste Trim (déjà fait) et on met initiale en majuscule si texte
        return CapFirst(v)
    }

    ; --- Qualite (6) ---
    if (name = "Qualite") {
        code := NormalizeCodeChar(v)
        switch code {
            case "A": return "Adéquate"
            case "B": return "Sous-optimal"
            case "C": return "Non diagnostique"
            default:  return v
        }
    }

    ; --- NoduleX_Localisation (13,20,27,34,41) ---
    if RegExMatch(name, "^Nodule[1-5]_Localisation$") {
        code := NormalizeCodeChar(v)
        switch code {
            case "A": return "Parenchymateux"
            case "B": return "Sous-pleural"
            case "C": return "Scissural"
            default:  return v
        }
    }

    ; --- NoduleX_Type (14,21,28,35,42) ---
    if RegExMatch(name, "^Nodule[1-5]_Type$") {
        code := NormalizeCodeChar(v)
        switch code {
            case "A": return "Solide"
            case "B": return "Verre dépoli"
            case "C": return "Calcifié"
            case "D": return "Mixte"
            default:  return v
        }
    }

    ; --- NoduleX_Comparaison (17,24,31,38,45) ---
    if RegExMatch(name, "^Nodule[1-5]_Comparaison$") {
        code := NormalizeCodeChar(v)
        switch code {
            case "A": return "Stable"
            case "B": return "Augmenté"
            case "C": return "Diminué"
            case "D": return "Nouveau"
            default:  return v
        }
    }

    ; --- Lobes nodulaires (NoduleX_Lobe) -> idem, 1re lettre en majuscule ---
    if RegExMatch(name, "^Nodule[1-5]_Lobe$") {
        return CapFirst(v)
    }

    ; --- N1D/N1S, N2D/N2S, ... -> ajouter " mm" ---
    if RegExMatch(name, "^[N][1-5][DS]$") {
        ; si vide, on ne met rien
        if (v = "")
            return ""

        ; si "mm" est déjà présent (ex: "6,7 mm"), on ne double pas
        if RegExMatch(v, "i)\bmm\b")
        return v

        ; sinon on ajoute simplement " mm"
        return v " mm"
    }


    ; --- Emphyseme (47), Pneumopathie (48), Calcif_Coronaire (50) ---
    if (name = "Emphyseme" || name = "Pneumopathie" || name = "Calcif_Coronaire") {
        code := NormalizeCodeChar(v)
        switch code {
            case "A": return "Aucun"
            case "B": return "Léger"
            case "C": return "Modéré"
            case "D": return "Sévère"
            default:  return v
        }
    }

    ; --- Epanchement (49) ---
    if (name = "Epanchement") {
        ; A = Aucun
        if (low = "a")
            return "Aucun"

        ; B/C/D/E + localisation (droit/gauche/bilatérale)
        ; ex : "B droit", "C gauche", "D bilatérale", etc.
        if RegExMatch(low, "^\s*([a-e])\s+(\w+)", &m) {
            sevCode := m[1]
            side    := m[2]

            sevText := ""
            switch sevCode {
                case "b": sevText := "Minime"
                case "c": sevText := "Léger"
                case "d": sevText := "Modéré"
                case "e": sevText := "Sévère"
                default:  sevText := ""
            }

            if (sevText = "")
                return v

            ; côté
            if InStr(side, "droit")
                sideText := "Droit"
            else if InStr(side, "gauch")
                sideText := "Gauche"
            else if InStr(side, "bilat")
                sideText := "Bilatéral"
            else
                sideText := CapFirst(side)

            return sevText " " sideText
        }

        return v
    }

    ; --- Nodules_Commentaire (46), Autre_Comm (51), Decouverte_Liste (54), Decouverte_recom (55) ---
    if (name = "Nodules_Commentaire"
        || name = "Autre_Comm"
        || name = "Decouverte_Liste"
        || name = "Decouverte_recom") {
        return CapFirst(v)
    }

    ; --- PLCO, Date_Exam, Nb_Total, Nb_Plus, Serie, DLP, etc. (idem) ---
    ; Règle générale "idem" :
    ; - si c'est surtout des chiffres -> on laisse tel quel
    ; - si ça contient des lettres -> on met juste la 1re lettre en majuscule
    if RegExMatch(v, "[A-Za-zÀ-ÖØ-öø-ÿ]") {
        return CapFirst(v)
    }

    return v
}

CapFirst(s) {
    s := Trim(s)
    if (s = "")
        return s
    return StrUpper(SubStr(s, 1, 1)) . SubStr(s, 2)
}

CheckTAFDValues(vars) {
    invalid := []  ; tableau des noms de variables dont la valeur est non conforme

    ; --- Helper interne pour tester "valeur non vide mais pas dans liste" ---
    IsInList(val, allowed) {
        for _, a in allowed
            if (val = a)
                return true
        return false
    }

    ; Récupère une valeur nettoyée (ou "")
    GetVal(name) {
        if !vars.Has(name)
            return ""
        return Trim(vars[name])
    }

    ; --- Info_TAFD ---
    v := GetVal("Info_TAFD")
    if (v != "" && !IsInList(v
        , ["TAFD de base", "TAFD annuelle suivi", "TAFD intervalle 6 mois", "TAFD intervalle 3 mois"])) {
        invalid.Push("Info_TAFD")
    }

    ; --- Comparaison / Decouverte / Double_Lecture ---
    for _, name in ["Comparaison", "Decouverte", "Double_Lecture"] {
        v := GetVal(name)
        if (v = "")
            continue
        if (v != "Oui" && v != "Non")
            invalid.Push(name)
    }

    ; --- Qualite ---
    v := GetVal("Qualite")
    if (v != "" && !IsInList(v, ["Adéquate", "Sous-optimal", "Non diagnostique"])) {
        invalid.Push("Qualite")
    }

    ; --- Emphyseme / Pneumopathie / Calcif_Coronaire ---
    for _, name in ["Emphyseme", "Pneumopathie", "Calcif_Coronaire"] {
        v := GetVal(name)
        if (v = "")
            continue
        if !IsInList(v, ["Aucun", "Léger", "Modéré", "Sévère"])
            invalid.Push(name)
    }

    ; --- NoduleX_Localisation ---
    for n in [1,2,3,4,5] {
        name := "Nodule" n "_Localisation"
        v := GetVal(name)
        if (v = "")
            continue
        if !IsInList(v, ["Parenchymateux", "Sous-pleural", "Scissural"])
            invalid.Push(name)
    }

    ; --- NoduleX_Type ---
    for n in [1,2,3,4,5] {
        name := "Nodule" n "_Type"
        v := GetVal(name)
        if (v = "")
            continue
        if !IsInList(v, ["Solide", "Verre dépoli", "Calcifié", "Mixte", "Kyste pulmonaire atypique"])
            invalid.Push(name)
    }

    ; --- NoduleX_Comparaison ---
    for n in [1,2,3,4,5] {
        name := "Nodule" n "_Comparaison"
        v := GetVal(name)
        if (v = "")
            continue
        if !IsInList(v, ["Stable", "Augmenté", "Diminué", "Nouveau"])
            invalid.Push(name)
    }

    ; --- Epanchement ---
    v := GetVal("Epanchement")
    if (v != "") {
        if (v = "Aucun") {
            ; OK
        } else {
            ; doit être "Minime/Léger/Modéré/Sévère" + "Droit/Gauche/Bilatéral"
            if !RegExMatch(v, "^(Minime|Léger|Modéré|Sévère)\s+(droit|gauche|bilatéral)$") {
                invalid.Push("Epanchement")
            }
        }
    }

       ; --- LungRADS (52) ---
        v := GetVal("LungRADS")
        if (v != "") {
            v := StrUpper(v)
            ok := false

            ; 00 à 06
            if RegExMatch(v, "^0[0-6]$")
               ok := true
            ; 1, 2, 3
            else if (v = "1" || v = "2" || v = "3")
                ok := true
            ; A0, A3, B, X
            else if (v = "A0" || v = "A3" || v = "B" || v = "X")
                ok := true

            if !ok
                invalid.Push("LungRADS")
        }

    ; Les zones "idem" (PLCO, Nb_Total, Nb_Plus, etc.)
    ; ne sont PAS testées ici : elles passent telles quelles.

    return invalid
}


ComputeLungRadsVars(tafdVars) {
    global LUNGRADS_MARK
    mark := LUNGRADS_MARK  ; ce qu'on met dans les cases cochées (ex: "X")

    ; initialiser toutes les variables liées au tableau LungRADS
    for name in [
        "LR0", "LR1", "LR2", "LR3",
        "LR4A", "LR4B", "LR4X",
        "L0a", "L0b", "L0c",
        "L4a", "L4b"
    ] {
        tafdVars[name] := ""  ; toutes les cases vides par défaut
    }

    if !tafdVars.Has("LungRADS")
        return

    v := Trim(tafdVars["LungRADS"])
    if (v = "")
        return

    v := StrUpper(v)

    ; 00 = LR0, L0a
    if (v = "00") {
        tafdVars["LR0"] := mark
        tafdVars["L0a"] := mark
        return
    }

    ; 01 à 06 = LR0, L0b, L0c = chiffre
    if RegExMatch(v, "^0([1-6])$", &m) {
        tafdVars["LR0"] := mark
        tafdVars["L0b"] := mark
        tafdVars["L0c"] := m[1]    ; "1" à "6"
        return
    }

    ; 1 = LR1
    if (v = "1") {
        tafdVars["LR1"] := mark
        return
    }

    ; 2 = LR2
    if (v = "2") {
        tafdVars["LR2"] := mark
        return
    }

    ; 3 = LR3
    if (v = "3") {
        tafdVars["LR3"] := mark
        return
    }

    ; A0 = LR4A, L4a
    if (v = "A0") {
        tafdVars["LR4A"] := mark
        tafdVars["L4a"]  := mark
        return
    }

    ; A3 = LR4A, L4b
    if (v = "A3") {
        tafdVars["LR4A"] := mark
        tafdVars["L4b"]  := mark
        return
    }

    ; B = LR4B
    if (v = "B") {
        tafdVars["LR4B"] := mark
        return
    }

    ; X = LR4X
    if (v = "X") {
        tafdVars["LR4X"] := mark
        return
    }

    ; Normalement, on ne devrait jamais se rendre ici,
    ; car CheckTAFDValues a déjà filtré les valeurs non conformes.
}


FillDocVariablesFromMap(doc, varsMap) {
    vars := doc.Variables

    for name, value in varsMap {
        ; Word aime mal les DocVariables = "" → on met un espace à la place
        safe := (value = "") ? " " : value

        try {
            vars.Item(name).Value := safe
        } catch as err {
            vars.Add(name, safe)
        }
    }

    doc.Fields.Update()
    try doc.ActiveWindow.View.ShowFieldCodes := False
}


DebugDumpTAFDMap(tafdVars, filePath := "") {
    if (filePath = "") {
        filePath := A_ScriptDir "\DebugTAFDVars_" FormatTime(A_Now, "yyyyMMdd_HHmmss") ".txt"
    }

    ;FileDelete filePath

    out := "Debug TAFD vars - " FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") "`r`n"
        . "Total variables dans la map : " tafdVars.Count "`r`n`r`n"

    for name, value in tafdVars {
        valText := value
        valText := StrReplace(valText, "`r`n", "\r\n")
        valText := StrReplace(valText, "`n", "\n")
        out .= name " = [" valText "]`r`n"
    }

    FileAppend(out, filePath, "UTF-8")
    return filePath
}

MergeTAFDWithWebForm(tafdVars, downloadsDir) {
    global reqnb_full

    ; Si aucune reqnb_full → rien à faire
    if (reqnb_full = "" || reqnb_full = 0) {
        return tafdVars
    }

    ; Ex.: TAFD_123456.txt, TAFD_123456 (1).txt...
    pattern := "TAFD_" reqnb_full "*.txt"
    files   := []  ; liste des chemins trouvés

    ; --- Chercher dans Téléchargements ---
    try {
        if DirExist(downloadsDir) {
            Loop Files, downloadsDir "\" pattern, "F"  ; "F" = fichiers seulement
                files.Push(A_LoopFileFullPath)
        }
    } catch as err {
        ; problème d'accès → on laisse tomber la fusion
        return tafdVars
    }

    ; Aucun fichier trouvé → proposer d'annuler ou de continuer
    if (files.Length = 0) {
        msg := "Données WebForm non trouvées pour cette requête." . "`n`n"
            . "Requête : " reqnb_full "`n"
            . "Fichiers recherchés : " pattern "`n`n"
            . "Voulez-vous continuer sans les données du WebForm ?" . "`n`n"
            . "Oui = continuer le transfert (RadEdit seulement)" . "`n"
            . "Non = annuler (CancelTransfer)"

        resp := MsgBox(msg, "Transfert TAFD - Données WebForm manquantes", "YesNo Icon?")

        if (resp = "No") {
            ; On annule proprement puis on quitte le script
            CancelTransfer()
            ExitApp
        }

        ; L'utilisateur accepte de continuer sans WebForm
        return tafdVars
    }

    ; --- Choisir le fichier le plus récent ---
    bestFile := ""
    bestTime := ""
    for _, path in files {
        t := FileGetTime(path, "M")  ; date/heure modification
        if (t = "") {
            continue
        }
        if (bestFile = "" || t > bestTime) {
            bestFile := path
            bestTime := t
        }
    }

    if (bestFile = "") {
        return tafdVars
    }

    ; --- Lire le fichier TAFD_xxx.txt dans fileVars ---
    fileVars := Map()
    try {
        txt := FileRead(bestFile, "UTF-8")
    } catch as err {
        ; problème de lecture → on garde simplement tafdVars
        return tafdVars
    }

    ; Découpage ligne par ligne
    for line in StrSplit(txt, "`n", "`r") {
        line := Trim(line, " `t`r`n")
        if (line = "" || SubStr(line, 1, 1) = ";")
            continue

        pos := InStr(line, "=")
        if (!pos)
            continue

        name := Trim(SubStr(line, 1, pos - 1), " `t")
        val  := Trim(SubStr(line, pos + 1), " `t")

        ; enlever les crochets [..] si présents
        if (SubStr(val, 1, 1) = "[" && SubStr(val, -1) = "]") {
            val := SubStr(val, 2, StrLen(val) - 2)
        }

        if (name != "") {
            fileVars[name] := val
        }
    }

    ; --- Fusion simple : les valeurs du fichier écrasent celles de tafdVars ---
    for name, fileVal in fileVars {
        tafdVars[name] := fileVal
    }

    ; --- Nettoyage : effacer tous les fichiers TAFD_reqnb_full*.txt ---
    for _, path in files {
        try FileDelete(path)
        catch as err {
            ; non bloquant : si suppression impossible, on ignore
        }
    }

    return tafdVars
}