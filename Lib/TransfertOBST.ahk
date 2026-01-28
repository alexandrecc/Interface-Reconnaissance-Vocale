TransfertOBST(word, doc, report_file, examType, mode) {

    ; report_file n'est plus utilisé pour OBST (variables = fichier OBST_Data)
    rawDir       := "R:\CSSSNL\Bureautique\Imagerie Medicale\Partage\Technologues\Obstétrique\RAW DATA"
    downloadsDir := EnvGet("USERPROFILE") "\Downloads"

    usedFile := ""
    obstVars := LoadOBSTVarsFromDataFile(rawDir, downloadsDir, reqnb_full, &usedFile)

    ; Si aucun fichier data trouvé : proposer continuer / annuler
    if (usedFile = "") {
        msg := "Fichier OBST_Data introuvable pour reqnb_full=" reqnb_full "`n`n"
            . "Cherché dans:`n- " rawDir "`n- " downloadsDir "`n`n"
            . "Voulez-vous continuer quand même (sans variables) ?"
        resp := MsgBox(msg, "Transfert OBST - Fichier data manquant", "YesNo Icon?")
        if (resp = "No") {
            CancelTransfer()
            return
        }
        ; continuer avec une map vide (les DocVariables seront quand même créées si besoin)
        obstVars := Map()
    }

    ; 2) Dump debug dans un fichier texte (map après fusion)
    ;debugPath := DebugDumpOBSTMap(obstVars)

    ; 3) Info à l’écran
    ;MsgBox "Debug OBST créé :`n" debugPath



    ; 2) Vérifier variables manquantes
    missing := CheckOBSTVars(obstVars, examType)

    if (missing.Length > 0) {
        msg := "Les variables suivantes sont vides ou manquantes:`n`n"
        for _, name in missing
            msg .= "- " name "`n"

        msg .= "`nVoulez-vous continuer malgré tout ?`n"
            . "Oui = continuer le transfert`n"
            . "Non = annuler (CancelTransfer)"

        resp := MsgBox(msg, "Transfert Obstétrique - Variables manquantes", "YesNo Icon?")

        if (resp = "No") {
            CancelTransfer()
            return
        }
    }

    ; 3) Vérifier valeurs non conformes
    invalidVals := CheckOBSTValues(obstVars)

    if (invalidVals.Length > 0) {
        msg := "Les variables suivantes ont une valeur non conforme (hors choix possibles):`n`n"
        for _, name in invalidVals {
            curVal := obstVars.Has(name) ? obstVars[name] : "<absente>"
            msg .= "- " name " = " curVal "`n"
        }

        msg .= "`nVoulez-vous continuer malgré tout ?`n"
            . "Oui = continuer le transfert`n"
            . "Non = annuler (CancelTransfer)"

        resp := MsgBox(msg, "Transfert Obstétrique - Valeurs non conformes", "YesNo Icon?")

        if (resp = "No") {
            CancelTransfer()
            return
        }
    }

    ; 4) Choisir le bon modèle selon examType
    templateDir := A_ScriptDir "\Textes\Corps\US\OBSTETRIQUE"
    fileName := ""

    switch examType {
        case "OBST_Morpho":
            fileName := "ModeleOBST_Morpho.rtf"
        case "OBST_Crois":
            fileName := "ModeleOBST_Crois.rtf"
        case "OBST_Doppler":
            fileName := "ModeleOBST_Doppler.rtf"
        case "OBST_Crois_Doppler":
            fileName := "ModeleOBST_Crois_Doppler.rtf"
        default:
            MsgBox "Type d'examen obstétrical inconnu : " examType
                , "Erreur TransfertOBST", "Icon!"
            CancelTransfer()
            return
    }

    templatePath := templateDir "\" fileName

    if !FileExist(templatePath) {
        MsgBox "Gabarit OBST introuvable :`n" templatePath
            , "Erreur gabarit OBST"
        CancelTransfer()
        return
    }

    ; 5) Insérer le modèle dans Word (comme pour TAFD)
    doc.Content.Delete

    try {
        r := doc.Range(0, 0)
        r.InsertFile(templatePath)
    } catch as err {
        MsgBox "Erreur lors de l'insertion du modèle OBST :`n" err.Message
            , "Erreur gabarit OBST", "Icon!"
        CancelTransfer()
        return
    }

    ; 6) Remplir les DocVariables à partir de la map
    FillDocVariablesFromMap(doc, obstVars)
    AddText(mode)

    ; 7) Optionnel : tout en Arial (comme pour TAFD)
    try {
        rngAll := doc.Content
        rngAll.Font.Name := "Arial"
    } catch as err {
        ; pas bloquant
    }
}




FormatWeeksDays(v) {
    v := Trim(v)
    if (v = "")
        return ""

    ; x,y → x sem y jrs
    if RegExMatch(v, "^\s*(\d+)\s*[,\.]\s*(\d+)\s*$", &m) {
        sa := m[1]
        j  := m[2]
        return sa " sem " j " jrs"
    }

    ; x y → x sem y jrs (fallback)
    if RegExMatch(v, "^\s*(\d+)\s+(\d+)\s*$", &m2) {
        sa := m2[1]
        j  := m2[2]
        return sa " sem " j " jrs"
    }

    ; x → x sem 0 jrs (entier seul)
    if RegExMatch(v, "^\d+$") {
        return v " sem 0 jrs"
    }

    ; Sinon on laisse tel quel
    return v
}


FormatPercentSimple(v) {
    v := Trim(v)
    if (v = "")
        return ""
    if RegExMatch(v, "i)perc")
        return v
    return v "e perc."
}

FormatPercentWithAgeRef(v, ageRefText) {
    v := Trim(v)
    if (v = "")
        return ""
    if RegExMatch(v, "i)perc")
        return v   ; déjà formaté

    if (ageRefText != "")
        return v "e perc. de " ageRefText
    else
        return v "e perc."
}

TranslateOBSTValue(name, val) {
    v   := Trim(val)
    if (v = "")
        return ""

    low := StrLower(v)

    ; --- Booléens Oui/Non ---
    if (name = "Mvt_Foet" || name = "Coeur_Foet" || name = "Doppler_Cordon") {
        if (low = "oui" || v = "8")
            return "Oui"
        if (low = "non")
            return "Non"
        return CapFirst(v)
    }

    ; --- Mesures biométriques : + " mm" ---
    if RegExMatch(name, "^Mesure_(BPD|CC|AC|FL|Hum)$") {
        if RegExMatch(v, "i)\bmm\b")
            return v
        return v " mm"
    }

    ; --- NbSem_xx : x,y → x sem y jrs ---
    if RegExMatch(name, "^NbSem_(BPD|CC|AC|FL|Hum)$") {
        return FormatWeeksDays(v)
    }

    ; --- Percentiles de mesure : "z" → "ze perc." ---
    if RegExMatch(name, "^Perc_(BPD|CC|AC|FL|Hum)$") {
        return FormatPercentSimple(v)
    }

    ; --- Liquide_Amnio : A/B/C ou "normal(e)" → texte ---
    if (name = "Liquide_Amnio") {
        if (v = "")
            return ""  ; PPP complétera éventuellement

        ; Texte "normal" / "normale" insensible à la casse
        if RegExMatch(low, "^\s*normal(e)?\s*$")
            return "Normal"

        code := NormalizeCodeChar(v)
        switch code {
            case "A": return "Normal"
            case "B": return "Oligohydramnios"
            case "C": return "Polyhydramnios"
            default:  return CapFirst(v)
        }
    }

    ; --- Amnio_PPP : + " cm" ---
    if (name = "Amnio_PPP") {
        if RegExMatch(v, "i)\bcm\b")
            return v
        return v " cm"
    }

    ; --- Placenta : A/B/C/D ou "normal(e)" ---
    if (name = "Placenta") {
        ; Texte "normal" / "normale" insensible à la casse
        if RegExMatch(low, "^\s*normal(e)?\s*$")
            return "Normal"

        code := NormalizeCodeChar(v)
        switch code {
            case "A": return "Normal"
            case "B": return "Bas inséré"
            case "C": return "Marginal"
            case "D": return "Praevia"
            default:  return CapFirst(v)
        }
    }

    ; Placenta_Loc / Placenta_Grade / Presentation / Renseignements / Examen / Remarques :
    ; → texte "idem" : 1re lettre majuscule
    if (name = "Placenta_Loc"
        || name = "Placenta_Grade"
        || name = "Presentation"
        || name = "Renseignements"
        || name = "Examen"
        || name = "Remarques") {
        return CapFirst(v)
    }

    ; --- GA / Age_Ref / Age_Echo : "x,y" → "x sem y jrs" ---
    if (name = "Age_Ref" || name = "Age_Echo") {
        return FormatWeeksDays(v)
    }

    ; --- Poids : + " g" ---
    if (name = "Poids") {
        if RegExMatch(v, "i)\bg\b")
            return v
        return v " g"
    }

    ; --- Anatomie : Tout_Morpho / Crois_Std / items 17-28 ---
    if (name = "Tout_Morpho" || name = "Crois_Std") {
        return StrUpper(v)   ; "OK" ou vide
    }
    if RegExMatch(name, "^Anat_") {
        ; NormalizeCodeChar gère le cas "à" → "A", etc.
        code := NormalizeCodeChar(v)
        return StrUpper(code)   ; A / B / C
    }

    ; --- Doppler PI / Result ---
    if (name = "UA_PI" || name = "MCA_PI") {
        return v      ; on laisse brut, c’est un chiffre
    }

; --- Résultats Doppler : normal/anormal ---
if (name = "UA_Result" || name = "MCA_Result" || name = "Ratio_CP_Result") {
    ; normal / normale
    if RegExMatch(low, "^\s*normal(e)?\s*$")
        return "Normal"

    ; anormal / anormale
    if RegExMatch(low, "^\s*anormal(e)?\s*$")
        return "Anormal"

    ; Dragon : "a normal" / "a normale" → Anormal
    if RegExMatch(low, "^\s*a\s*normal(e)?\s*$")
        return "Anormal"

    return CapFirst(v)
}


    ; --- Ratio_CP_Perc, UA_Perc, MCA_Perc, Perc (45) ---
    ; Pour l'instant on laisse brut ici, on ajoutera "de [Age_Ref]" en 2e passe
    if (name = "UA_Perc" || name = "MCA_Perc" || name = "Ratio_CP_Perc" || name = "Perc") {
        return Trim(v)
    }

    ; --- Champs simples G/P/A/Foetus_Nb/DDM/DPA etc. ("idem" générique) ---
    if RegExMatch(v, "[A-Za-zÀ-ÖØ-öø-ÿ]") {
        return CapFirst(v)
    }

    ; Sinon (valeurs purement numériques) → on les laisse telles quelles
    return v
}

TranslateOBSTCrossFields(obstVars) {
    ; --- Age de référence déjà traduit ---
    ageRef := ""
    if obstVars.Has("Age_Ref")
        ageRef := Trim(obstVars["Age_Ref"])

    ; --- Percentiles dépendant d'Age_Ref ---
    for name in ["UA_Perc", "MCA_Perc", "Ratio_CP_Perc", "Perc"] {
        if obstVars.Has(name) {
            obstVars[name] := FormatPercentWithAgeRef(obstVars[name], ageRef)
        }
    }

    ; --- Liquide_Amnio via Amnio_PPP si Liquide_Amnio vide ---
    liq := ""
    if obstVars.Has("Liquide_Amnio")
        liq := Trim(obstVars["Liquide_Amnio"])

    if (liq = "" && obstVars.Has("Amnio_PPP")) {
        rawPPP := obstVars["Amnio_PPP"]
        ; Extraire le nombre (ex: "4,5 cm" → "4,5")
        numText := RegExReplace(rawPPP, "[^0-9,\.]", "")
        if (numText != "") {
            valStr := StrReplace(numText, ",", ".")
            ppp := valStr + 0.0
            if (ppp < 2) {
                obstVars["Liquide_Amnio"] := "Oligohydramnios"
            } else if (ppp > 8) {
                obstVars["Liquide_Amnio"] := "Polyhydramnios"
            } else { ; 2 à 8 inclus
                obstVars["Liquide_Amnio"] := "Normal"
            }
        }
    }

    ; --- Doppler (92) dépend de Doppler_Cordon (94) ---
    dopplerFlag := ""
    if obstVars.Has("Doppler_Cordon")
        dopplerFlag := StrLower(Trim(obstVars["Doppler_Cordon"]))

    if (dopplerFlag = "oui") {
        obstVars["Doppler"] := "DOPPLER EVALUATION DE RETARD DE CROISSANCE"
    } else {
        obstVars["Doppler"] := ""
    }

    ; --- Anatomie : remplir 17N/17V/17NV... 28N/28V/28NV ---
    if !OBST_HasAnyAnatomieMark(obstVars)
        ComputeOBSTAnatomieVars(obstVars)
}

OBST_HasAnyAnatomieMark(obstVars) {
    ; Si le fichier data a déjà mis des X (ou autre) dans 17..28 N/V/NV,
    ; on considère ces valeurs comme "source de vérité" et on NE RECALCULE PAS.
    loop 12 {
        idx  := 16 + A_Index  ; 17..28
        base := String(idx)
        for _, suf in ["N", "V", "NV"] {
            k := base suf
            if obstVars.Has(k) && Trim(obstVars[k]) != ""
                return true
        }
    }
    return false
}

ComputeOBSTAnatomieVars(obstVars) {
    ; indices anatomiques 17..28
    static anatomieNames := Map(
        17, "Anat_BPD",
        18, "Anat_FosPos",
        19, "Anat_Atrium",
        20, "Anat_TMO",
        21, "Anat_Rachis",
        22, "Anat_Thorax_4ch",
        23, "Anat_Thorax_Crois",
        24, "Anat_Abdo_Estomac",
        25, "Anat_ParoiAbdo",
        26, "Anat_Reins",
        27, "Anat_Vessie",
        28, "Anat_4Membres"
    )

    ; 1) Tout vider : 17N/17V/17NV ... 28N/28V/28NV
    loop 12 {
        idx := 16 + A_Index  ; 17..28
        base := String(idx)
        obstVars[base "N"]  := ""
        obstVars[base "V"]  := ""
        obstVars[base "NV"] := ""
    }

    ; 2) Tout_Morpho (15) : si OK → X sur 17-28 N et V
    tout := obstVars.Has("Tout_Morpho") ? StrUpper(Trim(obstVars["Tout_Morpho"])) : ""
    if (tout = "OK") {
        loop 12 {
            idx := 16 + A_Index
            base := String(idx)
            obstVars[base "N"] := "X"
            obstVars[base "V"] := "X"
        }
    }

    ; 3) Crois_Std (16) : si OK → X sur 17,18,19,22,23,24,26,27 N et V
    crois := obstVars.Has("Crois_Std") ? StrUpper(Trim(obstVars["Crois_Std"])) : ""
    if (crois = "OK") {
        for idx in [17,18,19,22,23,24,26,27] {
            base := String(idx)
            obstVars[base "N"] := "X"
            obstVars[base "V"] := "X"
        }
    }

    ; 4) Ajustement fin par variables 17..28 (A / B / C)
    for idx, varName in anatomieNames {
        if !obstVars.Has(varName)
            continue

        code := StrUpper(Trim(obstVars[varName]))
        base := String(idx)

        if (code = "")
            continue

        ; on écrase le réglage précédent pour cet item précis
        obstVars[base "N"]  := ""
        obstVars[base "V"]  := ""
        obstVars[base "NV"] := ""

        ; A = "normal sur tout" → X sur N et V
        if (code = "A") {
            obstVars[base "N"] := "X"
            obstVars[base "V"] := "X"
        }
        ; B = "visualisé seulement" → X sur V
        else if (code = "B") {
            obstVars[base "V"] := "X"
        }
        ; C = "non visualisé / non vu" → X sur NV
        else if (code = "C") {
            obstVars[base "NV"] := "X"
        }
        ; autres codes → ignorés
    }
}


TranslateOBSTVars(obstVars) {
    ; 1) Traduction champ par champ
    for name, value in obstVars {
        obstVars[name] := TranslateOBSTValue(name, value)
    }

    ; 2) Ajustements qui dépendent d'autres variables
    TranslateOBSTCrossFields(obstVars)
}

CheckOBSTVars(vars, examType := "") {
    missing := []  ; tableau des noms de variables manquantes

    ; Helper : existe et non vide
    HasNonEmpty(name) {
        return vars.Has(name) && Trim(vars[name]) != ""
    }

    ; --- CAS SPÉCIAL : OBST_Doppler ---
    ; Pour cet examType, seules ces variables doivent être obligatoires :
    ; 4,6,34-41,82,42
    ; => Foetus_Nb, Coeur_Foet, UA_*, MCA_*, Ratio_CP_*, RCP, Age_Ref
    if (examType = "OBST_Doppler") {
        required := [
            "Foetus_Nb",      ; 4
            "Coeur_Foet",     ; 6
            "Age_Ref",        ; 42
            "UA_PI",          ; 34
            "UA_Perc",        ; 35
            "UA_Result",      ; 36
            "MCA_PI",         ; 37
            "MCA_Perc",       ; 38
            "MCA_Result",     ; 39
            "Ratio_CP_Perc",  ; 40
            "Ratio_CP_Result", ; 41
            "RCP"             ; 82
        ]

        for name in required
            RequireVar(vars, name, missing)

        return missing
    }

    ; --- CAS GÉNÉRAL (Standard, OBST_Morpho, OBST_Crois, OBST_Crois_Doppler) ---
    ; (ton ancienne logique continue ici)
    ; --- 1) Variables toujours obligatoires ---
    alwaysReq := [
        "Foetus_Nb",      ; 4
        "Coeur_Foet",     ; 6
        "Presentation",   ; 9

        "Mesure_BPD",     ; 10
        "Mesure_CC",      ; 11
        "Mesure_AC",      ; 12
        "Mesure_FL",      ; 13

        "NbSem_BPD",      ; 50
        "NbSem_CC",       ; 51
        "NbSem_AC",       ; 52
        "NbSem_FL",       ; 53

        "Perc_BPD",       ; 60
        "Perc_CC",        ; 61
        "Perc_AC",        ; 62
        "Perc_FL",        ; 63

        "Liquide_Amnio",  ; 29
        "Placenta",       ; 31
        "Placenta_Loc",   ; 32
        "Placenta_Grade", ; 33

        "Age_Ref",        ; 42
        "Age_Echo"        ; 43
    ]

    for name in alwaysReq
        RequireVar(vars, name, missing)

    ; --- 2) Règles conditionnelles basées sur "Examen" (91) ---
    exam := ""
    if vars.Has("Examen")
        exam := Trim(vars["Examen"])

    ; 2a) MORPHO : 91 = "ECHO OBSTETRICALE 16 SEM ET PLUS(MORPHO)"
    if (exam = "ECHO OBSTETRICALE 16 SEM ET PLUS(MORPHO)") {
        ; Règle : pour chaque indice 17 à 28, au moins un des champs
        ; 17N / 17V / 17NV (etc.) doit être non vide.

        missingIdx := []  ; indices 17..28 sans aucune info

        loop 12 {
            idx  := 16 + A_Index   ; 17..28
            base := String(idx)
            hasThis := false

            for suffix in ["N", "V", "NV"] {
                name := base . suffix  ; "17N", "17V", "17NV", etc.
                if HasNonEmpty(name) {
                    hasThis := true
                    break
                }
            }

            if !hasThis
                missingIdx.Push(idx)
        }

        if (missingIdx.Length > 0) {
            ; Message simple, sans liste détaillée des indices
            missing.Push("Anatomie_Morpho (un ou plusieurs segments 17-28 n'ont aucune valeur N/V/NV)")
        }
    }



    ; 2b) CROISSANCE : 
    ; 91 = "ECHO OBSTETRICALE 16 SEM ET PLUS(CROISSANCE)"
    ;   ou "ECHO OBSTETRICALE 16 SEM ET PLUS(CROISSANCE) (T)"
    if (exam = "ECHO OBSTETRICALE 16 SEM ET PLUS(CROISSANCE)"
        || exam = "ECHO OBSTETRICALE 16 SEM ET PLUS(CROISSANCE) (T)") {

        for name in ["Amnio_PPP", "Poids", "Perc"]  ; 30, 44, 45
            RequireVar(vars, name, missing)
    }

    ; --- 3) Doppler (94 = Oui) ---
    dop := ""
    if vars.Has("Doppler_Cordon")
        dop := StrLower(Trim(vars["Doppler_Cordon"]))

    if (dop = "oui") {
        for name in [
            "UA_PI", "UA_Perc", "UA_Result",
            "MCA_PI", "MCA_Perc", "MCA_Result",
            "Ratio_CP_Perc", "Ratio_CP_Result",
            "RCP"   ; 82
        ] {
            RequireVar(vars, name, missing)
        }
    }

    return missing
}


CheckOBSTValues(vars) {
    invalid := []

    ; Helpers
    GetVal(name) {
        if !vars.Has(name)
            return ""
        return Trim(vars[name])
    }

    IsInteger(v) {
        return RegExMatch(v, "^\d+$")
    }

    IsNumber(v) {
        return RegExMatch(v, "^\d+([.,]\d+)?$")
    }

    IsInList(val, allowed) {
        for _, a in allowed
            if (val = a)
                return true
        return false
    }

    ; --- 1) Doit être nombre entier : 1,2,3,4 (G,P,A,Foetus_Nb) ---
    for name in ["G", "P", "A", "Foetus_Nb"] {
        v := GetVal(name)
        if (v = "")
            continue
        if !IsInteger(v)
            invalid.Push(name)
    }

    ; --- 2) Doit être "nombre entier g" : 44 (Poids) ---
    v := GetVal("Poids")
    if (v != "" && !RegExMatch(v, "^\d+\s*g$"))
        invalid.Push("Poids")

    ; --- 3) Doit être Oui ou Non : 5,6 (Mvt_Foet, Coeur_Foet) ---
    for name in ["Mvt_Foet", "Coeur_Foet"] {
        v := GetVal(name)
        if (v = "")
            continue
        if (v != "Oui" && v != "Non")
            invalid.Push(name)
    }

    ; --- 4) Doit être Normal ou Anormal : 36,39,41 ---
    for name in ["UA_Result", "MCA_Result", "Ratio_CP_Result"] {
        v := GetVal(name)
        if (v = "")
            continue
        if (v != "Normal" && v != "Anormal")
            invalid.Push(name)
    }

    ; --- 5) Format "## sem # jrs" : 50,51,52,53,54,42,43 ---
    for name in ["NbSem_BPD", "NbSem_CC", "NbSem_AC", "NbSem_FL", "NbSem_Hum",
                 "Age_Ref", "Age_Echo"] {
        v := GetVal(name)
        if (v = "")
            continue
        ; Exemple: "30 sem 2 jrs"
        if !RegExMatch(v, "^\d+\s+sem\s+\d+\s+jrs$")
            invalid.Push(name)
    }

    ; --- 6) "nombre entier e perc. de [Age_Ref]" : 45,35,38,40 ---
    for name in ["Perc", "UA_Perc", "MCA_Perc", "Ratio_CP_Perc"] {
        v := GetVal(name)
        if (v = "")
            continue
        ; ex: "12e perc. de 30 sem 2 jrs"
        if !RegExMatch(v, "^\d+e perc\. de .+$")
            invalid.Push(name)
    }

    ; --- 7) "nombre entier e perc." : 60 à 64 ---
    for name in ["Perc_BPD", "Perc_CC", "Perc_AC", "Perc_FL", "Perc_Hum"] {
        v := GetVal(name)
        if (v = "")
            continue
        ; ex: "12e perc."
        if !RegExMatch(v, "^\d+e perc\.$")
            invalid.Push(name)
    }

    ; --- 8) "nombre cm" : 30 (Amnio_PPP) ---
    v := GetVal("Amnio_PPP")
    if (v != "" && !RegExMatch(v, "^\d+([.,]\d+)?\s*cm$"))
        invalid.Push("Amnio_PPP")

    ; --- 9) "nombre mm" : 10 à 14 ---
    for name in ["Mesure_BPD", "Mesure_CC", "Mesure_AC", "Mesure_FL", "Mesure_Hum"] {
        v := GetVal(name)
        if (v = "")
            continue
        if !RegExMatch(v, "^\d+([.,]\d+)?\s*mm$")
            invalid.Push(name)
    }

    ; --- 10) Doit être 0,1,2 ou 3 : 33 (Placenta_Grade) ---
    v := GetVal("Placenta_Grade")
    if (v != "" && !IsInList(v, ["0","1","2","3"]))
        invalid.Push("Placenta_Grade")

    ; --- 11) Placenta : Normal, Bas Inséré, Marginal, Praevia ---
    v := GetVal("Placenta")
    if (v != "" && !IsInList(v, ["Normal", "Bas inséré", "Marginal", "Praevia"]))
        invalid.Push("Placenta")

    ; --- 12) Liquide_Amnio : Normal, Oligohydramnios, Polyhydramnios ---
    v := GetVal("Liquide_Amnio")
    if (v != "" && !IsInList(v, ["Normal", "Oligohydramnios", "Polyhydramnios"]))
        invalid.Push("Liquide_Amnio")

    ; --- 13) Doit être nombre : 34,37,82 (UA_PI, MCA_PI, RCP) ---
    for name in ["UA_PI", "MCA_PI", "RCP"] {
        v := GetVal(name)
        if (v = "")
            continue
        if !IsNumber(v)
            invalid.Push(name)
    }

    return invalid
}

LoadOBSTVarsFromDataFile(rawDir, downloadsDir, reqnb_full, &usedFile := "") {
    usedFile := ""

    if (reqnb_full = "" || reqnb_full = 0)
        return Map()

    pattern := "OBST_Data_" reqnb_full "*.txt"
    files := []

    ; Chercher dans RAW DATA
    try {
        if DirExist(rawDir) {
            Loop Files, rawDir "\" pattern, "F"
                files.Push(A_LoopFileFullPath)
        }
    } catch as err {
        ; accès réseau problématique -> on continue quand même avec Downloads
    }

    ; Chercher dans Downloads
    try {
        if DirExist(downloadsDir) {
            Loop Files, downloadsDir "\" pattern, "F"
                files.Push(A_LoopFileFullPath)
        }
    } catch as err {
        ; rien
    }

    if (files.Length = 0)
        return Map()

    ; Choisir le fichier le plus récent
    bestFile := ""
    bestTime := ""
    for _, path in files {
        t := FileGetTime(path, "M")
        if (t = "")
            continue
        if (bestFile = "" || t > bestTime) {
            bestFile := path
            bestTime := t
        }
    }

    if (bestFile = "")
        return Map()

    ; Canonicaliser dans RAW DATA (copie + suppression original si venait de Downloads)
    try {
        SplitPath(bestFile, &fileName)
        canonical := rawDir "\" fileName
        if (bestFile != canonical) {
            if DirExist(rawDir) {
                FileCopy(bestFile, canonical, true)
                try FileDelete(bestFile)
                bestFile := canonical
            }
        }
    } catch as err {
        ; non bloquant
    }

    ; Supprimer les autres doublons
    for _, path in files {
        if (path = bestFile)
            continue
        try FileDelete(path)
    }

    usedFile := bestFile
    return ReadOBSTVarsFromDataTxt(bestFile)
}

ReadOBSTVarsFromDataTxt(path) {
    vars := Map()

    try {
        txt := FileRead(path, "UTF-8")
    } catch as err {
        MsgBox "Erreur lecture OBST_Data :`n" path "`n`n" err.Message, "Transfert OBST", "Icon!"
        return vars
    }

    for line in StrSplit(txt, "`n", "`r") {
        line := Trim(line, " `t`r`n")
        if (line = "" || SubStr(line, 1, 1) = ";")
            continue

        pos := InStr(line, "=")
        if (!pos)
            continue

        name := Trim(SubStr(line, 1, pos - 1), " `t")
        val  := Trim(SubStr(line, pos + 1), " `t")

        ; Si valeur entre [..] -> enlever crochets
        if (SubStr(val, 1, 1) = "[" && SubStr(val, -1) = "]")
            val := SubStr(val, 2, StrLen(val) - 2)

        ; "[]" ou vide -> ""
        if (val = "[]")
            val := ""

        vars[name] := val
    }

    ; IMPORTANT :
    ; Si tu veux garder toutes tes conversions métier (semaines/jours, mm, perc., etc.)
    ; on réutilise ton TranslateOBSTVars existant.
    TranslateOBSTVars(vars)

    return vars
}



DebugDumpOBSTMap(obstVars, filePath := "") {
    if (filePath = "") {
        filePath := A_ScriptDir "\DebugOBSTVars_" FormatTime(A_Now, "yyyyMMdd_HHmmss") ".txt"
    }

    out := "Debug OBST vars - " FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") "`r`n"
        . "Total variables dans la map : " obstVars.Count "`r`n`r`n"

    for name, value in obstVars {
        valText := value
        valText := StrReplace(valText, "`r`n", "\r\n")
        valText := StrReplace(valText, "`n", "\n")
        out .= name " = [" valText "]`r`n"
    }

    FileAppend(out, filePath, "UTF-8")
    return filePath
}

DebugShowMap(label, m) {
    txt := label "`n`n"
    for name, val in m {
        txt .= name " = [" val "]`n"
    }
    MsgBox txt
}