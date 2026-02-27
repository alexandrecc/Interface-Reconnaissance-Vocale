AddText(mode := "") {
    try word := ComObjActive("Word.Application")
    catch {
        MsgBox "Word n'est pas ouvert.", "Erreur", "T3"
        return
    }
    doc := word.ActiveDocument
    sel := word.Selection

    sel.SetRange(doc.Content.End, doc.Content.End)
    sel.ParagraphFormat.Alignment := 0

    chosenRtf := (mode = "ToutSUG" or mode = "OuvrirSUG") ? finalRtfSUG : finalRtf
    if AddFinalRtf(doc, sel, chosenRtf) {
	sel.SetRange(doc.Content.End, doc.Content.End)
        sel.ParagraphFormat.Alignment := 0
        sel.Font.Name := "Arial"
        sel.Font.Size := 10
        return
    }


    ; --- Textes fixes ---
    textSUG    := GetTextByKey(textFile, "textSUG",	"Examen demandé, réalisé et interprété en urgence sur l'horaire de garde à la demande du médecin prescripteur.")
    authorLine := GetTextByKey(textFile, "authorLine",	"Dictée faite par Dr")
    sent1      := GetTextByKey(textFile, "sent1",	"Ce rapport a été généré à l'aide d'un logiciel de reconnaissance vocale.")
    sent2      := GetTextByKey(textFile, "sent2",	"Pour cette raison il pourrait contenir des erreurs grammaticales, orthographiques ou syntaxiques.")

    ; Bloc en fonction du mode
    if (mode = "ToutSUG" or mode = "OuvrirSUG") {
        fullText := "`r`n`r`n" textSUG "`r`n`r`n" authorLine " " GetFullName_API() "`r`n`r`n" sent1 " " sent2
    } else {
        fullText := "`r`n`r`n" authorLine " " GetFullName_API() "`r`n`r`n" sent1 " " sent2
    }
    ; --- Insertion ---
    sel.EndKey(6)          ; aller à la fin du doc (wdStory = 6)
    startPos := sel.Start  ; mémoriser début de l’insertion
    sel.TypeText(fullText) ; insérer tout le bloc

    ; --- Mise en forme ---
    endPos := sel.Start
    blockRng := doc.Range(startPos, endPos)

    ; Globale : Arial 10 italique
    blockRng.Font.Name   := "Arial"
    blockRng.Font.Size   := 10
    blockRng.Font.Italic := True

    ; Spécifique : première phrase en gras + souligné
    fr := blockRng.Duplicate
    fr.Find.ClearFormatting
    fr.Find.Text := sent1
    if fr.Find.Execute() {
        fr.Font.Bold := True
        ;fr.Font.Underline := 1
    }

    ; --- Fin ---
    sel.SetRange(doc.Content.End, doc.Content.End)
    sel.Font.Name := "Arial"
    sel.Font.Size := 10

}


AddFinalRtf(doc, sel, finalRtfPath := "") {
    global finalRtf, finalRtfSUG, textFile
    try {
        ; 1) Résoudre le chemin demandé
        if (finalRtfPath = "")
            finalRtfPath := finalRtf

        abs := finalRtfPath
        if !InStr(abs, ":")
            abs := A_ScriptDir "\" RegExReplace(finalRtfPath, "^\.(\\|/)", "")

        ; 2) Si le fichier demandé existe → insertion directe
        if FileExist(abs) {
            sel.SetRange(doc.Content.End, doc.Content.End)
            beforeEnd := doc.Content.End
            sel.InsertFile(abs)
            return (doc.Content.End > beforeEnd)
        }

        ; 3) Sinon, si on a demandé le SUG mais qu'il n'existe pas…
        ;    On tente: insérer textSUG + insérer le RTF standard si présent
        ;    (Si le standard n'existe pas non plus → return false, AddText fera son fallback.)
        ;    On compare avec le chemin SUG global.
        sugAbs := finalRtfSUG
        if !InStr(sugAbs, ":")
            sugAbs := A_ScriptDir "\" RegExReplace(finalRtfSUG, "^\.(\\|/)", "")

        isSUGRequest := (StrLower(abs) = StrLower(sugAbs))
        if isSUGRequest {
            ; Standard absolu
            stdAbs := finalRtf
            if !InStr(stdAbs, ":")
                stdAbs := A_ScriptDir "\" RegExReplace(finalRtf, "^\.(\\|/)", "")

            if FileExist(stdAbs) {
                ; 3a) Insérer la ligne textSUG (style: Arial 10 italique), puis le RTF standard
                textSUG := GetTextByKey(textFile, "textSUG"
                    , "Examen demandé, réalisé et interprété en urgence sur l'horaire de garde à la demande du médecin prescripteur.")

                sel.EndKey(6)
                startSUG := sel.Start
                sel.TypeText("`r`n`r`n" textSUG "")
                rngSUG := doc.Range(startSUG, sel.Start)
                rngSUG.Font.Name   := "Arial"
                rngSUG.Font.Size   := 10
                rngSUG.Font.Italic := True

                ; Insérer le RTF standard
                sel.SetRange(doc.Content.End, doc.Content.End)
                beforeEnd := doc.Content.End
                sel.InsertFile(stdAbs)
                return (doc.Content.End > beforeEnd)
            }
        }

        ; 4) Rien inséré ici → AddText fera le fallback (incluant textSUG si mode SUG)
        return false
    } catch as err {
        return false
    }
}

CleanupEnd(doc) {
    ; Nettoie les paragraphes vides en fin de document V1
    while (doc.Paragraphs.Count > 1) {
        lastPara := doc.Paragraphs.Item(doc.Paragraphs.Count)
        t := Trim(lastPara.Range.Text, "`r`n `t")
        if (t = "") {
            lastPara.Range.Delete()
        } else {
            break
        }
    }
}

AttVer(doc) {
    AlignLeft := 0, AlignCenter := 1

    paras := doc.Paragraphs
    cnt := paras.Count
    if (cnt < 1)
        return

    ; Dernier paragraphe non vide
    idx := 0
    Loop cnt {
      i := cnt - A_Index + 1
      t := Trim(paras.Item(i).Range.Text, "`r`n `t")
      if (t != "") {
          idx := i
          break
      }
    }
    if (idx = 0)
      return

    p := paras.Item(idx)

; --- normalisation ultra-légère et tolérante ---
txt := p.Range.Text
txt := StrReplace(txt, Chr(160), " ")       ; remplace NBSP par espace
txt := Trim(txt, "`r`n `t")
txt := RTrim(txt, " :")                     ; tolère un ":" final + espace

; --- match tolérant accents / casse ---
if !RegExMatch(txt, "i)^\s*attention\s+(?:à|a)\s+v(?:é|e)rifier\s*$")
    return


    ; --- Ajouter un Enter AVANT ---
    p.Range.InsertParagraphBefore()
    paras := doc.Paragraphs          ; refresh
    idx := idx + 1
    p := paras.Item(idx)
    if (idx > 1)
        paras.Item(idx - 1).Alignment := AlignLeft

    ; --- Mise en forme de CE paragraphe ---
    r := p.Range
    txtRange := doc.Range(r.Start, Max(r.Start, r.End - 1))
    try txtRange.Case := 1            ; wdUpperCase
    catch as err {
        up := StrUpper(txtRange.Text)
        txtRange.Text := up
    }
    txtRange.Font.Name      := "Arial"
    txtRange.Font.Size      := 12
    txtRange.Font.Bold      := True
    txtRange.Font.Underline := 1
    p.Alignment := AlignCenter

    ; Paragraphe suivant à gauche (pour ne pas polluer AddText)
    if (idx < paras.Count)
        paras.Item(idx + 1).Alignment := AlignLeft
    else {
        r2 := doc.Range(doc.Content.End, doc.Content.End)
        r2.InsertParagraphAfter()
        paras := doc.Paragraphs
        paras.Item(paras.Count).Alignment := AlignLeft
    }
}

InsertManualAttVer(doc) {
    ; Ancre sûre à la fin du document, même s’il est vide
    endPos := doc.Content.End
    startPos := (endPos > 0) ? (endPos - 1) : 0
    r := doc.Range(startPos, startPos)

    ; 2 lignes vides AVANT
    r.InsertParagraphAfter()
    r.Collapse(0)
    r.InsertParagraphAfter()
    r.Collapse(0)

    ; Insérer le texte et le formater
    r.Text := "ATTENTION À VÉRIFIER"
    r.Font.Name      := "Arial"
    r.Font.Size      := 12
    r.Font.Bold      := True
    r.Font.Underline := 1
    r.ParagraphFormat.Alignment := 1
    
    r.InsertParagraphAfter()

    ; Revenir en Arial 10 pour la suite (à la nouvelle fin)
    endPos2 := doc.Content.End
    tailPos := (endPos2 > 0) ? (endPos2 - 1) : 0
    tail := doc.Range(tailPos, tailPos)
    tail.Font.Name := "Arial"
    tail.Font.Size := 10
}