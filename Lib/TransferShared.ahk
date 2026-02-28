#Requires AutoHotkey v2

; Convert all RTF \'xx (hex) sequences to real characters.
DecodeRTFHex(s) {
    out := ""
    pos := 1
    while RegExMatch(s, "\\'([0-9A-Fa-f]{2})", &m, pos) {
        out .= SubStr(s, pos, m.Pos(0) - pos)
        out .= Chr("0x" m[1])
        pos := m.Pos(0) + m.Len(0)
    }
    return out . SubStr(s, pos)
}

; Decode \u#### (including negative values) to Unicode characters.
DecodeRTFUnicode(s) {
    out := ""
    pos := 1
    while RegExMatch(s, "\\u(-?\d+)\??", &m, pos) {
        out .= SubStr(s, pos, m.Pos(0) - pos)
        code := m[1] + 0
        if (code < 0)
            code += 65536
        out .= Chr(code)
        pos := m.Pos(0) + m.Len(0)
    }
    return out . SubStr(s, pos)
}

RensMaj(doc) {
    try {
        p1  := doc.Paragraphs.Item(1).Range
        txt := p1.Text

        if !RegExMatch(txt, "i)^\s*renseignements cliniques\s*:")
            return

        posColon := InStr(txt, ":")
        if (posColon = 0)
            return

        i := posColon + 1
        while (i <= StrLen(txt)) {
            ch := SubStr(txt, i, 1)
            if (ch = " " || ch = "`t" || Ord(ch) = 160) {
                i++
                continue
            }
            if (ch = "[" && SubStr(txt, i, 2) = "[]") {
                i += 2
                continue
            }
            break
        }
        if (i > StrLen(txt))
            return

        absStart := p1.Start + i - 1
        absEnd   := absStart + 1

        ; Uppercase without visible selection.
        r := doc.Range(absStart, absEnd)
        r.Case := 1

        ; Fallback if Word Case assignment did not apply.
        one := r.Text
        up  := StrUpper(one)
        if (up != one)
            r.Text := up

    } catch as err {
        ; Ignore non-critical formatting errors.
    }
}

CapitalizeParagraphStarts(doc, maxScan := 40) {
    wdUpperCase := 1
    paras := doc.Paragraphs
    count := paras.Count

    Loop count {
        para := paras.Item(A_Index)
        r := para.Range.Duplicate

        ; Empty paragraph / paragraph mark only.
        if (r.End - r.Start <= 1)
            continue

        ; Remove trailing paragraph marker.
        r.End -= 1

        start := r.Start
        end   := r.End
        if (end <= start)
            continue

        sampleEnd := start + maxScan
        if (sampleEnd > end)
            sampleEnd := end

        ; Read only paragraph start (single COM call).
        sample := doc.Range(start, sampleEnd).Text
        if (sample = "")
            continue

        ; First letter (not a number) inside sampled text.
        if !RegExMatch(sample, "\p{L}", &m)
            continue

        ; Absolute document position of first letter.
        pos := start + (m.Pos[0] - 1)

        chR := doc.Range(pos, pos + 1)
        ch  := chR.Text

        ; Lowercase letter -> uppercase.
        if RegExMatch(ch, "^\p{Ll}$")
            chR.Case := wdUpperCase
    }
}

HasArg(name) {
    for v in A_Args
        if (StrLower(v) = StrLower(name))
            return true
    return false
}
