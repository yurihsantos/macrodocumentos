Function ORDINAL(ByVal NUM As Long) As String

    If NUM < 1 Or NUM > 9999 Then
        ORDINAL = "Fora do Escopo"
        Exit Function
    End If

    Dim UNIS As Integer 'Unidade Simples
    Dim DEZS As Integer 'Dezena Simples
    Dim CENS As Integer 'Centena Simples
    Dim UNMI As Integer 'Unidade de Milhar
    Dim DEMI As Integer 'Dezena de Milhar
    Dim CEMI As Integer 'Centena de Milhar
    Dim UNMM As Integer 'Unidade de Milhão
    Dim DEMM As Integer 'Dezena de Milhão
    Dim CEMM As Integer 'Centena de Milhão

    UNIS = Int((NUM Mod (10 ^ 1)) / (10 ^ 0))
    DEZS = Int((NUM Mod (10 ^ 2)) / (10 ^ 1))
    CENS = Int((NUM Mod (10 ^ 3)) / (10 ^ 2))
    UNMI = Int((NUM Mod (10 ^ 4)) / (10 ^ 3))
    DEMI = Int((NUM Mod (10 ^ 5)) / (10 ^ 4))
    CEMI = Int((NUM Mod (10 ^ 6)) / (10 ^ 5))
    UNMM = Int((NUM Mod (10 ^ 7)) / (10 ^ 6))
    DEMM = Int((NUM Mod (10 ^ 8)) / (10 ^ 7))
    CEMM = Int((NUM Mod (10 ^ 9)) / (10 ^ 8))

    Select Case UNMM: Case 1 To 9
        ORDINAL = _
        Choose(UNMM, _
        "milionésimo", _
        "segundo milionésimo", _
        "terceiro milionésimo", _
        "quarto milionésimo", _
        "quinto milionésimo", _
        "sexto milionésimo", _
        "sétimo milionésimo", _
        "oitavo milionésimo", _
        "nono milionésimo")
    End Select

    Select Case CEMI: Case 1 To 9
        ORDINAL = ORDINAL & " " & _
        Choose(CEMI, _
        "centésimo", _
        "ducentésimo", _
        "tricentésimo", _
        "quadringentésimo", _
        "quingentésimo", _
        "sexcentésimo", _
        "septingentésimo", _
        "octingentésimo", _
        "nongentésimo")
    End Select

    Select Case DEMI: Case 1 To 9
        ORDINAL = ORDINAL & " " & _
        Choose(DEMI, _
        "décimo", _
        "vigésimo", _
        "trigésimo", _
        "quadragésimo", _
        "quinquagésimo", _
        "sexagésimo", _
        "septuagésimo", _
        "octogésimo", _
        "nonagésimo")
    End Select

    Select Case UNMI: Case 1 To 9
        ORDINAL = ORDINAL & " " & _
        Choose(UNMI, _
        "primeiro milésimo", _
        "segundo milésimo", _
        "terceiro milésimo", _
        "quarto milésimo", _
        "quinto milésimo", _
        "sexto milésimo", _
        "sétimo milésimo", _
        "oitavo milésimo", _
        "nono milésimo")
    End Select

    Select Case CENS: Case 1 To 9
        ORDINAL = ORDINAL & " " & _
        Choose(CENS, _
        "centésimo", _
        "ducentésimo", _
        "tricentésimo", _
        "quadringentésimo", _
        "quingentésimo", _
        "sexcentésimo", _
        "septingentésimo", _
        "octingentésimo", _
        "nongentésimo")
    End Select

    Select Case DEZS: Case 1 To 9
        ORDINAL = ORDINAL & " " & _
        Choose(DEZS, _
        "décimo", _
        "vigésimo", _
        "trigésimo", _
        "quadragésimo", _
        "quinquagésimo", _
        "sexagésimo", _
        "septuagésimo", _
        "octogésimo", _
        "nonagésimo")
    End Select
    
    Select Case UNIS: Case 1 To 9
        ORDINAL = ORDINAL & " " & _
        Choose(UNIS, _
        "primeiro", _
        "segundo", _
        "terceiro", _
        "quarto", _
        "quinto", _
        "sexto", _
        "sétimo", _
        "oitavo", _
        "nono")
    End Select

End Function