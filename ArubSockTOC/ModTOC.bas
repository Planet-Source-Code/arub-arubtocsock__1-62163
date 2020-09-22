Attribute VB_Name = "ModTOC"
'some of the binary functions are taken from Xeon, thanks to him for that
Option Explicit
Dim i As Long
Public Function EncryptPW(ByRef strPass As String) As String
    
    Dim arrTable() As Variant, strEncrypted As String
    Dim lngX As Long, strHex As String
    arrTable = Array("84", "105", "99", "47", "84", "111", "99")
    strEncrypted$ = "0x"
    For lngX& = 0 To Len(strPass$) - 1
        strHex$ = Hex(Asc(Mid(strPass$, lngX& + 1, 1)) Xor CLng(arrTable((lngX& Mod 7))))
        If CLng("&H" & strHex$) < 16 Then strEncrypted$ = strEncrypted$ & "0"
        strEncrypted$ = strEncrypted$ & strHex$
    Next
    EncryptPW$ = LCase(strEncrypted$)
End Function
'Put a value up to 65535 into this, and get a 2 byte integer
Public Function Word(ByVal lngVal As Long) As String
    Dim Lo As Single
    Dim Hi As Single

    Lo = Fix(lngVal / 256)
    Hi = lngVal Mod 256

    Word = Chr(Lo) & Chr(Hi)
End Function

'Input a 2 byte integer into this, and get a value out
Public Function GetWord(ByVal strVal As String) As Long
    Dim Lo As Long
    Dim Hi As Long
    
    Lo = Asc(Mid(strVal, 1, 1))
    Hi = Asc(Mid(strVal, 2, 1))
    
    GetWord = (Lo * 256) + Hi
End Function
Public Function Normalize(strShit As String)
    'normalize it for the toc server
    Dim strString
    strString = Array(Chr(34), Chr(92), "[", "]", "{", "}", "(", ")")
        For i = 0 To UBound(strString)
            strShit = Replace(strShit, strString(i), Chr(92) & strString(i))
        Next i
    Normalize = strShit
End Function
Public Function Qt(strString)
    'encircle a string in quotes
    Qt = Chr(34) & strString & Chr(34)
End Function
Public Function Minimal(strString)
    'get lower case without spaces
    Minimal = LCase(Replace(strString, " ", vbNullString))
End Function

