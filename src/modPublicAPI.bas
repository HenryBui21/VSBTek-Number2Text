Attribute VB_Name = "modPublicAPI"
' Author    : VSBTek
' Purpose   : Giao dien ham cho nguoi dung Excel
' FontType  : 1 = Unicode, 2 = VNI, 3 = TCVN3
' CurrType  : 1 = VND, 2 = USD
Option Explicit

' Ham lay chuoi Unicode cho don vi tien te
Private Function U_Curr(ByVal Key As String) As String
    Select Case LCase(Key)
        Case "dong": U_Curr = ChrW(273) & ChrW(7891) & "ng"
        Case "chan": U_Curr = "ch" & ChrW(7861) & "n"
        Case "xu": U_Curr = "xu"
        Case "dola": U_Curr = ChrW(273) & ChrW(244) & " la M" & ChrW(7923)
        Case "cents": U_Curr = "cents"
        Case Else: U_Curr = ""
    End Select
End Function

Private Function DocSoInternal(ByVal Amount As Variant, _
                               ByVal CurrType As Integer, _
                               ByVal FontType As Integer, _
                               ByVal UseLe As Boolean) As String
    Dim Result As String
    Dim PartLe As Double
    Dim sNguyen As String, sLeStr As String
    Dim sFull As String, sep As String
    Dim parts() As String
    Dim i As Integer, dgt As Integer

    On Error GoTo ErrorHandler

    ' Lay dau thap phan cua he thong (vi du: . hoac ,)
    sep = Mid(Format(0, "0.0"), 2, 1)
    
    ' Chuyen so sang chuoi chuan
    sFull = CStr(Amount)
    If InStr(sFull, sep) > 0 Then
        parts = Split(sFull, sep)
        sNguyen = ConvertNumberToText(parts(0), UseLe)
        sLeStr = parts(1)
    Else
        sNguyen = ConvertNumberToText(sFull, UseLe)
        sLeStr = ""
    End If

    ' Bat dau doc tung kieu
    If CurrType = 1 Then ' VND
        If Len(sLeStr) > 0 Then
            Result = sNguyen & " " & U("phay") & " "
            For i = 1 To Len(sLeStr)
                dgt = Val(Mid(sLeStr, i, 1))
                Result = Result & ChonChu(dgt) & " "
            Next i
            Result = Result & U_Curr("dong")
        Else
            Result = sNguyen & " " & U_Curr("dong") & " " & U_Curr("chan")
        End If
    ElseIf CurrType = 2 Then ' USD (Accounting English)
        Dim nNguyen As Double, nLe As Integer
        nNguyen = Fix(Val(sFull))
        If Len(sLeStr) > 0 Then
            nLe = Val(Left(sLeStr & "00", 2))
        Else
            nLe = 0
        End If
        
        Dim sPart1 As String, sPart2 As String
        If nNguyen > 0 Then
            sPart1 = ConvertNumberToTextEng(nNguyen) & IIf(nNguyen <= 1, " Dollar", " Dollars")
        End If
        
        If nLe > 0 Then
            sPart2 = ConvertNumberToTextEng(nLe) & IIf(nLe <= 1, " Cent", " Cents")
        End If
        
        If nNguyen > 0 And nLe > 0 Then
            Result = sPart1 & " and " & sPart2 & " only."
        ElseIf nNguyen > 0 Then
            Result = sPart1 & " only."
        ElseIf nLe > 0 Then
            Result = sPart2 & " only."
        Else
            Result = "Zero Dollar only."
        End If
    ElseIf CurrType = 3 Then ' General English
        Result = ConvertNumberToTextEng(Amount) & " only."
    Else
        ' Kieu so thong thuong (Vietnamese): Doc "phay" + tung chu so
        If Len(sLeStr) > 0 Then
            Result = sNguyen & " " & U("phay") & " "
            For i = 1 To Len(sLeStr)
                dgt = Val(Mid(sLeStr, i, 1))
                Result = Result & ChonChu(dgt) & " "
            Next i
        Else
            Result = sNguyen
        End If
    End If

    Result = Trim(Result)
    Result = UCase(Left(Result, 1)) & Mid(Result, 2)
    
    If CurrType = 2 Or CurrType = 3 Then
        DocSoInternal = Result ' English logic doesn't need Vietnamese font conversion
    Else
        DocSoInternal = ConvertFromUnicode(Result, FontType)
    End If
    Exit Function

ErrorHandler:
    DocSoInternal = "#ERROR!"
End Function

''' Doc so thanh VND: =VND(so) hoac =VND(so, kieu_chu)
Public Function VND(ByVal Amount As Variant, _
                    Optional ByVal FontType As Integer = 1, _
                    Optional ByVal UseLe As Boolean = True) As String
    VND = DocSoInternal(Amount, 1, FontType, UseLe)
End Function

''' Doc so thanh USD (English): =USD(so)
Public Function USD(ByVal Amount As Variant) As String
    USD = DocSoInternal(Amount, 2, 1, True)
End Function

''' Doc so sang tieng Anh: =ENG(so)
Public Function ENG(ByVal Amount As Variant) As String
    ENG = DocSoInternal(Amount, 3, 1, True)
End Function

''' Doc so tuy chinh: =DocSo(so, loai_tien, kieu_chu)
''' loai_tien: 1=VND, 2=USD (Eng), 3=English, 0=Common VN
Public Function DocSo(ByVal Amount As Variant, _
                      Optional ByVal CurrType As Integer = 0, _
                      Optional ByVal FontType As Integer = 1, _
                      Optional ByVal UseLe As Boolean = True) As String
    DocSo = DocSoInternal(Amount, CurrType, FontType, UseLe)
End Function
