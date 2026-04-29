Attribute VB_Name = "modNumber2TextCore"
' Author    : VSBTek
Option Explicit

' Ham tro giup lay chuoi Unicode (Tranh loi encoding)
Public Function U(ByVal Key As String) As String
    Select Case LCase(Trim(Key))
        Case "khong": U = "kh" & ChrW(244) & "ng"
        Case "mot": U = "m" & ChrW(7897) & "t"
        Case "hai": U = "hai"
        Case "ba": U = "ba"
        Case "bon": U = "b" & ChrW(7889) & "n"
        Case "nam": U = "n" & ChrW(259) & "m"
        Case "sau": U = "s" & ChrW(225) & "u"
        Case "bay": U = "b" & ChrW(7843) & "y"
        Case "tam": U = "t" & ChrW(225) & "m"
        Case "chin": U = "ch" & ChrW(237) & "n"
        Case "muoi_": U = "m" & ChrW(432) & ChrW(417) & "i"
        Case "muoi": U = "m" & ChrW(432) & ChrW(7901) & "i"
        Case "tram": U = "tr" & ChrW(259) & "m"
        Case "nghin": U = "ngh" & ChrW(236) & "n"
        Case "trieu": U = "tri" & ChrW(7879) & "u"
        Case "ty": U = "t" & ChrW(7927)
        Case "linh": U = "linh"
        Case "le": U = "l" & ChrW(7867)
        Case "lam": U = "l" & ChrW(259) & "m"
        Case "mot_": U = "m" & ChrW(7889) & "t"
        Case "am": U = ChrW(194) & "m"
        Case "phay": U = "ph" & ChrW(7849) & "y"
        Case Else: U = ""
    End Select
End Function

Public Function ChonChu(ByVal n As Integer) As String
    Dim Chu As Variant
    Chu = Array(U("khong"), U("mot"), U("hai"), U("ba"), U("bon"), U("nam"), U("sau"), U("bay"), U("tam"), U("chin"))
    ChonChu = Chu(n)
End Function

Private Function Doc3So(ByVal n As Integer, ByVal DayDu As Boolean, ByVal IsLe As Boolean) As String
    Dim Tram As Integer, Chuc As Integer, DonVi As Integer
    Dim sTram As String, sChuc As String, sDonVi As String
    
    Tram = n \ 100
    Chuc = (n Mod 100) \ 10
    DonVi = n Mod 10
    
    If Tram = 0 And Not DayDu Then
        sTram = ""
    Else
        sTram = ChonChu(Tram) & " " & U("tram") & " "
    End If
    
    Select Case Chuc
        Case 0
            If Tram = 0 And Not DayDu Then
                sChuc = ""
            ElseIf DonVi = 0 Then
                sChuc = ""
            Else
                sChuc = IIf(IsLe, U("le") & " ", U("linh") & " ")
            End If
        Case 1
            sChuc = U("muoi") & " "
        Case Else
            sChuc = ChonChu(Chuc) & " " & U("muoi_") & " "
    End Select
    
    Select Case DonVi
        Case 0: sDonVi = ""
        Case 1: sDonVi = IIf(Chuc > 1, U("mot_") & " ", U("mot") & " ")
        Case 5: sDonVi = IIf(Chuc > 0, U("lam") & " ", U("nam") & " ")
        Case Else: sDonVi = ChonChu(DonVi) & " "
    End Select
    
    Doc3So = Trim(sTram & sChuc & sDonVi)
End Function

Public Function ConvertNumberToText(ByVal MyNumber As Variant, Optional ByVal IsLe As Boolean = True) As String
    Dim sNumber As String
    Dim PartTy As String, PartConLai As String
    Dim Result As String
    Dim i As Integer, Nhom(1 To 3) As Integer
    Dim HasValue As Boolean
    
    ' Chuyen sang chuoi va loai bo khoang trang, dau phay (neu co)
    sNumber = Trim(Replace(CStr(MyNumber), ",", ""))
    If sNumber = "" Or sNumber = "0" Then
        ConvertNumberToText = UCase(Left(U("khong"), 1)) & Mid(U("khong"), 2)
        Exit Function
    End If
    
    ' Xu ly so am
    If Left(sNumber, 1) = "-" Then
        Result = U("am") & " "
        sNumber = Mid(sNumber, 2)
    End If
    
    ' Bo cac so 0 o dau
    Do While Left(sNumber, 1) = "0" And Len(sNumber) > 1
        sNumber = Mid(sNumber, 2)
    Loop
    
    ' Chia de quy theo don vi Ty (9 chu so)
    If Len(sNumber) > 9 Then
        PartConLai = Left(sNumber, Len(sNumber) - 9)
        PartTy = Right(sNumber, 9)
        
        Dim sBacCao As String, sBacThap As String
        sBacCao = ConvertNumberToText(PartConLai, IsLe)
        sBacThap = Doc9So(PartTy, True, IsLe)
        
        Result = Result & sBacCao & " " & U("ty")
        If sBacThap <> "" Then
            Result = Result & " " & sBacThap
        End If
    Else
        Result = Result & Doc9So(sNumber, False, IsLe)
    End If
    
    ConvertNumberToText = Trim(Result)
End Function

' Ham tro giup doc nhom 9 chu so (3 nhom 3)
Private Function Doc9So(ByVal s9 As String, ByVal DayDu As Boolean, ByVal IsLe As Boolean) As String
    Dim n As Double
    Dim N3(1 To 3) As Integer
    Dim Units As Variant
    Dim i As Integer, Result As String
    Dim HasValue As Boolean
    
    n = Val(s9)
    If n = 0 Then
        Doc9So = ""
        Exit Function
    End If
    
    Units = Array("", U("nghin"), U("trieu"))
    
    ' Tach thanh 3 nhom 3
    For i = 1 To 3
        N3(i) = n Mod 1000
        n = Int(n / 1000)
    Next i
    
    HasValue = DayDu
    For i = 3 To 1 Step -1
        If N3(i) > 0 Then
            Result = Result & Doc3So(N3(i), HasValue, IsLe) & " " & Units(i - 1) & " "
            HasValue = True
        ElseIf HasValue And i > 1 Then
            ' Neu dang do dang ma gap nhom 0, can xem nhom tiep theo co gia tri khong
            ' Thuc te Doc3So da xu ly kha tot, o day chi can dam bao HasValue
        End If
    Next i
    
    Doc9So = Trim(Result)
End Function
