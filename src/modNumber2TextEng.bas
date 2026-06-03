Attribute VB_Name = "modNumber2TextEng"
' Author    : VSBTek
' Purpose   : Core English Number to Text Logic
Option Explicit

Private Units As Variant
Private Tens As Variant
Private Groups As Variant

Private Sub InitializeArrays()
    Units = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", _
                  "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen")
    Tens = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")
    Groups = Array("", "Thousand", "Million", "Billion", "Trillion", "Quadrillion")
End Sub

Public Function ConvertNumberToTextEng(ByVal MyNumber As Variant) As String
    Dim sNumber As String
    Dim Result As String
    Dim i As Integer, GroupIdx As Integer
    Dim n As Double
    
    If IsEmpty(Units) Then InitializeArrays
    
    sNumber = Trim(Replace(CStr(MyNumber), ",", ""))
    If sNumber = "" Or sNumber = "0" Then
        ConvertNumberToTextEng = "Zero"
        Exit Function
    End If
    
    If Left(sNumber, 1) = "-" Then
        Result = "Minus "
        sNumber = Mid(sNumber, 2)
    End If
    
    ' Split into groups of 3
    Dim Part As String
    GroupIdx = 0
    Do While Len(sNumber) > 0
        If Len(sNumber) > 3 Then
            Part = Right(sNumber, 3)
            sNumber = Left(sNumber, Len(sNumber) - 3)
        Else
            Part = sNumber
            sNumber = ""
        End If
        
        Dim sPart As String
        sPart = Doc3SoEng(Val(Part))
        
        If sPart <> "" Then
            If Groups(GroupIdx) <> "" Then
                Result = sPart & " " & Groups(GroupIdx) & " " & Result
            Else
                Result = sPart & " " & Result
            End If
        End If
        GroupIdx = GroupIdx + 1
    Loop
    
    ConvertNumberToTextEng = Trim(Result)
End Function

Private Function Doc3SoEng(ByVal n As Integer) As String
    Dim h As Integer, t As Integer, u As Integer
    Dim Result As String
    
    h = n \ 100
    t = n Mod 100
    
    If h > 0 Then
        Result = Units(h) & " Hundred"
        If t > 0 Then Result = Result & " and "
    End If
    
    If t > 0 Then
        If t < 20 Then
            Result = Result & Units(t)
        Else
            Result = Result & Tens(t \ 10)
            If (t Mod 10) > 0 Then
                Result = Result & "-" & Units(t Mod 10)
            End If
        End If
    End If
    
    Doc3SoEng = Result
End Function

Public Function ChonChuEng(ByVal n As Integer) As String
    Dim Chu As Variant
    Chu = Array("Zero", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine")
    ChonChuEng = Chu(n)
End Function
