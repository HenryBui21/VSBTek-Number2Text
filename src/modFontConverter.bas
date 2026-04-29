Attribute VB_Name = "modFontConverter"
Option Explicit

' Ham lay chuoi Unicode chuan tu ma so
Private Function GetUniChars() As String
    Dim s As String
    s = ChrW(224) & "," & ChrW(225) & "," & ChrW(7843) & "," & ChrW(227) & "," & ChrW(7841) & "," & ChrW(226) & "," & ChrW(7847) & "," & ChrW(7845) & "," & ChrW(7849) & "," & ChrW(7851) & "," & ChrW(7853) & "," & ChrW(259) & "," & ChrW(7857) & "," & ChrW(7855) & "," & ChrW(7859) & "," & ChrW(7861) & "," & ChrW(7863) & "," & _
        ChrW(232) & "," & ChrW(233) & "," & ChrW(7867) & "," & ChrW(7869) & "," & ChrW(7865) & "," & ChrW(234) & "," & ChrW(7873) & "," & ChrW(7871) & "," & ChrW(7875) & "," & ChrW(7877) & "," & ChrW(7879) & "," & ChrW(236) & "," & ChrW(237) & "," & ChrW(7881) & "," & ChrW(297) & "," & ChrW(7883) & "," & _
        ChrW(242) & "," & ChrW(243) & "," & ChrW(7887) & "," & ChrW(245) & "," & ChrW(7885) & "," & ChrW(244) & "," & ChrW(7891) & "," & ChrW(7889) & "," & ChrW(7893) & "," & ChrW(7895) & "," & ChrW(7897) & "," & ChrW(417) & "," & ChrW(7901) & "," & ChrW(7899) & "," & ChrW(7903) & "," & ChrW(7905) & "," & ChrW(7907) & "," & _
        ChrW(249) & "," & ChrW(250) & "," & ChrW(7911) & "," & ChrW(361) & "," & ChrW(7909) & "," & ChrW(432) & "," & ChrW(7915) & "," & ChrW(7913) & "," & ChrW(7917) & "," & ChrW(7919) & "," & ChrW(7921) & "," & ChrW(7923) & "," & ChrW(253) & "," & ChrW(7927) & "," & ChrW(7925) & "," & ChrW(7929) & "," & ChrW(273) & "," & _
        ChrW(192) & "," & ChrW(193) & "," & ChrW(7842) & "," & ChrW(195) & "," & ChrW(7840) & "," & ChrW(194) & "," & ChrW(7846) & "," & ChrW(7844) & "," & ChrW(7848) & "," & ChrW(7850) & "," & ChrW(7852) & "," & ChrW(258) & "," & ChrW(7856) & "," & ChrW(7854) & "," & ChrW(7858) & "," & ChrW(7860) & "," & ChrW(7862) & "," & _
        ChrW(200) & "," & ChrW(201) & "," & ChrW(7866) & "," & ChrW(7868) & "," & ChrW(7864) & "," & ChrW(202) & "," & ChrW(7872) & "," & ChrW(7870) & "," & ChrW(7874) & "," & ChrW(7876) & "," & ChrW(7878) & "," & ChrW(204) & "," & ChrW(205) & "," & ChrW(7880) & "," & ChrW(296) & "," & ChrW(7882) & "," & _
        ChrW(210) & "," & ChrW(211) & "," & ChrW(7886) & "," & ChrW(213) & "," & ChrW(7884) & "," & ChrW(204) & "," & ChrW(7890) & "," & ChrW(7888) & "," & ChrW(7892) & "," & ChrW(7894) & "," & ChrW(7896) & "," & ChrW(416) & "," & ChrW(7900) & "," & ChrW(7898) & "," & ChrW(7902) & "," & ChrW(7904) & "," & ChrW(7906) & "," & _
        ChrW(217) & "," & ChrW(218) & "," & ChrW(7910) & "," & ChrW(360) & "," & ChrW(7908) & "," & ChrW(431) & "," & ChrW(7914) & "," & ChrW(7912) & "," & ChrW(7916) & "," & ChrW(7918) & "," & ChrW(7920) & "," & ChrW(7922) & "," & ChrW(221) & "," & ChrW(7926) & "," & ChrW(7924) & "," & ChrW(7928) & "," & ChrW(272)
    GetUniChars = s
End Function

' Ham lay chuoi VNI (Ma hoa ANSI)
Private Function GetVniChars() As String
    GetVniChars = "a" & Chr(248) & ",a" & Chr(249) & ",a" & Chr(251) & ",a" & Chr(245) & ",a" & Chr(239) & ",a" & Chr(226) & ",a" & Chr(224) & ",a" & Chr(225) & ",a" & Chr(229) & ",a" & Chr(227) & ",a" & Chr(228) & ",a" & Chr(234) & ",a" & Chr(232) & ",a" & Chr(233) & ",a" & Chr(250) & ",a" & Chr(252) & ",a" & Chr(235) & ",e" & Chr(248) & ",e" & Chr(249) & ",e" & Chr(251) & ",e" & Chr(245) & ",e" & Chr(239) & ",e" & Chr(226) & ",e" & Chr(224) & ",e" & Chr(225) & ",e" & Chr(229) & ",e" & Chr(227) & ",e" & Chr(228) & "," & Chr(236) & "," & Chr(237) & "," & Chr(78) & ",ó,ị,o" & Chr(248) & ",o" & Chr(249) & ",o" & Chr(251) & ",o" & Chr(245) & ",o" & Chr(239) & ",o" & Chr(226) & ",o" & Chr(224) & ",o" & Chr(225) & ",o" & Chr(229) & ",o" & Chr(227) & ",o" & Chr(228) & "," & Chr(244) & Chr(248) _
        & "," & Chr(244) & Chr(249) & "," & Chr(244) & Chr(251) & "," & Chr(244) & Chr(245) & "," & Chr(244) & Chr(239) & ",u" & Chr(248) & ",u" & Chr(249) & ",u" & Chr(251) & ",u" & Chr(245) & ",u" & Chr(239) & "," & Chr(246) & Chr(248) & "," & Chr(246) & Chr(249) & "," & Chr(246) & Chr(251) & "," & Chr(246) & Chr(245) & "," & Chr(246) & Chr(239) & ",y" & Chr(248) & ",y" & Chr(249) & ",y" & Chr(251) & ",y" & Chr(245) & ",y" & Chr(239) & "," & Chr(241) & ",A" & Chr(216) & ",A" & Chr(217) & ",A" & Chr(219) & ",A" & Chr(213) & ",A" & Chr(207) & ",A" & Chr(194) & ",A" & Chr(192) & ",A" & Chr(193) & ",A" & Chr(197) & ",A" & Chr(195) & ",A" & Chr(196) & ",A" & Chr(202) & ",A" & Chr(200) & ",A" & Chr(201) & ",A" & Chr(218) & ",A" & Chr(220) & ",A" & Chr(203) & ",E" & Chr(216) & ",E" & Chr(217) _
        & ",E" & Chr(219) & ",E" & Chr(213) & ",E" & Chr(207) & ",E" & Chr(194) & ",E" & Chr(192) & ",E" & Chr(193) & ",E" & Chr(197) & ",E" & Chr(195) & ",E" & Chr(196) & "," & Chr(204) & "," & Chr(205) & "," & Chr(73) & ",Ó,Ị,O" & Chr(216) & ",O" & Chr(217) & ",O" & Chr(219) & ",O" & Chr(213) & ",O" & Chr(207) & ",O" & Chr(194) & ",O" & Chr(192) & ",O" & Chr(193) & ",O" & Chr(197) & ",O" & Chr(195) & ",O" & Chr(196) & "," & Chr(212) & Chr(216) & "," & Chr(212) & Chr(217) & "," & Chr(212) & Chr(219) & "," & Chr(212) & Chr(213) & "," & Chr(212) & Chr(207) & ",U" & Chr(216) & ",U" & Chr(217) & ",U" & Chr(219) & ",U" & Chr(213) & ",U" & Chr(207) & "," & Chr(214) & Chr(216) & "," & Chr(214) & Chr(217) & "," & Chr(214) & Chr(219) & "," & Chr(214) & Chr(213) & "," & Chr(214) & Chr(207) & ",Y" _
        & Chr(216) & ",Y" & Chr(217) & ",Y" & Chr(219) & ",Y" & Chr(213) & ",Y" & Chr(207) & "," & Chr(209)
End Function

' Ham lay chuoi TCVN3 (Ma hoa ANSI)
Private Function GetTcvnChars() As String
    GetTcvnChars = Chr(181) & "," & Chr(184) & "," & Chr(182) & "," & Chr(183) & "," & Chr(185) & "," & Chr(169) & "," & Chr(229) & "," & Chr(232) & "," & Chr(230) & "," & Chr(231) & "," & Chr(234) & "," & Chr(168) & "," & Chr(187) & "," & Chr(190) & "," & Chr(188) & "," & Chr(189) & "," & Chr(198) & "," & Chr(204) & "," & Chr(208) & "," & Chr(206) & "," & Chr(207) & "," & Chr(209) & "," & Chr(170) & "," & Chr(210) & "," & Chr(213) & "," & Chr(211) & "," & Chr(212) & "," & Chr(214) & "," & Chr(221) & "," & Chr(223) & "," & Chr(222) & "," & Chr(224) & "," & Chr(225) & "," & Chr(242) & "," & Chr(243) & "," & Chr(244) & "," & Chr(245) & "," & Chr(246) & "," & Chr(171) & "," & Chr(248) & "," & Chr(253) & "," & Chr(249) & "," & Chr(250) & "," & Chr(251) & "," & Chr(172) & "," & Chr(254) & "," _
        & Chr(255) & "," & Chr(241) & "," & Chr(242) & "," & Chr(243) & "," & Chr(247) & "," & Chr(248) & "," & Chr(249) & "," & Chr(250) & "," & Chr(251) & "," & Chr(173) & "," & Chr(236) & "," & Chr(238) & "," & Chr(237) & "," & Chr(239) & "," & Chr(243) & "," & Chr(250) & "," & Chr(253) & "," & Chr(254) & "," & Chr(255) & "," & Chr(174) & "," & Chr(181) & "," & Chr(184) & "," & Chr(182) & "," & Chr(183) & "," & Chr(185) & "," & Chr(169) & "," & Chr(229) & "," & Chr(232) & "," & Chr(230) & "," & Chr(231) & "," & Chr(234) & "," & Chr(168) & "," & Chr(187) & "," & Chr(190) & "," & Chr(188) & "," & Chr(189) & "," & Chr(198) & "," & Chr(204) & "," & Chr(208) & "," & Chr(206) & "," & Chr(207) & "," & Chr(209) & "," & Chr(170) & "," & Chr(210) & "," & Chr(213) & "," & Chr(211) & "," & Chr(212) _
        & "," & Chr(214) & "," & Chr(221) & "," & Chr(223) & "," & Chr(222) & "," & Chr(224) & "," & Chr(225) & "," & Chr(242) & "," & Chr(243) & "," & Chr(244) & "," & Chr(245) & "," & Chr(246) & "," & Chr(171) & "," & Chr(248) & "," & Chr(253) & "," & Chr(249) & "," & Chr(250) & "," & Chr(251) & "," & Chr(172) & "," & Chr(254) & "," & Chr(255) & "," & Chr(241) & "," & Chr(242) & "," & Chr(243) & "," & Chr(247) & "," & Chr(248) & "," & Chr(249) & "," & Chr(250) & "," & Chr(251) & "," & Chr(173) & "," & Chr(236) & "," & Chr(238) & "," & Chr(237) & "," & Chr(239) & "," & Chr(243) & "," & Chr(250) & "," & Chr(253) & "," & Chr(254) & "," & Chr(255) & "," & Chr(174)
End Function

Public Function ConvertFromUnicode(ByVal Text As String, ByVal TargetFont As Integer) As String
    Dim uniArr() As String, targetArr() As String
    Dim i As Integer
    Dim Result As String
    
    If TargetFont = 1 Or Text = "" Then
        ConvertFromUnicode = Text
        Exit Function
    End If
    
    uniArr = Split(GetUniChars(), ",")
    If TargetFont = 2 Then
        targetArr = Split(GetVniChars(), ",")
    Else
        targetArr = Split(GetTcvnChars(), ",")
    End If
    
    Result = Text
    For i = 0 To UBound(uniArr)
        Result = Replace(Result, uniArr(i), targetArr(i))
    Next i
    
    ConvertFromUnicode = Result
End Function
