Attribute VB_Name = "Text-Normalization-For-Dharma"
'Author: Nguyen Son

Sub FixOldWords()
    'Freeze screen
    'Application.ScreenUpdating = False
    Dim xFind As String
    Dim xReplace As String
    Dim xFindArr As Variant
    Dim xReplaceArr As Variant
    xFind = "bijnh, Bijnh, hoojt, Hoojt, nhown, Nhown, nhuwst, Nhuwst"
    xReplace = "beejnh, Beejnh, hajt, Hajt, nhaan, Nhaan, nhaast, Nhaast"
    xFindArr = Split(xFind, ", ")
    xReplaceArr = Split(xReplace, ", ")
    Dim i As Long
    Dim xTempFind As String
    Dim xTempReplace As String
    For i = 0 To UBound(xFindArr)
        xTempFind = xFindArr(i)
        xTempReplace = xReplaceArr(i)
        Selection.HomeKey Unit:=wdStory
        With Selection.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = UniConvert(xTempFind, "Telex")
            .Replacement.Text = UniConvert(xTempReplace, "Telex")
            .MatchCase = True
            .MatchWholeWord = True
            .Forward = True
            .Execute Replace:=wdReplaceAll
        End With
    Next
    'Unfreeze screen
    'Application.ScreenUpdating = True
    'Application.Assistant.DoAlert "Thông báo", "Hoàn thành!", msoAlertButtonOK, msoAlertIconInfo, 0, 0, 0
End Sub

Sub FixFormat()
    Dim xFind As String
    Dim xReplace As String
    Dim xFindArr As Variant
    Dim xReplaceArr As Variant
    xFind = " :\ ?\ !\ ,\ .\ -\- \  "
    xReplace = ":\?\!\,\.\-\-\ "
    xFindArr = Split(xFind, "\")
    xReplaceArr = Split(xReplace, "\")
    Dim i As Long
    Selection.HomeKey Unit:=wdStory
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = False
        .MatchWholeWord = True
        .Forward = True
        .Wrap = wdFindContinue
        For i = 0 To UBound(xFindArr)
            .Text = xFindArr(i)
            .Replacement.Text = xReplaceArr(i)
            .Execute Replace:=wdReplaceAll
            Dim iSafe As Integer
            iSafe = 0
            Do While .Found And iSafe < 10000
                iSafe = iSafe + 1
                Selection.HomeKey Unit:=wdStory
                Selection.Find.Execute
                If .Found Then
                    .Execute Replace:=wdReplaceAll
                End If
            Loop
        Next i
    End With
    'Application.Assistant.DoAlert "Thông báo", "Hoàn thành!", msoAlertButtonOK, msoAlertIconInfo, 0, 0, 0
End Sub

Sub Format()
    'Application.ScreenUpdating = False
    Selection.HomeKey Unit:=wdStory
    ActiveDocument.Select
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 14
    End With
    With Selection.PageSetup
        .PaperSize = wdPaperA4
        .TopMargin = CentimetersToPoints(2)
        .RightMargin = CentimetersToPoints(2.5)
        .BottomMargin = CentimetersToPoints(1.5)
        .LeftMargin = CentimetersToPoints(2)
        
    End With
    With ActiveDocument.Sections(1)
        .Footers(wdHeaderFooterPrimary).PageNumbers.Add _
        PageNumberAlignment:=wdAlignPageNumberRight, _
        FirstPage:=True
    End With
    ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary).PageNumbers(1).Select
    With Selection.Font
        .Name = "Times New Roman"
        .Size = 14
    End With
    ActiveWindow.View.Type = wdPrintView
    'Application.ScreenUpdating = True
    'Application.Assistant.DoAlert "Thông báo", "Hoàn thành!", msoAlertButtonOK, msoAlertIconInfo, 0, 0, 0
End Sub

Sub FixQoutes()
    Dim xFind As Variant
    Dim xReplace As Variant
    xFind = Array("([^0147]) ", " ([^0148])", ".([^0148])", "([:a-zA-Z0-9])([^0147])", "([^0148])([:a-zA-Z0-9])")
    xReplace = Array("\1", "\1", "\1.", "\1 \2", "\1 \2")
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .MatchWildcards = True
        .Forward = True
        .Wrap = wdFindContinue
        For i = 0 To UBound(xFind)
            .Text = xFind(i)
            .Replacement.Text = xReplace(i)
            .Execute Replace:=wdReplaceAll
            Dim iSafe As Integer
            iSafe = 0
            Do While .Found And iSafe < 10000
                iSafe = iSafe + 1
                Selection.HomeKey Unit:=wdStory
                Selection.Find.Execute
                If .Found = True Then
                    .Execute Replace:=wdReplaceAll
                End If
            Loop
        Next i
    End With
    'MsgBox "Done!"
End Sub

Sub FixAll()
    Application.ScreenUpdating = False
    FixOldWords
    FixFormat
    FixQoutes
    Format
    Application.ScreenUpdating = True
    Application.Assistant.DoAlert "Thông báo", "Hoàn thành!", msoAlertButtonOK, msoAlertIconInfo, 0, 0, 0
End Sub

Function UniConvert(Text As String, InputMethod As String) As String
  Dim VNI_Type, Telex_Type, CharCode, Temp, i As Long
  UniConvert = Text
  VNI_Type = Array("a81", "a82", "a83", "a84", "a85", "a61", "a62", "a63", "a64", "a65", "e61", _
      "e62", "e63", "e64", "e65", "o61", "o62", "o63", "o64", "o65", "o71", "o72", "o73", "o74", _
      "o75", "u71", "u72", "u73", "u74", "u75", "a1", "a2", "a3", "a4", "a5", "a8", "a6", "d9", _
      "e1", "e2", "e3", "e4", "e5", "e6", "i1", "i2", "i3", "i4", "i5", "o1", "o2", "o3", "o4", _
      "o5", "o6", "o7", "u1", "u2", "u3", "u4", "u5", "u7", "y1", "y2", "y3", "y4", "y5")
  Telex_Type = Array("aws", "awf", "awr", "awx", "awj", "aas", "aaf", "aar", "aax", "aaj", "ees", _
      "eef", "eer", "eex", "eej", "oos", "oof", "oor", "oox", "ooj", "ows", "owf", "owr", "owx", _
      "owj", "uws", "uwf", "uwr", "uwx", "uwj", "as", "af", "ar", "ax", "aj", "aw", "aa", "dd", _
      "es", "ef", "er", "ex", "ej", "ee", "is", "if", "ir", "ix", "ij", "os", "of", "or", "ox", _
      "oj", "oo", "ow", "us", "uf", "ur", "ux", "uj", "uw", "ys", "yf", "yr", "yx", "yj")
  CharCode = Array(ChrW(7855), ChrW(7857), ChrW(7859), ChrW(7861), ChrW(7863), ChrW(7845), ChrW(7847), _
      ChrW(7849), ChrW(7851), ChrW(7853), ChrW(7871), ChrW(7873), ChrW(7875), ChrW(7877), ChrW(7879), _
      ChrW(7889), ChrW(7891), ChrW(7893), ChrW(7895), ChrW(7897), ChrW(7899), ChrW(7901), ChrW(7903), _
      ChrW(7905), ChrW(7907), ChrW(7913), ChrW(7915), ChrW(7917), ChrW(7919), ChrW(7921), ChrW(225), _
      ChrW(224), ChrW(7843), ChrW(227), ChrW(7841), ChrW(259), ChrW(226), ChrW(273), ChrW(233), ChrW(232), _
      ChrW(7867), ChrW(7869), ChrW(7865), ChrW(234), ChrW(237), ChrW(236), ChrW(7881), ChrW(297), ChrW(7883), _
      ChrW(243), ChrW(242), ChrW(7887), ChrW(245), ChrW(7885), ChrW(244), ChrW(417), ChrW(250), ChrW(249), _
      ChrW(7911), ChrW(361), ChrW(7909), ChrW(432), ChrW(253), ChrW(7923), ChrW(7927), ChrW(7929), ChrW(7925))
  Select Case InputMethod
    Case Is = "VNI": Temp = VNI_Type
    Case Is = "Telex": Temp = Telex_Type
  End Select
  For i = 0 To UBound(CharCode)
    UniConvert = Replace(UniConvert, Temp(i), CharCode(i))
    UniConvert = Replace(UniConvert, UCase(Temp(i)), UCase(CharCode(i)))
  Next i
End Function

