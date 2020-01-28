Attribute VB_Name = "Assistant"
Function SetMycell(ByVal s As String)
    Worksheets("Accueil").Cells(2, 9).Value = s
End Function

Function GetMycell() As String
    GetMycell = Worksheets("Accueil").Cells(2, 9).Value
End Function

Sub ProjetIntendance()
    If MsgBox("Bienvenue" + vbCrLf + "Voulez-vous remettre l'assistant à zero ?", vbYesNo, Title:="Assistant Intendance") = vbYes Then
        For Each s In ActiveWorkbook.Worksheets
            If (s.Name <> "BaseAliments" And s.Name <> "BaseRecettes") Then
                If Sheets.Count = 1 Then Call CreateSheet("BaseAliments", 1)
                s.Delete
            End If
        Next
        
        If CreateSheet("Accueil", 1) Then SetAccueil
        If CreateSheet("Menu", 2) Then SetMenu
        If CreateSheet("Liste Sec", 3) Then SetSec
        If CreateSheet("Liste Frais", 4) Then SetFrais
        If CreateSheet("Recettes", 5) Then SetRecettes

    End If
    If MsgBox("Voulez-vous mettre à jour la feuille de BaseAliments sur Internet ?", vbYesNo, Title:="Mise a jour ?") = vbYes Then
        Dim c: c = GetPingResult("google.fr")
        If (c = "Connected") Then
            If sheetExists("BaseAliments") Then Worksheets("BaseAliments").Delete
            If sheetExists("BaseRecettes") Then Worksheets("BaseRecettes").Delete
            If CreateSheet("BaseAliments", 6) Then SetBaseAliments
            If CreateSheet("BaseRecettes", 7) Then SetBaseRecettes
        Else
            MsgBox "Impossible de se connecter : " & vbCrLf & c
        End If
    End If
End Sub




Private Sub SetAccueil()
    With Worksheets("Accueil")
        .Range("A1:I1").MergeCells = True
        .Range("A1:I1").Value = "Assistant d'Intendance"
        .Range("A1:I1").Font.size = 64
        .Range("A1:I1").Font.Color = RGB(255, 0, 255)
        
        .Rows("1:3").HorizontalAlignment = xlCenter
        .Rows("2:3").Font.Bold = True
        .Columns("A:I").ColumnWidth = 20
        
        .Range("A2").Value = "Nb de jours :"
        .Range("A2:A3").Style = "Entrée"
        .Range("A2:A3").Interior.Color = RGB(255, 0, 0)
        .Range("A2:A3").FormatConditions.Add(Type:=xlExpression, Formula1:="=$A$3 <> """"").Interior.Color = RGB(220, 255, 220)
                
        .Range("C2").Value = "Date de Début :"
        .Range("C2:C3").Style = "Entrée"
        .Range("C2:C3").NumberFormat = "dd/mm/yyyy"
        .Range("C2:C3").Interior.Color = RGB(255, 0, 0)
        .Range("C2:C3").FormatConditions.Add(Type:=xlExpression, Formula1:="=$C$3 <> """"").Interior.Color = RGB(220, 255, 220)
         
        .Range("E2").Value = "Nb de gros mangeurs :"
        .Range("E2:E3").Style = "Entrée"
        .Range("E2:E3").Interior.Color = RGB(255, 0, 0)
        .Range("E2:E3").FormatConditions.Add(Type:=xlExpression, Formula1:="=$E$3 <> """"").Interior.Color = RGB(220, 255, 220)
        
        .Range("G2").Value = "Jours courses :"
        .Range("G3").Formula = "=C3"
        .Range("G4").Value = "Remplir les jours dans cette colonne"
        .Range("G2:G3").Style = "Entrée"
        .Range("G2:G3").Interior.Color = RGB(255, 0, 0)
        
        .Columns(7).FormatConditions.Add(Type:=xlExpression, Formula1:="=G1<> """"").Interior.Color = RGB(220, 255, 220)
        .Columns(7).NumberFormat = "ddd d mmmm yy"
        
        .Range("A5:E6").MergeCells = True
        .Range("A5:E6").Value = "proportions moyenne de mileu de camps (A ajuster selon vos estimation, en general avec -20% en début et +20% en fin de camp)"
        .Range("A7:E7").MergeCells = True
        .Range("A7:E7").Value = "les quantité sont prévu en pièce (standard) si non précisé"
        .Range("A9:E9").MergeCells = True
        .Range("A9:E9").Value = "Comment utiliser l'assistant :"
        


        .Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 220, 300, 200).TextFrame.Characters.Text = _
"Pour tout reinitialiser : " & vbCrLf & _
"*Clic droit sur le ruban" & vbCrLf & _
"*Personnaliser le ruban" & vbCrLf & _
"*Dans la liste ""Categorie"" Choisir ""Macros""" & vbCrLf & _
"*En bas a droite,  clic sur ""Nouveau groupe""" & vbCrLf & _
"*Puis clic sur ""Ajouter""" & vbCrLf & _
"*L'icone devrait mtn apparaitre a droite du ruban" & vbCrLf & _
"*Cliquer dessus vous proposeras de  le remettre a zero (feuille par feuille),  puis vous proposera de mettre a jour les bases"

        .Shapes.AddTextbox(msoTextOrientationHorizontal, 400, 220, 400, 200).TextFrame.Characters.Text = _
"Essayez de remplir les 4 informations ci-dessus d'abord, c'est mieux :-)" & vbCrLf & _
"Vous devez remplir la feuille ""Menu"", les listes se génèrent automatiquement." & vbCrLf & _
"Les propositions de la feuille Menu sont celles des bases aliments et recettes" & vbCrLf & _
"Vous pouvez ajouter des aliments et recettes dans les bases" & vbCrLf & _
"Vous pourrez, à la fin, ajouter ce que vous voulez à la main dans les listes" & vbCrLf
        
        Dim hlink As Shape: Set hlink = .Shapes.AddLabel(msoTextOrientationHorizontal, 500, 400, 200, 10)
        hlink.TextFrame.Characters.Text = "Questions, idées, bugs : Contactez-moi !"
        hlink.TextFrame.Characters.Font.Color = RGB(6, 69, 173)
        .Hyperlinks.Add hlink, "mailto:fx73000@yahoo.fr?subject=(AssistantIntendance) - Help"
        
        .Range("J1").Value = "1.1"
        .Shapes.AddTextbox(msoTextOrientationHorizontal, 1040, 82, 80, 400).TextFrame.Characters.Text = _
"Version Log :" & vbCrLf & _
"Menu et Listes Fonctionnels, à tester" & vbCrLf & _
"Recettes not ok"


        
         Set btn = .Buttons.Add(870, 83, 110, 30)
    End With
    btn.Name = "Valider"
    btn.Caption = "Valider"
    btn.OnAction = "Actualiser"
End Sub

Private Sub SetSec()
 With Worksheets("Liste Sec")
    .Columns(2).Font.Bold = True
    .Columns(2).ColumnWidth = .Columns(2).ColumnWidth * 2
    .Columns(4).NumberFormat = "# ?/?"
    .Columns(4).HorizontalAlignment = xlHAlignCenter
    .Columns(6).Style = "Entrée"
    .Columns(6).NumberFormat = "# ?/?"
    .Rows(1).Font.Bold = True
    .Range("A1").Value = "Liste de course sec"
    .Range("C1").Value = "Nb de repas"
    .Range("D1").Value = "Proportion"
    .Range("F1").Value = "A acheter !"
    .Range("G1").Value = "Unité"
 End With
End Sub

Private Sub SetFrais()
 With Worksheets("Liste Frais")
    .Columns(2).Font.Bold = True
    .Columns(2).ColumnWidth = 30
    .Columns(4).NumberFormat = "# ?/?"
    .Columns(4).HorizontalAlignment = xlHAlignCenter
    .Columns(6).Style = "Entrée"
    .Columns(6).NumberFormat = "# ?/?"
    .Rows(1).Font.Bold = True
    .Range("A1").Value = "Liste de course fraiche"
    .Range("C1").Value = "Nb de repas"
    .Range("D1").Value = "Proportion"
    .Range("F1").Value = "A acheter !"
    .Range("G1").Value = "Unité"
 End With
End Sub

Private Sub SetRecettes()
 With Worksheets("Recettes")
    .Columns(1).Font.Bold = True
 End With
End Sub


Private Sub SetBaseAliments()
    Call GetDataFromGoogle("BaseAliments", "https://docs.google.com/spreadsheets/d/1Rudp78FjmbWhtPMRLr6ItgIq0p_alzeSQKTVKlkOVA4/edit?usp=sharing")
    With Worksheets("BaseAliments")
        .Range("A1:A3").EntireRow.Delete
        .Range("A1").EntireColumn.Delete
        .Columns(2).NumberFormat = "# ?/?"
    End With
End Sub

Private Sub SetBaseRecettes()
    Call GetDataFromGoogle("BaseRecettes", "https://docs.google.com/spreadsheets/d/1J2PQ1NK6bKJ3sykCZon31PPnAkeOO8qYifyUxvTBDJs/edit?usp=sharing")
    With Worksheets("BaseRecettes")
        .Range("A1:A3").EntireRow.Delete
        .Range("A1").EntireColumn.Delete
        .Columns(3).NumberFormat = "# ?/?"
    End With
End Sub





Private Sub SetMenu()
 With Worksheets("Menu")
    .Columns("A:AZ").ColumnWidth = 30
    
    .Rows("1:7").HorizontalAlignment = xlCenter
    .Rows("1:7").Style = "Calcul"
    .Rows("2:7").Font.Color = RGB(0, 0, 0)
    .Rows(1).Interior.Color = RGB(100, 120, 255)
    .Rows(1).Font.Color = RGB(255, 255, 255)
    .Rows(1).NumberFormat = "dddd"
    .Rows(2).Interior.Color = RGB(200, 210, 255)
    .Rows(2).Font.Italic = True
    .Rows(3).NumberFormat = "dd mmm"
    .Rows(3).Interior.Color = RGB(200, 210, 255)
    .Rows(3).Font.Bold = True
    .Rows(4).Interior.Color = RGB(255, 210, 210)
    .Rows(5).Interior.Color = RGB(255, 250, 200)
    .Rows(6).Interior.Color = RGB(255, 210, 250)
    .Rows(7).Interior.Color = RGB(230, 255, 230)
    
    .Columns(1).Font.Bold = True
    
    .Range("A2").Value = "Activité"
    .Range("A4").Value = "Pti Dej"
    .Range("A5").Value = "Midi"
    .Range("A6").Value = "Goûter"
    .Range("A7").Value = "Dîner"
    
    Dim code As String
    code = "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbCrLf & "    SetMycell (ActiveCell.Value)" & vbCrLf
    code = code & "    Dim n As Integer: n = 7" & vbCrLf & "    While (Range(""A"" & n).Interior.Color <> RGB(255, 255, 255))" & vbCrLf & "        n = n + 1" & vbCrLf & "    Wend" & vbCrLf
    code = code & "If Target.Count = 1 And Not Intersect(Evaluate(""4:"" & n-1), Target) Is Nothing And Intersect([A:A], Target) Is Nothing Then" & vbCrLf & "        Dim lst1 As Range" & vbCrLf & "        Dim lst2 As Range" & vbCrLf & "On Error GoTo ErrorHandler" & vbCrLf & "        With Worksheets(""BaseAliments"")" & vbCrLf & "        Set lst1 = .Range(.Range(""A1"").address, .Range(""A"" & .Rows.Count).End(xlUp).address)" & vbCrLf & "        End With" & vbCrLf & "        With Worksheets(""BaseRecettes"")" & vbCrLf & "        Set lst2 = .Range(.Range(""A1"").address, .Range(""A"" & .Rows.Count).End(xlUp).address)" & vbCrLf & "        End With" & vbCrLf
    code = code & "         If (Me.OLEObjects.Count <> 0) Then AddLineOnObject" & vbCrLf & "        Set Ctrl = ActiveSheet.OLEObjects.Add(ClassType:=""Forms.ComboBox.1"", _" & vbCrLf & "                Link:=False, DisplayAsIcon:=False, Left:=Target.Left, Top:=Target.Top, Width:=165, Height:=16)" & vbCrLf
    code = code & "         With Ctrl" & vbCrLf & "            .Name = ""CB""" & vbCrLf & "            For Each Cell In lst1" & vbCrLf & "                .Object.AddItem Cell.Value" & vbCrLf & "            Next Cell" & vbCrLf & "            For Each Cell In lst2" & vbCrLf & "        If Cell.Value <> """" Then .Object.AddItem Cell.Value" & vbCrLf & "            Next Cell" & vbCrLf & "            .LinkedCell = ActiveCell.address" & vbCrLf & "            .Object.Font.size = 8" & vbCrLf & "            .Activate" & vbCrLf & "            End With" & vbCrLf & "    SendKeys ""%{DOWN}""" & vbCrLf & "    End If" & vbCrLf
    code = code & "Exit Sub" & vbCrLf & "ErrorHandler:" & vbCrLf & "MsgBox (""Il faut creer les Bases de Donnée et les remplir d'au moins une valeur !"")" & vbCrLf & "End Sub" & vbCrLf
    code = code & "Private Sub CB_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)" & vbCrLf & "If KeyCode = 27 Then Worksheets(""Menu"").OLEObjects(""CB"").Object.Value = """" " & vbCrLf & "If KeyCode = 9 Or KeyCode = 13 Or KeyCode = 27 Then AddLineOnObject" & vbCrLf & "End Sub"
    On Error GoTo ErrorHandler
    ActiveWorkbook.VBProject.VBComponents(.CodeName).CodeModule.AddFromString code
    

    End With
    Exit Sub
ErrorHandler:
MsgBox ("Pour pouvoir remettre l'assistant à zero, il faut : " + vbCrLf + "Dans les Options" + vbCrLf + " - Centre de Gestion de la Confidentialité" + vbCrLf + " - Paramètre du Centre de Gestion de la Confidentialité" + vbCrLf + " - Paramètres des macros" + vbCrLf + "il faut cocher Acces approuvé au modèle d'objet du projet VBA" + vbCrLf + "(ou alors rerécuperer l'assistant sur internet)")
End Sub

'Appelé à la validation de la combobox
Public Function AddLineOnObject()
    ListRepartition
    Dim a: a = Worksheets("Menu").OLEObjects("CB").LinkedCell
    If (Range(a) <> "") Then
        a = Cells(Range(a).Row + 1, 1).address(RowAbsolute:=False, ColumnAbsolute:=False)
        If (Range(a).Value <> "" Or Range(a).Interior.Color = RGB(255, 255, 255)) Then
            Range(a).EntireRow.Insert
        End If
        
    Else
         For i = 1 To WorksheetFunction.max(Worksheets("Accueil").Range("A3").Value, 1)
            If Range(Cells(Range(a).Row, i).address(RowAbsolute:=False, ColumnAbsolute:=False)) <> "" Then
                Worksheets("Menu").OLEObjects.Delete
                Exit Function
            End If
         Next
         For i = 2 To WorksheetFunction.max(Worksheets("Accueil").Range("A3").Value, 1)
            If Range(Cells(Range(a).Row - 1, i).address(RowAbsolute:=False, ColumnAbsolute:=False)) <> "" Then
                Worksheets("Menu").OLEObjects.Delete
                Exit Function
            End If
         Next
         Range(Cells(Range(a).Row, 1).address(RowAbsolute:=False, ColumnAbsolute:=False)).EntireRow.Delete
    End If
    Worksheets("Menu").OLEObjects.Delete
End Function

Private Function ListRepartition()
    Dim a As String: a = Worksheets("Menu").OLEObjects("CB").Object.Value
    Dim d As Date: d = Worksheets("Menu").Columns(Worksheets("Menu").OLEObjects("CB").TopLeftCell.Column).Cells(1).Value
    
    If a <> "" Then
        If GetMycell = "" Then
            Addtolist a, d
        Else
            If GetMycell <> a Then
                RemoveFromList d
                Addtolist a, d
            End If
        End If
    Else
        If GetMycell <> "" Then RemoveFromList d
    End If
End Function


Private Function Addtolist(item As String, Optional dat As Date = "00:00:00")
    'On Error GoTo EH
    If dat = "00:00:00" Then Exit Function
    
    Dim ba As Range: Set ba = Worksheets("BaseAliments").Columns(1).Cells.Find(what:=item)
    Dim br As Range: Set br = Worksheets("BaseRecettes").Columns(1).Cells.Find(what:=item)
    
    'Calcul de la zone de course pour le frais
    Dim daterange As Range: Set daterange = GetDateRange(dat)
    Dim c: c = CountInMenu(item, daterange)
        
    'Classement
    If Not ba Is Nothing Then 'Est dans la base Aliments
        If (Worksheets("BaseAliments").Cells(ba.Row, 4).Value = "Sec") Then
            Set b = Sheets("Liste Sec").Columns(2).Cells.Find(what:=item)
            If Not b Is Nothing Then Worksheets("Liste Sec").Cells(b.Row, 3).Value = c Else Call ListSetAlimentSec(ba)
        Else
            Set b = daterange.Find(what:=item)
            If Not b Is Nothing Then Worksheets("Liste Frais").Cells(b.Row, 3).Value = c Else Call ListSetAlimentFrais(ba, dat)
        End If
       
    ElseIf Not br Is Nothing Then 'Est dans la base Recettes
        Dim r As Range: Set r = Sheets("BaseRecettes").Range("B" & br.Row & ":" & "B" & NextRow(Range(br.address), "BaseRecettes"))
        For i = br.Row To NextRow(Range(br.address), "BaseRecettes") 'Pour chacun des ingédients
            If (Worksheets("BaseRecettes").Cells(i, 5).Value = "Sec") Then
                Set b = Sheets("Liste Sec").Columns(2).Cells.Find(what:=r(i) & " - " & item)
                If Not b Is Nothing Then Worksheets("Liste Sec").Cells(b.Row, 3).Value = c Else Call ListSetAlimentSec(r(i), " - " & item)
            Else
                Set b = daterange.Find(what:=r(i) & " - " & item)
                If Not b Is Nothing Then Worksheets("Liste Frais").Cells(b.Row, 3).Value = c Else Call ListSetAlimentFrais(r(i), dat, " - " & item)
            End If
        Next
        
    Else 'N'est pas dans une base
        NewAliment
    End If
    
    Exit Function
EH:
    MyErrorHandler
End Function


Private Sub RemoveFromList(Optional dat As Date = "00:00:00")
    If dat = "00:00:00" Then Exit Sub
    
    Dim Mycell As String: Mycell = GetMycell
    'Calcul de la zone de course pour le frais
    Dim daterange As Range: Set daterange = GetDateRange(dat)
    Dim c As Integer: c = CountInMenu(Mycell, daterange)
    
    If Worksheets("BaseRecettes").Columns(1).Cells.Find(what:=Mycell) Is Nothing Then
        Set todelete = Worksheets("Liste Sec").Columns(2).Cells.Find(what:=Mycell)
        If Not todelete Is Nothing Then
            If c = 0 Then todelete.EntireRow.Delete Else todelete.Next.Value = c
        End If
        Set todelete = daterange.Find(what:=Mycell)
        If Not todelete Is Nothing Then
            If c = 0 Then todelete.EntireRow.Delete Else todelete.Next.Value = c
        End If
    Else
        If c = 0 Then
            Call DelRowsContaining("- " & Mycell, "Liste Sec")
            Call DelRowsContaining("- " & Mycell, "Liste Frais")
        Else
            Set t = Worksheets("Liste Sec").Columns(2).Cells.Find(what:="- " & Mycell, LookAt:=xlPart)
            If Not t Is Nothing Then
                t.Next.Value = c
                Set tt = Worksheets("Liste Sec").Columns(2).Cells.FindNext(t)
                While t <> tt
                    tt.Next.Value = c
                    Set tt = Worksheets("Liste Sec").Columns(2).Cells.FindNext(tt)
                Wend
            End If
            Set t = daterange.Find(what:="- " & Mycell, LookAt:=xlPart)
            If Not t Is Nothing Then
                t.Next.Value = c
                Set tt = daterange.FindNext(t)
                While t <> tt
                    tt.Next.Value = c
                    Set tt = daterange.FindNext(tt)
                Wend
            End If
        End If
    End If
End Sub

Private Function GetDateRange(dat As Date) As Range
    If dat = "00:00:00" Or dat >= Worksheets("Liste Frais").Cells(Rows.Count, "A").End(xlUp).Value Then
        Set GetDateRange = Worksheets("Liste Frais").Range("B" + CStr(Worksheets("Liste Frais").Cells(Rows.Count, "A").End(xlUp).Row) + ":B" + CStr(Worksheets("Liste Frais").Cells(Rows.Count, "B").End(xlUp).Row))
    Else
        Dim dcourse As Range: Set dcourse = Worksheets("Liste Frais").Range("A2")
        If dcourse.Value > dat Then Exit Function
        Do While True
            If Worksheets("Liste Frais").Columns(1).Find(what:="*", after:=dcourse).Value > dat Then
                Set GetDateRange = Worksheets("Liste Frais").Range("B" + CStr(dcourse.Row) + ":B" + CStr(Worksheets("Liste Frais").Columns(1).Find(what:="*", after:=dcourse).Row - 1))
                Exit Do
            Else
                dcourse = Worksheets("Liste Frais").Columns(1).Find(what:="*", after:=dcourse)
            End If
        Loop
    End If
    
End Function

Private Function CountInMenu(item As String, Optional dr As Range = Nothing) As Integer
If dr Is Nothing Then
    CountInMenu = Application.WorksheetFunction.CountIf(Worksheets("Menu").Cells, item)
Else
    Dim colA As String, colB As String
     colA = Col_Letter(Worksheets("Menu").Rows(3).Find(what:=dr.End(xlToLeft)).Column)
     colB = Col_Letter(Worksheets("Menu").Rows(3).Find(Worksheets("Liste Frais").Range("A" + CStr(dr.Row + dr.Rows.Count)).Value).Column - 1)
     CountInMenu = Application.WorksheetFunction.CountIf(Worksheets("Menu").Range(colA + "4:" + colB + CStr(LastMenuRow)), item)
    
End If
End Function

Private Sub Actualiser()
 With Worksheets("Accueil")
    'Mise a Jour du menu
    If (.Range("A3") <> "" And .Range("C3") <> "") Then
        Dim d As Date: d = .Cells(3, 3)
        For i = 0 To .Range("A3")
            Worksheets("Menu").Cells(3, 2 + i).Value = d + i
            Worksheets("Menu").Cells(1, 2 + i).Value = d + i
        Next
        'Mise a jour de la liste Frais
        If (.Range("G3") <> "") Then
            Worksheets("Liste Frais").Cells.ClearContents
            SetFrais
            RepartirDate
            Dim n As Integer: n = LastMenuRow
            
            Dim previous As Long: previous = 2
            While Worksheets("Liste Frais").Range("A2").Value > Worksheets("Menu").Range(Col_Letter(previous) + "3").Value
            previous = previous + 1
            Wend
            Dim dat As Date: dat = Worksheets("Liste Frais").Range("A2")
            Do While dat <> "00:00:00"
                While dat > Worksheets("Menu").Range(Col_Letter(previous) + "3").Value
                    For Each cel In Worksheets("Menu").Range(Col_Letter(previous) + "4:" + Col_Letter(previous) + CStr(n))
                        If Not cel Is Nothing And Not cel.Value = "" Then Addtolist cel.Value, Worksheets("Menu").Columns(cel.Column).Cells(1).Value
                    Next
                previous = previous + 1
                Wend
                If Worksheets("Liste Frais").Columns(1).Find(what:="*", after:=Worksheets("Liste Frais").Columns(1).Find(what:=dat)).Row < Worksheets("Liste Frais").Columns(1).Find(what:=dat).Row Then Exit Do
                dat = Worksheets("Liste Frais").Columns(1).Find(what:="*", after:=Worksheets("Liste Frais").Columns(1).Find(what:=dat)).Value
            Loop
            While Worksheets("Accueil").Range("C3").Value + Worksheets("Accueil").Range("A3").Value > Worksheets("Menu").Range(Col_Letter(previous) + "3").Value
                For Each cel In Worksheets("Menu").Range(Col_Letter(previous) + "4:" + Col_Letter(previous) + CStr(n))
                    If Not cel Is Nothing And Not cel.Value = "" Then Addtolist cel.Value, Worksheets("Menu").Columns(cel.Column).Cells(1).Value
                Next
            previous = previous + 1
            Wend
        End If
    End If
 End With
End Sub


Private Sub RepartirDate()
Dim n As Integer: n = 2
Dim max As Integer: max = 100
For Each cel In Worksheets("Accueil").Range("G:G")
    If IsDate(cel) Then
    Worksheets("Liste Frais").Range("A" + CStr(n)) = cel
    Worksheets("Liste Frais").Range("A" + CStr(n)).NumberFormat = "dd/mm/yyyy"
    n = n + 1
    Else
    max = max - 1
    If max = 0 Then Exit For
    End If
Next cel
End Sub

Private Function ListSetAlimentFrais(a As Range, Optional dat As Date = "00:00:00", Optional ing As String = "")
    If (dat <> "00:00:00") Then Insertlineatdate (dat)
    With Sheets("Liste Frais").Columns(2).Cells.Find(what:="")
                    .Value = a.Value & ing
                    .Next.Value = 1
                    .Next.Next.Value = a.Next.Value
                    .Next.Next.Next.Next.Formula = "= $" & Col_Letter(.Next.Column) & "$" & .Row & " * $" & Col_Letter(.Next.Next.Column) & "$" & .Row & " *Accueil!E3"
                    .Next.Next.Next.Next.Next = a.Next.Next.Value
    End With
End Function

Private Sub Insertlineatdate(dat As Date)
    Dim n As Integer: n = 2
    While Sheets("Liste Frais").Columns(1).Cells(n).Value = "" Or Sheets("Liste Frais").Columns(1).Cells(n).Value <= dat
        n = n + 1
        If (n > 1000) Then Exit Sub
    Wend
    If Sheets("Liste Frais").Range("B" + CStr(n - 1)).Value <> "" Then
        Sheets("Liste Frais").Range(CStr(n) + ":" + CStr(n)).Insert
    End If
End Sub

Private Sub ListSetAlimentSec(a As Range, Optional ing As String = "")
    With Sheets("Liste Sec").Columns(2).Cells.Find(what:="")
                    .Value = a.Value & ing
                    .Next.Value = 1
                    .Next.Next.Value = a.Next.Value
                    .Next.Next.Next.Next.Formula = "= $" & Col_Letter(.Next.Column) & "$" & .Row & " * $" & Col_Letter(.Next.Next.Column) & "$" & .Row & " *Accueil!E3"
                    .Next.Next.Next.Next.Next = a.Next.Next.Value
    End With
End Sub


Private Sub NewAliment()
    Dim rep: rep = MsgBox("Aliment non répertorié" & vbCrLf & "Vous devez l'ajouter si vous voulez qu'il apparaisse dans la liste", vbYesNo, "Ajout")
    If (rep = vbYes) Then
        With Worksheets("BaseAliments").Columns(1).Cells.Find(what:="")
                .Value = InputBox("Nom de l'aliment ?", "Nom", item)
                .Next.Next.Value = InputBox("Unité de mesure ?", "Unité")
                .Next.Value = InputBox("Portion par personne ?", "Portion")
                .Next.Next.Next.Value = InputBox("Sec ou Frais", "Conservation", "Sec/Frais")
        End With
    Else
    Worksheets("Menu").CB.Value = ""
    End If
End Sub

Private Sub DelRowsContaining(s As String, ws As String)
    Set t = Sheets(ws).Columns(2).Cells.Find(what:=s, LookAt:=xlPart)
    While Not t Is Nothing
        t.EntireRow.Delete
        Set t = Sheets(ws).Columns(2).Cells.Find(what:=s, LookAt:=xlPart)
    Wend
End Sub

Private Function LastMenuRow() As Integer
    For Each cel In Worksheets("Menu").Range("A:A")
        If cel.Interior.Color = RGB(255, 255, 255) Then
            LastMenuRow = cel.Row - 1
            Exit For
        End If
    Next cel
End Function

Private Sub MyErrorHandler()
If MsgBox("Je suis désolé, il y a eu une erreur quelque part." + vbCrLf + vbCrLf + "Un envoi de mail auto via oulook est prévu pour que je puisse l'examiner (Si vous n'avez pas outlook, je serais ravi que vous preniez le temps de me l'envoyer vous meme)." + vbCrLf + vbCrLf + "Envoyer le rapport d'erreur ?", vbYesNo, Title:="Rapport d'erreur") = vbYes Then
    Dim path As String: path = Environ("AppData") & "\AI_ErroringFile.xls"
    CopySaveAt path
    
    'Dim fileData As String: fileData = ReadBinaryFile(path)

    Set Mail = CreateObject("Outlook.Application").CreateItem(0)
    With Mail
        .To = "fx73000@yahoo.fr"
        .SentOnBehalfOfName = "aierrorhandler@gmail.com"
        .Subject = "AIError Report"
        .body = "Error" + Err.Number + vbCrLf + Err.Description
        .Attachments.Add path
        .Send
End With
End If
End Sub



Private Sub junk()

    Dim url As String: url = "https://drive.google.com/"
    Set gdrive = New clsGdrive
    gdrive.Token = GetAuthCode("aierrorhandler@gmail.com", "ErrorH1234")
    gdrive.SimpleUpload path
    
    
    
    Dim Cdo_Message As New CDO.Message
    Set Cdo_Message = CreateObject("CDO.Message")

    Set Cdo_Message.Configuration = GetSMTPGmailServerConfig()
    With Cdo_Message
        .To = "fx73000@yahoo.fr"
        .From = "AIErrorHandler@gmail.com"
        .Subject = "Test envoi mail via EXCEL"
        .TextBody = "Bonjour"
        .Send
    End With

    Set Cdo_Message = Nothing
End Sub
