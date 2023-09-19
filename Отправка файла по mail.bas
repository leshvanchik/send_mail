Attribute VB_Name = "Module1"
Option Explicit
Sub Granica()

Dim myBook, loadBook As String
Dim k As Long
Dim i As Date
i = Date
Dim a, b, z, y As String
z = ThisWorkbook.Worksheets("Титул").Range("N9")
y = ThisWorkbook.Worksheets("Титул").Range("N12")
a = z & "02_3 Границы 11-00 " & i & ".xlsx"
b = z & "02_4 Мониторинг размещения 11-00 " & i & ".xlsx"

ThisWorkbook.Worksheets("Титул").Cells(18, 10) = "Размещено в обсерваторе «Балтиец»            (занято мест на " & i & ")"

    myBook = ThisWorkbook.Name
    loadBook = Dir(a)
    GetObject (a)
    k = Workbooks(loadBook).Worksheets("Для диаграммы").Cells(Rows.Count, 5).End(xlUp).Row
    Workbooks(myBook).Worksheets("Границы").Range("B5:Q6") = Workbooks(loadBook).Worksheets(1).Range("B8:Q9").Value
    Workbooks(myBook).Worksheets("Границы").Range("B38:D44") = Workbooks(loadBook).Worksheets("Для диаграммы").Range("E" & k - 6 & ":G" & k + 1).Value
    Workbooks(myBook).Worksheets("Границы").Range("F38:G44") = Workbooks(loadBook).Worksheets("Для диаграммы").Range("H" & k - 6 & ":I" & k + 1).Value
    Workbooks(loadBook).Close (False)

    myBook = ThisWorkbook.Name
    loadBook = Dir(b)
    GetObject (b)
    k = Workbooks(loadBook).Worksheets("данные").Cells(Rows.Count, 9).End(xlUp).Row
    Workbooks(myBook).Worksheets("Титул").Range("J19") = Workbooks(loadBook).Worksheets("данные").Range("I" & k).Value
    Workbooks(myBook).Worksheets("Титул").Range("K19") = Workbooks(loadBook).Worksheets(1).Range("A5").Value
    Application.DisplayAlerts = 0
    Workbooks(myBook).Worksheets("Балтиец").Delete
    Application.DisplayAlerts = 1
    Workbooks(loadBook).Worksheets("Балтиец").Copy After:=Workbooks(myBook).Sheets(2)
    Workbooks(loadBook).Close (False)
    
    With Workbooks(myBook).Worksheets("Титул")
    .Range("L19").Formula = "=TRIM(LEFTB(SUBSTITUTE(R19C[-1],"","",REPT("" "",999)),999))"
    .Range("L19") = .Range("L19").Value
    .Range("K19") = Extract_Number_from_Text(.Range("L19"), 0)
    .Range("L19").Delete
    .Select
    End With

'    Workbooks.Add
'    ActiveWorkbook.Worksheets(1).Range("A1:P4") = Workbooks("СверхГраница").Worksheets("Загрузка1").Range("A1:P4").Value
'    ActiveWorkbook.SaveAs Filename:=y & " " & i & ".xlsx"
'    ActiveWorkbook.Close (True)

    Dim OutApp As Object
    Dim OutMail As Object
    Dim IsOultOpen As Boolean
    Dim cell As Range
     
    Application.ScreenUpdating = False
    On Error Resume Next
    Set OutApp = GetObject(, "Outlook.Application")   'запускаем Outlook в скрытом режиме
        If Err = 0 Then
            IsOultOpen = True
            Else
            Err.Clear
            Set OutApp = CreateObject("Outlook.Application")
    End If
    OutApp.Session.Logon ' "a.tihonov@iac.spb.ru", "infocenter@iac.spb.ru", False, True
    Set OutMail = OutApp.CreateItem(0)   'создаем новое сообщение
    If Err.Number <> 0 Then Set OutApp = Nothing: Set OutMail = Nothing: Exit Sub
    'заполняем поля сообщения
    With OutMail
        .To = ThisWorkbook.Worksheets(1).Range("N2").Value
        .Subject = ThisWorkbook.Worksheets(1).Range("N3").Value
        .Body = ThisWorkbook.Worksheets(1).Range("N4").Value
        .Attachments.Add ThisWorkbook.Worksheets(1).Range("N5").Value
        'команду Send можно заменить на Display, чтобы посмотреть сообщение перед отправкой
        .Display
    End With
 
    If IsOultOpen = False Then OutApp.Quit
    Set OutApp = Nothing: Set OutMail = Nothing
    DoEvents
 
End Sub
     
Function Extract_Number_from_Text(sWord As String, Optional Metod As Integer)
'sWord = ссылка на ячейку или непосредственно текст
'Metod = 0 – числа
'Metod = 1 – текст
    Dim sSymbol As String, sInsertWord As String
    Dim i As Integer
 
    If sWord = "" Then Extract_Number_from_Text = "Нет данных!": Exit Function
    sInsertWord = ""
    sSymbol = ""
    For i = 1 To Len(sWord)
        sSymbol = Mid(sWord, i, 1)
        If Metod = 1 Then
            If Not LCase(sSymbol) Like "*[0-9]*" Then
                If (sSymbol = "," Or sSymbol = "." Or sSymbol = " ") And i > 1 Then
                    If Mid(sWord, i - 1, 1) Like "*[0-9]*" And Mid(sWord, i + 1, 1) Like "*[0-9]*" Then
                        sSymbol = ""
                    End If
                End If
                sInsertWord = sInsertWord & sSymbol
            End If
        Else
            If LCase(sSymbol) Like "*[0-9]*" Then
                If LCase(sSymbol) Like "*[.,]*" And i > 1 Then
                    If Not Mid(sWord, i - 1, 1) Like "*[0-9]*" Or Not Mid(sWord, i + 1, 1) Like "*[0-9]*" Then
                        sSymbol = ""
                    End If
                End If
                sInsertWord = sInsertWord & sSymbol
            End If
        End If
    Next i
    Extract_Number_from_Text = sInsertWord
End Function
