VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} KARTELAPROM 
   Caption         =   "KARTELAPROM"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7950
   OleObjectBlob   =   "KARTELAPROM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "KARTELAPROM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Click()

End Sub

Private Sub CheckBox2_Click()

End Sub

Private Sub CommandButton2_Click()
Me.Hide
ARHIKI.Show vbModeless
    

End Sub

Private Sub Label6_Click()

End Sub

Private Sub TextBox1_Change()

End Sub



Private Sub CommandButton1_Click()

Dim test As Worksheet



    Sheets("EXODA").Copy After:=Sheets(Sheets.Count)
    Set test = ActiveSheet
    test.Name = "TEST1"





Dim RngOne As Range, cell As Range
Dim LastCell As Long
Dim arrList() As String, lngCnt As Long
Dim FMDATE As Date, TODATE As Date


FMDATE = TextBox3.Value
TODATE = TextBox4.Value


With test
    LastCell = .Range("A" & test.Rows.Count).End(xlUp).Row
    Set RngOne = .Range("A2:A" & LastCell)
End With


'load values into an array
lngCnt = 0
For Each cell In RngOne
    ReDim Preserve arrList(lngCnt)
    arrList(lngCnt) = cell.Text
    lngCnt = lngCnt + 1
Next


With test

    If .FilterMode Then .ShowAllData
    
    
  
If TextBox1.Value <> vbNullString Then
    .Range("A:K").AutoFilter Field:=1, Criteria1:=TextBox1.Value, Operator:=xlFilterValues
    
    End If
    
 

   
  
   
    If TextBox5.Value <> vbNullString Then
    .Range("A:K").AutoFilter Field:=2, Criteria1:=TextBox5.Value, Operator:=xlFilterValues
   
End If
   
   
   
   
   
   If TextBox3.Value & TextBox4.Value <> vbNullString Then
    .Range("A:K").AutoFilter Field:=3, Criteria1:=">=" & FMDATE, Criteria2:="<=" & TODATE, Operator:=xlFilterValues
   
End If
   
 
 
 
 
 If CheckBox1.Value = True Then
  .Range("A:K").AutoFilter Field:=7, Criteria1:=">0", Operator:=xlFilterValues
   End If
   
   
    If CheckBox2.Value = True Then
  .Range("A:K").AutoFilter Field:=8, Criteria1:=">0", Operator:=xlFilterValues
   End If
   
   
   
   
   
End With







    Dim oRow As Range, rng As Range
    Dim myRows As Range
    With test
        Set myRows = Intersect(.Range("A:A").EntireRow, .UsedRange)
        If myRows Is Nothing Then Exit Sub
    End With

    For Each oRow In myRows.Columns(1).Cells
        If oRow.EntireRow.Hidden Then
            If rng Is Nothing Then
                Set rng = oRow
            Else
                Set rng = Union(rng, oRow)
            End If
        End If
    Next

    If Not rng Is Nothing Then rng.EntireRow.Delete







test.AutoFilterMode = False









Dim testA As Worksheet



    Sheets("PLIROMES").Copy After:=Sheets(Sheets.Count)
    Set testA = ActiveSheet
    testA.Name = "TEST2"



With testA
    LastCell = .Range("A" & testA.Rows.Count).End(xlUp).Row
    Set RngOne = .Range("A2:A" & LastCell)
End With


'load values into an array
lngCnt = 0
For Each cell In RngOne
    ReDim Preserve arrList(lngCnt)
    arrList(lngCnt) = cell.Text
    lngCnt = lngCnt + 1
Next


With testA

    If .FilterMode Then .ShowAllData
    
    
  
If TextBox1.Value <> vbNullString Then
    .Range("A:F").AutoFilter Field:=1, Criteria1:=TextBox1.Value, Operator:=xlFilterValues
    
    End If
    
 

   
  
   
    If TextBox5.Value <> vbNullString Then
    .Range("A:F").AutoFilter Field:=2, Criteria1:=TextBox5.Value, Operator:=xlFilterValues
   
End If
   
   
   
   
   
 
 
   
   
   
   
End With






Dim oRowA As Range, rngA As Range
    Dim myRowsA As Range
    
    
    
  
    With testA
        Set myRowsA = Intersect(.Range("A:A").EntireRow, .UsedRange)
        If myRowsA Is Nothing Then Exit Sub
    End With

    For Each oRowA In myRowsA.Columns(1).Cells
        If oRowA.EntireRow.Hidden Then
            If rngA Is Nothing Then
                Set rngA = oRowA
            Else
                Set rngA = Union(rngA, oRowA)
            End If
        End If
    Next

    If Not rngA Is Nothing Then rngA.EntireRow.Delete







testA.AutoFilterMode = False




















   Sheets.Add After:=Sheets(Sheets.Count)
   Dim wks As Worksheet
   Dim lastrow As Long
   Dim lastrow3 As Long
   
   Set wks = Sheets(Sheets.Count)

   wks.Name = TextBox1.Value & "_" & TextBox5.Value & Format(Now(), "_yyyy-mm-dd_hh-mm-ss")


   
   
   
   With Sheets("TEST1")

    
    lastrow = .Range("A" & .Rows.Count).End(xlUp).Row
    
    .Range("F" & lastrow + 1).Value = WorksheetFunction.Sum(.Range("F1:F" & lastrow))
    .Range("G" & lastrow + 1).Value = WorksheetFunction.Sum(.Range("G1:G" & lastrow))
    .Range("H" & lastrow + 1).Value = WorksheetFunction.Sum(.Range("H1:H" & lastrow))
    .Range("I" & lastrow + 1).Value = WorksheetFunction.Sum(.Range("I1:I" & lastrow))

    .Range("A1:K" & lastrow + 1).Copy wks.Range("A" & wks.Rows.Count).End(xlUp)

   End With

   
   
   
   



With Sheets("TEST2")

    lastrow3 = .Range("A" & .Rows.Count).End(xlUp).Row

    .Range("F" & lastrow3 + 1).Value = WorksheetFunction.Sum(.Range("F1:F" & lastrow3))
    .Range("A1:F" & lastrow3 + 1).Copy wks.Range("A" & wks.Rows.Count).End(xlUp).Offset(3)

   End With

'Stopping Application Alerts
Application.DisplayAlerts = False

Sheets("TEST1").Delete
Sheets("TEST2").Delete

'Enabling Application alerts once we are done with our task
Application.DisplayAlerts = True




End Sub










Private Sub TextBox2_Change()

End Sub

Private Sub TextBox3_Change()

End Sub
