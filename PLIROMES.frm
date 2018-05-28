VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PLIROMES 
   Caption         =   "PLIROMES"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9405
   OleObjectBlob   =   "PLIROMES.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PLIROMES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Me.Hide
ARHIKI.Show vbModeless
End Sub

Private Sub CommandButton2_Click()

Dim PROM_ID As String, CODE_ID As Double, DATE_ID As Date, INVOICE_ID As Double, PAYDATE_ID As Date, VALUE_ID As Double





PROM_ID = TextBox1


CODE_ID = TextBox2

DATE_ID = TextBox3

INVOICE_ID = TextBox4

PAYDATE_ID = TextBox5

VALUE_ID = TextBox6





Worksheets("PLIROMES").Select

Worksheets("PLIROMES").Range("A1").Select

If Worksheets("PLIROMES").Range("A1").Offset(1, 0) <> "" Then

Worksheets("PLIROMES").Range("A1").End(xlDown).Select

End If

ActiveCell.Offset(1, 0).Select

ActiveCell.Value = PROM_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = CODE_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = DATE_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = INVOICE_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = PAYDATE_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = VALUE_ID






TextBox1 = 0
TextBox2 = 0
TextBox3 = 0
TextBox4 = 0
TextBox5 = 0
TextBox6 = 0





ActiveWorkbook.Save



End Sub

