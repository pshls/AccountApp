VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EXODA 
   Caption         =   "EXODA"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9555
   OleObjectBlob   =   "EXODA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EXODA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then TextBox7 = TextBox6 * 24 / 100
If CheckBox1.Value = False Then TextBox7 = 0
End Sub

Private Sub CheckBox2_Click()

If CheckBox2.Value = True Then TextBox8 = TextBox6 * 20 / 100
If CheckBox2.Value = False Then TextBox8 = 0

End Sub






Private Sub CheckBox3_Click()
If CheckBox3.Value = True Then TextBox9 = "GRAFEIOU" Else TextBox9 = ""


End Sub

Private Sub CommandButton1_Click()
Dim PROM_ID As String, CODE_ID As Double, DATE_ID As Date, INVOICE_ID As Double, PERIGRAFI_ID As String, VALUE_ID As Double, VAT_ID As Double, TAX_ID As Double, PELATIS_ID As String, PELCODE_ID As Double





PROM_ID = TextBox1


CODE_ID = TextBox2

DATE_ID = TextBox3

INVOICE_ID = TextBox4

PERIGRAFI_ID = TextBox5

VALUE_ID = TextBox6


VAT_ID = TextBox7

TAX_ID = TextBox8

PELATIS_ID = TextBox9

PELCODE_ID = TextBox10


Worksheets("EXODA").Select

Worksheets("EXODA").Range("A1").Select

If Worksheets("EXODA").Range("A1").Offset(1, 0) <> "" Then

Worksheets("EXODA").Range("A1").End(xlDown).Select

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

ActiveCell.Value = PERIGRAFI_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = VALUE_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = VAT_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = TAX_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = VALUE_ID + VAT_ID - TAX_ID


ActiveCell.Offset(0, 1).Select

ActiveCell.Value = PELATIS_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = PELCODE_ID




TextBox1 = 0
TextBox2 = 0
TextBox3 = 0
TextBox4 = 0
TextBox5 = 0
TextBox6 = 0
TextBox7 = 0
TextBox8 = 0
TextBox9 = 0
TextBox10 = 0


CheckBox1.Value = False
CheckBox2.Value = False
CheckBox3.Value = False

ActiveWorkbook.Save

End Sub



Private Sub CommandButton2_Click()


TextBox1 = Worksheets("EXODA").Range("A1").End(xlDown)
TextBox2 = Worksheets("EXODA").Range("B1").End(xlDown)
TextBox3 = Worksheets("EXODA").Range("C1").End(xlDown)
TextBox4 = Worksheets("EXODA").Range("D1").End(xlDown)

End Sub






Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub


Private Sub CommandButton3_Click()


Me.Hide
ARHIKI.Show vbModeless
    



End Sub

Private Sub CommandButton4_Click()

Dim PROM_ID As String, CODE_ID As Double, DATE_ID As Date, INVOICE_ID As Double, PERIGRAFI_ID As String, VALUE_ID As Double, VAT_ID As Double, TAX_ID As Double, PELATIS_ID As String, PELCODE_ID As Double





PROM_ID = TextBox1


CODE_ID = TextBox2

DATE_ID = TextBox3

INVOICE_ID = TextBox4

PERIGRAFI_ID = TextBox5

VALUE_ID = TextBox6


VAT_ID = TextBox7

TAX_ID = TextBox8

PELATIS_ID = TextBox9

PELCODE_ID = TextBox10


Worksheets("EXODA").Select

Worksheets("EXODA").Range("A1").Select

If Worksheets("EXODA").Range("A1").Offset(1, 0) <> "" Then

Worksheets("EXODA").Range("A1").End(xlDown).Select

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

ActiveCell.Value = PERIGRAFI_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = VALUE_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = VAT_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = TAX_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = VALUE_ID + VAT_ID - TAX_ID


ActiveCell.Offset(0, 1).Select

ActiveCell.Value = PELATIS_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = PELCODE_ID













Dim PROMA_ID As String, CODEA_ID As Double, DATEA_ID As Date, INVOICEA_ID As Double, VALUEA_ID As Double





PROMA_ID = TextBox1


CODEA_ID = TextBox2

DATEA_ID = TextBox3

INVOICEA_ID = TextBox4



VALUEA_ID = TextBox6





Worksheets("PLIROMES").Select

Worksheets("PLIROMES").Range("A1").Select

If Worksheets("PLIROMES").Range("A1").Offset(1, 0) <> "" Then

Worksheets("PLIROMES").Range("A1").End(xlDown).Select

End If

ActiveCell.Offset(1, 0).Select

ActiveCell.Value = PROMA_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = CODEA_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = DATEA_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = INVOICEA_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = DATEA_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = VALUEA_ID + VAT_ID - TAX_ID








TextBox1 = 0
TextBox2 = 0
TextBox3 = 0
TextBox4 = 0
TextBox5 = 0
TextBox6 = 0
TextBox7 = 0
TextBox8 = 0
TextBox9 = 0
TextBox10 = 0


CheckBox1.Value = False
CheckBox2.Value = False
CheckBox3.Value = False

ActiveWorkbook.Save




End Sub

Private Sub Label11_Click()

End Sub

Private Sub Label9_Click()

End Sub
