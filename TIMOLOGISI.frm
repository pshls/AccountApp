VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TIMOLOGISI 
   Caption         =   "TIMOLOGISI"
   ClientHeight    =   10245
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10485
   OleObjectBlob   =   "TIMOLOGISI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TIMOLOGISI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CheckBox1_Click()
If CheckBox1.Value = True Then TextBox6 = TextBox5 * 24 / 100
If CheckBox1.Value = False Then TextBox6 = 0
End Sub

Private Sub CheckBox2_Click()

If CheckBox2.Value = True Then TextBox7 = TextBox5 * 20 / 100
If CheckBox2.Value = False Then TextBox7 = 0

End Sub






Private Sub CommandButton1_Click()
Dim PELATIS_ID As String, CODE_ID As Double, DATE_ID As Date, INVOICE_ID As Double, PERIGRAFI_ID As String, VALUE_ID As Double, VAT_ID As Double, TAX_ID As Double





PELATIS_ID = TextBox2


CODE_ID = TextBox1

DATE_ID = TextBox3

INVOICE_ID = TextBox4

PERIGRAFI_ID = TextBox9

VALUE_ID = TextBox5


VAT_ID = TextBox6

TAX_ID = TextBox7



Worksheets("PELATES").Select

Worksheets("PELATES").Range("A1").Select

If Worksheets("PELATES").Range("A1").Offset(1, 0) <> "" Then

Worksheets("PELATES").Range("A1").End(xlDown).Select

End If

ActiveCell.Offset(1, 0).Select

ActiveCell.Value = PELATIS_ID

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




TextBox1 = 0
TextBox2 = 0
TextBox3 = 0
TextBox4 = 0
TextBox5 = 0
TextBox6 = 0
TextBox7 = 0
TextBox9 = 0


CheckBox1.Value = False
CheckBox2.Value = False

ActiveWorkbook.Save

End Sub



Private Sub CommandButton2_Click()


TextBox2 = Worksheets("PELATES").Range("A1").End(xlDown)
TextBox1 = Worksheets("PELATES").Range("B1").End(xlDown)
TextBox3 = Worksheets("PELATES").Range("C1").End(xlDown)
TextBox4 = Worksheets("PELATES").Range("D1").End(xlDown)

End Sub






Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub


Private Sub CommandButton3_Click()


Me.Hide
ARHIKI.Show vbModeless
    



End Sub



Private Sub CommandButton4_Click()
Dim PELATIS_ID As String, CODE_ID As Double, DATE_ID As Date, INVOICE_ID As Double, PERIGRAFI_ID As String, VALUE_ID As Double, VAT_ID As Double, TAX_ID As Double





PELATIS_ID = TextBox2


CODE_ID = TextBox1

DATE_ID = TextBox3

INVOICE_ID = TextBox4

PERIGRAFI_ID = TextBox9

VALUE_ID = TextBox5


VAT_ID = TextBox6

TAX_ID = TextBox7



Worksheets("PELATES").Select

Worksheets("PELATES").Range("A1").Select

If Worksheets("PELATES").Range("A1").Offset(1, 0) <> "" Then

Worksheets("PELATES").Range("A1").End(xlDown).Select

End If

ActiveCell.Offset(1, 0).Select

ActiveCell.Value = PELATIS_ID

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







Worksheets("PARAGELIES").Select

Worksheets("PARAGELIES").Range("A1").Select

If Worksheets("PARAGELIES").Range("A1").Offset(1, 0) <> "" Then

Worksheets("PARAGELIES").Range("A1").End(xlDown).Select

End If

ActiveCell.Offset(1, 0).Select

ActiveCell.Value = PELATIS_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = CODE_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = DATE_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = INVOICE_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = PERIGRAFI_ID

ActiveCell.Offset(0, 1).Select

ActiveCell.Value = -VALUE_ID












TextBox1 = 0
TextBox2 = 0
TextBox3 = 0
TextBox4 = 0
TextBox5 = 0
TextBox6 = 0
TextBox7 = 0
TextBox9 = 0


CheckBox1.Value = False
CheckBox2.Value = False

ActiveWorkbook.Save
End Sub
