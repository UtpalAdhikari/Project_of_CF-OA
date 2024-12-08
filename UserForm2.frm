VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   5550
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   12480
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Sheet1.Activate
Cells(Rows.Count, "A").End(xlUp).Offset(1, 0).Value = TextBox1.Value
Cells(Rows.Count, "B").End(xlUp).Offset(1, 0).Value = TextBox2.Value
Cells(Rows.Count, "C").End(xlUp).Offset(1, 0).Value = ComboBox1.Value
Cells(Rows.Count, "D").End(xlUp).Offset(1, 0).Value = ComboBox2.Value
Cells(Rows.Count, "E").End(xlUp).Offset(1, 0).Value = TextBox3.Value
Cells(Rows.Count, "F").End(xlUp).Offset(1, 0).Value = ComboBox3.Value
Cells(Rows.Count, "G").End(xlUp).Offset(1, 0).Value = ComboBox4.Value
Cells(Rows.Count, "H").End(xlUp).Offset(1, 0).Value = TextBox4.Value
Cells(Rows.Count, "I").End(xlUp).Offset(1, 0).Value = TextBox5.Value
Cells(Rows.Count, "J").End(xlUp).Offset(1, 0).Value = TextBox6.Value

If CheckBox1.Value = True Then
Cells(Rows.Count, "K").End(xlUp).Offset(1, 0).Value = "Yes"
Else
Cells(Rows.Count, "K").End(xlUp).Offset(1, 0).Value = "No"
End If

Me.Hide
UserForm3.Show
End Sub

Private Sub SpinButton1_Change()
TextBox2.Value = SpinButton1.Value
End Sub

Private Sub SpinButton2_Change()
TextBox4.Value = SpinButton2.Value
End Sub
