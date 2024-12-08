VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4660
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7110
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
      
       
If TextBox1.Value = "utpal" And TextBox2.Value = "12345" Then
Me.Hide
UserForm2.Show
Else
MsgBox ("Incorrect id or password")
End If


End Sub



Private Sub TextBox1_Change()

End Sub
