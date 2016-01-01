VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formVardas 
   Caption         =   "Vardas"
   ClientHeight    =   1065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4095
   OleObjectBlob   =   "formVardas.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formVardas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Call EnterName(TextBox1.Text)
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call EnterName(TextBox1.Text)
    End If
End Sub

Function EnterName(ByVal strText As String)
    If strText = "" Then
        shResults.Cells(lrow, 3) = "-"
        Else
            lrow = 3
            Do Until IsEmpty(shResults.Cells(lrow, 2))
                lrow = lrow + 1
            Loop
            shResults.Cells(lrow, 3) = strText
            formVardas.Hide
    End If
End Function
