VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8196.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   16116
   OleObjectBlob   =   "progman.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Dim r As String
    Dim rr As String
    Dim t As Long
    Dim y As Long
    
    Dim a
    On Error GoTo err2
    rr = Me.TextBox2.Text
    Me.ListBox1.AddItem (rr)
    
    
    
    Open "out.data" For Output As #1
    rr = ""
    y = Me.ListBox1.ListCount
    For t = 0 To y
        
        Print #1, Me.ListBox1.List(t)
       
        
    Next
err2:
On Error GoTo err3
Close #1
err3:
    
    
    
    
    
End Sub

Private Sub ListBox1_Click()
Dim a

a = Shell(Me.ListBox1.List(Me.ListBox1.ListIndex))
End Sub

Private Sub UserForm_Activate()
On Error GoTo err1
MsgBox (CurDir())
Open "out.data" For Input As #1
    rr = ""
    While Not (EOF(1))
        Line Input #1, r
        Me.ListBox1.AddItem (r)
    Wend
err1:
On Error GoTo err4
Close #1
err4:

End Sub
