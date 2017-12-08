VERSION 5.00
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   4560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8895
   LinkTopic       =   "Form8"
   ScaleHeight     =   4560
   ScaleWidth      =   8895
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   6480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   6735
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Form_Load()
Dim a, b As String
con.Open "dsn=xyz", "other", "laxmi", -1
rs.Open "Select * from Customer_Details ", con, adOpenDynamic, adLockOptimistic
Text1.Text = Format(Date, "dd-mm-yyyy")
Text2.Text = Left(Text1.Text, 5)

 While Not rs.EOF
MsgBox Left(rs.Fields("DOB").Value, 5)
If Text2.Text = Left(rs.Fields("DOB").Value, 5) Then
 List1.AddItem rs.Fields("DOB").Value
  End If
rs.MoveNext
 Wend
 End Sub
