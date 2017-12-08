VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log In"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3915
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   3915
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
rs.Open "Select * from OTHER", con, adOpenDynamic, adLockOptimistic
If StrReverse(rs.Fields("tyre").Value) = Text1.Text Then
Unload Me
Form1.Visible = True
Else
MsgBox "Invalid password.", vbCritical, "Error"
Text1.SetFocus
End If
rs.Close
End Sub

Private Sub Form_Load()
con.Open "dsn=xyz", "other", "laxmi", -1

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
End
End If
If KeyAscii = 13 Then
rs.Open "Select * from OTHER", con, adOpenDynamic, adLockOptimistic
If StrReverse(rs.Fields("tyre").Value) = Text1.Text Then
Unload Me
Form1.Visible = True
Else
MsgBox "Invalid password.", vbCritical, "Error"
Text1.SetFocus
End If
rs.Close
End If

End Sub

