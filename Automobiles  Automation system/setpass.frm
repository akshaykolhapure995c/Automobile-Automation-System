VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4545
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Change"
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox Text3 
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
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox Text2 
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
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
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
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "Confirm password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "New password"
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
      TabIndex        =   4
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Old password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Enter old password. ", vbInformation, "Message"
Text1.SetFocus
Else
If Text2.Text <> Text3.Text Then
MsgBox "Comfirmation failed.", vbCritical, "Message"
Text3.SetFocus
Else
rs.Open "Select * from OTHER", con, adOpenDynamic, adLockOptimistic
rs.Fields("Tyre").Value = StrReverse(Text2.Text)
rs.Update
MsgBox "Changed successfully", vbExclamation, "Message"
rs.Close
Unload Me
End If
End If
End Sub
Private Sub Form_Load()
con.Open "dsn=xyz", "other", "laxmi", -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
Form1.Enabled = True
End Sub

Private Sub Text2_GotFocus()
rs.Open "Select * from OTHER", con, adOpenDynamic, adLockOptimistic
If rs.Fields("Tyre").Value <> StrReverse(Text1.Text) Then
MsgBox "Enter valid old password.", vbCritical, "Message"
Text1.SetFocus
End If
rs.Close
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Text1.Text = "" Then
MsgBox "Enter old password. ", vbInformation, "Message"
Text1.SetFocus
Else
If Text2.Text <> Text3.Text Then
MsgBox "Comfirmation failed.", vbCritical, "Message"
Text3.SetFocus
Else
rs.Open "Select * from OTHER", con, adOpenDynamic, adLockOptimistic
rs.Fields("Tyre").Value = StrReverse(Text2.Text)
rs.Update
MsgBox "Changed successfully", vbExclamation, "Message"
rs.Close
Unload Me
End If
End If
End If
End Sub
