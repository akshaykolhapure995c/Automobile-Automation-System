VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Job"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4005
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Refresh"
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Addnew"
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   975
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
      Left            =   1440
      TabIndex        =   3
      Top             =   840
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
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Charges"
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
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Job"
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
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()

If Not IsNumeric(Text2.Text) Or Text2.Text = "" Or Text1.Text = "" Or IsNumeric(Text1.Text) Then
MsgBox "Enter proper job type.", vbInformation, "Message"
Text2.SetFocus
Else
rs.Open "Select * from Job_details", con, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields("Type_of_job").Value = Val(Text1.Text)
rs.Fields("Charges").Value = Val(Text2.Text)
rs.Update
MsgBox "Added successfully.", vbInformation, "Message"
rs.Close
Me.ref

End If

End Sub
Private Sub Command1_GotFocus()
If Not IsNumeric(Text2.Text) Or Text2.Text = "" Then
MsgBox "Enter proper amount charge.", vbInformation, "Message"
Text2.SetFocus
End If
End Sub
Private Sub Command2_Click()
Me.ref
End Sub
Private Sub Form_Load()
con.Open "dsn=xyz", "other", "laxmi", -1
Me.BackColor = &HE0E0E0
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
Form1.Enabled = True
End Sub

Private Sub Text2_GotFocus()
If IsNumeric(Text1.Text) Or Text1.Text = "" Then
MsgBox "Enter proper job type.", vbInformation, "Message"
Text1.SetFocus
End If
End Sub

Sub ref()
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End Sub
