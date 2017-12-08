VERSION 5.00
Begin VB.Form Form19 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete Job"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4950
   LinkTopic       =   "Form19"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Delete"
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Job No"
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
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
If Combo1.Text = "" Then
MsgBox "Select proper job no.", vbInformation, "Message"
Else
If MsgBox("Delete order ?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
con.Execute "delete * from Order_dtls where jbs='" & Combo1.Text & "'"
MsgBox "Delete.", vbInformation, "Message"
End If
End If
End Sub
Private Sub Form_Load()
con.Open "dsn=xyz"
Combo1.Clear
rs.Open "Select jbs from Order_dtls", con, adOpenDynamic, adLockOptimistic
While Not rs.EOF
Combo1.AddItem rs.Fields("jbs").Value
rs.MoveNext
Wend
rs.Close
End Sub
Private Sub Form_Unload(Cancel As Integer)
con.Close
Form1.Enabled = True
End Sub
