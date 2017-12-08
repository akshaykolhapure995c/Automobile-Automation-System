VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Birthday Reminder"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7035
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
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
      Height          =   2670
      ItemData        =   "bdrm.frx":0000
      Left            =   120
      List            =   "bdrm.frx":0002
      TabIndex        =   3
      Top             =   720
      Width           =   6855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Show"
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Today's Birthdays..."
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
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim f As Boolean
Private Sub Command1_Click()
List1.Clear
rs.Open "Select Full_Name,DOB from Customer_Details", con, adOpenDynamic, adLockOptimistic
While Not rs.EOF
If Left(rs.Fields("DOB").Value, 5) = Left(Label2.Caption, 5) Then
'If rs.Fields("DOB").Value = Label2.Caption Then
List1.AddItem rs.Fields("Full_Name").Value & "      -------------      " & rs.Fields("DOB").Value
f = True
End If
rs.MoveNext
Wend
If Not f Then
MsgBox "No birthdays today...!", vbExclamation, "Message"
End If
rs.Close
End Sub
Private Sub Form_Load()
con.Open "dsn=xyz"
Label2.Caption = Date
End Sub
Private Sub Form_Unload(Cancel As Integer)
con.Close
Form1.Enabled = True
End Sub
