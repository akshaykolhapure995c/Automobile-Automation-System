VERSION 5.00
Begin VB.Form Form21 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Job Type"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5415
   LinkTopic       =   "Form21"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
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
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Edit"
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
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
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
      ItemData        =   "edit_job.frx":0000
      Left            =   1680
      List            =   "edit_job.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   3615
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
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label4 
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
      TabIndex        =   6
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Type Of Job"
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
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Combo1_Click()
rs.Open "Select * from Job_details where Type_of_job='" & Combo1.Text & "'", con, adOpenDynamic, adLockOptimistic
Text3.Text = rs.Fields("Charges").Value
rs.Close
End Sub
Private Sub Command1_Click()
If Combo1.Text = "" Or Text3.Text = "" Or Not IsNumeric(Text3.Text) Then
MsgBox "Fill fields properly.", vbInformation, "Message"
Me.ref
Else
rs.Open "Select * from Job_details where Type_of_job='" & Combo1.Text & "'", con, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
rs.Fields("Type_of_job").Value = Combo1.Text
rs.Fields("Charges").Value = Text3.Text
rs.Update
MsgBox "Edited Successfully.", vbInformation, "Message"
'Me.loadJobs
Me.ref
Else
MsgBox "Do not change contents of job type." & vbNewLine & _
"Just Edit charges.", vbInformation, "Message"
Me.ref
End If
rs.Close
End If
End Sub
Private Sub Command2_Click()
Me.ref
End Sub
Private Sub Command3_Click()
If Combo1.Text = "" Or Text3.Text = "" Then
MsgBox "Fill fields properly.", vbInformation, "Message"
Me.ref
Else
If MsgBox("Sure to delete ?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
con.Execute "Delete * from Job_details where Type_of_job='" & Combo1.Text & "'"
MsgBox "Deleted Successfully.", vbInformation, "Message"
Me.loadJobs
Me.ref
End If
End If
End Sub
Private Sub Form_Load()
con.Open "dsn=xyz"
Me.loadJobs
End Sub
Private Sub Form_Unload(Cancel As Integer)
con.Close
Form1.Enabled = True
End Sub
Sub ref()
Combo1.Text = ""
Text3.Text = ""
Combo1.SetFocus
End Sub
Sub loadJobs()
rs.Open "Select * from Job_details", con, adOpenDynamic, adLockOptimistic
Combo1.Clear
While Not rs.EOF
Combo1.AddItem rs.Fields("Type_of_job").Value
rs.MoveNext
Wend
rs.Close
End Sub
