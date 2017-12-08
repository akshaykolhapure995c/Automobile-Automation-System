VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Labour"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4605
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4605
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
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   2535
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
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   2
      Top             =   2040
      Width           =   2535
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
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
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
      Height          =   885
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Full Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Address "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Contact Number"
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
      Top             =   2040
      Width           =   1695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
If Not IsNumeric(Text3.Text) Or Text1.Text = "" Or Text2.Text = "" Then
MsgBox "Fill fields properly.", vbInformation, "Message"
Text1.SetFocus
Else
rs.Open "Select * from Labour_Details where Contact_no='" & Text3.Text & "'", con, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
MsgBox "Duplicate contact number." & vbNewLine & "Not Valid.", vbCritical, "Message"
rs.Close
Exit Sub
Else
GoTo dn
End If
dn:
rs.Close
rs.Open "Select * from Labour_Details", con, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields("full_name").Value = Text1.Text
rs.Fields("Address").Value = Text2.Text
rs.Fields("Contact_no").Value = Text3.Text
rs.Update
MsgBox "Added successfully.", vbInformation, "Message"
rs.Close
Unload Me
End If
End Sub
Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub
Private Sub Form_Load()
con.Open "dsn=xyz", "other", "laxmi", -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
Form1.Enabled = True
End Sub

Private Sub Text2_GotFocus()
If IsNumeric(Text1.Text) Or Text1.Text = "" Then
MsgBox "Enter proper name.", vbInformation, "Message"
Text1.SetFocus
End If
End Sub

'Private Sub Text3_GotFocus()
'If Not IsNumeric(Text3.Text) Then
'MsgBox "Enter proper contact number.", vbInformation, "Message"
'End If
'If Text2.Text = "" Then
'MsgBox "Enter proper contact number.", vbInformation, "Message"
'Text2.SetFocus
'End If

'End Sub
Private Sub Text3_Change()

End Sub
