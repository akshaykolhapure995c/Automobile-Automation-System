VERSION 5.00
Begin VB.Form Form10 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Labour"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4950
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   3
      Top             =   1800
      Width           =   2895
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Edit"
      Enabled         =   0   'False
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
      TabIndex        =   5
      Top             =   3720
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
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   4
      Top             =   3120
      Width           =   2895
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
      Height          =   270
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   2895
   End
   Begin VB.TextBox Text4 
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
      Height          =   270
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Search"
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1095
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
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Address"
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
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
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
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Not IsNumeric(Text3.Text) Or IsNumeric(Text1.Text) Then
MsgBox "Fill fields properly.", vbInformation, "Message"
Else
rs.Open "Select * from Labour_Details where Contact_no='" & Text4.Text & "'", con, adOpenDynamic, adLockOptimistic
rs.Fields("full_name").Value = Text1.Text
rs.Fields("Address").Value = Text2.Text
rs.Fields("Contact_no").Value = Text3.Text
rs.Update
MsgBox "Edited Successfully.", vbInformation, "Message"
rs.Close
Command1.Enabled = False
Unload Me
End If
End Sub
Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Command1.Enabled = False
Text4.SetFocus
End Sub
Private Sub Form_Load()
con.Open "dsn=xyz"
End Sub
Private Sub Command3_Click()
If Not IsNumeric(Text4.Text) Then
MsgBox "Enter proper contact number.", vbInformation, "Message"
Text4.SetFocus
Else
Command1.Enabled = True
rs.Open "Select * from Labour_Details where Contact_no='" & Text4.Text & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF Then
MsgBox "Record not found.", vbCritical, "Message"
Text4.SetFocus
Else
Text1.Text = rs.Fields("full_name").Value
Text2.Text = rs.Fields("Address").Value
Text3.Text = rs.Fields("Contact_no").Value
Text1.SetFocus
End If
rs.Close
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
Form1.Enabled = True
End Sub

Private Sub Text3_Change()
'If Not IsNumeric(Text3.Text) Then
'MsgBox "Enter proper contact number.", vbInformation, "Message"
'End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Not IsNumeric(Text4.Text) Then
MsgBox "Enter proper contact number.", vbInformation, "Message"
Else
Command1.Enabled = True
rs.Open "Select * from Labour_Details where Contact_no='" & Text4.Text & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF Then
MsgBox "Record not found.", vbCritical, "Message"
Else
Text1.Text = rs.Fields("full_name").Value
Text2.Text = rs.Fields("Address").Value
Text3.Text = rs.Fields("Contact_no").Value
Text1.SetFocus
End If
rs.Close
End If
End If
End Sub
