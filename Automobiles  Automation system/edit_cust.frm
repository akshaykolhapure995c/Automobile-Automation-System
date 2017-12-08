VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Customer"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5955
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1455
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
      Height          =   285
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   0
      Top             =   240
      Width           =   2295
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
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   3855
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
      TabIndex        =   7
      Top             =   3360
      Width           =   3855
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
      ItemData        =   "edit_cust.frx":0000
      Left            =   1920
      List            =   "edit_cust.frx":0061
      TabIndex        =   4
      Top             =   2640
      Width           =   615
   End
   Begin VB.ComboBox Combo2 
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
      ItemData        =   "edit_cust.frx":00E1
      Left            =   2640
      List            =   "edit_cust.frx":0109
      TabIndex        =   5
      Top             =   2640
      Width           =   735
   End
   Begin VB.ComboBox Combo3 
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
      ItemData        =   "edit_cust.frx":013D
      Left            =   3480
      List            =   "edit_cust.frx":01F8
      TabIndex        =   6
      Top             =   2640
      Width           =   855
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
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
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
      Height          =   765
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label Label6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   15
      Top             =   2640
      Width           =   1335
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
      TabIndex        =   14
      Top             =   240
      Width           =   1695
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
      TabIndex        =   13
      Top             =   840
      Width           =   1575
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
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "DOB"
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
      TabIndex        =   11
      Top             =   2640
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
      TabIndex        =   10
      Top             =   3360
      Width           =   1695
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Combo4_Change()
rs.Open "Select * from Customer_Details where Contact='" & Combo4.Text & "'", con, adOpenDynamic, adLockOptimistic
Text1.Text = rs.Fields("Full_Name").Value
Text2.Text = rs.Fields("Address").Value
'Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text = rs.Fields("DOB").Value
Text3.Text = rs.Fields("Contact").Value

rs.Close
End Sub

Private Sub Command1_Click()

If Combo1.Text = "" Or Combo2.Text = "" Or Combo3.Text = "" Then
MsgBox "Select proper birth date.", vbInformation, "Message"
Else

rs.Open "Select * from Customer_Details", con, adOpenDynamic, adLockOptimistic
 rs.Fields("Full_Name").Value = Text1.Text
 rs.Fields("Address").Value = Text2.Text
 rs.Fields("DOB").Value = Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text
 rs.Fields("Contact").Value = Text3.Text
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
Label6.Caption = ""
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Command1.Enabled = False
End Sub
Private Sub Form_Load()
con.Open "dsn=xyz"
End Sub
Private Sub Command3_Click()

If Not IsNumeric(Text4.Text) Then
MsgBox "Enter proper date of birth.", vbInformation, "Message"
Else
Command1.Enabled = True
rs.Open "Select * from Customer_Details where contact='" & Text4.Text & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF Then
MsgBox "Record not found.", vbCritical, "Message"
Else
Text1.Text = rs.Fields("Full_Name").Value
Text2.Text = rs.Fields("Address").Value
Label6.Caption = rs.Fields("DOB").Value
Text3.Text = rs.Fields("contact").Value
End If
rs.Close
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
con.Close
End Sub

Private Sub Text3_Change()
'If Not IsNumeric(Text3.Text) Then
    'MsgBox "Enter proper contact number.", vbInformation, "Message"
'End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Not IsNumeric(Text3.Text) Then
MsgBox "Enter proper contact number.", vbInformation, "Message"
Else
Command1.Enabled = True
rs.Open "Select * from Customer_Details where contact='" & Text4.Text & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF Then
MsgBox "Record not found.", vbCritical, "Message"
Else
Text1.Text = rs.Fields("Full_Name").Value
Text2.Text = rs.Fields("Address").Value
Label6.Caption = rs.Fields("DOB").Value
Text3.Text = rs.Fields("contact").Value
End If
rs.Close
End If
End If
End Sub
