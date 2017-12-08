VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Customer"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7680
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
      Height          =   1170
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   840
      Width           =   1335
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox Combo3 
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
      Height          =   360
      ItemData        =   "cust_details.frx":0000
      Left            =   3600
      List            =   "cust_details.frx":00BB
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.ComboBox Combo2 
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
      Height          =   360
      ItemData        =   "cust_details.frx":022D
      Left            =   2640
      List            =   "cust_details.frx":0255
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
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
      Height          =   360
      ItemData        =   "cust_details.frx":0289
      Left            =   1800
      List            =   "cust_details.frx":02EA
      TabIndex        =   2
      Top             =   2280
      Width           =   855
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
      Height          =   360
      Left            =   1920
      LinkTimeout     =   0
      MaxLength       =   10
      TabIndex        =   5
      Top             =   2880
      Width           =   2775
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
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   2775
   End
   Begin VB.Label Label5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   12
      Top             =   2280
      Width           =   2175
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
      TabIndex        =   11
      Top             =   2880
      Width           =   1695
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
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   1095
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
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1095
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
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs2 As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim e, vin As Boolean
Dim tn, tv As String
Dim nr As Long
Private Sub Combo1_GotFocus()
If Text2.Text = "" Then
MsgBox "Enter proper address.", vbInformation, "Message"
Text2.SetFocus
End If
End Sub
Private Sub Combo4_Click()
tv = Combo4.Text
rs2.Open "Select * from Order_dtls where vehno='" & tv & "'", con, adOpenDynamic, adLockOptimistic
tn = rs2.Fields("owner_name").Value
rs2.Close
rs2.Open "Select * from Customer_Details where Full_Name='" & tn & "'", con, adOpenDynamic, adLockOptimistic
If Not rs2.EOF Then
Text1.Text = rs2.Fields("Full_Name").Value
Text2.Text = rs2.Fields("Address").Value
Label5.Caption = rs2.Fields("DOB").Value
Text3.Text = rs2.Fields("Contact").Value
'Command8.Enabled = True

'Command4.Enabled = False
'Command3.Enabled = False
'Command6.Enabled = False
'Command5.Enabled = False

Else
MsgBox "Record not found.", vbCritical, "Message"
Combo4.Text = ""
Combo4.SetFocus
End If
rs2.Close
End Sub

Private Sub Command1_Click()
If Combo1.Text = "" Or Combo2.Text = "" Or Combo3.Text = "" Or Text3.Text = "" Or IsNumeric(Text1.Text) Or Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Not IsNumeric(Text3.Text) Then
MsgBox "Invalid input.", vbCritical, "Message"
Text1.SetFocus
Else
rs.Open "Select * from Customer_Details where Contact='" & Text3.Text & "'", con, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
MsgBox "Contact number already exists.", vbCritical, "Message"
Text3.SetFocus
rs.Close
Else
'rs.Close
rs.AddNew
rs.Fields("Full_Name").Value = Text1.Text
rs.Fields("Address").Value = Text2.Text
rs.Fields("DOB").Value = Combo1.Text & "/" & Combo2.Text & "/" & Combo3.Text
rs.Fields("Contact").Value = Text3.Text
rs.Update
MsgBox "Added successfully.", vbInformation, "Message"
Me.refr
End If
End If
End Sub
Private Sub Command2_Click()
Me.refr
End Sub



Private Sub Text3_Change()
'If Text3.Text = " " Then
'MsgBox "Enter proper contact number.", vbInformation, "Message"
'End If
End Sub

Private Sub Form_Load()
'rs.Open
con.Open "dsn=xyz"


End Sub
Private Sub Form_Unload(Cancel As Integer)
rs.Close
con.Close
Form1.Enabled = True
End Sub




Private Sub Text2_GotFocus()
If IsNumeric(Text1.Text) Or Text1.Text = "" Then
MsgBox "Enter proper name.", vbInformation, "Message"
Text1.SetFocus
End If
End Sub
Sub filldata()
If Not rs.EOF And Not rs.BOF Then
Text1.Text = rs.Fields("Full_Name").Value
Text2.Text = rs.Fields("Address").Value
Label5.Caption = rs.Fields("DOB").Value
Text3.Text = rs.Fields("Contact").Value
End If
End Sub
Sub refr()
'Command8.Enabled = False

'Command4.Enabled = True
'Command3.Enabled = True
'Command6.Enabled = True
'Command5.Enabled = True
'Combo4.Text = ""
Combo3.Text = ""
Combo1.Text = ""
Combo2.Text = ""
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Label5.Caption = ""
Text1.SetFocus
End Sub

Private Sub ewe_Change()
If Not IsNumeric(ewe.Text) Or Text3.Text = "" Then
MsgBox "Enter proper contact number.", vbInformation, "Message"
Text3.SetFocus
Text3.Text = ""
End If

End Sub
