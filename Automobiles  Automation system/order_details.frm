VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Order"
   ClientHeight    =   5145
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   8730
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   8730
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Add"
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.ComboBox Combo9 
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
      Left            =   6240
      TabIndex        =   9
      Top             =   4560
      Width           =   2415
   End
   Begin VB.ComboBox Combo8 
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
      TabIndex        =   8
      Top             =   4560
      Width           =   2415
   End
   Begin VB.ComboBox Combo7 
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
      ItemData        =   "order_details.frx":0000
      Left            =   1920
      List            =   "order_details.frx":0061
      TabIndex        =   4
      Top             =   3840
      Width           =   2415
   End
   Begin VB.ComboBox Combo6 
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
      ItemData        =   "order_details.frx":011F
      Left            =   1920
      List            =   "order_details.frx":0141
      TabIndex        =   3
      Top             =   3240
      Width           =   2415
   End
   Begin VB.ComboBox Combo5 
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
      ItemData        =   "order_details.frx":018B
      Left            =   1920
      List            =   "order_details.frx":024C
      TabIndex        =   2
      Top             =   2640
      Width           =   2415
   End
   Begin VB.ComboBox Combo4 
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
      TabIndex        =   1
      Top             =   2040
      Width           =   2415
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
      Left            =   6240
      TabIndex        =   7
      Top             =   3240
      Width           =   2415
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
      Left            =   6240
      TabIndex        =   6
      Top             =   2640
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
      Left            =   6240
      TabIndex        =   5
      Top             =   2040
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
      Left            =   1920
      TabIndex        =   0
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   8760
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8760
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label14 
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
      Left            =   1920
      TabIndex        =   25
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   1920
      TabIndex        =   24
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label12 
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
      Height          =   375
      Left            =   4440
      TabIndex        =   23
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Labour Name"
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
      TabIndex        =   22
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label Label10 
      Caption         =   "Chasis No"
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
      Left            =   4440
      TabIndex        =   21
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "Key No"
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
      Left            =   4440
      TabIndex        =   20
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "Engine No"
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
      Left            =   4440
      TabIndex        =   19
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Model Year"
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
      TabIndex        =   18
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Color"
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
      TabIndex        =   17
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Model"
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
      TabIndex        =   16
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "Owner Name"
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
      TabIndex        =   15
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label3 
      Caption         =   "Vehicle Number"
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
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Date"
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
      Width           =   735
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
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim e As Boolean

Private Sub Combo2_Change()

End Sub

Private Sub Combo4_GotFocus()
If Text1.Text = "" Then
MsgBox "Enter proper vehicle number.", vbInformation, "Message"
Text1.SetFocus
Else
rs.Open "Select * from Order_dtls where vehno='" & Text1.Text & "'", con, adOpenDynamic, adLockOptimistic
If rs.EOF Then
Else
e = True
Combo4.Locked = True
Combo5.Locked = True
Combo6.Locked = True
Combo7.Locked = True
Text2.Locked = True
Text3.Locked = True
Text4.Locked = True
Combo4.Text = rs.Fields("owner_name").Value
Combo5.Text = rs.Fields("model").Value
Combo6.Text = rs.Fields("color").Value
Combo7.Text = rs.Fields("model_yr").Value
Text2.Text = rs.Fields("engine_no").Value
Text3.Text = rs.Fields("key_no").Value
Text4.Text = rs.Fields("chasis_no").Value
End If
rs.Close
End If
End Sub
Private Sub Combo5_GotFocus()
If Combo4.Text = "" Then
MsgBox "Select proper owner name.", vbInformation, "Message"
Combo4.SetFocus
End If
End Sub
Private Sub Combo6_GotFocus()
If Combo5.Text = "" Then
MsgBox "Select proper model.", vbInformation, "Message"
Combo5.SetFocus
End If
End Sub
Private Sub Combo7_GotFocus()
If Combo6.Text = "" Then
MsgBox "Select proper color.", vbInformation, "Message"
Combo6.SetFocus
End If
End Sub
Private Sub Command1_Click()
If (Combo4.Text = "" Or Combo5.Text = "" Or Combo6.Text = "" Or Combo7.Text = "" Or Combo8.Text = "" Or Combo9.Text = "") And Not e Then
MsgBox "Select combo boxes properly.", vbInformation, "Message"
Else
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
MsgBox "Fill textboxes proper birth date.", vbInformation, "Message"
Else
rs.Open "Select * from Order_dtls", con, adOpenDynamic, adLockOptimistic
rs.AddNew
rs.Fields("dt").Value = Date
rs.Fields("vehno").Value = Text1.Text
rs.Fields("owner_name").Value = Combo4.Text
rs.Fields("model").Value = Combo5.Text
rs.Fields("color").Value = Combo6.Text
rs.Fields("model_yr").Value = Combo7.Text
rs.Fields("engine_no").Value = Text2.Text
rs.Fields("key_no").Value = Text3.Text
rs.Fields("chasis_no").Value = Text4.Text
rs.Fields("labour_name").Value = Combo8.Text
rs.Fields("type_of_job").Value = Combo9.Text
rs.Fields("Job_no").Value = Label13.Caption
rs.Update
rs.Close
MsgBox "Added successfully.", vbInformation, "Message"
Unload Me
End If
End If
End Sub
Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo4.Text = ""
Combo5.Text = ""
Combo6.Text = ""
Combo7.Text = ""
Combo8.Text = ""
Combo9.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text1.SetFocus
End Sub



Private Sub Form_Load()
con.Open "dsn=xyz", "other", "laxmi", -1
Label14.Caption = Date
rs.Open "Select * from Order_dtls", con, adOpenDynamic, adLockOptimistic
If Not rs.EOF Then
rs.MoveLast
Label13.Caption = rs.Fields("Job_no") + 1
Else
Label13.Caption = 1
End If
rs.Close
Combo4.Clear
rs.Open "Select * from Customer_Details", con, adOpenDynamic, adLockOptimistic
While Not rs.EOF
Combo4.AddItem (rs.Fields("Full_Name").Value)
rs.MoveNext
Wend
rs.Close
Combo8.Clear
rs.Open "Select * from Labour_Details", con, adOpenDynamic, adLockOptimistic
While Not rs.EOF
Combo8.AddItem (rs.Fields("full_name").Value)
rs.MoveNext
Wend
rs.Close
Combo9.Clear
rs.Open "Select * from Job_details", con, adOpenDynamic, adLockOptimistic
While Not rs.EOF
Combo9.AddItem (rs.Fields("Type_of_job").Value)
rs.MoveNext
Wend
rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
Form1.Enabled = True
End Sub

Private Sub Text1_Change()
If Right(Text1.Text, 1) = " " Then
MsgBox ("Dont use space in vehicle number."), vbInformation, "Message"
Text1.Text = Trim(Text1.Text)
End If
End Sub

Private Sub Text2_GotFocus()
If Combo7.Text = "" Then
MsgBox ("Select proper model year."), vbInformation, "Message"
Combo7.SetFocus
End If
End Sub

Private Sub Text3_GotFocus()
If Text2.Text = "" Then
MsgBox ("Enter proper engine number."), vbInformation, "Message"
Text2.SetFocus
End If
End Sub

Private Sub Text4_GotFocus()
If Text3.Text = "" Then
MsgBox ("Select proper key number."), vbInformation, "Message"
Text3.SetFocus
End If
End Sub

Private Sub Combo8_GotFocus()
If Text4.Text = "" Then
MsgBox "Enter proper chasis number.", vbInformation, "Message"
Text4.SetFocus
End If
End Sub
Private Sub Combo9_GotFocus()
If Combo8.Text = "" Then
MsgBox "Select proper labour name.", vbInformation, "Message"
Combo8.SetFocus
End If
End Sub

Private Sub Command1_GotFocus()
If Combo9.Text = "" Then
MsgBox "Select proper type of job.", vbInformation, "Message"
Combo9.SetFocus
End If
End Sub
