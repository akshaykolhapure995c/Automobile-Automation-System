VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form17 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servicing Reminder"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   LinkTopic       =   "Form17"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "SEND SMS TO ALL"
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
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5640
      TabIndex        =   4
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
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
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid table1 
      Height          =   5175
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9128
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorSel    =   -2147483643
      BackColorBkg    =   16777215
      GridColor       =   0
      AllowUserResizing=   3
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      Caption         =   "Dates Before 60 Days From Today..."
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
      TabIndex        =   3
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label3 
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
      Left            =   2040
      TabIndex        =   2
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Today                :"
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
      Width           =   1815
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset

Dim f As Boolean
Private Sub Command1_Click()
table1.Clear
table1.Rows = 2
Me.givnem
Dim i As Integer
i = 1
rs.Open "Select * from Order_dtls", con, adOpenDynamic, adLockOptimistic
While Not rs.EOF
If LCase(rs.Fields("type_of_job").Value) = "servicing" Then
If Date - 60 >= rs.Fields("dt").Value Then
table1.TextMatrix(i, 0) = rs.Fields("Job_no").Value
table1.TextMatrix(i, 1) = rs.Fields("dt").Value
table1.TextMatrix(i, 2) = rs.Fields("dt").Value + 60
table1.TextMatrix(i, 3) = rs.Fields("vehno").Value
table1.TextMatrix(i, 4) = rs.Fields("owner_name").Value
rs2.Open "Select * from Customer_Details where Full_Name='" & rs.Fields("owner_name").Value & "'", con, adOpenDynamic, adLockOptimistic
'table1.TextMatrix(i, 5) = rs2.Fields("Address").Value
'table1.TextMatrix(i, 6) = rs2.Fields("Contact").Value
rs2.Close
i = i + 1
table1.Rows = table1.Rows + 1
f = True
End If
End If
rs.MoveNext
Wend
If Not f Then
MsgBox "No servicings...!", vbExclamation, "Message"
End If
rs.Close
End Sub
Private Sub Form_Load()
con.Open "dsn=xyz"
Label3.Caption = Date
Me.givnem
End Sub
Private Sub Form_Unload(Cancel As Integer)
con.Close
Form1.Enabled = True
End Sub
Sub givnem()
table1.TextMatrix(0, 0) = "JOB NO"
table1.TextMatrix(0, 1) = "LAST DATE"
table1.TextMatrix(0, 2) = "NEXT DATE"
table1.TextMatrix(0, 3) = "VEH NO"
table1.TextMatrix(0, 4) = "OWNER"
table1.TextMatrix(0, 5) = "ADDRESS"
table1.TextMatrix(0, 6) = "CONTACT"
End Sub

