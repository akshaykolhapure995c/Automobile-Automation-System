VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipt Maker"
   ClientHeight    =   6510
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   10305
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10305
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Make"
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid table1 
      Height          =   4935
      Left            =   0
      TabIndex        =   5
      Top             =   1440
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   8705
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   16777215
      ForeColor       =   -2147483643
      ForeColorFixed  =   16777215
      BackColorBkg    =   16777215
      GridColor       =   0
      TextStyle       =   1
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
      Left            =   2040
      TabIndex        =   1
      Top             =   840
      Width           =   2655
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
      Left            =   2040
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
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
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim ar(7) As String
Private Sub Combo1_Click()
rs.Open "Select * from Order_dtls where vehno='" & Combo1.Text & "'", con, adOpenDynamic, adLockOptimistic
Combo2.Clear
table1.Clear
Me.givnem
i = 1
table1.Rows = 2
While Not rs.EOF
Combo2.AddItem rs.Fields("Job_no").Value
table1.TextMatrix(i, 0) = rs.Fields("Job_no").Value
table1.TextMatrix(i, 1) = rs.Fields("dt").Value
table1.TextMatrix(i, 2) = rs.Fields("vehno").Value
table1.TextMatrix(i, 3) = rs.Fields("owner_name").Value
table1.TextMatrix(i, 4) = rs.Fields("model").Value
table1.TextMatrix(i, 5) = rs.Fields("color").Value
table1.TextMatrix(i, 6) = rs.Fields("type_of_job").Value
rs.MoveNext
table1.Rows = table1.Rows + 1
i = i + 1
Wend
rs.Close
End Sub
Private Sub Command1_Click()
If Combo1.Text = "" Then
MsgBox "Select proper vehicle number.", vbInformation, "Message"
Else
If Combo2.Text = "" Then
MsgBox "Select proper job number.", vbInformation, "Message"
Else
rs.Open "Select * from Order_dtls where Job_no='" & Combo2.Text & "'", con, adOpenDynamic, adLockOptimistic
ar(0) = rs.Fields("Job_no").Value
ar(1) = rs.Fields("dt").Value
ar(2) = rs.Fields("vehno").Value
ar(3) = rs.Fields("owner_name").Value
ar(4) = rs.Fields("model").Value
ar(5) = rs.Fields("color").Value
ar(6) = rs.Fields("type_of_job").Value
rs.Close
rs.Open "Select * from Job_details where Type_of_job='" & ar(6) & "'", con, adOpenDynamic, adLockOptimistic
ar(7) = rs.Fields("Charges").Value

rs.Close
MsgBox "Ready to print.", vbInformation, "Message"
Open "temp.ktf" For Output As #1
i = 0
While i < 8
Write #1, ar(i)
i = i + 1
Wend
Close #1
Form18.Visible = True
End If
End If
End Sub
Private Sub Form_Load()
con.Open "dsn=xyz"
Combo1.Clear
rs.Open "Select * from Order_dtls", con, adOpenDynamic, adLockOptimistic
While Not rs.EOF
If UCase(Combo1.List(Combo1.ListCount - 1)) <> UCase(rs.Fields("vehno").Value) Then
Combo1.AddItem (rs.Fields("vehno").Value)
End If
rs.MoveNext
Wend
rs.Close
Me.givnem
End Sub
Sub givnem()
table1.TextMatrix(0, 0) = "JOB NO"
table1.TextMatrix(0, 1) = "DATE"
table1.TextMatrix(0, 2) = "VEH NO"
table1.TextMatrix(0, 3) = "OWNER"
table1.TextMatrix(0, 4) = "MODEL"
table1.TextMatrix(0, 5) = "COLOR"
table1.TextMatrix(0, 6) = "JOB"
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
Form1.Enabled = True
End Sub
