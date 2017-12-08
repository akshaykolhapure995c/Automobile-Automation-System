VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form16 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "History of vehicles"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14625
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   14625
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "his_tab.frx":0000
      Left            =   3000
      List            =   "his_tab.frx":0007
      TabIndex        =   2
      Text            =   "selsect the date"
      Top             =   360
      Width           =   2415
   End
   Begin MSFlexGridLib.MSFlexGrid table1 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   13361
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
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
   Begin VB.Label Label1 
      Caption         =   "date  of services"
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
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim s As String
Private Sub Combo1_Click()
Dim i As Integer
rs.Open "Select * from Order_dtls where dt='" & Combo1.Text & "'", con, adOpenDynamic, adLockOptimistic
i = 1
table1.Rows = 2
If rs.EOF Then
MsgBox "Record not found.", vbCritical, "Message"
table1.Clear
Me.givnem
Else
While Not rs.EOF

table1.TextMatrix(i, 0) = rs.Fields("Job_no").Value
table1.TextMatrix(i, 1) = rs.Fields("dt").Value
table1.TextMatrix(i, 2) = rs.Fields("vehno").Value
table1.TextMatrix(i, 3) = rs.Fields("owner_name").Value
table1.TextMatrix(i, 4) = rs.Fields("model").Value
table1.TextMatrix(i, 5) = rs.Fields("color").Value
table1.TextMatrix(i, 6) = rs.Fields("model_yr").Value
table1.TextMatrix(i, 7) = rs.Fields("engine_no").Value
table1.TextMatrix(i, 8) = rs.Fields("key_no").Value
table1.TextMatrix(i, 9) = rs.Fields("chasis_no").Value
table1.TextMatrix(i, 10) = rs.Fields("labour_name").Value
table1.TextMatrix(i, 11) = rs.Fields("type_of_job").Value
table1.Rows = table1.Rows + 1
i = i + 1
rs.MoveNext
Wend
End If
rs.Close
End Sub
Private Sub Form_Load()
con.Open "dsn=xyz"
Combo1.Clear
rs.Open "Select * from Order_dtls", con, adOpenDynamic, adLockOptimistic
While Not rs.EOF
Combo1.AddItem rs.Fields("dt").Value
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
table1.TextMatrix(0, 6) = "MODEL_YR"
table1.TextMatrix(0, 7) = "ENGINE"
table1.TextMatrix(0, 8) = "KEY"
table1.TextMatrix(0, 9) = "CHASIS"
table1.TextMatrix(0, 10) = "LABOUR"
table1.TextMatrix(0, 11) = "JOB"
End Sub
Private Sub Form_Unload(Cancel As Integer)
con.Close
Form1.Enabled = True
End Sub
