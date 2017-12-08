VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form15 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Job Details"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8190
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid table1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   10
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
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Form_Load()
con.Open "dsn=xyz"
Me.givnem
Dim i As Integer
i = 1
rs.Open "Select * from Job_details", con, adOpenDynamic, adLockOptimistic
While Not rs.EOF
table1.TextMatrix(i, 0) = rs.Fields("Type_of_job").Value
table1.TextMatrix(i, 1) = rs.Fields("Charges").Value
table1.Rows = table1.Rows + 1
rs.MoveNext
i = i + 1
Wend
rs.Close
End Sub
Sub givnem()
table1.TextMatrix(0, 0) = "JOB"
table1.TextMatrix(0, 1) = "CHARGES"
End Sub
Private Sub Form_Unload(Cancel As Integer)
con.Close
Form1.Enabled = True
End Sub


