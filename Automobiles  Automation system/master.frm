VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Laxmipati Automobiles"
   ClientHeight    =   10635
   ClientLeft      =   120
   ClientTop       =   375
   ClientWidth     =   15240
   Icon            =   "master.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "master.frx":AA1E0
   ScaleHeight     =   10635
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   800
      Left            =   5760
      Top             =   5160
   End
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   7080
      Top             =   5160
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   " Automobile Service Center Management"
      BeginProperty Font 
         Name            =   "Stencil"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Width           =   13815
   End
   Begin VB.Menu nw 
      Caption         =   "New"
      Begin VB.Menu cust 
         Caption         =   "Customer"
      End
      Begin VB.Menu lab 
         Caption         =   "Labour"
      End
      Begin VB.Menu ord 
         Caption         =   "Order"
      End
      Begin VB.Menu jb 
         Caption         =   "Job"
      End
      Begin VB.Menu rct 
         Caption         =   "Receipt"
      End
      Begin VB.Menu ex 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu vw 
      Caption         =   "View"
      Begin VB.Menu cd 
         Caption         =   "Customer Details"
      End
      Begin VB.Menu ld 
         Caption         =   "Labour Details"
      End
      Begin VB.Menu jd 
         Caption         =   "Job Description"
      End
      Begin VB.Menu hv 
         Caption         =   "History of vehicle"
      End
   End
   Begin VB.Menu ed 
      Caption         =   "Edit"
      Begin VB.Menu cs 
         Caption         =   "Customer"
      End
      Begin VB.Menu lb 
         Caption         =   "Labour"
      End
      Begin VB.Menu jobtype 
         Caption         =   "Job Type"
      End
   End
   Begin VB.Menu op 
      Caption         =   "Options"
      Begin VB.Menu sp 
         Caption         =   "Set Password"
      End
      Begin VB.Menu br 
         Caption         =   "Birthday Reminder"
      End
      Begin VB.Menu sr 
         Caption         =   "Servicing Reminder"
      End
      Begin VB.Menu bcp 
         Caption         =   "Backup"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub abt_Click()
Form20.Visible = True
Me.Enabled = False
End Sub
Private Sub bcp_Click()
Form12.Visible = True
Me.Enabled = False
End Sub
Private Sub br_Click()
Form8.Visible = True
Me.Enabled = False
End Sub
Private Sub cd_Click()
Form13.Visible = True
Me.Enabled = False
End Sub
Private Sub cs_Click()
Form9.Visible = True
Me.Enabled = True
End Sub
Private Sub cust_Click()
Form2.Visible = True
Me.Enabled = False
End Sub

Private Sub dj_Click()
Form19.Visible = True
Me.Enabled = False
End Sub

Private Sub ex_Click()
End
End Sub
Private Sub Form_Unload(Cancel As Integer)
End
End Sub
Private Sub hv_Click()
Form16.Visible = True
Me.Enabled = False
End Sub
Private Sub jb_Click()
Form4.Visible = True
Me.Enabled = False
End Sub
Private Sub jd_Click()
Form15.Visible = True
Me.Enabled = False
End Sub
Private Sub jobtype_Click()
Form21.Visible = True
Me.Enabled = False
End Sub
Private Sub lab_Click()
Form3.Visible = True
Me.Enabled = False
End Sub
Private Sub lb_Click()
Form10.Visible = True
Me.Enabled = True
End Sub
Private Sub ld_Click()
Form14.Visible = True
Me.Enabled = True
End Sub
Private Sub ord_Click()
Form6.Visible = True
Me.Enabled = False
End Sub
Private Sub rct_Click()
Form7.Visible = True
Me.Enabled = False
End Sub
Private Sub sp_Click()
Form11.Visible = True
Me.Enabled = False
End Sub

Private Sub sr_Click()
Form17.Visible = True
Me.Enabled = False

End Sub

Private Sub Timer1_Timer()
Label1.ForeColor = &HFF&
Timer2.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Label1.ForeColor = &H40C0&
Timer1.Enabled = True
Timer2.Enabled = False
End Sub
