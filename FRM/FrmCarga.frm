VERSION 5.00
Begin VB.Form FrmCarga 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6720
      Top             =   3720
   End
   Begin VB.PictureBox pb 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      Height          =   495
      Left            =   480
      ScaleHeight     =   435
      ScaleWidth      =   6915
      TabIndex        =   0
      Top             =   4920
      Width           =   6975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "Gestión Integral de Alumnos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   555
      Left            =   720
      TabIndex        =   1
      Top             =   0
      Width           =   6600
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   0
      Picture         =   "FrmCarga.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8040
   End
End
Attribute VB_Name = "FrmCarga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    pb.Value = 0
    Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
    If pb.Value = 200 Then
        frmClave.Show
        Timer1.Enabled = False
        Unload Me
    Else
        pb.Value = pb.Value + 1
    End If
End Sub
