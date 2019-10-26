VERSION 5.00
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Status de Inicio de Mes"
   ClientHeight    =   3405
   ClientLeft      =   4905
   ClientTop       =   2070
   ClientWidth     =   3360
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmStatus.frx":324A
   ScaleHeight     =   3405
   ScaleWidth      =   3360
   Begin VB.Frame Frame1 
      BackColor       =   &H00884400&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin isButtonTest.isButton cmdDetalles 
         Height          =   420
         Left            =   1680
         TabIndex        =   9
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmStatus.frx":AC67
         Style           =   8
         Caption         =   "       Detalles"
         IconSize        =   18
         IconAlign       =   1
         CaptionAlign    =   1
         iNonThemeStyle  =   7
         HighlightColor  =   4194304
         FontHighlightColor=   14737632
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   0
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblUltimoPlanCartera 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label lblUltimoPlanDePago 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblUltimoCartera 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblAlumnosDelMes 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Último Alumno en Cartera: "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   525
         Left            =   120
         TabIndex        =   4
         Top             =   2100
         Width           =   1560
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Último Alumno con Plan de Pago: "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   420
         Left            =   120
         TabIndex        =   3
         Top             =   1500
         Width           =   1560
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Último Alumno en Cartera: "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   420
         Left            =   120
         TabIndex        =   2
         Top             =   900
         Width           =   1560
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nuevos Alumnos en Cartera: "
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   420
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDetalles_Click()
    frmDetalleStatus.Show
    Me.Enabled = False
End Sub

Private Sub Form_Load()
    Centrar Me
    With rsStatus
        If .State = 1 Then .Close
        '''tabla alumnos del mes
        .Open "SELECT count(codalumno) FROM alumnosdelmes", Cn, adOpenDynamic, adLockPessimistic
        lblAlumnosDelMes.Caption = !expr1000
        
        '''ultimo alumno en la cartera
        .Close
        .Open "SELECT max(codalumno) FROM marcas", Cn, adOpenDynamic, adLockPessimistic
        lblUltimoCartera.Caption = !expr1000
        
        '''ultimo plan de pago
        .Close
        .Open "SELECT max(codalumno) FROM plandepago", Cn, adOpenDynamic, adLockPessimistic
        lblUltimoPlanDePago.Caption = !expr1000
        .Close
        .Open "SELECT max(codalumno) FROM marcas WHERE deuda>1", Cn, adOpenDynamic, adLockPessimistic
        lblUltimoPlanCartera.Caption = !expr1000
        .Close
    End With
End Sub
