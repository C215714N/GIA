VERSION 5.00
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmReingresos 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Reingresos"
   ClientHeight    =   1665
   ClientLeft      =   5160
   ClientTop       =   2400
   ClientWidth     =   4020
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmReingresos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   4020
   Begin VB.Frame Frame1 
      BackColor       =   &H00662200&
      Caption         =   "Codigo"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   975
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3765
      Begin VB.TextBox txtCodNuevo 
         Height          =   360
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtCodViejo 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   855
      End
      Begin isButtonTest.isButton cmdReingresar 
         Height          =   420
         Left            =   2280
         TabIndex        =   7
         Top             =   405
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmReingresos.frx":10CA
         Style           =   8
         Caption         =   "     Reingreso"
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "--->"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   540
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nuevo"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Actual"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ATENCION: RECUERDE QUE EL REINGRESO  APLICA SOLAMENTE A LOS LIBROS DE AULA."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   450
      Left            =   225
      TabIndex        =   8
      Top             =   1065
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ATENCION: RECUERDE QUE EL REINGRESO  APLICA SOLAMENTE A LOS LIBROS DE AULA."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   450
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   3735
   End
End
Attribute VB_Name = "frmReingresos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReingresar_Click()
    If txtCodViejo.text = "" Then MsgBox "Ingrese el codigo actual del alumno", vbCritical, "Control de Reingresos de Alumnos": txtCodViejo.SetFocus: Exit Sub
    If txtCodNuevo.text = "" Then MsgBox "Ingrese el nuevo codigo del alumno", vbCritical, "Control de Reingresos de Alumnos": txtCodNuevo.SetFocus: Exit Sub
    
    With rsLibro
        If .State = 1 Then .Close
        .Open "SELECT * FROM librodeaula WHERE codalumno=" & Int(txtCodViejo.text), Cn, adOpenDynamic, adLockPessimistic
        If .BOF Or .EOF Then Exit Sub
        .MoveFirst
        Do Until .EOF
            !CodAlumno = Int(txtCodNuevo.text)
            .UpdateBatch
            .MoveNext
        Loop
        MsgBox "Se ha reingresado al alumno", , "Control de Reingresos de Alumnos"
        txtCodViejo.text = ""
        txtCodNuevo.text = ""
        txtCodViejo.SetFocus
    End With
End Sub

Private Sub Form_Load()
    Centrar Me
End Sub

Private Sub txtCodNuevo_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub

Private Sub txtCodViejo_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
