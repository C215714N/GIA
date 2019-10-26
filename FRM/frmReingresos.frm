VERSION 5.00
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmReingresos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Reingresos"
   ClientHeight    =   1755
   ClientLeft      =   5160
   ClientTop       =   2400
   ClientWidth     =   4020
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReingresos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmReingresos.frx":324A
   ScaleHeight     =   1755
   ScaleWidth      =   4020
   Begin VB.Frame Frame1 
      BackColor       =   &H00884400&
      Caption         =   "Codigo"
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
         Icon            =   "frmReingresos.frx":AC67
         Style           =   8
         Caption         =   "       Reingresar"
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
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "--->"
         ForeColor       =   &H8000000F&
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   480
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ATENCIÓN: RECUERDE QUE EL REINGRESO  APLICA SOLAMENTE A LOS LIBROS DE AULA."
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   120
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
    If txtCodViejo.Text = "" Then MsgBox "Ingrese el código actual del alumno", vbCritical, "Control de Reingresos de Alumnos": txtCodViejo.SetFocus: Exit Sub
    If txtCodNuevo.Text = "" Then MsgBox "Ingrese el nuevo código del alumno", vbCritical, "Control de Reingresos de Alumnos": txtCodNuevo.SetFocus: Exit Sub
    
    With rsLibro
        If .State = 1 Then .Close
        .Open "SELECT * FROM librodeaula WHERE codalumno=" & Int(txtCodViejo.Text), Cn, adOpenDynamic, adLockPessimistic
        If .BOF Or .EOF Then Exit Sub
        .MoveFirst
        Do Until .EOF
            !CodAlumno = Int(txtCodNuevo.Text)
            .UpdateBatch
            .MoveNext
        Loop
        MsgBox "Se ha reingresado al alumno", , "Control de Reingresos de Alumnos"
        txtCodViejo.Text = ""
        txtCodNuevo.Text = ""
        txtCodViejo.SetFocus
    End With
End Sub

Private Sub Form_Load()
    Centrar Me
End Sub

Private Sub txtCodNuevo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtCodViejo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
