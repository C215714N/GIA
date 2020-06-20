VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmMarcar 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Marcar"
   ClientHeight    =   1200
   ClientLeft      =   4380
   ClientTop       =   3630
   ClientWidth     =   2640
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmMarcar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1200
   ScaleWidth      =   2640
   Begin isButtonTest.isButton cmdGrabar 
      Height          =   420
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmMarcar.frx":324A
      Style           =   8
      Caption         =   "     Aceptar"
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
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   91750401
      CurrentDate     =   41341
   End
   Begin VB.CheckBox chkLlamar 
      BackColor       =   &H00662200&
      Caption         =   "Llamar"
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
      Left            =   120
      MaskColor       =   &H00662200&
      TabIndex        =   0
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin VB.CheckBox chkPasa 
      BackColor       =   &H00662200&
      Caption         =   "Pasa"
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
      Left            =   120
      MaskColor       =   &H00662200&
      TabIndex        =   1
      Top             =   420
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin VB.CheckBox chkAbona 
      BackColor       =   &H00662200&
      Caption         =   "Abona"
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
      Left            =   120
      MaskColor       =   &H00662200&
      TabIndex        =   2
      Top             =   720
      UseMaskColor    =   -1  'True
      Width           =   1200
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
End
Attribute VB_Name = "frmMarcar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGrabar_Click()
    Marcar
    With rsMarcar
        .Requery
        .Find "Codalumno=" & CodAlumno
        If .BOF Or .EOF Then MsgBox "El alumno no se puede gestionar", vbCritical, "Marcar": Exit Sub
            !fechacompromiso = DTPFecha.Value
            !fechagestion = Date
            If chkLlamar.Value = 1 Then
                !LPA = "L"
            End If
            If chkPasa.Value = 1 Then
                !LPA = "P"
            End If
            If chkAbona.Value = 1 Then
                !LPA = "A"
            End If
            If chkLlamar.Value = 0 And chkPasa.Value = 0 And chkAbona.Value = 0 Then
                !LPA = ""
            End If
            .UpdateBatch
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    Centrar Me
    DTPFecha.Value = Date
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ''' muestra la ventana desde donde viene
    If Label1.Caption = "frmCuotasXFecha" Then
        frmCuotasXFecha.Enabled = True
        rsCuotasXFecha.Requery
        frmCuotasXFecha.formatoGrilla
    ElseIf Label1.Caption = "frmAnalisisDeCuotas" Then
        frmAnalisisDeCuotas.Enabled = True
    ElseIf Label1.Caption = "frmAnalisisSituacion" Then
        frmAnalisisSituacion.Enabled = True
        rsAnalisisSituacionDeDeuda.Requery
    ElseIf Label1.Caption = "frmMarcas" Then
        frmMarcas.Enabled = True
        rsMarcas.Requery
    End If
    
    '''si se llega desde cuotas x fecha actualiza grilla
    If CuotasXFecha = True Then rsCuotasXFecha.Requery: frmCuotasXFecha.formatoGrilla
End Sub
