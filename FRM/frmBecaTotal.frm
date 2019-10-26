VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmBecaTotal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alumnos 100 %"
   ClientHeight    =   4830
   ClientLeft      =   3255
   ClientTop       =   1950
   ClientWidth     =   9465
   Icon            =   "frmBecaTotal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmBecaTotal.frx":324A
   ScaleHeight     =   4830
   ScaleMode       =   0  'User
   ScaleWidth      =   9530.206
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   21
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00662200&
      Caption         =   "Totales"
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
      Left            =   4680
      TabIndex        =   8
      Top             =   0
      Width           =   4695
      Begin VB.Label lblTotalDebido 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   14
         Top             =   480
         Width           =   1350
      End
      Begin VB.Label lblTotalPagado 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1350
      End
      Begin VB.Label Label8 
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Gasto Adm."
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
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Debido"
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
         Height          =   255
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblAlumnos 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3240
         TabIndex        =   9
         Top             =   480
         Width           =   1350
      End
      Begin VB.Label Label10 
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Alumnos"
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
         Height          =   495
         Left            =   3240
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00552233&
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
      Height          =   3735
      Left            =   7800
      TabIndex        =   15
      Top             =   960
      Width           =   1575
      Begin VB.TextBox txtDebe 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   19
         Top             =   1200
         Width           =   1355
      End
      Begin VB.TextBox txtMatricula 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1355
      End
      Begin VB.TextBox txtComision 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   1355
      End
      Begin MSComCtl2.DTPicker dtpCancelacion 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   44171265
         CurrentDate     =   42093
      End
      Begin isButtonTest.isButton cmdGrabar 
         Height          =   420
         Left            =   120
         TabIndex        =   25
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmBecaTotal.frx":AC67
         Style           =   8
         Caption         =   "   Aceptar"
         IconAlign       =   1
         iNonThemeStyle  =   0
         HighlightColor  =   16744576
         FontHighlightColor=   12632256
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
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelación"
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
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Comisión"
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
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Matricula"
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
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Debe"
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
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.TextBox txtObservaciones 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   7575
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   44171265
      CurrentDate     =   42089
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   44171265
      CurrentDate     =   42089
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00884400&
      Caption         =   "Búsqueda"
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
      TabIndex        =   4
      Top             =   0
      Width           =   4455
      Begin isButtonTest.isButton cmdBuscar 
         Height          =   420
         Left            =   3000
         TabIndex        =   24
         Top             =   400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmBecaTotal.frx":B541
         Style           =   8
         Caption         =   "       Buscar"
         IconSize        =   18
         IconAlign       =   1
         CaptionAlign    =   1
         iNonThemeStyle  =   0
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
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
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
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
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Observaciones"
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
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "frmBecaTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
    Buscar
End Sub

Private Sub cmdGrabar_Click()
    On Error GoTo LineaError
    
    With rsAlumnosBecados
        .Find "idreferencial=" & grilla.Columns(0).Text
        !matricula = txtMatricula.Text
        !Debe = txtDebe.Text
        !comision = txtComision.Text
        !cancelacion = dtpCancelacion.Value
        !observaciones = txtObservaciones.Text
        .UpdateBatch
    End With
    
    '''Restablece los parametros
    txtMatricula.Text = ""
    txtDebe.Text = ""
    txtComision.Text = ""
    dtpCancelacion.Value = Date
       
    Buscar
    
LineaError:
    Select Case Err.Number
        Case 3021
            Resume Next
        End Select
End Sub
Private Sub Form_Load()
    Centrar Me
    dtpDesde.Value = Date
    dtpHasta.Value = Date
End Sub

Private Sub grilla_Click()
    txtMatricula.Text = grilla.Columns(5).Text
    txtDebe.Text = grilla.Columns(6).Text
    txtComision.Text = grilla.Columns(7).Text
    dtpCancelacion.Value = grilla.Columns(9).Text
    txtObservaciones.Text = grilla.Columns(10).Text
    cmdGrabar.Enabled = True
End Sub

Private Sub Buscar()
    Dim desde As Date
    Dim hasta As Date
    
    desde = Format(dtpDesde.Value, "mm/dd/yyyy")
    hasta = Format(dtpHasta.Value, "mm/dd/yyyy")
    
    
    With rsAlumnosBecados
        If .State = 1 Then .Close
        .Open "SELECT sum(matricula),sum(debe),count(*) FROM alumnosbecados WHERE cancelacion>=#" & desde & "# and cancelacion<=#" & hasta & "#", Cn, adOpenDynamic, adLockPessimistic
        lblTotalPagado.Caption = Format(!expr1000, "currency")
        lblTotalDebido.Caption = Format(!expr1001, "currency")
        lblAlumnos.Caption = !expr1002
        .Close
        .Open "SELECT idreferencial,nya as Alumno, tel1 as Telefono, capac as Curso, Asistente, Matricula,Debe,Comision, Fechasus as Fecha, Cancelacion, b.Observaciones FROM suscripciones as s, alumnosbecados as b WHERE b.idreferencial=s.id and cancelacion>=#" & desde & "# and cancelacion<=#" & hasta & "#", Cn, adOpenDynamic, adLockPessimistic
    End With
    
    Set grilla.DataSource = rsAlumnosBecados
    formatoGrilla
End Sub

Sub formatoGrilla()
    Dim w As Integer
    For N = 0 To 10
        If N = 1 Then
            w = 3200
        ElseIf N = 5 Or N = 6 Then
            w = 800
            grilla.Columns(N).NumberFormat = "$ #####"
        ElseIf N = 8 Or N = 9 Then
            w = 1150
        Else:
            w = 0
        End If
        grilla.Columns(N).Width = w
    Next
End Sub

Private Sub txtComision_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDebe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtMatricula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
