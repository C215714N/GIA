VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmMatriculas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Matrículas"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9285
   Icon            =   "frmMatriculas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmMatriculas.frx":324A
   ScaleHeight     =   3720
   ScaleWidth      =   9285
   Begin VB.Frame Frame2 
      BackColor       =   &H00552233&
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
      Height          =   2535
      Left            =   7560
      TabIndex        =   7
      Top             =   960
      Width           =   1605
      Begin VB.Label lblTotalMatricula 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   500
         Width           =   1335
      End
      Begin VB.Label lblTotalAbonado 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblTotalDebe 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2000
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Matrículas"
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
         Height          =   360
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Abonado"
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
         Height          =   360
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1350
      End
      Begin VB.Label Label5 
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
         Height          =   360
         Left            =   120
         TabIndex        =   11
         Top             =   1750
         Width           =   1350
      End
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
      TabIndex        =   3
      Top             =   0
      Width           =   4500
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   360
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
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
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   360
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
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
      Begin isButtonTest.isButton cmdBuscar 
         Height          =   420
         Left            =   3000
         TabIndex        =   6
         Top             =   400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmMatriculas.frx":AC67
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
      Begin VB.Label Label1 
         BackColor       =   &H00662200&
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
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00662200&
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
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   615
      End
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
End
Attribute VB_Name = "frmMatriculas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActualizarMatriculas As Boolean
    Dim desde As Date
    Dim hasta As Date

Private Sub cmbBuscar_Click()
    
    desde = Format(dtpDesde.Value, "mm/dd/yyyy")
    hasta = Format(dtpHasta.Value, "mm/dd/yyyy")
    
    With rsMatriculas
        If ActualizarMatriculas = True Then
            If .State = 1 Then .UpdateBatch: .Close: ActualizarMatriculas = False
        Else
            If .State = 1 Then .Close: ActualizarMatriculas = False
        End If

        .Open "SELECT sum(matriculas.totalMatricula) as [Matricula],sum(Abonado) as [Abono], sum(Debe) as [Debe] FROM matriculas,suscripciones WHERE matriculas.id=suscripciones.id and fechasus>=#" & desde & "# and fechasus<=#" & hasta & "#", Cn, adOpenDynamic, adLockPessimistic
        lblTotalMatricula.Caption = FormatCurrency(!matricula)
        lblTotalAbonado.Caption = FormatCurrency(!Abono)
        lblTotalDebe.Caption = FormatCurrency(!Debe)
        .Close
        .Open "SELECT nya as [Alumno],matriculas.totalMatricula as [Matricula],Abonado, Debe, matriculas.Observaciones,matriculas.id,suscripciones.id FROM matriculas,suscripciones WHERE matriculas.id=suscripciones.id and fechasus>=#" & desde & "# and fechasus<=#" & hasta & "#", Cn, adOpenDynamic, adLockPessimistic
        Set grilla.DataSource = rsMatriculas
    End With
    formatoGrilla
End Sub

Private Sub Form_Load()
    Centrar Me
    dtpDesde.Value = Date
    dtpHasta.Value = Date
    
End Sub

Private Sub grilla_KeyPress(KeyAscii As Integer)
    If grilla.Col = 0 Then KeyAscii = 0: Exit Sub
    ActualizarMatriculas = True
    If KeyAscii = 13 And grilla.Col = 2 Then
        grilla.Columns(3).Text = grilla.Columns(1).Text - grilla.Columns(2).Text
    ElseIf KeyAscii = 13 And grilla.Col = 1 Then
        grilla.Columns(3).Text = grilla.Columns(1).Text - grilla.Columns(2).Text
    End If
End Sub

Sub formatoGrilla()
    Dim w As Integer
    For N = 0 To 6
    If N = 0 Then
        w = 3200
    ElseIf N > 4 Then
        w = 0
    Else:
        w = 800
        grilla.Columns(N).NumberFormat = "$ #####"
    End If
    grilla.Columns(N).Width = w
    Next
End Sub
