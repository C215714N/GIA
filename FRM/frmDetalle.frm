VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmDetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Cuenta"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmDetalle.frx":0000
   ScaleHeight     =   4980
   ScaleWidth      =   10365
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   4400
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8400
      _ExtentX        =   14817
      _ExtentY        =   7752
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00552233&
      Caption         =   "Saldo"
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
      Height          =   2775
      Left            =   8640
      TabIndex        =   2
      Top             =   240
      Width           =   1575
      Begin isButtonTest.isButton cmdImprimir 
         Height          =   420
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmDetalle.frx":7A1D
         Style           =   8
         Caption         =   "       Imprimir"
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
      Begin isButtonTest.isButton cmdSalir 
         Height          =   420
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmDetalle.frx":82F7
         Style           =   8
         Caption         =   "       Volver"
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
      Begin VB.Label lblSaldoActual 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Actual"
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
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblSaldoAnterior 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Anterior"
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
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label lblCuenta 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle de la Cuenta"
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
      TabIndex        =   1
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "frmDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text

Private Sub cmdImprimir_Click()
    Set dtrCuenta.DataSource = rsContabilidad
    dtrCuenta.Sections("Sección4").Controls("lbldesde").Caption = frmConsultarCuentas.dtpDesde.Value
    dtrCuenta.Sections("Sección4").Controls("lblhasta").Caption = frmConsultarCuentas.dtpHasta.Value
    dtrCuenta.Sections("Sección4").Controls("etiqueta13").Caption = lblSaldoAnterior.Caption
    dtrCuenta.Sections("Sección5").Controls("etiqueta15").Caption = lblSaldoActual.Caption
    dtrCuenta.Sections("Sección4").Controls("lblinforme").Caption = lblCuenta.Caption
    dtrCuenta.Show
    dtrCuenta.Caption = lblCuenta.Caption
    Me.Enabled = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Centrar Me
    Dim fecha1 As Date
    Dim fecha2 As Date
    
    fecha1 = Format(frmConsultarCuentas.dtpDesde.Value, "mm/dd/yyyy")
    fecha2 = Format(frmConsultarCuentas.dtpHasta.Value, "mm/dd/yyyy")
    ''' carga saldos del form anterior
    lblSaldoActual.Caption = FormatCurrency(frmConsultarCuentas.lblSaldoActual.Caption)
    lblSaldoAnterior.Caption = FormatCurrency(frmConsultarCuentas.lblSaldoAnterior.Caption)

    ''' consulta de detalles de la cuenta dentro de la fecha
        With rsContabilidad
            If .State = 1 Then .Close
            If Clave = "cobranza" Then
                .Open "SELECT Asiento, Fecha, NroFactura as Factura, Detalle, Debe, Haber FROM contabilidad WHERE cuenta='" & frmConsultarCuentas.grilla.Columns(0).Text & "' And  Fecha>= #" & fecha1 & "# And Fecha<= #" & fecha2 & "# and detalle like 'ALUMNO %' ORDER BY Fecha,asiento"
            Else
                .Open "SELECT Asiento, Fecha, NroFactura as Factura, Detalle, Debe, Haber FROM contabilidad WHERE cuenta='" & frmConsultarCuentas.grilla.Columns(0).Text & "' And  Fecha>= #" & fecha1 & "# And Fecha<= #" & fecha2 & "# ORDER BY Fecha,asiento"
            End If
        End With
        ''' carga consulta en grilla y le aplica un formato
        Set grilla.DataSource = rsContabilidad
        grilla.Columns(0).Width = 800
        grilla.Columns(1).Width = 1150
        grilla.Columns(2).Width = 800
        grilla.Columns(3).Width = 3500
        grilla.Columns(4).Width = 800
        grilla.Columns(5).Width = 800
        grilla.Columns(4).NumberFormat = "$ ######"
        grilla.Columns(5).NumberFormat = "$ ######"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmConsultarCuentas.Enabled = True
End Sub
