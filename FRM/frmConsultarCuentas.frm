VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmConsultarCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar Cuenta"
   ClientHeight    =   4080
   ClientLeft      =   3720
   ClientTop       =   2040
   ClientWidth     =   10215
   Icon            =   "frmConsultarCuentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmConsultarCuentas.frx":324A
   ScaleHeight     =   4080
   ScaleWidth      =   10215
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
      Height          =   3135
      Left            =   8520
      TabIndex        =   7
      Top             =   720
      Width           =   1575
      Begin isButtonTest.isButton cmdDetalle 
         Height          =   420
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmConsultarCuentas.frx":AC67
         Style           =   8
         Caption         =   "       Detalle"
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
      Begin isButtonTest.isButton cmdCerrar 
         Height          =   420
         Left            =   120
         TabIndex        =   13
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmConsultarCuentas.frx":B541
         Style           =   8
         Caption         =   "       Volver"
         IconSize        =   18
         IconAlign       =   1
         CaptionAlign    =   1
         iNonThemeStyle  =   7
         ShowFocus       =   -1  'True
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
         TabIndex        =   10
         Top             =   1320
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
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label5 
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
         Height          =   300
         Left            =   120
         TabIndex        =   11
         Top             =   1080
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
         Height          =   300
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   7200
      Top             =   360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   21
      RowDividerStyle =   0
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
   Begin MSComCtl2.DTPicker dtpDesde 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-MMM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
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
      Format          =   85458945
      CurrentDate     =   41334
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   360
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
      Format          =   85458945
      CurrentDate     =   41332
   End
   Begin isButtonTest.isButton cmdBuscar 
      Height          =   420
      Left            =   3000
      TabIndex        =   6
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmConsultarCuentas.frx":BE1B
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
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   300
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label3 
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
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
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
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmConsultarCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim fecha1 As Date
    Dim fecha2 As Date
    Dim SaldoAnterior As Currency
    Dim SaldoDeudor As Currency
    Dim SaldoAcreedor As Currency
    
    Option Compare Text

Private Sub Form_Load()
    Centrar Me
    Contabilidad
    Cuentas
    dtpDesde.Value = Date
    dtpHasta.Value = Date
    Adodc.CursorLocation = adUseClient
    Adodc.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=T:\base.mdb;Persist Security Info=False;Jet OLEDB:Database Password=ascir"
End Sub

Private Sub cmdBuscar_Click()
    ''error al colocar fechas en orden inverso
    If dtpHasta.Value < dtpDesde.Value Then MsgBox "Fechas incorrectas", vbCritical + vbOKOnly, "Consultar Cuentas": dtpDesde.SetFocus: Exit Sub
    
    ''' cambio formato de fecha a las variables
    fecha1 = Format(dtpDesde.Value, "mm/dd/yyyy")
    fecha2 = Format(dtpHasta.Value, "mm/dd/yyyy")
    
''' consulta cuentas en las fechas y muestra en grilla
    
    If Clave = "cobranza" Then
        Adodc.RecordSource = "SELECT Cuenta, round(sum(Debe),2) as [Debe], round(sum(Haber),2) as [Haber] FROM contabilidad WHERE cuenta='CAJA ADMINISTRACION' and Fecha>= #" & fecha1 & "# And Fecha<= #" & fecha2 & "# and detalle like 'alumno %' or cuenta='Descuento' and Fecha>=#" & fecha1 & "# And Fecha<=#" & fecha2 & "# and detalle like 'alumno %' or cuenta='ARANCELES CURSOS' and Fecha>=#" & fecha1 & "# and Fecha<= #" & fecha2 & "# group by cuenta"
        Adodc.Refresh
        Set grilla.DataSource = Adodc
        grilla.Columns(0).Width = 5000
    Else
        Adodc.RecordSource = "SELECT Cuenta, round(sum(Debe),2) as [Debe], round(sum(Haber),2) as Haber FROM contabilidad WHERE  Fecha>= #" & fecha1 & "# And Fecha<= #" & fecha2 & "# group by cuenta"
        Adodc.Refresh
        Set grilla.DataSource = Adodc
        grilla.Columns(0).Width = 5000
    End If
    
End Sub


Private Sub cmdDetalle_Click()
    Me.Enabled = False
    frmDetalle.lblCuenta.Caption = frmDetalle.lblCuenta.Caption & " " & grilla.Columns(0).Text
    frmDetalle.Show
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub grilla_Click()

If Clave = "cobranza" Then cmdDetalle.Enabled = True: Exit Sub


'''busca la cuenta para saber si tiene saldo deudor o acreedor
With rsCuentas
    If .State = 1 Then .Close
    .Open "SELECT tipo FROM cuentas WHERE cuenta='" & grilla.Columns(0).Text & "'", Cn, adOpenDynamic, adLockPessimistic
    .MoveFirst
End With

    ''' asigna fecha como parametro de fecha anterior a:
    fecha1 = Format(dtpDesde.Value, "mm/dd/yyyy")
    '''habilita boton detalle
    cmdDetalle.Enabled = True
    
    '''asigna los saldos a las variables
    If grilla.Columns(1).Text = "" Then
        SaldoDeudor = 0
    Else
        SaldoDeudor = grilla.Columns(1).Text
    End If
    
    If grilla.Columns(2).Text = "" Then
        SaldoAcreedor = 0
    Else
        SaldoAcreedor = grilla.Columns(2).Text
    End If
    
'''calcula saldo anterior y saldo actual
If rsCuentas!tipo = "DEBE" Then
    
    With rsContabilidad
        If .State = 1 Then .Close
        .Open "SELECT sum(debe)-sum(haber) FROM contabilidad WHERE cuenta='" & grilla.Columns(0).Text & "' and fecha<#" & fecha1 & "#", Cn, adOpenDynamic, adLockPessimistic
        On Error GoTo Error
        SaldoAnterior = !expr1000
        lblSaldoActual.Caption = FormatCurrency(SaldoAnterior + SaldoDeudor - SaldoAcreedor)
        lblSaldoAnterior.Caption = FormatCurrency(SaldoAnterior)
        Exit Sub
    End With

Error:
        SaldoAnterior = 0
        lblSaldoActual.Caption = FormatCurrency(SaldoAnterior + SaldoDeudor - SaldoAcreedor)
        lblSaldoAnterior.Caption = FormatCurrency(SaldoAnterior)
        Exit Sub

Else
    With rsContabilidad
        If .State = 1 Then .Close
        .Open "SELECT sum(haber)-sum(debe) FROM contabilidad WHERE cuenta='" & grilla.Columns(0).Text & "' and fecha<#" & fecha1 & "#", Cn, adOpenDynamic, adLockPessimistic
        On Error GoTo error2
        SaldoAnterior = !expr1000
        lblSaldoActual.Caption = FormatCurrency(SaldoAnterior + SaldoAcreedor - SaldoDeudor)
        lblSaldoAnterior.Caption = FormatCurrency(SaldoAnterior)
        Exit Sub
    End With
error2:
        SaldoAnterior = 0
        lblSaldoActual.Caption = FormatCurrency(SaldoAnterior + SaldoAcreedor - SaldoDeudor)
        lblSaldoAnterior.Caption = FormatCurrency(SaldoAnterior)
        Exit Sub


End If
End Sub

