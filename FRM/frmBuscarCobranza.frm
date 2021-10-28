VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmBuscarCobranza 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Alumno"
   ClientHeight    =   4020
   ClientLeft      =   4545
   ClientTop       =   3405
   ClientWidth     =   9360
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
   Icon            =   "frmBuscarCobranza.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   9360
   Begin VB.TextBox txtBuscar 
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4000
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   495
      Left            =   7440
      Top             =   5400
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
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
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      TabIndex        =   1
      Top             =   840
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   21
      RowDividerStyle =   0
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
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
   Begin isButtonTest.isButton cmdAceptar 
      Height          =   420
      Left            =   4200
      TabIndex        =   2
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmBuscarCobranza.frx":10CA
      Style           =   8
      Caption         =   "     Aceptar"
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdCancelar 
      Height          =   420
      Left            =   5640
      TabIndex        =   3
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmBuscarCobranza.frx":19A4
      Style           =   8
      Caption         =   "     Cancelar"
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblAlumno 
      BackStyle       =   0  'Transparent
      Caption         =   "Ingrese Codigo, Documento o Nombre"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmBuscarCobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Centrar Me
    Dim busca As String
    Verificaciones
    Adodc.CursorLocation = adUseClient
    Adodc.ConnectionString = DbCon
End Sub
Private Sub cmdAceptar_Click()
    On Error GoTo LineaError:
    
    frmCobranza.lblCodAlumno.Caption = grilla.Columns(0).Text
    frmCobranza.lblNyA.Caption = grilla.Columns(1).Text
    If Trim(Len(frmCobranza.lblCodAlumno.Caption)) = 1 Then frmCobranza.lblCodAlumno.Caption = Format(frmCobranza.lblCodAlumno.Caption, "0000#")
    If Trim(Len(frmCobranza.lblCodAlumno.Caption)) = 2 Then frmCobranza.lblCodAlumno.Caption = Format(frmCobranza.lblCodAlumno.Caption, "000##")
    If Trim(Len(frmCobranza.lblCodAlumno.Caption)) = 3 Then frmCobranza.lblCodAlumno.Caption = Format(frmCobranza.lblCodAlumno.Caption, "00###")
    If Trim(Len(frmCobranza.lblCodAlumno.Caption)) = 4 Then frmCobranza.lblCodAlumno.Caption = Format(frmCobranza.lblCodAlumno.Caption, "0####")
    frmCobranza.Show
    Unload Me
    Exit Sub
    
LineaError:
    MsgBox "Debe realizar una busqueda", vbOKOnly + vbCritical, "Gestion Integral del Alumno"
    ErrCode Err
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub txtBuscar_Change()
    If txtBuscar.Text = "" Then
        cmdAceptar.Enabled = False
    Else
        cmdAceptar.Enabled = True
    End If
    BuscarAlumno
End Sub

Sub BuscarAlumno()
    busca = UCase(Trim(txtBuscar.Text)) & "%"
    Adodc.RecordSource = "SELECT  codalumno as [Cod], nya as [Alumno], tipodoc as [Tipo],DNI as [N°], capac as [Capacitacion] FROM verificaciones WHERE codAlumno LIKE '" & busca & "' OR dni  LIKE '" & busca & "' OR nya  LIKE '" & busca & "' ORDER BY codalumno DESC, nya, dni"
    Adodc.Refresh
    Set grilla.DataSource = Adodc
    formatoGrilla
End Sub

Sub formatoGrilla()
'''Establece las Dimensiones de las Columnas
    Dim w As Integer
    For N = 0 To 4 Step 1
    If N = 1 Or N = 4 Then
        w = 3400
    Else:
        w = 700 - N * (-5.5 ^ N)
        grilla.Columns(N).Alignment = dbgCenter
    End If
    grilla.Columns(N).Width = w
    Next
End Sub

Private Sub txtBuscar_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
