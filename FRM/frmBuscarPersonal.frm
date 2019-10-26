VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmBuscarPersonal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Personal"
   ClientHeight    =   3075
   ClientLeft      =   4155
   ClientTop       =   2055
   ClientWidth     =   7110
   Icon            =   "frmBuscarPersonal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmBuscarPersonal.frx":324A
   ScaleHeight     =   3029.557
   ScaleMode       =   0  'User
   ScaleWidth      =   7110
   Begin VB.TextBox txtBuscar 
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
      TabIndex        =   0
      Top             =   360
      Width           =   4000
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   0
      Top             =   3600
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=T:\base.mdb;Persist Security Info=False;Jet OLEDB:Database Password=ascir"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=T:\base.mdb;Persist Security Info=False;Jet OLEDB:Database Password=ascir"
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
   Begin MSDataGridLib.DataGrid Grilla 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   6850
      _ExtentX        =   12091
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
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
   Begin isButtonTest.isButton cmdAceptar 
      Height          =   420
      Left            =   4200
      TabIndex        =   3
      Top             =   304
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmBuscarPersonal.frx":AC67
      Style           =   8
      Caption         =   "       Aceptar"
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
   Begin isButtonTest.isButton cmdCancelar 
      Height          =   420
      Left            =   5640
      TabIndex        =   4
      Top             =   304
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmBuscarPersonal.frx":B541
      Style           =   8
      Caption         =   "       Cancelar"
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar Empleado"
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
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1395
   End
End
Attribute VB_Name = "frmBuscarPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    With rsPersonal
        If .BOF Or .EOF Then Exit Sub
        .Requery
        .Find "NyA='" & grilla.Columns(0).Text & "'"
        frmPersonal.lblID.Caption = !ID
        frmPersonal.txtNya.Text = !NyA
        frmPersonal.txtDNI.Text = !dni
        frmPersonal.cmbTipoDoc.Text = !tipodoc
        frmPersonal.txtDireccion.Text = !direccion
        frmPersonal.txtLocalidad.Text = !localidad
        frmPersonal.dtpFechaNacimiento.Value = !Fechanacimiento
        frmPersonal.dtcCargo.Text = !cargo
        frmPersonal.txtTelCasa.Text = !telcasa
        frmPersonal.txtTelCel.Text = !telcel
        frmPersonal.dtpFechaIngreso.Value = !fechaingreso
    End With
Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Centrar Me
    Dim busca As String
    Adodc.CursorLocation = adUseClient
    Adodc.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=T:\base.mdb;Persist Security Info=False;Jet OLEDB:Database Password=ascir"
    Adodc.RecordSource = "SELECT nya as [Apellido y Nombres],Direccion,Localidad, Telcasa as [Telefono Casa], telcel as Celular, Cargo FROM personal WHERE [nya] like '" & busca & "'"
    Set grilla.DataSource = Adodc
    formatoGrilla
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmPersonal.Enabled = True
End Sub

Private Sub txtBuscar_Change()
    If txtBuscar.Text = "" Then
        cmdAceptar.Enabled = False
    Else
        cmdAceptar.Enabled = True
    End If
    busca = UCase(Trim(txtBuscar.Text)) & "%"
    Adodc.RecordSource = "SELECT nya as [Apellido y Nombres],Direccion,Localidad, Telcasa as [Telefono Casa], telcel as Celular, Cargo FROM personal WHERE [nya] like '" & busca & "'"
    
    Adodc.Refresh
    formatoGrilla
End Sub

Sub formatoGrilla()
    Dim w As Integer
    For N = 0 To 5 Step 1
        If N = 0 Or N = 5 Then
            w = 3400 - (N * 100)
        Else:
            w = 0
        grilla.Columns(N).Width = w
        End If
    Next
End Sub
    
Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdAceptar_Click
End Sub
