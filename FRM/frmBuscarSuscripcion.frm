VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmBuscarSuscripcion 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Suscripcion"
   ClientHeight    =   4080
   ClientLeft      =   5475
   ClientTop       =   3150
   ClientWidth     =   9405
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmBuscarSuscripcion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   9405
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   120
      Top             =   5040
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
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5318
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
   Begin isButtonTest.isButton cmdAceptar 
      Height          =   420
      Left            =   4200
      TabIndex        =   3
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmBuscarSuscripcion.frx":10CA
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
      TabIndex        =   4
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmBuscarSuscripcion.frx":19A4
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Apellido y Nombre:"
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
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1920
   End
End
Attribute VB_Name = "frmBuscarSuscripcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Centrar Me
    Dim busca As String
    Adodc.CursorLocation = adUseClient
    Adodc.ConnectionString = DbCon
    Adodc.RecordSource = "SELECT id, Nya as [Apellido y Nombres], tipoDoc as [Tipo], DNI as [N�], Direccion, CP, Localidad, Nacionalidad, Edad, FechaNac, Capac as [Capacitacion], Asistente, Tel1, Tel2, Tel3, Tel4, PTel1, PTel2, PTel3, PTel4, TotalCurso, Cuotas, GastoAdm, Fechasus, Observaciones, Manuales, DchoExamen, TotalMatricula, NroFactura FROM suscripciones WHERE [Nya] like '" & busca & "'"
    Set grilla.DataSource = Adodc
    formatoGrilla
End Sub
Private Sub cmdAceptar_Click()
    On Error GoTo LineaError
    If Adodc.Recordset.RecordCount = 0 Then MsgBox "Debe realizar una busqueda", vbOKOnly + vbCritical, "Gestion Integral del Alumno": Exit Sub

    If Verificar = False Then
        frmSuscripciones.lblID.Caption = grilla.Columns(0).Text
        frmSuscripciones.txtNya.Text = grilla.Columns(1).Text
        frmSuscripciones.cmbTipoDoc.Text = grilla.Columns(2).Text
        frmSuscripciones.txtDocumento.Text = grilla.Columns(3).Text
        frmSuscripciones.txtDireccion.Text = grilla.Columns(4).Text
        frmSuscripciones.txtCP.Text = grilla.Columns(5).Text
        frmSuscripciones.dtcLocalidad.Text = grilla.Columns(6).Text
        frmSuscripciones.txtNacionalidad.Text = grilla.Columns(7).Text
    
        If Month(grilla.Columns(9).Text) < Month(Date) Then
                frmSuscripciones.txtEdad.Text = DateDiff("yyyy", grilla.Columns(9).Text, Date)
            ElseIf Day(grilla.Columns(9).Text) <= Day(Date) And Month(grilla.Columns(9).Text) = Month(Date) Then
                frmSuscripciones.txtEdad.Text = DateDiff("yyyy", grilla.Columns(9).Text, Date)
            ElseIf Day(grilla.Columns(9).Text) > Day(Date) And Month(grilla.Columns(9).Text) >= Month(Date) Then
                frmSuscripciones.txtEdad.Text = DateDiff("yyyy", grilla.Columns(9).Text, Date) - 1
            Else
              frmSuscripciones.txtEdad.Text = DateDiff("yyyy", grilla.Columns(9).Text, Date) - 1
        End If
    
        frmSuscripciones.dtpFechaNacimiento.Value = grilla.Columns(9).Text
        frmSuscripciones.dtcCapacitacion.Text = grilla.Columns(10).Text
        frmSuscripciones.dtcAsistente.Text = grilla.Columns(11).Text
        frmSuscripciones.txtTel1.Text = grilla.Columns(12).Text
        frmSuscripciones.txtTel2.Text = grilla.Columns(13).Text
        frmSuscripciones.txtTel3.Text = grilla.Columns(14).Text
        frmSuscripciones.txtTel4.Text = grilla.Columns(15).Text
        frmSuscripciones.txtPT1.Text = grilla.Columns(16).Text
        frmSuscripciones.txtPT2.Text = grilla.Columns(17).Text
        frmSuscripciones.txtPT3.Text = grilla.Columns(18).Text
        frmSuscripciones.txtPT4.Text = grilla.Columns(19).Text
        frmSuscripciones.txtTotalCurso.Text = grilla.Columns(20).Text
        frmSuscripciones.txtTotalCuotas.Text = grilla.Columns(21).Text
        frmSuscripciones.txtGastoAdm.Text = grilla.Columns(22).Text
        frmSuscripciones.dtpFechaSuscripcion.Value = grilla.Columns(23).Text
        frmSuscripciones.txtObservaciones.Text = grilla.Columns(24).Text
        frmSuscripciones.txtNroFactura.Text = grilla.Columns(28).Text
        frmSuscripciones.txtTotalMatricula.Text = grilla.Columns(27).Text
        
        If grilla.Columns(25).Text = "0" Then
                frmSuscripciones.chkManuales.Value = 0
            Else
                frmSuscripciones.chkManuales.Value = 1
            End If
            If grilla.Columns(26).Text = "0" Then
                frmSuscripciones.chkExamenes.Value = 0
            Else
            frmSuscripciones.chkExamenes.Value = 1
        End If
    
        frmSuscripciones.Enabled = True
        Unload Me
    Else
        frmVerificaciones.Enabled = True
        frmVerificaciones.HabilitarCuadros False, True
        frmVerificaciones.HabilitarBotones False, True
        frmVerificaciones.txtNya.SetFocus
        frmVerificaciones.DTPFechaVerificacion.Value = Date
        frmVerificaciones.lblCodAlumno.Caption = ""
        frmVerificaciones.txtNya.Text = grilla.Columns(1).Text
        frmVerificaciones.cmbTipoDoc.Text = grilla.Columns(2).Text
        frmVerificaciones.txtDocumento.Text = grilla.Columns(3).Text
        frmVerificaciones.txtDireccion.Text = grilla.Columns(4).Text
        frmVerificaciones.txtCP.Text = grilla.Columns(5).Text
        frmVerificaciones.dtcLocalidad.Text = grilla.Columns(6).Text
        frmVerificaciones.txtNacionalidad.Text = grilla.Columns(7).Text
        
        If Month(grilla.Columns(9).Text) < Month(Date) Then
                frmVerificaciones.txtEdad.Text = DateDiff("yyyy", grilla.Columns(9).Text, Date)
            ElseIf Day(grilla.Columns(9).Text) <= Day(Date) And Month(grilla.Columns(9).Text) = Month(Date) Then
                frmVerificaciones.txtEdad.Text = DateDiff("yyyy", grilla.Columns(9).Text, Date)
            ElseIf Day(grilla.Columns(9).Text) > Day(Date) And Month(grilla.Columns(9).Text) >= Month(Date) Then
                frmVerificaciones.txtEdad.Text = DateDiff("yyyy", grilla.Columns(9).Text, Date) - 1
            Else
                frmVerificaciones.txtEdad.Text = DateDiff("yyyy", grilla.Columns(9).Text, Date) - 1
        End If
        
        frmVerificaciones.dtpFechaNacimiento.Value = grilla.Columns(9).Text
        frmVerificaciones.dtcCapacitacion.Text = grilla.Columns(10).Text
        frmVerificaciones.dtcAsistente.Text = grilla.Columns(11).Text
        frmVerificaciones.txtTel1.Text = grilla.Columns(12).Text
        frmVerificaciones.txtTel2.Text = grilla.Columns(13).Text
        frmVerificaciones.txtTel3.Text = grilla.Columns(14).Text
        frmVerificaciones.txtTel4.Text = grilla.Columns(15).Text
        frmVerificaciones.txtPT1.Text = grilla.Columns(16).Text
        frmVerificaciones.txtPT2.Text = grilla.Columns(17).Text
        frmVerificaciones.txtPT3.Text = grilla.Columns(18).Text
        frmVerificaciones.txtPT4.Text = grilla.Columns(19).Text
        frmVerificaciones.txtTotalCurso.Text = grilla.Columns(20).Text
        frmVerificaciones.txtTotalCuotas.Text = grilla.Columns(21).Text
        frmVerificaciones.txtGastoAdm.Text = grilla.Columns(22).Text
        frmVerificaciones.dtpFechaSuscripcion.Value = grilla.Columns(23).Text
        frmVerificaciones.txtObservaciones.Text = grilla.Columns(24).Text
        
        If grilla.Columns(25).Text = "0" Then
                frmVerificaciones.chkManuales.Value = 0
            Else
                frmVerificaciones.chkManuales.Value = 1
            End If
            If grilla.Columns(26).Text = "0" Then
                frmVerificaciones.chkExamenes.Value = 0
            Else
                frmVerificaciones.chkExamenes.Value = 1
        End If
        Unload Me
    End If
    
LineaError: ErrCode
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
    If Verificar = False Then
        frmSuscripciones.Enabled = True
    Else
        frmVerificaciones.Enabled = True
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
    If Verificar = False Then
        frmSuscripciones.Enabled = True
    Else
        frmVerificaciones.Enabled = True
    End If
End Sub

Private Sub txtBuscar_Change()
    If txtBuscar.Text = "" Then
        cmdAceptar.Enabled = False
    Else
        cmdAceptar.Enabled = True
    End If
    busca = UCase(Trim(txtBuscar.Text)) & "%"
    Adodc.RecordSource = "SELECT id, Nya as [Apellido y Nombres], tipoDoc as [Tipo], DNI as [N�], Direccion, CP, Localidad, Nacionalidad, Edad, FechaNac, Capac as [Capacitacion], Asistente, Tel1, Tel2, Tel3, Tel4, PTel1, PTel2, PTel3, PTel4, TotalCurso, Cuotas, GastoAdm, Fechasus, Observaciones, Manuales, DchoExamen, TotalMatricula, NroFactura FROM suscripciones WHERE dni LIKE '" & busca & "' OR [nya] LIKE '" & busca & "' ORDER BY nya, dni"
    Adodc.Refresh
    Adodc.Refresh
    formatoGrilla
End Sub
Sub formatoGrilla()
Dim N As Integer
    Dim w As Integer
    For N = 0 To 28 Step 1
        If N = 1 Or N = 10 Then
            w = 3400
        ElseIf N = 2 Or N = 3 Then
            w = 700 - N * (-5.5 ^ N)
        Else:
            w = 0
        End If
        grilla.Columns(N).Width = w
    Next
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdAceptar_Click
End Sub
