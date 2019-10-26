VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmMarcas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Marcas"
   ClientHeight    =   5670
   ClientLeft      =   4395
   ClientTop       =   1680
   ClientWidth     =   5745
   Icon            =   "frmMarcas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmMarcas.frx":324A
   ScaleHeight     =   5670
   ScaleWidth      =   5745
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   375
      Left            =   4200
      Top             =   3600
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   661
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
      Height          =   4455
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   7858
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
      TabIndex        =   5
      Top             =   0
      Width           =   5535
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
         CurrentDate     =   41345
      End
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
         CurrentDate     =   41345
      End
      Begin VB.CheckBox chkAbona 
         BackColor       =   &H00884400&
         Caption         =   "Abona"
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
         Left            =   3000
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.CheckBox chkPasa 
         BackColor       =   &H00884400&
         Caption         =   "Pasa"
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
         Left            =   3000
         TabIndex        =   3
         Top             =   400
         Width           =   855
      End
      Begin VB.CheckBox chkLlama 
         BackColor       =   &H00884400&
         Caption         =   "Llama"
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
         Left            =   3000
         TabIndex        =   2
         Top             =   200
         Width           =   855
      End
      Begin isButtonTest.isButton cmdBuscar 
         Height          =   420
         Left            =   4080
         TabIndex        =   14
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmMarcas.frx":AC67
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
      Begin VB.Label Label5 
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
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
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
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin isButtonTest.isButton cmdMarcar 
      Height          =   420
      Left            =   4200
      TabIndex        =   10
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmMarcas.frx":B541
      Style           =   8
      Caption         =   "       Marcar"
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
   Begin isButtonTest.isButton cmdDatos 
      Height          =   420
      Left            =   4200
      TabIndex        =   11
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmMarcas.frx":BE1B
      Style           =   8
      Caption         =   "       Datos"
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
   Begin isButtonTest.isButton cmdCuotas 
      Height          =   420
      Left            =   4200
      TabIndex        =   12
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmMarcas.frx":C6F5
      Style           =   8
      Caption         =   "       Cuotas"
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
      Left            =   4200
      TabIndex        =   13
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmMarcas.frx":CFCF
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
   Begin VB.Label Label6 
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
      Height          =   420
      Left            =   4200
      TabIndex        =   9
      Top             =   4080
      Visible         =   0   'False
      Width           =   1350
   End
End
Attribute VB_Name = "frmMarcas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAbona_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpDesde.SetFocus
End Sub

Private Sub chkLlama_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkPasa.SetFocus
End Sub

Private Sub chkPasa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then chkAbona.SetFocus
End Sub

Private Sub cmdBuscar_Click()
    Dim fecha1, fecha2 As Date
    
    '''aplico formato a las variables de fecha para la consulta sql
    fecha1 = Format(dtpDesde.Value, "mm,dd,yyyy")
    fecha2 = Format(dtpHasta.Value, "mm,dd,yyyy")
    
    '''realiza búsqueda dependiendo de los parámetros escogidos
    With rsMarcas
        If .State = 1 Then .Close
        If chkLlama.Value = 1 And chkPasa.Value = 0 And chkAbona.Value = 0 Then
            .Open "SELECT codalumno as [Codigo],fechacompromiso as [Compromiso],LPA, Fechagestion as [Gestion] FROM marcas WHERE LPA = 'L' and fechacompromiso >= #" & fecha1 & "# and fechacompromiso <= #" & fecha2 & "#", Cn, adOpenDynamic, adLockPessimistic
        ElseIf chkLlama.Value = 0 And chkPasa.Value = 1 And chkAbona.Value = 0 Then
            .Open "SELECT codalumno as [Codigo],fechacompromiso as [Compromiso],LPA, Fechagestion as [Gestion] FROM marcas WHERE LPA = 'P' and fechacompromiso >= #" & fecha1 & "# and fechacompromiso <= #" & fecha2 & "#", Cn, adOpenDynamic, adLockPessimistic
        ElseIf chkLlama.Value = 0 And chkPasa.Value = 0 And chkAbona.Value = 1 Then
            .Open "SELECT codalumno as [Codigo],fechacompromiso as [Compromiso],LPA, Fechagestion as [Gestion] FROM marcas WHERE LPA = 'A' and fechacompromiso >= #" & fecha1 & "# and fechacompromiso <= #" & fecha2 & "#", Cn, adOpenDynamic, adLockPessimistic
        ElseIf chkLlama.Value = 1 And chkPasa.Value = 1 And chkAbona.Value = 0 Then
            .Open "SELECT codalumno as [Codigo],fechacompromiso as [Compromiso],LPA, Fechagestion as [Gestion] FROM marcas WHERE LPA = 'L' or LPA= 'P' and fechacompromiso >= #" & fecha1 & "# and fechacompromiso <= #" & fecha2 & "#", Cn, adOpenDynamic, adLockPessimistic
        ElseIf chkLlama.Value = 1 And chkPasa.Value = 0 And chkAbona.Value = 1 Then
            .Open "SELECT codalumno as [Codigo],fechacompromiso as [Compromiso],LPA, Fechagestion as [Gestion] FROM marcas WHERE LPA = 'L'or LPA='A' and fechacompromiso >= #" & fecha1 & "# and fechacompromiso <= #" & fecha2 & "#", Cn, adOpenDynamic, adLockPessimistic
        ElseIf chkLlama.Value = 0 And chkPasa.Value = 1 And chkAbona.Value = 1 Then
            .Open "SELECT codalumno as [Codigo],fechacompromiso as [Compromiso],LPA, Fechagestion as [Gestion] FROM marcas WHERE LPA = 'P'or LPA='A' and fechacompromiso >= #" & fecha1 & "# and fechacompromiso <= #" & fecha2 & "#", Cn, adOpenDynamic, adLockPessimistic
        Else
            .Open "SELECT codalumno as [Codigo],fechacompromiso as [Compromiso],LPA, Fechagestion as [Gestion] FROM marcas WHERE fechacompromiso >= #" & fecha1 & "# and fechacompromiso <= #" & fecha2 & "#", Cn, adOpenDynamic, adLockPessimistic
        End If
    End With

    Set grilla.DataSource = rsMarcas
        formatoGrilla
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdCuotas_Click()
    If Label6.Caption = "" Then
        MsgBox "Primero debe Elegir un Alumno", vbOKOnly + vbInformation, "Marcas"
    Else
        CodAlumno = grilla.Columns(0).Text
        frmAnalisisDeCuotas.Show
        frmAnalisisDeCuotas.Label11.Caption = Me.Name
        Me.Enabled = False
    End If
End Sub

Private Sub cmdDatos_Click()
    If Label6.Caption = "" Then
        MsgBox "Primero debe Elegir un Alumno", vbOKOnly + vbInformation, "Marcas": Exit Sub
    Else
        frmVerificaciones.Label20.Caption = Me.Name

        Verificaciones
        frmVerificaciones.Show
        CodAlumno = grilla.Columns(0).Text
        With rsVerificaciones
            .Requery
            .Find "CodAlumno=" & CodAlumno
            frmVerificaciones.lblCodAlumno.Caption = !CodAlumno
            frmVerificaciones.txtNya.Text = !NyA
            frmVerificaciones.cmbTipoDoc.Text = !tipodoc
            frmVerificaciones.txtDocumento.Text = !dni
            frmVerificaciones.txtDireccion.Text = !direccion
            frmVerificaciones.txtCP.Text = !cp
            frmVerificaciones.dtcLocalidad.Text = !localidad
            frmVerificaciones.txtNacionalidad.Text = !nacionalidad
            
            If Month(!fechanac) < Month(Date) Then
                frmVerificaciones.txtEdad.Text = DateDiff("yyyy", !fechanac, Date)
            ElseIf Day(!fechanac) <= Day(Date) And Month(!fechanac) = Month(Date) Then
                frmVerificaciones.txtEdad.Text = DateDiff("yyyy", !fechanac, Date)
            ElseIf Day(!fechanac) > Day(Date) And Month(!fechanac) >= Month(Date) Then
                frmVerificaciones.txtEdad.Text = DateDiff("yyyy", !fechanac, Date) - 1
            Else
                frmVerificaciones.txtEdad.Text = DateDiff("yyyy", !fechanac, Date) - 1
            End If

            
            frmVerificaciones.dtpFechaNacimiento.Value = !fechanac
            frmVerificaciones.dtcCapacitacion.Text = !capac
            frmVerificaciones.dtcAsistente.Text = !Asistente
            frmVerificaciones.txtTel1.Text = !tel1
            frmVerificaciones.txtTel2.Text = !tel2
            frmVerificaciones.txtTel3.Text = !tel3
            frmVerificaciones.txtTel4.Text = !tel4
            frmVerificaciones.txtPT1.Text = !ptel1
            frmVerificaciones.txtPT2.Text = !ptel2
            frmVerificaciones.txtPT3.Text = !ptel3
            frmVerificaciones.txtPT4.Text = !ptel4
            frmVerificaciones.txtTotalCurso.Text = !totalcurso
            frmVerificaciones.txtTotalCuotas.Text = !cuotas
            frmVerificaciones.txtGastoAdm.Text = !gastoadm
            frmVerificaciones.dtpFechaSuscripcion.Value = !fechasus
            frmVerificaciones.DTPFechaVerificacion.Value = !FechaVerif
            frmVerificaciones.txtObservaciones.Text = !observaciones
            If !manuales = False Then
                frmVerificaciones.chkManuales.Value = 0
            Else
                frmVerificaciones.chkManuales.Value = 1
            End If
            If !dchoexamen = False Then
                frmVerificaciones.chkExamenes.Value = 0
            Else
                frmVerificaciones.chkExamenes.Value = 1
            End If
        End With
        
        If Trim(Len(frmVerificaciones.lblCodAlumno.Caption)) = 1 Then frmVerificaciones.lblCodAlumno.Caption = Format(frmVerificaciones.lblCodAlumno.Caption, "0000#")
        If Trim(Len(frmVerificaciones.lblCodAlumno.Caption)) = 2 Then frmVerificaciones.lblCodAlumno.Caption = Format(frmVerificaciones.lblCodAlumno.Caption, "000##")
        If Trim(Len(frmVerificaciones.lblCodAlumno.Caption)) = 3 Then frmVerificaciones.lblCodAlumno.Caption = Format(frmVerificaciones.lblCodAlumno.Caption, "00###")
        If Trim(Len(frmVerificaciones.lblCodAlumno.Caption)) = 4 Then frmVerificaciones.lblCodAlumno.Caption = Format(frmVerificaciones.lblCodAlumno.Caption, "0####")
    End If
    
    Me.Enabled = False
    BotonMarcar = 2
End Sub

Private Sub cmdMarcar_Click()
    If Label6.Caption = "" Then
        MsgBox "Primero debe Elegir un Alumno", vbOKOnly + vbInformation, "Marcas"
    Else
        CodAlumno = Val(Label6.Caption)
        frmMarcar.Label1.Caption = Me.Name
        frmMarcar.Show
        Me.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Centrar Me
    Marcar
    chkLlama.Value = 0
    chkPasa.Value = 0
    chkAbona.Value = 0
    dtpDesde.Value = Date
    dtpHasta.Value = Date
    Adodc.CursorLocation = adUseClient
    Adodc.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=T:\Base.mdb;Persist Security Info=False;Jet OLEDB:Database Password=ascir"
    Label6.Caption = ""
End Sub

Private Sub grilla_Click()
    Label6.Caption = grilla.Columns(0).Text
End Sub

Sub formatoGrilla()
    Dim w As Integer
    For N = 0 To 3 Step 1
        If N = 0 Then
                w = 800
            ElseIf N = 2 Then
                w = 300
            Else:
                w = 1150
        End If
        grilla.Columns(N).Width = w
        grilla.Columns(N).Alignment = dbgCenter
    Next
End Sub
