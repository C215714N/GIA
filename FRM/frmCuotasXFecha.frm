VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmCuotasXFecha 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuotas Por Fecha"
   ClientHeight    =   6165
   ClientLeft      =   3615
   ClientTop       =   2100
   ClientWidth     =   6765
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
   Icon            =   "frmCuotasXFecha.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   6765
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   5160
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      Height          =   5055
      Left            =   120
      TabIndex        =   4
      Top             =   900
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8916
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
   Begin MSComCtl2.DTPicker dtpDesde 
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   154533889
      CurrentDate     =   41345
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   154533889
      CurrentDate     =   41345
   End
   Begin isButtonTest.isButton cmdBuscar 
      Height          =   420
      Left            =   3000
      TabIndex        =   11
      Top             =   350
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCuotasXFecha.frx":10CA
      Style           =   8
      Caption         =   "     Buscar"
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdMarcar 
      Height          =   420
      Left            =   5280
      TabIndex        =   12
      Top             =   850
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCuotasXFecha.frx":19A4
      Style           =   8
      Caption         =   "     Marcar"
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
   Begin isButtonTest.isButton cmdDatos 
      Height          =   420
      Left            =   5280
      TabIndex        =   13
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCuotasXFecha.frx":227E
      Style           =   8
      Caption         =   "     Datos"
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
   Begin isButtonTest.isButton cmdCuotas 
      Height          =   420
      Left            =   5280
      TabIndex        =   14
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCuotasXFecha.frx":2B58
      Style           =   8
      Caption         =   "     Cuotas"
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
   Begin isButtonTest.isButton cmdCerrar 
      Height          =   420
      Left            =   5280
      TabIndex        =   15
      Top             =   5520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCuotasXFecha.frx":3432
      Style           =   8
      Caption         =   "     Volver"
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
   Begin isButtonTest.isButton cmdExportar 
      Height          =   420
      Left            =   5280
      TabIndex        =   16
      Top             =   4920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCuotasXFecha.frx":3D0C
      Style           =   8
      Caption         =   "     Exportar"
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
   Begin VB.Label lblTotalAlumnos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5280
      TabIndex        =   9
      Top             =   4320
      Width           =   1350
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Alumnos"
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
      Left            =   5280
      TabIndex        =   10
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resta"
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
      Height          =   225
      Left            =   5280
      TabIndex        =   8
      Top             =   3360
      Width           =   435
   End
   Begin VB.Label lblResta 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5280
      TabIndex        =   7
      Top             =   3600
      Width           =   1350
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deuda Total"
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
      Height          =   225
      Left            =   5280
      TabIndex        =   6
      Top             =   2640
      Width           =   960
   End
   Begin VB.Label lblDeudaTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5280
      TabIndex        =   5
      Top             =   2880
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta"
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
      Height          =   225
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
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
      Height          =   225
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   510
   End
End
Attribute VB_Name = "frmCuotasXFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Centrar Me
    dtpDesde.Value = Date
    dtpHasta.Value = Date
End Sub

Private Sub cmdBuscar_Click()
If dtpHasta.Value < dtpDesde.Value Then MsgBox "Ingrese fechas validas", vbCritical, "Gestion Integral del Alumno": dtpDesde.SetFocus: Exit Sub
If DateDiff("m", Date, dtpDesde.Value) > 1 Then MsgBox "No se puede realizar esta consulta", vbCritical, "Gestion Integral del Alumno": dtpDesde.SetFocus: Exit Sub
    
    cmdMarcar.Enabled = True
    cmdCuotas.Enabled = True
    cmdDatos.Enabled = True
    
    Dim fecha1 As Date
    Dim fecha2 As Date
       
    fecha1 = Format(dtpDesde.Value, "mm/dd/yyyy")
    fecha2 = Format(dtpHasta.Value, "mm/dd/yyyy")
    
If DateDiff("m", Date, dtpDesde.Value) = 1 And DateDiff("m", Date, dtpHasta.Value) = 1 Then
    Dim total As Currency
    With rsCuotasXFecha
        If .State = 1 Then .Close
        .Open "SELECT sum(p.deudatotal) FROM plandepago as p INNER JOIN marcas as m ON p.codalumno=m.codalumno WHERE fechavto>=#" & fecha1 & "# and fechavto<=#" & fecha2 & "#and cantidadcuotas=1 and pago=1 union SELECT sum(p.deudatotal) FROM plandepago as p INNER JOIN alumnosdelmes as a ON p.codalumno=a.codalumno WHERE fechavto>=#" & fecha1 & "# and fechavto<=#" & fecha2 & "#", Cn, adOpenDynamic, adLockPessimistic
        total = 0
        .MoveFirst
        Do Until .EOF
            total = total + !expr1000
            .MoveNext
        Loop
        lblDeudaTotal.Caption = Format(total, "currency")
        lblResta.Caption = lblDeudaTotal.Caption
    End With
    With rsCuotasXFecha
        If .State = 1 Then .Close
        .Open "SELECT p.codalumno as Alumno,p.nrocuota as N�, p.fechavto as Vencimiento, p.deudatotal as Deuda FROM plandepago as p INNER JOIN marcas as m on p.codalumno=m.codalumno WHERE fechavto>=#" & fecha1 & "# and fechavto<=#" & fecha2 & "# and cantidadcuotas=1 and pago=1 ORDER BY p.codalumno union SELECT p.codalumno as Alumno,p.nrocuota as N�, p.fechavto as Vencimiento, p.deudatotal as Deuda FROM plandepago as p INNER JOIN alumnosdelmes as a on p.codalumno=a.codalumno WHERE fechavto>=#" & fecha1 & "# and fechavto<=#" & fecha2 & "#", Cn, adOpenDynamic, adLockPessimistic
    End With
    
    lblTotalAlumnos.Caption = rsCuotasXFecha.RecordCount
    Set Grilla.DataSource = rsCuotasXFecha

Else
    With rsCuotasXFecha
        If .State = 1 Then .Close
        .Open "SELECT sum(m.deuda) FROM plandepago as p,marcas as m WHERE fechavto>=#" & fecha1 & "# and fechavto<=#" & fecha2 & "# and nrocuota>1 and p.codalumno=m.codalumno and cantidadcuotas=1", Cn, adOpenDynamic, adLockPessimistic
        lblDeudaTotal.Caption = Format(!expr1000, "currency")
    End With
    
    With rsCuotasXFecha
        If .State = 1 Then .Close
        .Open "SELECT sum(deudatotal) FROM plandepago as p,marcas as m WHERE fechavto>=#" & fecha1 & "# and fechavto<=#" & fecha2 & "# and nrocuota>1 and p.codalumno=m.codalumno and cuotasdebidas=1 and cantidadcuotas=1", Cn, adOpenDynamic, adLockPessimistic
        lblResta.Caption = Format(!expr1000, "currency")
    End With
    With rsCuotasXFecha
        If .State = 1 Then .Close
        .Open "SELECT p.codalumno as Alumno,p.nrocuota as N�, p.fechavto as Vencimiento, p.deudatotal as Deuda,M.fechacompromiso as Compromiso,M.LPA FROM plandepago as p, marcas as m WHERE fechavto>=#" & fecha1 & "# and fechavto<=#" & fecha2 & "# and deudatotal >0 and cantidadcuotas=1 and p.codalumno=m.codalumno ORDER BY p.codalumno", Cn, adOpenDynamic, adLockPessimistic
    End With

    Set Grilla.DataSource = rsCuotasXFecha
    lblTotalAlumnos.Caption = rsCuotasXFecha.RecordCount
    formatoGrilla
End If
End Sub
Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Sub formatoGrilla()
    Dim w As Integer
    For N = 0 To 5 Step 1
        Grilla.Columns(N).Alignment = dbgCenter
        If N = 0 Or N = 3 Then
            w = 750
            ElseIf N = 1 Or N = 5 Then
            w = 300
            Else: w = 1150
        End If
        Grilla.Columns(N).Width = w
    Next
End Sub

Private Sub cmdCuotas_Click()
    CodAlumno = Grilla.Columns(0).Text
    frmAnalisisDeCuotas.Show
    frmAnalisisDeCuotas.Label11.Caption = Me.Name
    Me.Enabled = False
    CuotasXFecha = True
End Sub

Private Sub cmdDatos_Click()
    frmVerificaciones.Label20.Caption = Me.Name
    
    Verificaciones
    frmVerificaciones.Show
    CodAlumno = Grilla.Columns(0).Text
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

    Me.Enabled = False

End Sub

Private Sub cmdExportar_Click()
    Call Exportar_Datagrid(Grilla.ApproxCount)
End Sub

Private Sub Exportar_Datagrid(TotalFilas As Long)
    Me.MousePointer = vbHourglass
    Set obj_excel = CreateObject("Excel.Application")
    Set obj_Libro = obj_excel.workbooks.Open("T:\AlumnosDelMes.xls")
    Set obj_Hoja = obj_excel.ActiveSheet
       
    Columna = 0
    For X = 0 To Grilla.Columns.Count - 1
        If Grilla.Columns(X).Visible Then
            Columna = Columna + 1
            obj_Hoja.Cells(1, Columna) = Grilla.Columns(X).Caption
            For Y = 0 To TotalFilas - 1
                obj_Hoja.Cells(Y + 2, Columna) = Grilla.Columns(X).CellValue(Grilla.GetBookmark(Y))
            Next
        End If
    Next
    obj_excel.Visible = True
    With obj_Hoja
        .Columns("A:Z").autofit
    End With
    
    Me.MousePointer = vbDefault
    Set obj_Hoja = Nothing
    Set obj_Libro = Nothing
    Set obj_excel = Nothing
End Sub

Private Sub cmdMarcar_Click()
    CodAlumno = Grilla.Columns(0).Text
    frmMarcar.Label1.Caption = Me.Name
    frmMarcar.Show
    Me.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    CuotasXFecha = False
End Sub
