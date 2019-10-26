VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmAnalisisDeCuotas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Análisis de Cuotas"
   ClientHeight    =   5130
   ClientLeft      =   3885
   ClientTop       =   1845
   ClientWidth     =   8955
   Icon            =   "frmAnalisisDeCuotas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmAnalisisDeCuotas.frx":324A
   ScaleHeight     =   5130
   ScaleWidth      =   8955
   Begin isButtonTest.isButton cmdDatos 
      Height          =   420
      Left            =   7500
      TabIndex        =   11
      Top             =   2700
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmAnalisisDeCuotas.frx":AC67
      Style           =   8
      Caption         =   "       Datos"
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
   Begin isButtonTest.isButton cmdReingresar 
      Height          =   420
      Left            =   7500
      TabIndex        =   12
      Top             =   3900
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmAnalisisDeCuotas.frx":B541
      Style           =   8
      Caption         =   "       Reingresar"
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
   Begin isButtonTest.isButton cmdEgresado 
      Height          =   420
      Left            =   7500
      TabIndex        =   13
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmAnalisisDeCuotas.frx":BE1B
      Style           =   8
      Caption         =   "       Egresado"
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
   Begin isButtonTest.isButton cmdEditar 
      Height          =   420
      Left            =   7500
      TabIndex        =   14
      Top             =   3300
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmAnalisisDeCuotas.frx":C6F5
      Style           =   8
      Caption         =   "       Editar"
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
   Begin isButtonTest.isButton cmdBaja 
      Height          =   420
      Left            =   7500
      TabIndex        =   15
      Top             =   4500
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmAnalisisDeCuotas.frx":CFCF
      Style           =   8
      Caption         =   "       Baja"
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
   Begin VB.TextBox txtObservaciones 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   100
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3840
      Width           =   7300
   End
   Begin MSDataGridLib.DataGrid grilla2 
      Height          =   2895
      Left            =   3800
      TabIndex        =   1
      Top             =   840
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   21
      RowDividerStyle =   0
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9
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
      Caption         =   "HISTÓRICO"
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
   Begin MSDataGridLib.DataGrid grilla1 
      Height          =   2895
      Left            =   100
      TabIndex        =   0
      Top             =   825
      Width           =   3600
      _ExtentX        =   6350
      _ExtentY        =   5106
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   21
      RowDividerStyle =   0
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9
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
      Caption         =   "PLAN DE PAGO"
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
   Begin isButtonTest.isButton cmdBuscar 
      Height          =   420
      Left            =   7500
      TabIndex        =   8
      Top             =   900
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmAnalisisDeCuotas.frx":D8A9
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
   Begin isButtonTest.isButton cmdGrabar 
      Height          =   420
      Left            =   7500
      TabIndex        =   9
      Top             =   1500
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmAnalisisDeCuotas.frx":E183
      Style           =   8
      Caption         =   "       Aceptar"
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
   Begin isButtonTest.isButton cmdMarcar 
      Height          =   420
      Left            =   7500
      TabIndex        =   10
      Top             =   2100
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmAnalisisDeCuotas.frx":EA5D
      Style           =   8
      Caption         =   "       Marcar"
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
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alumno"
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
      Left            =   1080
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
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
      TabIndex        =   6
      Top             =   90
      Width           =   585
   End
   Begin VB.Label Label11 
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
      Left            =   7500
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblNya 
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
      Height          =   360
      Left            =   1080
      TabIndex        =   4
      Top             =   345
      Width           =   6300
   End
   Begin VB.Label lblCodAlumno 
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
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   345
      Width           =   855
   End
End
Attribute VB_Name = "frmAnalisisDeCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Compare Text


Private Sub cmdBaja_Click()
    a = MsgBox("¿Está seguro que desea dar la baja de este alumno?", vbYesNo + vbQuestion, "Análisis de Cuotas")
    If a = vbYes Then
        frmBajas.Show
        frmAnalisisDeCuotas.Enabled = False
        frmBajas.txtmotivo.Text = ""
        frmBajas.SetFocus
        
        '''Baja en plan de pago
        With rsPlanDePago
            If .State = 1 Then .Close
            .Open "SELECT * FROM plandepago WHERE codalumno=" & CodAlumno, Cn, adOpenDynamic, adLockPessimistic
            .MoveFirst
            Do Until .EOF
                If !tipodepago = "PAG" Then
                    .MoveNext
                ElseIf !tipodepago = "Par" Then
                    .MoveNext
                Else
                    !tipodepago = "BAJA"
                    !fechapago = Date
                    !DeudaTotal = 0
                    !CuotasDebidas = 0
                    .UpdateBatch
                    .MoveNext
                End If
            Loop
        End With

        ''' actualiza la grilla
        AnalisisDeCuota
        Set grilla1.DataSource = rsAnalisisDeCuenta
        formatoGrilla
    End If
    
    
End Sub

Private Sub cmdBuscar_Click()
    Unload Me
    Analisis = True
    frmBuscarVerificacion.Show
End Sub

Private Sub cmdDatos_Click()
    frmVerificaciones.Label20.Caption = Me.Name
    Me.Enabled = False
    Verificaciones
    frmVerificaciones.Show
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

    Analisis = True

End Sub

Private Sub cmdEditar_Click()
    frmModificarPlanDePago.Show
    Me.Enabled = False
End Sub

Private Sub cmdEgresado_Click()
    If MsgBox("Confirma que el alumno ha egresado?", vbQuestion + vbYesNo, "Análisis de Cuotas") = vbYes Then
        With rsVerificaciones
            If .State = 1 Then .Close
            .Open "SELECT codalumno as [Codigo], estado FROM verificaciones WHERE codalumno=" & frmAnalisisDeCuotas.lblCodAlumno.Caption, Cn, adOpenDynamic, adLockPessimistic
            .Requery
            .MoveFirst
            !estado = "Egresado"
            .UpdateBatch
        End With
    End If
End Sub

Private Sub cmdGrabar_Click()
    On Error GoTo LineaError
    With rsAnalisisDeCuenta
        !observaciones = txtObservaciones.Text
        .UpdateBatch
    End With

    With rsMarcar
        .Requery
        .Find "Codalumno=" & CodAlumno
            !fechagestion = Date
            .UpdateBatch
    End With

    cmdGrabar.Enabled = False

LineaError:
    Select Case Err.Number
        Case 3021
            Resume Next
        End Select
End Sub

Private Sub cmdMarcar_Click()
    frmMarcar.Label1.Caption = Me.Name
    frmMarcar.Show
    Me.Enabled = False
End Sub


Private Sub cmdReingresar_Click()
    grilla1.Row = rsAnalisisDeCuenta.RecordCount - 1
    frmPlanDePagoReingreso.txtNroCuota = Int(grilla1.Columns(0).Text) + 1
    frmPlanDePagoReingreso.txtCantidadCuotas = 1
    frmPlanDePagoReingreso.DTPFecha.Value = grilla1.Columns(1).Text
    frmPlanDePagoReingreso.txtMonto.Text = grilla1.Columns(3).Text
    
    If frmPlanDePagoReingreso.DTPFecha.Month = 12 Then
            frmPlanDePagoReingreso.DTPFecha.Month = 1
            frmPlanDePagoReingreso.DTPFecha.Year = frmPlanDePagoReingreso.DTPFecha.Year + 1
    Else
        frmPlanDePagoReingreso.DTPFecha.Month = frmPlanDePagoReingreso.DTPFecha.Month + 1
    End If
    
    frmPlanDePagoReingreso.Show
    Me.Enabled = False
End Sub

Private Sub Form_Load()
    Centrar Me
    AnalisisDeCuota
    If rsAnalisisDeCuenta.BOF Or rsAnalisisDeCuenta.EOF Then MsgBox "El alumno seleccionado no tiene creado un Plan de Pago", vbCritical, "Análisis de Cuotas": Exit Sub
    Historico
    Marcar
    Set grilla1.DataSource = rsAnalisisDeCuenta
    Set grilla2.DataSource = rsHistorico
    formatoGrilla
    lblCodAlumno.Caption = rsAnalisisDeCuenta!codigo
    lblNyA.Caption = rsAnalisisDeCuenta!Alumno
    If Trim(Len(lblCodAlumno.Caption)) = 1 Then lblCodAlumno.Caption = Format(lblCodAlumno.Caption, "0000#")
    If Trim(Len(lblCodAlumno.Caption)) = 2 Then lblCodAlumno.Caption = Format(lblCodAlumno.Caption, "000##")
    If Trim(Len(lblCodAlumno.Caption)) = 3 Then lblCodAlumno.Caption = Format(lblCodAlumno.Caption, "00###")
    If Trim(Len(lblCodAlumno.Caption)) = 4 Then lblCodAlumno.Caption = Format(lblCodAlumno.Caption, "0####")
End Sub

Sub formatoGrilla()
    Dim w As Integer
    For N = 0 To 6
        If N = 0 Then
            w = 300
        ElseIf N = 1 Or N = 2 Then
            w = 1150
        ElseIf N = 3 Then
            w = 800
            grilla1.Columns(N).NumberFormat = "$ #####"
            grilla2.Columns(N).NumberFormat = "$ #####"
            grilla1.Columns(N - 1).Width = w
            grilla2.Columns(N - 2).Width = w
        Else:
            w = 0
        End If
            grilla1.Columns(N).Width = w
        If N <= 3 Then
            grilla1.Columns(N).Alignment = dbgCenter
            grilla2.Columns(N).Alignment = dbgCenter
            grilla2.Columns(N).Width = w
        End If
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Label11.Caption = "frmCuotasXFecha" Then
        frmCuotasXFecha.Enabled = True
        rsCuotasXFecha.Requery
        frmCuotasXFecha.formatoGrilla
    ElseIf Label11.Caption = "frmAnalisisSituacion" Then
        frmAnalisisSituacion.Enabled = True
        rsAnalisisSituacionDeDeuda.Requery
        frmAnalisisSituacion.grilla.Columns(2).Width = 800
    ElseIf Label11.Caption = "frmMarcas" Then
        frmMarcas.Enabled = True
        rsMarcas.Requery
        frmMarcas.grilla.Columns(2).Width = 400
    End If

End Sub

Private Sub grilla1_Click()
    txtObservaciones.Text = grilla1.Columns(4).Text
End Sub

Private Sub txtObservaciones_Change()
    If cmdGrabar.Visible = True Then
        cmdGrabar.Enabled = True
    End If
End Sub
