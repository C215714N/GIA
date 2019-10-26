VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmAnalisisSituacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Análisis Situación de Deuda"
   ClientHeight    =   4095
   ClientLeft      =   6795
   ClientTop       =   2775
   ClientWidth     =   6390
   Icon            =   "frmAnalisisSituacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmAnalisisSituacion.frx":324A
   ScaleHeight     =   4095
   ScaleWidth      =   6390
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   3800
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6694
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
   Begin VB.TextBox txtResta 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
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
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   3480
      Width           =   1335
   End
   Begin isButtonTest.isButton cmdMarcar 
      Height          =   420
      Left            =   4900
      TabIndex        =   5
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmAnalisisSituacion.frx":AC67
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
      Left            =   4900
      TabIndex        =   6
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmAnalisisSituacion.frx":B541
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
      Left            =   4900
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmAnalisisSituacion.frx":BE1B
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
      Left            =   4900
      TabIndex        =   8
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmAnalisisSituacion.frx":C6F5
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
   Begin VB.Label Label3 
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
      Left            =   4920
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resta"
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
      Left            =   4920
      TabIndex        =   1
      Top             =   3240
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Analizado"
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
      Left            =   4920
      TabIndex        =   0
      Top             =   2520
      Width           =   810
   End
End
Attribute VB_Name = "frmAnalisisSituacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdCuotas_Click()
        CodAlumno = grilla.Columns(0).Text
        frmAnalisisDeCuotas.Show
        frmAnalisisDeCuotas.Label11.Caption = Me.Name
        Me.Enabled = False

End Sub

Private Sub cmdDatos_Click()
    frmVerificaciones.Label20.Caption = Me.Name
    CodAlumno = grilla.Columns(0).Text
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

    BotonMarcar = 4
End Sub

Private Sub cmdMarcar_Click()
    frmMarcar.Label1.Caption = Me.Name
    CodAlumno = grilla.Columns(0).Text
    frmMarcar.Show
    Me.Enabled = False

End Sub

Private Sub Form_Load()
    Centrar Me
    Dim CuotasDebidas As Integer
    CuotasDebidas = (Situacion + 30) / 30
    
    If Situacion = 0 Then
        Label1.Caption = "Analizado " & Situacion
    Else
        Label1.Caption = "Analizado - " & Situacion
    End If
    
    With rsAnalisisSituacionDeDeuda
        If .State = 1 Then .Close
        .Open "SELECT codalumno as Alumno, Cuota, Deuda, fechaCompromiso as Compromiso, LPA FROM marcas WHERE cantidadcuotas=" & CuotasDebidas & " and pago=0 ORDER BY codalumno", Cn, adOpenDynamic, adLockPessimistic
    End With
    Set grilla.DataSource = rsAnalisisSituacionDeDeuda
    formatoGrilla
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmSituacionDeCartera.Enabled = True
End Sub

Sub formatoGrilla()
    Dim w As Integer
    For N = 0 To 3 Step 1
        If N = 0 Or N = 2 Then
            w = 800
        Else:
            w = 400 * N
        End If
        grilla.Columns(N).Width = w
    Next
End Sub
