VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmDerechosExamenes 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Derechos de Examen"
   ClientHeight    =   4725
   ClientLeft      =   7245
   ClientTop       =   2280
   ClientWidth     =   5730
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmDerechosExamenes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   5730
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtAlumno 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   11
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox txtCurso 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00662200&
      Caption         =   "Derecho Examen"
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
      Height          =   4335
      Left            =   3960
      TabIndex        =   6
      Top             =   240
      Width           =   1635
      Begin VB.TextBox txtPrecio 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         TabIndex        =   4
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ComboBox cmbPago 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmDerechosExamenes.frx":10CA
         Left            =   150
         List            =   "frmDerechosExamenes.frx":10D7
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtRecibo 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         TabIndex        =   3
         Top             =   2280
         Width           =   1335
      End
      Begin VB.ComboBox cmbModulo 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmDerechosExamenes.frx":10F9
         Left            =   150
         List            =   "frmDerechosExamenes.frx":10FB
         TabIndex        =   1
         Top             =   1080
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   375
         Left            =   150
         TabIndex        =   5
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   169082881
         CurrentDate     =   41978
      End
      Begin isButtonTest.isButton cmdAgregar 
         Height          =   420
         Left            =   150
         TabIndex        =   18
         Top             =   3300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmDerechosExamenes.frx":10FD
         Style           =   8
         Caption         =   "     Aceptar"
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
      Begin isButtonTest.isButton cmdExamenes 
         Height          =   420
         Left            =   150
         TabIndex        =   19
         Top             =   3800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmDerechosExamenes.frx":19D7
         Style           =   8
         Caption         =   "     Examen"
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
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Recibo"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   150
         TabIndex        =   16
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Pago"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   150
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Modulo"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   150
         TabIndex        =   8
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   150
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   3135
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
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
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo"
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
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Alumno"
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
      Left            =   1080
      TabIndex        =   13
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Capacitacion"
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
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmDerechosExamenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Centrar Me
    Control
    txtPrecio.Text = Format(rsControl!derechoExamen, "currency")
    DTPFecha.Value = Date
End Sub

Private Sub txtCodigo_KeyPress(keyAscii As Integer)
    If keyAscii = 13 Then
        If txtCodigo.Text = "" Then MsgBox "Ingrese el codigo del alumno", vbOKOnly, "GIA - Examenes": txtCodigo.SetFocus: Exit Sub
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT nya,capac FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                If .BOF = True And .EOF = True Then
                    MsgBox "No se encuentra el Codigo de Alumno" & vbNewLine & "Controle que el codigo ingresado sea correcto", vbExclamation, "Gestion Integral del Alumno - Gestion Integral del Alumno"
                ElseIf .BOF = False Or .EOF = False Then
                    txtAlumno.Text = !NyA
                    txtCurso.Text = !capac
                End If
            End With
            With rsDerechosExamenes
                If .State = 1 Then .Close
                .Open "SELECT Fecha, Modulo FROM derechoexamen WHERE codalumno=" & Int(txtCodigo.Text) & " ORDER BY fecha", Cn, adOpenDynamic, adLockPessimistic
            End With
            
            Set grilla.DataSource = rsDerechosExamenes
            formatoGrilla
            CargarModulos
            txtPrecio.Text = Format(rsControl!derechoExamen, "currency")
            cmbModulo.Enabled = True
            DTPFecha.Enabled = True
            cmdAgregar.Enabled = True
            cmbModulo.SetFocus
            
            If rsDerechosExamenes.RecordCount >= 1 Then
                cmdExamenes.Enabled = True
            Else: cmdExamenes.Enabled = False
            End If
        End If
End Sub

Private Sub cmdAgregar_Click()
    On Error GoTo LineaError
    If cmbModulo.Text = "" Then MsgBox "Elija el modulo", vbOKOnly + vbCritical, "GIA - Examenes": cmbModulo.SetFocus: Exit Sub
    If cmbPago.Text = "" Then MsgBox "Elija el tipo de pago", vbOKOnly + vbCritical, "GIA - Examenes": cmbPago.SetFocus: Exit Sub
    If txtRecibo.Text = "" Then MsgBox "Ingrese el numero de recibo", vbOKOnly + vbCritical, "GIA - Examenes": txtRecibo.SetFocus: Exit Sub
'''CONSULTA - TABLA DERECHOS DE EXAMEN
    With rsDerechosExamenes
        .Close
        .Open "SELECT * FROM derechoexamen", Cn, adOpenDynamic, adLockPessimistic
        .Requery
        .AddNew
        !CodAlumno = Int(txtCodigo.Text)
        !fecha = DTPFecha.Value
        !modulo = cmbModulo.Text
        .Update
        .Close
        .Open "SELECT Fecha, Modulo as Modulo FROM derechoexamen WHERE codalumno=" & Int(txtCodigo.Text) & " ORDER BY fecha", Cn, adOpenDynamic, adLockPessimistic
        Set grilla.DataSource = rsDerechosExamenes
        formatoGrilla
    End With
'''GESTION CONTABLE - ASIENTO
    With rsContabilidad
        If .State = 1 Then .Close
        .Open "SELECT * FROM contabilidad", Cn, adOpenDynamic, adLockPessimistic
        .Requery
        .AddNew
        !fecha = Date
        !asiento = Null
        !NroCuota = Null
        !CodAlumno = Null
        !Cuenta = "DERECHO DE EXAMEN"
        !Detalle = txtAlumno.Text & " - " & cmbModulo.Text
        !nrofactura = txtRecibo.Text
        !Haber = CSng(txtPrecio.Text)
        !Debe = Null
        .Update
        .Requery
        .AddNew
        !fecha = Date
        
        If cmbPago.Text = "Efectivo" Then
            !Cuenta = "CAJA ADMINISTRACION"
        ElseIf cmbPago.Text = "Descuento" Then
            !Cuenta = "Descuento"
        Else
            !Cuenta = "DEBITO TARJETA CREDITO"
        End If
        !Detalle = txtAlumno.Text & " - Derecho de Examen de " & cmbModulo.Text
        !nrofactura = txtRecibo.Text
        !Debe = CSng(txtPrecio.Text)
        !asiento = Null
        !NroCuota = Null
        !CodAlumno = Null
        !Haber = Null
        .Update
    End With
LineaError: ErrCode Err
End Sub

Private Sub txtRecibo_KeyPress(keyAscii As Integer)
    continue keyAscii
End Sub

Private Sub cmbModulo_KeyPress(keyAscii As Integer)
    With txtPrecio
        If cmbModulo.Text = "Final" Then
            .Text = Format(rsControl!examenFinal, "currency")
        ElseIf cmbModulo.Text = "Modulo Final" Then
            .Text = Format(rsControl!moduloFinal, "currency")
        Else:
            .Text = Format(rsControl!derechoExamen, "currency")
        End If
    End With
    continue keyAscii
End Sub

Private Sub cmbPago_KeyPress(keyAscii As Integer)
    continue keyAscii
End Sub

Private Sub cmdexamenes_Click()
    frmExamenes.Show
    frmExamenes.txtCodigo.Text = txtCodigo.Text
End Sub

Private Sub CargarModulos()
    If txtCurso.Text = "Operador de Pc" Then
        With cmbModulo
            .Clear
            .AddItem ("Windows")
            .AddItem ("Word")
            .AddItem ("Excel")
            .AddItem ("Access")
            .AddItem ("Power Point")
        End With
        
    ElseIf txtCurso.Text = "Dise�o Gr�fico" Then
        With cmbModulo
            .Clear
            .AddItem ("Windows")
            .AddItem ("Corel Draw")
            .AddItem ("Photoshop")
            .AddItem ("Page Maker")
        End With
        
    ElseIf txtCurso.Text = "Dise�o Web" Then
        With cmbModulo
            .Clear
            .AddItem ("Front Page")
            .AddItem ("Fireworks")
            .AddItem ("Flash")
            .AddItem ("Dreamweaver")
        End With
            
    ElseIf txtCurso.Text = "Programaci�n + Access" Then
        With cmbModulo
            .Clear
            .AddItem ("Access")
            .AddItem ("Modulo I")
            .AddItem ("Modulo II")
        End With
        
    ElseIf txtCurso.Text = "Programaci�n" Then
        With cmbModulo
            .Clear
            .AddItem ("Modulo I")
            .AddItem ("Modulo II")
        End With
        
    ElseIf txtCurso.Text = "Telefon�a Celular" Then
        With cmbModulo
            .Clear
            .AddItem ("Modulo I")
            .AddItem ("Modulo II")
            .AddItem ("Final")
        End With
        
    ElseIf txtCurso.Text = "T�cnico en aire acondicionado" Or txtCurso.Text = "Electricidad domiciliaria" Then
        With cmbModulo
            .Clear
            .AddItem ("Modulo I")
            .AddItem ("Modulo II")
            .AddItem ("Modulo III")
            .AddItem ("Final")
        End With
    
    ElseIf txtCurso.Text = "Soporte Tecnico" Then
        With cmbModulo
            .Clear
            .AddItem ("Modulo I")
            .AddItem ("Modulo II")
            .AddItem ("Modulo III")
            .AddItem ("Modulo IV")
            .AddItem ("Modulo V")
            .AddItem ("Examen Final")
        End With
        
    ElseIf txtCurso.Text = "Cuidador Domiciliario" Or txtCurso.Text = "Asistente Terapeutico" Or txtCurso.Text = "Auxiliar de Farmacia" Or txtCurso.Text = "Emergencias M�dicas" Or txtCurso.Text = "Emergencias Medicas Sanitarias" Or txtCurso.Text = "Extracc. Adm. Y Asist. Tec. Laborat." Then
        With cmbModulo
            .Clear
            .AddItem ("Parcial I")
            .AddItem ("Parcial II")
            .AddItem ("Parcial III")
            .AddItem ("Final")
        End With
    
    ElseIf txtCurso.Text = "Mandatario Automotor" Then
        With cmbModulo
            .Clear
            .AddItem ("Parcial I")
            .AddItem ("Parcial II")
            .AddItem ("Parcial III")
            .AddItem ("Modulo Final")
        End With
    ElseIf txtCurso.Text = "Asistente en Cardiolog�a" Then
        With cmbModulo
            .Clear
            .AddItem ("Final")
        End With
    Else
        With cmbModulo
            .Clear
            .AddItem ("Modulo I")
            .AddItem ("Modulo II")
            .AddItem ("Modulo III")
            .AddItem ("Final")
        End With
    End If
End Sub

Sub formatoGrilla()
    For N = 0 To 1
        grilla.Columns(N).Width = 1150 + (N * 800)
    Next
End Sub
