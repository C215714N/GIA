VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmExamenes 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Examenes"
   ClientHeight    =   4380
   ClientLeft      =   11445
   ClientTop       =   1770
   ClientWidth     =   5715
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmExamenes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5715
   Begin VB.Frame Frame1 
      BackColor       =   &H00662200&
      Caption         =   "Examen"
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
      Height          =   3975
      Left            =   3960
      TabIndex        =   12
      Top             =   240
      Width           =   1600
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   360
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
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
         Format          =   125763585
         CurrentDate     =   41978
      End
      Begin VB.TextBox txtPromedio 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   5
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtPractico 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtTeorico 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox cmbModulo 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
      End
      Begin isButtonTest.isButton cmdAgregar 
         Height          =   420
         Left            =   120
         TabIndex        =   19
         Top             =   3360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmExamenes.frx":10CA
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
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Promedio"
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
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Practico"
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
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Teorico"
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
         Left            =   120
         TabIndex        =   14
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
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   2775
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4895
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
      TabIndex        =   7
      Top             =   960
      Width           =   3735
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
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
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
   Begin VB.Label lblOrigen 
      Height          =   495
      Left            =   7080
      TabIndex        =   18
      Top             =   240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00662200&
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
      TabIndex        =   11
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00662200&
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
      TabIndex        =   10
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00662200&
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
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmExamenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Centrar Me
    DTPFecha.Value = Date
End Sub

Private Sub txtCodigo_KeyPress(keyAscii As Integer)
    On Error GoTo LineaError
    If keyAscii = 13 Then
        If txtCodigo.text = "" Then MsgBox "Ingrese el Codigo del alumno", vbOKOnly, "GIA - Examenes": txtCodigo.SetFocus: Exit Sub
        With rsVerificaciones
            If .State = 1 Then .Close
            .Open "SELECT nya,capac FROM verificaciones WHERE codalumno=" & Int(txtCodigo.text), Cn, adOpenDynamic, adLockPessimistic
                If .BOF = True And .EOF = True Then
                    MsgBox "No se encuentra el Codigo de Alumno" & vbNewLine & "Controle que el codigo ingresado sea correcto", vbExclamation, "Gestion Integral del Alumno - Gestion Integral del Alumno"
                ElseIf .BOF = False Or .EOF = False Then
                    txtAlumno.text = !NyA
                    txtCurso.text = !capac
                End If
        End With
        With rsExamenes
            If .State = 1 Then .Close
            .Open "SELECT Fecha, Modulo, Teorico as [T], Practico as [P], Promedio as [F] FROM examenes WHERE codalumno=" & Int(txtCodigo.text) & " ORDER BY fecha,id", Cn, adOpenDynamic, adLockPessimistic
        End With
        Set grilla.DataSource = rsExamenes
        formatoGrilla
        CargarModulos
        If rsExamenes.RecordCount = cmbModulo.ListCount Then
            cmdAgregar.Enabled = False
        Else: cmdAgregar.Enabled = True
        End If
    End If
LineaError: ErrCode
End Sub

Private Sub cmbModulo_Change()
    txtTeorico.SetFocus
End Sub

Private Sub cmdAgregar_Click()
    On Error GoTo LineaError
    If cmbModulo.text = "" Then MsgBox "Elija el modulo", vbOKOnly + vbCritical, "GIA - Examenes": cmbModulo.SetFocus: Exit Sub
    If txtTeorico.text = "" Then MsgBox "Ingrese nota del Examen teorico", vbOKOnly + vbCritical, "GIA - Examenes": txtTeorico.SetFocus: Exit Sub
    If txtPractico.text = "" Then MsgBox "Ingrese nota del Examen Practico", vbOKOnly + vbCritical, "GIA - Examenes": txtPractico.SetFocus: Exit Sub
    If txtPromedio.text = "" Then MsgBox "Ingrese nota promedio", vbOKOnly + vbCritical, "GIA - Examenes": txtPromedio.SetFocus: Exit Sub
        
''' CARGAR EXAMEN - TABLA EXAMENES
    With rsExamenes
        .Find "Modulo='" & cmbModulo.text & "'"
        If .BOF Or .EOF Then
            .Close
            .Open "SELECT CodAlumno, Fecha, Modulo, Teorico as [T], Practico as [P], Promedio as [F]  FROM examenes", Cn, adOpenDynamic, adLockPessimistic
            .Requery
            .AddNew
            !CodAlumno = Int(txtCodigo.text)
            !fecha = DTPFecha.Value
            !T = txtTeorico.text
            !P = txtPractico.text
            !F = txtPromedio.text
            !modulo = cmbModulo.text
            .Update
            .Close
    '''CONTROL EXAMEN (MODULOS CARGADOS)
        Else:
            MsgBox "El alumno ya ha rendido este modulo", vbCritical, "Examenes"
            txtPromedio.text = ""
            txtTeorico.text = ""
            txtPractico.text = ""
            cmbModulo.SetFocus
            Exit Sub
        End If
        .Open "SELECT Fecha, Modulo, Teorico as [T], Practico as [P], Promedio as [F] FROM examenes WHERE codalumno=" & Int(txtCodigo.text) & " ORDER BY fecha,id", Cn, adOpenDynamic, adLockPessimistic
        Set grilla.DataSource = rsExamenes
        formatoGrilla
    
    '''ALUMNO EGRESADO
        If rsExamenes.RecordCount = cmbModulo.ListCount Then
            cmdAgregar.Enabled = False
            Egresado = True
        Else:
            cmdAgregar.Enabled = True
            Egresado = False
        End If
    '''CARGAR ALUMNO - TABLA EGRESADOS
        If Egresado = True Then
            MsgBox "El alumno " & txtAlumno.text & " ha egresado del curso de " & txtCurso.text, vbInformation, "Examenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.text)
                !fecha = DTPFecha.Value
                .Update
            End With
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
        End If
    End With
    
''' REESTABLECE EL FORMULARIO
    txtPractico.text = ""
    txtTeorico.text = ""
    txtPromedio.text = ""
    Egresado = False

LineaError: ErrCode
End Sub

Private Sub txtTeorico_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub

Private Sub txtPractico_KeyPress(keyAscii As Integer)
    If keyAscii = 13 Then
        txtPromedio.text = (Int(txtTeorico.text) + Int(txtPractico.text)) / 2
        Sendkeys "{TAB}"
    End If
End Sub

Private Sub txtPromedio_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lblOrigen.Caption = "Egresados" Then frmEgresados.Enabled = True
End Sub
Private Sub CargarModulos()
    If txtCurso.text = "Operador de Pc" Then
        With cmbModulo
            .Clear
            .AddItem ("Windows")
            .AddItem ("Word")
            .AddItem ("Excel")
            .AddItem ("Access")
            .AddItem ("Power Point")
        End With
        
    ElseIf txtCurso.text = "Diseño Gráfico" Then
        With cmbModulo
            .Clear
            .AddItem ("Windows")
            .AddItem ("Corel Draw")
            .AddItem ("Photoshop")
            .AddItem ("Page Maker")
        End With
        
    ElseIf txtCurso.text = "Diseño Web" Then
        With cmbModulo
            .Clear
            .AddItem ("Front Page")
            .AddItem ("Fireworks")
            .AddItem ("Flash")
            .AddItem ("Dreamweaver")
        End With
            
    ElseIf txtCurso.text = "Programación + Access" Then
        With cmbModulo
            .Clear
            .AddItem ("Access")
            .AddItem ("Modulo I")
            .AddItem ("Modulo II")
        End With
        
    ElseIf txtCurso.text = "Programación" Or txtCurso.text = "Telefonía Celular" Then
        With cmbModulo
            .Clear
            .AddItem ("Modulo I")
            .AddItem ("Modulo II")
        End With
        
    ElseIf txtCurso.text = "Técnico en aire acondicionado" Or txtCurso.text = "Electricidad domiciliaria" Then
        With cmbModulo
            .Clear
            .AddItem ("Modulo I")
            .AddItem ("Modulo II")
            .AddItem ("Modulo III")
            .AddItem ("Final")
        End With
    
    ElseIf txtCurso.text = "Soporte Tecnico" Then
        With cmbModulo
            .Clear
            .AddItem ("Modulo I")
            .AddItem ("Modulo II")
            .AddItem ("Modulo III")
            .AddItem ("Modulo IV")
            .AddItem ("Modulo V")
            .AddItem ("Examen Final")
        End With
        
    ElseIf txtCurso.text = "Cuidador Domiciliario" Or txtCurso.text = "Asistente Terapeutico" Or txtCurso.text = "Auxiliar de Farmacia" Then
        With cmbModulo
            .Clear
            .AddItem ("Parcial I")
            .AddItem ("Parcial II")
            .AddItem ("Parcial III")
            .AddItem ("Final")
        End With
    
    ElseIf txtCurso.text = "Emergencias Médicas" Or txtCurso.text = "Extracc. Adm. Y Asist. Tec. Laborat." Then
        With cmbModulo
            .Clear
            .AddItem ("Parcial I")
            .AddItem ("Parcial II")
            .AddItem ("Final")
        End With
    End If
End Sub

Sub formatoGrilla()
    Dim w As Integer
    For N = 0 To 4 Step 1
        If N < 2 Then
            w = N * 2000
        Else: w = 350
        End If
        grilla.Columns(N).Width = w
    Next
End Sub
