VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmExamenes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exámenes"
   ClientHeight    =   4380
   ClientLeft      =   11445
   ClientTop       =   1770
   ClientWidth     =   5715
   Icon            =   "frmExamenes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmExamenes.frx":324A
   ScaleHeight     =   4380
   ScaleWidth      =   5715
   Begin VB.Frame Frame1 
      BackColor       =   &H00662200&
      Caption         =   "Examen"
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   85393409
         CurrentDate     =   41978
      End
      Begin VB.TextBox txtPromedio 
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
         TabIndex        =   5
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox txtPractico 
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
         TabIndex        =   4
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtTeorico 
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
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ComboBox cmbModulo 
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
         Icon            =   "frmExamenes.frx":AC67
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
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
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
            Size            =   8.25
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
         Caption         =   "Práctico"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
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
         Caption         =   "Teórico"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
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
         Caption         =   "Módulo"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
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
      RowHeight       =   21
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
   Begin VB.TextBox txtCurso 
      Enabled         =   0   'False
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
      TabIndex        =   7
      Top             =   960
      Width           =   3735
   End
   Begin VB.TextBox txtAlumno 
      Enabled         =   0   'False
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
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   1  'Right Justify
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
      Caption         =   "Capacitación"
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
      TabIndex        =   11
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00662200&
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
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00662200&
      BackStyle       =   0  'Transparent
      Caption         =   "Código"
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

Private Sub cmbModulo_Change()
    txtTeorico.SetFocus
End Sub

Private Sub cmdAgregar_Click()
    If cmbModulo.Text = "" Then MsgBox "Elija el módulo", vbOKOnly + vbCritical, "GIA - Exámenes": cmbModulo.SetFocus: Exit Sub
    If txtTeorico.Text = "" Then MsgBox "Ingrese nota del examen teórico", vbOKOnly + vbCritical, "GIA - Exámenes": txtTeorico.SetFocus: Exit Sub
    If txtPractico.Text = "" Then MsgBox "Ingrese nota del examen práctico", vbOKOnly + vbCritical, "GIA - Exámenes": txtPractico.SetFocus: Exit Sub
    If txtPromedio.Text = "" Then MsgBox "Ingrese nota promedio", vbOKOnly + vbCritical, "GIA - Exámenes": txtPromedio.SetFocus: Exit Sub
    
    With rsExamenes
        .Find "Modulo='" & cmbModulo.Text & "'"
        If .BOF Or .EOF Then
        .Close
        .Open "SELECT CodAlumno, Fecha, Modulo, Teorico as [T], Practico as [P], Promedio as [F]  FROM examenes", Cn, adOpenDynamic, adLockPessimistic
        .Requery
        .AddNew
        !CodAlumno = Int(txtCodigo.Text)
        !fecha = DTPFecha.Value
        !T = txtTeorico.Text
        !P = txtPractico.Text
        !F = txtPromedio.Text
        !modulo = cmbModulo.Text
        .Update
        .Close
        
    '''CONTROL DE EXAMENES:
    '''Verifica que no se vuelvan a ingresar las notas de un Modulo
        Else
            MsgBox "El alumno ya ha rendido este módulo", vbCritical, "Exámenes"
            txtPromedio.Text = ""
            txtTeorico.Text = ""
            txtPractico.Text = ""
            cmbModulo.SetFocus
            Exit Sub
        End If
        .Open "SELECT Fecha, Modulo, Teorico as [T], Practico as [P], Promedio as [F] FROM examenes WHERE codalumno=" & Int(txtCodigo.Text) & " ORDER BY fecha,id", Cn, adOpenDynamic, adLockPessimistic
        Set grilla.DataSource = rsExamenes
        formatoGrilla
        
    '''EGRESADOS
        '''Operador de PC (5 Examenes)
        If txtCurso.Text = "Operador de Pc" And rsExamenes.RecordCount = 5 Then
            MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            cmdAgregar.Enabled = False
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
        
        '''Diseño Grafico (4 Examenes)
        ElseIf txtCurso.Text = "Diseño Gráfico" And rsExamenes.RecordCount = 4 Then
            MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            cmdAgregar.Enabled = False
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
        
        '''Programacion (2 Examenes)
        ElseIf txtCurso.Text = "Programación" And rsExamenes.RecordCount = 2 Then
            MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            cmdAgregar.Enabled = False
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
        
        '''Programacion & Access (3 Examenes)
        ElseIf txtCurso.Text = "Programación + Access" And rsExamenes.RecordCount = 3 Then
            MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
            cmdAgregar.Enabled = False
        
        '''Telefonia Celular (1 Examen)
        ElseIf txtCurso.Text = "Telefonía Celular" And rsExamenes.RecordCount = 1 Then
            MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
            cmdAgregar.Enabled = False
            
        '''Reparacion de PC y Redes (7 Examenes)
        ElseIf txtCurso.Text = "Armado y Reparación de PC y Redes" And rsExamenes.RecordCount = 7 Then
            MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
            cmdAgregar.Enabled = False
            
        '''Armado y Reparacion de PC (4 Examenes)
        ElseIf txtCurso.Text = "Armado y Reparación de PC" And rsExamenes.RecordCount = 4 Then
            MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
            cmdAgregar.Enabled = False
            
        '''Redes (3 Examenes)
        ElseIf txtCurso.Text = "Redes" And rsExamenes.RecordCount = 3 Then
            MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
            cmdAgregar.Enabled = False
            
            '''Tecnico PC I (3 Examenes)
            ElseIf txtCurso.Text = "Técnico en Pc nivel I" And rsExamenes.RecordCount = 3 Then
            MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
            cmdAgregar.Enabled = False
        
        '''Tecnico PC II (5 Examenes)
        ElseIf txtCurso.Text = "Técnico en Pc nivel II" And rsExamenes.RecordCount = 5 Then
            MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
            cmdAgregar.Enabled = False
        
        '''Ingles (3 Examenes)
        ElseIf txtCurso.Text = "Inglés" And rsExamenes.RecordCount = 3 Then
            MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
            cmdAgregar.Enabled = False
            
        '''Diseño Web (4 Examenes)
        ElseIf txtCurso.Text = "Diseño Web" And rsExamenes.RecordCount = 4 Then
            MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
            cmdAgregar.Enabled = False
            
        '''Electronica (4 Examenes)
        ElseIf txtCurso.Text = "Electronica" And rsExamenes.RecordCount = 4 Then
            MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
            cmdAgregar.Enabled = False
            
        '''Telefonia Celular (1 Examen)
        ElseIf txtCurso.Text = "Telefonía Celular" And rsExamenes.RecordCount = 1 Then
            MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
            cmdAgregar.Enabled = False
            
        '''Paneles Solares (3 Examenes)
        ElseIf txtCurso.Text = "Paneles Solares" And rsExamenes.RecordCount = 3 Then
            MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
            cmdAgregar.Enabled = False
            
            ElseIf ((txtCurso.Text = "Asistente en Salud" Or txtCurso.Text = "Asistente Terapeutico") And rsExamenes.RecordCount = 3) Or (txtCurso.Text = "Auxiliar de Farmacia" And rsExamenes.RecordCount = 2) Then
                        MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
            cmdAgregar.Enabled = False
            
            ElseIf txtCurso.Text = "Técnico en aire acondicionado" And rsExamenes.RecordCount = 2 Or txtCurso.Text = "Electricidad domiciliaria" And rsExamenes.RecordCount = 2 Then
            MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
            cmdAgregar.Enabled = False
            
        ElseIf txtCurso.Text = "Extracc. Adm. Y Asist. Tec. Laborat." And rsExamenes.RecordCount = 3 Then
            MsgBox "El alumno " & txtAlumno.Text & " ha egresado del curso de " & txtCurso.Text, vbInformation, "Exámenes"
            With rsEgresados
                If .State = 1 Then .Close
                .Open "SELECT * FROM egresados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !CodAlumno = Int(txtCodigo.Text)
                !fecha = DTPFecha.Value
                .Update
            End With
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
            cmdAgregar.Enabled = False
            With rsVerificaciones
                If .State = 1 Then .Close
                .Open "SELECT * FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                !estado = "Egresado"
                .UpdateBatch
            End With
            cmdAgregar.Enabled = False
        End If
    End With
  
    
    '''limpiar todo
        txtPractico.Text = ""
        txtTeorico.Text = ""
        txtPromedio.Text = ""
   
End Sub

Private Sub Form_Load()
    Centrar Me
    DTPFecha.Value = Date
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lblOrigen.Caption = "Egresados" Then frmEgresados.Enabled = True
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtCodigo.Text = "" Then MsgBox "Ingrese el código del alumno", vbOKOnly, "GIA - Exámenes": txtCodigo.SetFocus: Exit Sub
      
        With rsVerificaciones
            If .State = 1 Then .Close
            .Open "SELECT nya,capac FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
                If .BOF = True And .EOF = True Then
                    MsgBox "No se encuentra el Código de Alumno" & vbNewLine & "Controle que el codigo ingresado sea correcto", vbExclamation, "Gestion Integral del Alumno - Gestion Integral del Alumno"
                ElseIf .BOF = False Or .EOF = False Then
                    txtAlumno.Text = !NyA
                    txtCurso.Text = !capac
                End If
        End With
    
        With rsExamenes
            If .State = 1 Then .Close
            .Open "SELECT Fecha, Modulo, Teorico as [T], Practico as [P], Promedio as [F] FROM examenes WHERE codalumno=" & Int(txtCodigo.Text) & " ORDER BY fecha,id", Cn, adOpenDynamic, adLockPessimistic
        End With
        
        If txtCurso.Text = "Operador de Pc" And rsExamenes.RecordCount = 5 Then
            cmdAgregar.Enabled = False
        ElseIf txtCurso.Text = "Diseño Gráfico" And rsExamenes.RecordCount = 4 Then
            cmdAgregar.Enabled = False
        ElseIf txtCurso.Text = "Diseño Web" And rsExamenes.RecordCount = 4 Then
            cmdAgregar.Enabled = False
        ElseIf txtCurso.Text = "Programación + Access" And rsExamenes.RecordCount = 3 Then
            cmdAgregar.Enabled = False
        ElseIf txtCurso.Text = "Programación" And rsExamenes.RecordCount = 2 Then
            cmdAgregar.Enabled = False
        ElseIf (txtCurso.Text = "Técnico en aire acondicionado" And rsExamenes.RecordCount = 2) Or (txtCurso.Text = "Electricidad domiciliaria" And rsExamenes.RecordCount = 2) Then
            cmdAgregar.Enabled = False
        ElseIf txtCurso.Text = "Telefonía Celular" And rsExamenes.RecordCount = 1 Then
            cmdAgregar.Enabled = False
        ElseIf txtCurso.Text = "Armado y Reparación de PC y Redes" And rsExamenes.RecordCount = 7 Then
            cmdAgregar.Enabled = False
        ElseIf txtCurso.Text = "Armado y Reparación de PC" And rsExamenes.RecordCount = 4 Then
            cmdAgregar.Enabled = False
        ElseIf txtCurso.Text = "Redes" And rsExamenes.RecordCount = 3 Then
            cmdAgregar.Enabled = False
        ElseIf txtCurso.Text = "Técnico en Pc nivel I" And rsExamenes.RecordCount = 3 Then
            cmbagregar.Enabled = False
        ElseIf txtCurso.Text = "Técnico en Pc nivel II" And rsExamenes.RecordCount = 5 Then
            cmbagregar.Enabled = False
        ElseIf (txtCurso.Text = "Inglés" And rsExamenes.RecordCount = 3) Or (txtCurso.Text = "Inglés II" And rsExamenes.RecordCount = 3) Then
            cmdAgregar.Enabled = False
        ElseIf (txtCurso.Text = "Extracc. Adm. Y Asist. Tec. Laborat." And rsExamenes.RecordCount = 3) Or (((txtCurso.Text = "Asistente en Salud" Or txtCurso.Text = "Asistente Terapeutico") Or txtCurso.Text = "Asistente Terapeutico") And rsExamenes.RecordCount = 3) Then
            cmdAgregar.Enabled = False
        ElseIf (txtCurso.Text = "Paneles Solares" And rsExamenes.RecordCount = 3) Or (txtCurso.Text = "Auxiliar de Farmacia" Or rsExamenes.RecordCount = 2) Then
            cmdAgregar.Enabled = False
        ElseIf txtCurso.Text = "" Then
        
        ElseIf txtCurso.Text = "" Then
        
        Else
            cmdAgregar.Enabled = True
        End If
        
        Set grilla.DataSource = rsExamenes
        formatoGrilla
    
    If txtCurso.Text = "Operador de Pc" Then
        With cmbModulo
            .Clear
            .AddItem ("Windows")
            .AddItem ("Word")
            .AddItem ("Excel")
            .AddItem ("Access")
            .AddItem ("Power Point")
        End With
    ElseIf txtCurso.Text = "Diseño Gráfico" Then
        With cmbModulo
            .Clear
            .AddItem ("Windows")
            .AddItem ("Corel Draw")
            .AddItem ("Photoshop")
            .AddItem ("Page Maker")
        End With
    ElseIf txtCurso.Text = "Programación" Then
        With cmbModulo
            .Clear
            .AddItem ("Módulo I")
            .AddItem ("Módulo II")
        End With
    ElseIf txtCurso.Text = "Programación + Access" Then
        With cmbModulo
            .Clear
            .AddItem ("Access")
            .AddItem ("Módulo I")
            .AddItem ("Módulo II")
        End With
    ElseIf txtCurso.Text = "Telefonía Celular" Then
        With cmbModulo
            .Clear
            .AddItem ("Módulo I")
        End With
    ElseIf txtCurso.Text = "Armado y Reparación de PC y Redes" Then
        With cmbModulo
            .Clear
            .AddItem ("Armado I")
            .AddItem ("Armado II")
            .AddItem ("Armado III")
            .AddItem ("Armado IV")
            .AddItem ("Redes I")
            .AddItem ("Redes II")
            .AddItem ("Redes III")
        End With
    ElseIf txtCurso.Text = "Armado y Reparación de PC" Then
        With cmbModulo
            .Clear
            .AddItem ("Armado I")
            .AddItem ("Armado II")
            .AddItem ("Armado III")
            .AddItem ("Armado IV")
        End With
    ElseIf txtCurso.Text = "Redes" Then
        With cmbModulo
            .Clear
            .AddItem ("Redes I")
            .AddItem ("Redes II")
            .AddItem ("Redes III")
        End With
    ElseIf txtCurso.Text = "Técnico en Pc nivel I" Then
        With cmbModulo
            .Clear
            .AddItem ("Modulo I")
            .AddItem ("Modulo II")
            .AddItem ("Examen Final")
        End With
    ElseIf txtCurso.Text = "Técnico en Pc nivel II" Then
        With cmbModulo
            .Clear
            .AddItem ("Modulo I")
            .AddItem ("Modulo II")
            .AddItem ("Modulo III")
            .AddItem ("Modulo IV")
            .AddItem ("Examen Final")
        End With
        
    
    ElseIf txtCurso.Text = "Técnico en aire acondicionado" Or txtCurso.Text = "Electricidad domiciliaria" Then
        With cmbModulo
            .Clear
            .AddItem ("Modulo I")
            .AddItem ("Modulo II")
        End With
        
    ElseIf txtCurso.Text = "Inglés" Or txtCurso.Text = "Inglés II" Then
        With cmbModulo
            .Clear
            .AddItem ("Inglés I")
            .AddItem ("Inglés II")
            .AddItem ("Inglés III")
        End With
        
    ElseIf txtCurso.Text = "Diseño Web" Then
        With cmbModulo
            .Clear
            .AddItem ("Front Page")
            .AddItem ("Fireworks")
            .AddItem ("Flash")
            .AddItem ("Dreamweaver")
        End With
        
    ElseIf txtCurso.Text = "Extracc. Adm. Y Asist. Tec. Laborat." Then
        With cmbModulo
            .Clear
            .AddItem ("Extraccionista I")
            .AddItem ("Extraccionista II")
            .AddItem ("Extraccionista III")
        End With
        
    ElseIf txtCurso.Text = "Auxiliar de Farmacia" Then
        With cmbModulo
            .Clear
            .AddItem ("Auxiliar I")
            .AddItem ("Auxiliar II")
        End With
    '''Paneles Solares
    ElseIf txtCurso.Text = "Paneles Solares" Then
        With cmbModulo
            .Clear
            .AddItem ("Paneles I")
            .AddItem ("Paneles II")
            .AddItem ("Paneles III")
        End With
    '''Asistente en Salud & Asistente Terapeutico
    ElseIf (txtCurso.Text = "Asistente en Salud" Or txtCurso.Text = "Asistente Terapeutico") Then
        With cmbModulo
            .Clear
            .AddItem ("Salud I")
            .AddItem ("Salud II")
            .AddItem ("Salud III")
        End With
    End If
    cmbModulo.SetFocus
    End If
    '''carga de modulos

    
End Sub

Private Sub txtPractico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtPromedio.Text = Int((txtTeorico.Text) + Int(txtPractico.Text)) / 2
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtPromedio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtTeorico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
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
