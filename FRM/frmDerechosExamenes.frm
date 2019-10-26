VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmDerechosExamenes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Derechos de Exámenes"
   ClientHeight    =   4725
   ClientLeft      =   7245
   ClientTop       =   2280
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmDerechosExamenes.frx":0000
   ScaleHeight     =   4725
   ScaleWidth      =   5730
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
      TabIndex        =   11
      Top             =   360
      Width           =   2775
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
      TabIndex        =   10
      Top             =   1080
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00662200&
      Caption         =   "Derecho Examen"
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
         Left            =   120
         TabIndex        =   4
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ComboBox cmbPago 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmDerechosExamenes.frx":7A1D
         Left            =   120
         List            =   "frmDerechosExamenes.frx":7A2A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin VB.TextBox txtRecibo 
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
         Top             =   2280
         Width           =   1335
      End
      Begin VB.ComboBox cmbModulo 
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
         Height          =   345
         ItemData        =   "frmDerechosExamenes.frx":7A4C
         Left            =   120
         List            =   "frmDerechosExamenes.frx":7A4E
         TabIndex        =   1
         Top             =   1080
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   89456641
         CurrentDate     =   41978
      End
      Begin isButtonTest.isButton cmdAgregar 
         Height          =   420
         Left            =   120
         TabIndex        =   18
         Top             =   3300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmDerechosExamenes.frx":7A50
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
      Begin isButtonTest.isButton cmdExamenes 
         Height          =   420
         Left            =   120
         TabIndex        =   19
         Top             =   3800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmDerechosExamenes.frx":832A
         Style           =   8
         Caption         =   "       Examenes"
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
         BackStyle       =   0  'Transparent
         Caption         =   "Recibo"
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
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio"
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
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago"
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
         TabIndex        =   8
         Top             =   840
         Width           =   1335
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
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   3015
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   20
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
         Size            =   8.25
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
   Begin VB.Label Label1 
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
      TabIndex        =   14
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
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
      TabIndex        =   13
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label3 
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
      TabIndex        =   12
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmDerechosExamenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAgregar_Click()
    If cmbModulo.Text = "" Then MsgBox "Elija el módulo", vbOKOnly + vbCritical, "GIA - Exámenes": cmbModulo.SetFocus: Exit Sub
    If cmbPago.Text = "" Then MsgBox "Elija el tipo de pago", vbOKOnly + vbCritical, "GIA - Exámenes": cmbPago.SetFocus: Exit Sub
    If txtRecibo.Text = "" Then MsgBox "Ingrese el número de recibo", vbOKOnly + vbCritical, "GIA - Exámenes": txtRecibo.SetFocus: Exit Sub
    
    
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
        .Open "SELECT Fecha, Modulo as Módulo FROM derechoexamen WHERE codalumno=" & Int(txtCodigo.Text) & " ORDER BY fecha", Cn, adOpenDynamic, adLockPessimistic
        Set grilla.DataSource = rsDerechosExamenes
        formatoGrilla
    End With
    
    If MsgBox("¿Abona el total del Derecho de Examen?", vbYesNo + vbQuestion, "Derechos de Exámenes") = vbYes Then
        With rsContabilidad
            If .State = 1 Then .Close
            .Open "SELECT * FROM contabilidad", Cn, adOpenDynamic, adLockPessimistic
            .Requery
            .AddNew
            !fecha = Date
            !asiento = Null
            !NroCuota = Null
            !CodAlumno = Null
            !cuenta = "DERECHO DE EXAMEN"
            !Detalle = txtAlumno.Text & " - " & cmbModulo.Text
            !nrofactura = txtRecibo.Text
            !Haber = CSng(txtPrecio.Text)
            !Debe = Null
            .Update
            .Requery
            .AddNew
            !fecha = Date
            
            If cmbPago.Text = "Efectivo" Then
                !cuenta = "CAJA ADMINISTRACION"
            ElseIf cmbPago.Text = "Descuento" Then
                !cuenta = "Descuento"
            Else
                !cuenta = "DEBITO TARJETA CREDITO"
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
    Else
        MsgBox "Recuerde realizar el asiento contable correspondiente a esta operación", vbExclamation, "Derechos de Exámenes"
    End If

End Sub

Private Sub cmdexamenes_Click()
    frmExamenes.Show
    frmExamenes.txtCodigo.Text = txtCodigo.Text
End Sub

Private Sub Form_Load()
    Centrar Me
    Control
    txtPrecio.Text = Format(rsControl!derechoExamen, "currency")
    DTPFecha.Value = Date
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
    
        With rsDerechosExamenes
            If .State = 1 Then .Close
            .Open "SELECT Fecha, Modulo as Módulo FROM derechoexamen WHERE codalumno=" & Int(txtCodigo.Text) & " ORDER BY fecha", Cn, adOpenDynamic, adLockPessimistic
        End With
        
        Set grilla.DataSource = rsDerechosExamenes
        formatoGrilla
   
'''EXAMENES Y MODULOS
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
        
    ElseIf txtCurso.Text = "Programación + Access" Then
        With cmbModulo
            .Clear
            .AddItem ("Access")
            .AddItem ("Módulo I")
            .AddItem ("Módulo II")
        End With
        
    ElseIf txtCurso.Text = "Programación" Or txtCurso.Text = "Telefonía Celular" Then
        With cmbModulo
            .Clear
            .AddItem ("Módulo I")
            .AddItem ("Módulo II")
        End With
        
    ElseIf txtCurso.Text = "Técnico en aire acondicionado" Or txtCurso.Text = "Electricidad domiciliaria" Then
        With cmbModulo
            .Clear
            .AddItem ("Módulo I")
            .AddItem ("Módulo II")
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
        
    ElseIf txtCurso.Text = "Paneles Solares" Then
        With cmbModulo
            .Clear
            .AddItem ("Paneles I")
            .AddItem ("Paneles II")
            .AddItem ("Paneles III")
        End With
        
    ElseIf txtCurso.Text = "Asistente en Salud" Or txtCurso.Text = "Asistente Terapeutico" Then
        With cmbModulo
            .Clear
            .AddItem ("Salud I")
            .AddItem ("Salud II")
            .AddItem ("Salud III")
        End With
        
    ElseIf txtCurso.Text = "Auxiliar de Farmacia" Then
        With cmbModulo
            .Clear
            .AddItem ("Auxiliar I")
            .AddItem ("Auxiliar II")
        End With
    End If
    
    txtPrecio.Text = Format(rsControl!derechoExamen, "currency")
    cmbModulo.Enabled = True
    DTPFecha.Enabled = True
    cmdAgregar.Enabled = True
    cmbModulo.SetFocus
    
    End If
    
    If KeyAscii = 13 Then
        If rsDerechosExamenes.RecordCount >= 1 Then
            cmdExamenes.Enabled = True
            Else: cmdExamenes.Enabled = False
        End If
    End If
End Sub

Private Sub txtRecibo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbModulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Sub formatoGrilla()
    For N = 0 To 1
        grilla.Columns(N).Width = 1150 + (N * 800)
    Next
End Sub
