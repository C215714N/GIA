VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmVentaManuales 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Venta de Manuales"
   ClientHeight    =   4380
   ClientLeft      =   5385
   ClientTop       =   450
   ClientWidth     =   5715
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
   Icon            =   "frmVentaManuales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5715
   Begin VB.Frame Frame1 
      BackColor       =   &H00662200&
      Caption         =   "Venta Manual"
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
      TabIndex        =   13
      Top             =   240
      Width           =   1600
      Begin VB.TextBox txtRecibo 
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2880
         Width           =   1335
      End
      Begin VB.ComboBox cmbPago 
         Height          =   360
         ItemData        =   "frmVentaManuales.frx":10CA
         Left            =   120
         List            =   "frmVentaManuales.frx":10D7
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2280
         Width           =   1335
      End
      Begin VB.ComboBox cmbManual 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox txtStock 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtPrecio 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1335
      End
      Begin isButtonTest.isButton cmdVender 
         Height          =   420
         Left            =   120
         TabIndex        =   6
         Top             =   3360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmVentaManuales.frx":10F9
         Style           =   8
         Caption         =   "     Aceptar"
         IconSize        =   18
         IconAlign       =   1
         CaptionAlign    =   1
         iNonThemeStyle  =   0
         ShowFocus       =   -1  'True
         BackColor       =   6693376
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
         Value           =   -1  'True
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Pago"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Manual"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Recibo"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2640
         Width           =   1215
      End
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
      Height          =   375
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
      Height          =   375
      Left            =   1080
      TabIndex        =   9
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
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   3735
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   2775
      Left            =   120
      TabIndex        =   7
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
      TabIndex        =   12
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
      TabIndex        =   11
      Top             =   120
      Width           =   975
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
      TabIndex        =   10
      Top             =   720
      Width           =   1335
   End
End
Attribute VB_Name = "frmVentaManuales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbManual_KeyPress(keyAscii As Integer)
If keyAscii = 13 Then
    With rsManuales
        If .State = 1 Then .Close
        .Open "SELECT stock,precio FROM manuales WHERE manual='" & cmbManual.Text & "'", Cn, adOpenDynamic, adLockPessimistic
        If .BOF Or .EOF Then MsgBox "El manual no se ha registrado", vbCritical, "Venta de Manuales": cmbManual.SetFocus: Exit Sub
        txtStock.Text = !stock
        txtPrecio.Text = Format(!precio, "currency")
        cmbPago.SetFocus
    End With
End If
End Sub

Private Sub cmbPago_KeyPress(keyAscii As Integer)
    continue keyAscii
End Sub

Private Sub cmdVender_Click()
    On Error GoTo LineaError
    
    If cmbManual.Text = "" Then MsgBox "Elija manual", vbCritical, "Venta de Manuales": cmbManual.SetFocus: Exit Sub
    If txtStock.Text = "" Then MsgBox "Controle el stock del manual", vbCritical, "Venta de Manuales": cmbManual.SetFocus: Exit Sub
    If txtRecibo.Text = "" Then MsgBox "Ingrese numero de recibo", vbCritical, "Venta de Manuales": txtRecibo.SetFocus: Exit Sub
    If cmbPago.Text = "" Then MsgBox "Elija forma de pago", vbCritical, "Venta de Manuales": cmbPago.SetFocus: Exit Sub
    If Int(txtStock.Text) < 1 Then MsgBox "No hay manuales disponibles para vender", vbCritical, "Venta de Manuales": cmbManual.SetFocus: Exit Sub
    
    Dim Cuenta
'''CONTROL DE STOCK
    With rsManuales
        If .State = 1 Then .Close
        .Open "SELECT * FROM manuales", Cn, adOpenDynamic, adLockPessimistic
        .Find "manual='" & cmbManual.Text & "'"
        !stock = !stock - 1
        .UpdateBatch
    End With
    
''' CARGA MANUAL - TABLA VENTA DE MANUALES
    With rsVentaManuales
        .Close
        .Open "SELECT * FROM ventamanuales", Cn, adOpenDynamic, adLockPessimistic
        .Requery
        .AddNew
        !fecha = Date
        !CodAlumno = Int(txtCodigo.Text)
        !manual = cmbManual.Text
        .Update
        .Close
        .Open "SELECT Fecha, Manual FROM ventamanuales WHERE codalumno=" & Int(txtCodigo.Text) & " ORDER BY fecha", Cn, adOpenDynamic, adLockPessimistic
        Set grilla.DataSource = rsVentaManuales
    End With
    
''' GESTION CONTABLE - ASIENTO
    With cmbManual
        If .Text = "Materiales 01" Or .Text = "Materiales 02" Or .Text = "Materiales 03" Or .Text = "Lazo" Then
            Cuenta = "INSUMOS CURSOS"
        Else
            Cuenta = "MANUALES"
        End If
    End With
    
    With rsContabilidad
        If .State = 1 Then .Close
        .Open "SELECT * FROM contabilidad", Cn, adOpenDynamic, adLockPessimistic
        .Requery
        .AddNew
        !fecha = Date
        !asiento = Null
        !NroCuota = Null
        !CodAlumno = Null
        !Detalle = txtAlumno.Text & " - " & cmbManual.Text
        !nrofactura = txtRecibo.Text
        !Haber = CSng(txtPrecio.Text)
        !Debe = Null
        !Cuenta = Cuenta
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
        
        !Detalle = txtAlumno.Text & " - Manual de " & cmbManual.Text
        !nrofactura = txtRecibo.Text
        !Debe = CSng(txtPrecio.Text)
        !asiento = Null
        !NroCuota = Null
        !CodAlumno = Null
        !Haber = Null
        .Update
    End With
    
'''REESTABLECE LOS VALORES
    txtStock.Text = ""
    txtPrecio.Text = ""
    txtRecibo.Text = ""
    txtCodigo.SetFocus
    formatoGrilla
LineaError: ErrCode Err
End Sub

Private Sub Form_Load()
    Centrar Me
    formatoGrilla
End Sub


Private Sub txtCodigo_KeyPress(keyAscii As Integer)
    On Error GoTo LineaError
    If keyAscii = 13 Then
        If txtCodigo.Text = "" Then MsgBox "Ingrese el codigo del alumno", vbOKOnly, "GIA - Examenes": txtCodigo.SetFocus: Exit Sub
      
        With rsVerificaciones
            If .State = 1 Then .Close
            .Open "SELECT nya,capac FROM verificaciones WHERE codalumno=" & Int(txtCodigo.Text), Cn, adOpenDynamic, adLockPessimistic
            txtAlumno.Text = !NyA
            txtCurso.Text = !capac
        End With
        With rsVentaManuales
            If .State = 1 Then .Close
            .Open "SELECT Fecha, Manual FROM ventamanuales WHERE codalumno=" & Int(txtCodigo.Text) & " ORDER BY fecha", Cn, adOpenDynamic, adLockPessimistic
        End With
        Set grilla.DataSource = rsVentaManuales
        cmbManual.Clear
        cargarManuales
        cmbManual.SetFocus
    End If
    formatoGrilla
LineaError: ErrCode Err
End Sub

Sub cargarManuales()

    If txtCurso.Text = "Operador de Pc" Then
        With cmbManual
            .AddItem ("Windows")
            .AddItem ("Word")
            .AddItem ("Excel")
            .AddItem ("Access")
            .AddItem ("Power Point")
        End With
    
    ElseIf txtCurso.Text = "Redes Sociales" Then
        With cmbManual
            .AddItem ("Windows")
        End With
    
    ElseIf txtCurso.Text = "Diseño Gráfico" Then
        With cmbManual
            .AddItem ("Windows")
            .AddItem ("Corel Draw")
            .AddItem ("Photoshop")
            .AddItem ("Page Maker")
        End With
    
    ElseIf txtCurso.Text = "Diseño Web" Then
        With cmbManual
            .AddItem ("FrontPage - Fireworks")
            .AddItem ("Flash")
            .AddItem ("Dreamweaver")
        End With
    
    ElseIf txtCurso.Text = "Programación" Then
        With cmbManual
            .AddItem ("Programación")
        End With
    
    ElseIf txtCurso.Text = "Programación + Access" Then
        With cmbManual
            .AddItem ("Access")
            .AddItem ("Programación")
        End With
    
    ElseIf txtCurso.Text = "Técnico en aire acondicionado" Or txtCurso.Text = "Refrigeración" Then
        With cmbManual
            .AddItem ("Refrigeracion I")
            .AddItem ("Refrigeracion II")
            .AddItem ("Refrigeracion III")
        End With
    
    ElseIf txtCurso.Text = "Electricidad domiciliaria" Then
        With cmbManual
            .AddItem ("Electricidad")
        End With
    
    ElseIf txtCurso.Text = "Auxiliar de Farmacia" Then
        With cmbManual
            .AddItem ("Farmacia")
            .AddItem ("Primeros Auxilios")
            .AddItem ("Covid")
            .AddItem ("Lazo")
        End With
        
    ElseIf txtCurso.Text = "Extracc. Adm. Y Asist. Tec. Laborat." Then
        With cmbManual
            .AddItem ("Extraccionista")
            .AddItem ("Guia Practica")
            .AddItem ("Primeros Auxilios")
            .AddItem ("Covid")
            .AddItem ("Materiales 01")
            .AddItem ("Materiales 02")
            .AddItem ("Materiales 03")
            .AddItem ("Lazo")
        End With
        
    ElseIf txtCurso.Text = "Asistente Terapeutico" Or txtCurso.Text = "Cuidador Domiciliario" Then
        With cmbManual
            .AddItem ("Cuidador Dom. I")
            .AddItem ("Cuidador Dom. II")
            .AddItem ("Covid")
            .AddItem ("Lazo")
        End With
    
    ElseIf txtCurso.Text = "Emergencias Médicas" Then
        With cmbManual
            .AddItem ("Primeros Auxilios")
            .AddItem ("Covid")
        End With
    ElseIf txtCurso.Text = "Mandatario Automotor" Then
        With cmbManual
            .AddItem ("Mandatario")
        End With
    ElseIf txtCurso.Text = "Tecnico en Reparacion de Lavarropas y Secarropas" Then
        With cmbManual
            .AddItem ("Lavarropas y Secarropas")
        End With
    End If
End Sub

Sub formatoGrilla()
    For N = 0 To 1
        grilla.Columns(N).Width = 1150 + (N * 800)
    Next
End Sub
