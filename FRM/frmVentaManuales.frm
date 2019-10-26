VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmVentaManuales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Venta de Manuales"
   ClientHeight    =   4380
   ClientLeft      =   5385
   ClientTop       =   450
   ClientWidth     =   5715
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVentaManuales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmVentaManuales.frx":324A
   ScaleHeight     =   4380
   ScaleWidth      =   5715
   Begin VB.Frame Frame1 
      BackColor       =   &H00662200&
      Caption         =   "Venta Manual"
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
         ItemData        =   "frmVentaManuales.frx":AC67
         Left            =   120
         List            =   "frmVentaManuales.frx":AC74
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
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtPrecio 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         Icon            =   "frmVentaManuales.frx":AC96
         Style           =   8
         Caption         =   "       Aceptar"
         IconSize        =   18
         IconAlign       =   1
         CaptionAlign    =   1
         iNonThemeStyle  =   0
         ShowFocus       =   -1  'True
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
         Value           =   -1  'True
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago"
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
         Size            =   8.25
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
         Size            =   8.25
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   3735
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   2655
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4683
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
      TabIndex        =   12
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
      TabIndex        =   11
      Top             =   120
      Width           =   975
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
      TabIndex        =   10
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmVentaManuales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbManual_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

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

Private Sub cmdVender_Click()
    If cmbManual.Text = "" Then MsgBox "Elija manual", vbCritical, "Venta de Manuales": cmbManual.SetFocus: Exit Sub
    If txtStock.Text = "" Then MsgBox "Controle el stock del manual", vbCritical, "Venta de Manuales": cmbManual.SetFocus: Exit Sub
    If txtRecibo.Text = "" Then MsgBox "Ingrese numero de recibo", vbCritical, "Venta de Manuales": txtRecibo.SetFocus: Exit Sub
    If cmbPago.Text = "" Then MsgBox "Elija forma de pago", vbCritical, "Venta de Manuales": cmbPago.SetFocus: Exit Sub
    If Int(txtStock.Text) < 1 Then MsgBox "No hay manuales disponibles para vender", vbCritical, "Venta de Manuales": cmbManual.SetFocus: Exit Sub
    
    '''descuenta el manual del stock
    With rsManuales
        If .State = 1 Then .Close
        .Open "SELECT * FROM manuales", Cn, adOpenDynamic, adLockPessimistic
        .Find "manual='" & cmbManual.Text & "'"
        !stock = !stock - 1
        .UpdateBatch
    End With
    
    'asigna el manual al alumno
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
    
    
    'hace o no asiento contable dependiento si pago total o no
    If MsgBox("¿Abona el total del precio?", vbYesNo + vbQuestion, "Venta de Manuales") = vbYes Then
        With rsContabilidad
            If .State = 1 Then .Close
            .Open "SELECT * FROM contabilidad", Cn, adOpenDynamic, adLockPessimistic
            .Requery
            .AddNew
            !fecha = Date
            !asiento = Null
            !NroCuota = Null
            !CodAlumno = Null
            !cuenta = "MANUALES"
            !Detalle = txtAlumno.Text & " - " & cmbManual.Text
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
            
            !Detalle = txtAlumno.Text & " - Manual de " & cmbManual.Text
            !nrofactura = txtRecibo.Text
            !Debe = CSng(txtPrecio.Text)
            !asiento = Null
            !NroCuota = Null
            !CodAlumno = Null
            !Haber = Null
            .Update
        End With
    Else
        MsgBox "Recuerde realizar el asiento contable correspondiente a esta operación", vbExclamation, "Venta de Manuales"
    End If
    
    'limpia cuadros
    txtStock.Text = ""
    txtPrecio.Text = ""
    txtRecibo.Text = ""
    txtCodigo.SetFocus
    
    formatoGrilla
End Sub

Private Sub Form_Load()
    Centrar Me
    formatoGrilla
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtCodigo.Text = "" Then MsgBox "Ingrese el código del alumno", vbOKOnly, "GIA - Exámenes": txtCodigo.SetFocus: Exit Sub
      
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
    
    '''MANUALES OPERADOR DE PC
    If txtCurso.Text = "Operador de Pc" Then
        With cmbManual
            .Clear
            .AddItem ("Windows")
            .AddItem ("Word")
            .AddItem ("Excel")
            .AddItem ("Access")
            .AddItem ("Power Point")
        End With
    '''MANUALES REDES SOCIALES
    ElseIf txtCurso.Text = "Redes Sociales" Then
        With cmbManual
            .Clear
            .AddItem ("Windows")
        End With
    '''MANUALES DISEÑO GRAFICO
    ElseIf txtCurso.Text = "Diseño Gráfico" Then
        With cmbManual
            .Clear
            .AddItem ("Windows")
            .AddItem ("Corel Draw")
            .AddItem ("Photoshop")
            .AddItem ("Page Maker")
        End With
    '''MANUALES DISEÑO WEB
    ElseIf txtCurso.Text = "Diseño Web" Then
        With cmbManual
            .Clear
            .AddItem ("FrontPage - Fireworks")
            .AddItem ("Flash")
            .AddItem ("Dreamweaver")
        End With
    '''MANUALES PROGRAMACION
    ElseIf txtCurso.Text = "Programación" Then
        With cmbManual
            .Clear
            .AddItem ("Programación")
        End With
    '''MANUALES PROGRAMACION + ACCESS
    ElseIf txtCurso.Text = "Programación + Access" Then
        With cmbManual
            .Clear
            .AddItem ("Access")
            .AddItem ("Programación")
        End With
    '''MANUALES TELEFONIA CELULAR
    ElseIf txtCurso.Text = "Telefonía Celular" Then
        With cmbManual
            .Clear
            .AddItem ("Telefonía Celular")
        End With
    '''MANUALES CURSO COMPLETO ARMADO
    ElseIf txtCurso.Text = "Armado y Reparación de PC y Redes" Then
        With cmbManual
            .Clear
            .AddItem ("Armado I")
            .AddItem ("Armado II")
            .AddItem ("Armado III")
            .AddItem ("Armado IV")
            .AddItem ("Redes I")
            .AddItem ("Redes II")
            .AddItem ("Redes III")
        End With
    '''MANUALES ARMADO Y REPARACION DE PC
    ElseIf txtCurso.Text = "Armado y Reparación de PC" Then
        With cmbManual
            .Clear
            .AddItem ("Armado I")
            .AddItem ("Armado II")
            .AddItem ("Armado III")
            .AddItem ("Armado IV")
        End With
    '''MANUALES REDES
    ElseIf txtCurso.Text = "Redes" Then
        With cmbManual
            .Clear
            .AddItem ("Redes I")
            .AddItem ("Redes II")
            .AddItem ("Redes III")
        End With
    '''MANUALES TECNICO PC I
    ElseIf txtCurso.Text = "Técnico en Pc nivel I" Then
        With cmbManual
            .Clear
            .AddItem ("Armado I")
            .AddItem ("Armado II")
        End With
    '''MANUALES TECNICO PC II
    ElseIf txtCurso.Text = "Técnico en Pc nivel II" Then
        With cmbManual
            .Clear
            .AddItem ("Armado III")
            .AddItem ("Armado IV")
            .AddItem ("Redes I")
            .AddItem ("Redes II")
            .AddItem ("Redes III")
        End With
    End If
    cmbManual.SetFocus
    End If
    formatoGrilla
End Sub

Sub formatoGrilla()
    For N = 0 To 1
        grilla.Columns(N).Width = 1150 + (N * 800)
    Next
End Sub
