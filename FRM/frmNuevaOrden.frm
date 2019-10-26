VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmNuevaOrden 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nueva Orden de Trabajo"
   ClientHeight    =   2700
   ClientLeft      =   3960
   ClientTop       =   2340
   ClientWidth     =   9000
   Icon            =   "frmNuevaOrden.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   9000
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
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
      Height          =   495
      Left            =   7800
      TabIndex        =   13
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdNueva 
      Caption         =   "Nueva Orden"
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
      Height          =   495
      Left            =   7800
      TabIndex        =   12
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   11
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   10
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtFalla 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3960
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox txtEquipo 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   3735
   End
   Begin MSDataListLib.DataCombo dtcCliente 
      Height          =   360
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   "DataCombo1"
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
      Caption         =   "Falla presentada"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblNroOrden 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Equipo a revisar"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1305
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   6360
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   6360
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nº de Orden"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmNuevaOrden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdGrabar_Click()
    If dtcCliente.Text = "" Then MsgBox "Elija un cliente", vbCritical, "Nueva Orden de Trabajo": dtcCliente.SetFocus: Exit Sub
    If txtEquipo.Text = "" Then MsgBox "Identifique el equipo a revisar", vbCritical, "Nueva Orden de Trabajo": txtEquipo.SetFocus: Exit Sub
    If txtFalla.Text = "" Then MsgBox "Identifique la falla presentada por el equipo", vbCritical, "Nueva Orden de Trabajo": txtFalla.SetFocus: Exit Sub
    
    With rsBuscarClientes
        If .State = 1 Then .Close
        .Open "select * from clientes where cliente='" & dtcCliente.Text & "'", Cn, adOpenDynamic, adLockPessimistic
        .MoveFirst
    End With
    
    
    With rsOrdenesDeTrabajo
        If .State = 1 Then .Close
        .Open "select * from ordenesdetrabajo", Cn, adOpenDynamic, adLockPessimistic
        .AddNew
        !nroorden = Int(lblNroOrden.Caption)
        !codcliente = rsBuscarClientes!codcliente
        !fecharecibido = Date
        !equipo = Trim(txtEquipo.Text)
        !problema = Trim(txtFalla.Text)
        !estado = "A REVISAR"
        .Update
    End With
    
    With rsGestionDeOrdenes
        If .State = 1 Then .Close
        .Open "select * from gestiondeordenes", Cn, adOpenDynamic, adLockPessimistic
        .Requery
        .AddNew
        !nroordentrabajo = Int(lblNroOrden.Caption)
        !fecha = Date
        !gestion = "Ingreso"
        !Personal = Usuario
        .Update
    End With
    
    rsControl!nroorden = rsControl!nroorden + 1
    rsControl.UpdateBatch
    
    HabilitarCuadros False
    HabilitarBotones True, False
End Sub

Private Sub cmdImprimir_Click()
    With rsBuscarClientes
        .Close
        .Open "select * from clientes where cliente='" & dtcCliente.Text & "'", Cn, adOpenDynamic, adLockPessimistic
        .MoveFirst
    End With

    Set dtrOrdenDeTrabajo.DataSource = rsBuscarClientes
    dtrOrdenDeTrabajo.Sections("Sección2").Controls("lblcliente").Caption = rsBuscarClientes!cliente
    dtrOrdenDeTrabajo.Sections("Sección2").Controls("lbldireccion").Caption = rsBuscarClientes!direccion
    dtrOrdenDeTrabajo.Sections("Sección2").Controls("lbllocalidad").Caption = rsBuscarClientes!localidad
    dtrOrdenDeTrabajo.Sections("Sección2").Controls("lbltelefono").Caption = rsBuscarClientes!tel1
    dtrOrdenDeTrabajo.Sections("Sección2").Controls("lblcelular").Caption = rsBuscarClientes!tel2
    dtrOrdenDeTrabajo.Sections("Sección2").Controls("lblnroorden").Caption = lblNroOrden.Caption
    dtrOrdenDeTrabajo.Sections("Sección2").Controls("lblfecha").Caption = lblFecha.Caption
    dtrOrdenDeTrabajo.Sections("Sección1").Controls("lblEquipo").Caption = txtEquipo.Text
    dtrOrdenDeTrabajo.Sections("Sección1").Controls("lblFalla").Caption = txtFalla.Text

                        
    dtrOrdenDeTrabajo.Show
End Sub

Private Sub cmdNueva_Click()
    HabilitarCuadros True
    HabilitarBotones False, True
    Limpiar
    lblNroOrden.Caption = rsControl!nroorden
End Sub

Private Sub cmdSalir_click()
    Unload Me
End Sub

Private Sub dtcCliente_Change()
    txtEquipo.SetFocus
End Sub

Private Sub Form_Load()
    Centrar Me
    lblFecha.Caption = Date
    Control
    ''lblNroOrden.Caption = rsControl!nroorden
    Limpiar
End Sub

Sub Limpiar()
txtEquipo.Text = ""
txtFalla.Text = ""
End Sub

Private Sub txtEquipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtFalla.SetFocus
End Sub

Private Sub txtFalla_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdGrabar.SetFocus
End Sub

Sub HabilitarCuadros(a As Boolean)
    dtcCliente.Enabled = a
    txtFalla.Enabled = a
    txtEquipo.Enabled = a
End Sub

Sub HabilitarBotones(a As Boolean, b As Boolean)
    cmdNueva.Enabled = a
    cmdImprimir.Enabled = a
    cmdGrabar.Enabled = b
    
End Sub
