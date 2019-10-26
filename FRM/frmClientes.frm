VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clientes"
   ClientHeight    =   5985
   ClientLeft      =   5310
   ClientTop       =   2040
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9105
   Begin VB.TextBox txtObservaciones 
      Height          =   1215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   29
      Top             =   4320
      Width           =   7215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Soporte"
      Height          =   2655
      Left            =   7440
      TabIndex        =   23
      Top             =   0
      Width           =   1575
      Begin VB.CommandButton cmdQuitarEquipo 
         Caption         =   "Quitar Equipo"
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
         Height          =   420
         Left            =   120
         TabIndex        =   31
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton cmdAgregarEquipo 
         Caption         =   "Agregar Equipo"
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
         Height          =   420
         Left            =   120
         TabIndex        =   30
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CheckBox chkSoporte 
         Caption         =   "Soporte"
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
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtCantidadEquipos 
         Alignment       =   2  'Center
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
         TabIndex        =   25
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtPrecio 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   24
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Cant. Equipos"
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
         TabIndex        =   28
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label8 
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
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   1815
      Left            =   120
      TabIndex        =   21
      Top             =   2400
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin VB.Frame Frame1 
      Caption         =   "Cliente"
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   7215
      Begin VB.TextBox txtCliente 
         Height          =   285
         Left            =   1920
         TabIndex        =   12
         Top             =   480
         Width           =   4935
      End
      Begin VB.TextBox txtDireccion 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox txtLocalidad 
         Height          =   285
         Left            =   4080
         TabIndex        =   10
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtTel1 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtTel2 
         Height          =   285
         Left            =   2520
         TabIndex        =   8
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtCuit 
         Height          =   285
         Left            =   4920
         TabIndex        =   7
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblCodCliente 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Localidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   16
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CUIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7560
      TabIndex        =   5
      Top             =   3720
      Width           =   1335
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
      Height          =   420
      Left            =   7560
      TabIndex        =   4
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      Height          =   420
      Left            =   7560
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
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
      Height          =   420
      Left            =   7560
      TabIndex        =   2
      Top             =   4200
      Width           =   1335
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7560
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7560
      TabIndex        =   0
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Equipos"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   2160
      Width           =   1575
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkSoporte_Click()
    If chkSoporte.Value = 1 Then
        txtPrecio.Enabled = True
        txtCantidadEquipos.Enabled = True
        cmdAgregarEquipo.Enabled = True
        cmdQuitarEquipo.Enabled = True
'        txtPrecio.SetFocus
    Else
        txtPrecio.Enabled = False
        txtCantidadEquipos.Enabled = False
        cmdAgregarEquipo.Enabled = False
        cmdQuitarEquipo.Enabled = False

    End If
End Sub

Private Sub cmdAgregarEquipo_Click()
    frmAgregarEquipo.Show
    frmAgregarEquipo.lblCliente.Caption = txtCliente.Text
    Me.Enabled = False
End Sub

Private Sub cmdBuscar_Click()
    frmBuscarCliente.Show
    Me.Enabled = False
End Sub

Private Sub cmdCancelar_Click()
    HabilitarBotones True, False
    HabilitarCuadros False
    Limpiar
    chkSoporte.Enabled = False
    cmdAgregarEquipo.Enabled = False
    cmdQuitarEquipo.Enabled = False
    txtPrecio.Enabled = False
    txtCantidadEquipos.Enabled = False
End Sub

Private Sub cmdEditar_Click()
    If lblCodCliente.Caption = "" Then
        MsgBox "Primero debe buscar al cliente", vbCritical, "Cóndor - Clientes"
    Else
        HabilitarBotones False, True
        HabilitarCuadros True
        txtCliente.SetFocus
        ModiClientes = True
        chkSoporte.Enabled = True
        If chkSoporte.Value = 1 Then
            txtPrecio.Enabled = True
            txtCantidadEquipos.Enabled = True
            cmdAgregarEquipo.Enabled = True
            cmdQuitarEquipo.Enabled = True
        End If
    End If
End Sub

Private Sub cmdGrabar_Click()
    
    If txtCliente.Text = "" Then MsgBox "Agregue el nombre del cliente", vbCritical, "Cóndor - Clientes": txtCliente.SetFocus: Exit Sub
    If txtDireccion.Text = "" Then MsgBox "Agregue la direccion del cliente", vbCritical, "Cóndor - Clientes": txtDireccion.SetFocus: Exit Sub
    If txtLocalidad.Text = "" Then MsgBox "Agregue la localidad del cliente", vbCritical, "Cóndor - Clientes": txtLocalidad.SetFocus: Exit Sub
    If txtTel1.Text = "" Then MsgBox "Agregue el teléfono del cliente", vbCritical, "Cóndor - Clientes": txtTel1.SetFocus: Exit Sub
    
    If chkSoporte.Value = 1 Then
        If txtPrecio.Text = "" Then MsgBox "Ingrese el precio del soporte", vbCritical, "Cóndor - Clientes": txtPrecio.SetFocus: Exit Sub
        If Not IsNumeric(txtPrecio.Text) Then MsgBox "Ingrese el precio del soporte", vbCritical, "Cóndor - Clientes": txtPrecio.SetFocus: Exit Sub
        If txtCantidadEquipos.Text = "" Then MsgBox "Ingrese cantidad de equipos", vbCritical, "Cóndor - Clientes": txtCantidadEquipos.SetFocus: Exit Sub
        If Not IsNumeric(txtCantidadEquipos.Text) Then MsgBox "Ingrese cantidad de equipos", vbCritical, "Cóndor - Clientes": txtCantidadEquipos.SetFocus: Exit Sub
    End If
    
    'Clientes
    Control
    
    'If ModiClientes = False Then
     '   With rsClientes
      '      .Requery
       '     .AddNew
        '    !cliente = txtCliente.Text
         '   !direccion = txtDireccion.Text
          '  !localidad = txtLocalidad.Text
            '!tel1 = txtTel1.Text
           ' !tel2 = txtTel2.Text
           ' !cuit = txtCuit.Text
            '!codcliente = rsControl!codcliente
            'lblCodCliente.Caption = rsControl!codcliente
            '.Update
            'rsControl!codcliente = rsControl!codcliente + 1
            'rsControl.UpdateBatch
        'End With
    'Else
     '   With rsClientes
      '      .Requery
       '     .Find "codcliente=" & Int(lblCodCliente.Caption)
        '    !cliente = txtCliente.Text
         '   !direccion = txtDireccion.Text
          '  !localidad = txtLocalidad.Text
           ' !tel1 = txtTel1.Text
           ' !tel2 = txtTel2.Text
           ' !cuit = txtCuit.Text
           ' .UpdateBatch
        'End With
   ' End If
    
'    If chkSoporte.Value = 1 Then
 '       With rsSoportes
  '          If .State = 1 Then .Close
   '         .Open "select * from soportes", cn, adOpenDynamic, adLockPessimistic
    '        .Find "codcliente=" & Int(lblCodCliente.Caption)
     '       If .BOF Or .EOF Then
      '          .Requery
       '         .AddNew
        '        !codcliente = Int(lblCodCliente.Caption)
         '       !precio = txtPrecio.Text
          '      !cantidadequipos = Int(txtCantidadEquipos.Text)
           '     .Update
           ' Else
            '    .MoveFirst
             '   !precio = txtPrecio.Text
              '  !cantidadequipos = Int(txtCantidadEquipos.Text)
               ' .UpdateBatch
           ' End If
        'End With
    'End If
    
    chkSoporte.Enabled = False
    txtPrecio.Enabled = False
    txtCantidadEquipos.Enabled = False
    cmdAgregarEquipo.Enabled = False
    cmdQuitarEquipo.Enabled = False
    
    HabilitarBotones True, False
    HabilitarCuadros False
End Sub

Private Sub cmdNuevo_Click()
    HabilitarBotones False, True
    HabilitarCuadros True
    Limpiar
    txtCliente.SetFocus
    ModiClientes = False
    txtPrecio.Text = ""
    txtCantidadEquipos.Text = ""
    chkSoporte.Enabled = True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    HabilitarBotones True, False
    HabilitarCuadros False
End Sub

Private Sub HabilitarBotones(a As Boolean, b As Boolean)
    cmdNuevo.Enabled = a
    cmdBuscar.Enabled = a
    cmdEditar.Enabled = a
    cmdGrabar.Enabled = b
    cmdCancelar.Enabled = b
    cmdSalir.Enabled = a
End Sub

Private Sub HabilitarCuadros(a As Boolean)
    txtCliente.Enabled = a
    txtDireccion.Enabled = a
    txtLocalidad.Enabled = a
    txtTel1.Enabled = a
    txtTel2.Enabled = a
    txtCuit.Enabled = a
End Sub

Private Sub Limpiar()
    
    lblCodCliente.Caption = ""
    txtCliente.Text = ""
    txtDireccion.Text = ""
    txtLocalidad.Text = ""
    txtTel1.Text = ""
    txtTel2.Text = ""
    txtCuit.Text = ""
    chkSoporte.Value = 0
    txtPrecio.Text = ""
    txtCantidadEquipos.Text = ""
End Sub


Private Sub grilla_Click()
    On Error GoTo error
    txtObservaciones.Text = grilla.Columns(2).Text
    
error:
Exit Sub
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDireccion.SetFocus
End Sub

Private Sub txtCuit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdGrabar.SetFocus
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtLocalidad.SetFocus
End Sub

Private Sub txtLocalidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTel1.SetFocus
End Sub

Private Sub txtTel1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTel2.SetFocus
End Sub

Private Sub txtTel2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtCuit.SetFocus
End Sub
