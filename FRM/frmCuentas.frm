VERSION 5.00
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmCuentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmCuentas.frx":0000
   ScaleHeight     =   2670
   ScaleWidth      =   5205
   Begin VB.Frame Frame1 
      BackColor       =   &H00884400&
      Caption         =   "Buscar"
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
      Height          =   735
      Left            =   1080
      TabIndex        =   8
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton cmdBuscar1 
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscar2 
         Caption         =   "Codigo"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.ComboBox cmbTipoCta 
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
      ItemData        =   "frmCuentas.frx":7A1D
      Left            =   120
      List            =   "frmCuentas.frx":7A27
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2160
      Width           =   3495
   End
   Begin VB.TextBox txtDetalle 
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
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox txtNombreCuenta 
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
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   3495
   End
   Begin isButtonTest.isButton cmdGrabar 
      Height          =   420
      Left            =   3720
      TabIndex        =   11
      Top             =   1100
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCuentas.frx":7A38
      Style           =   8
      Caption         =   "       Guardar"
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
   Begin isButtonTest.isButton cmdCancelar 
      Height          =   420
      Left            =   3720
      TabIndex        =   12
      Top             =   1600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCuentas.frx":8312
      Style           =   8
      Caption         =   "       Cancelar"
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
   Begin isButtonTest.isButton cmdNuevo 
      Height          =   420
      Left            =   3720
      TabIndex        =   13
      Top             =   100
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCuentas.frx":8BEC
      Style           =   8
      Caption         =   "       Nuevo"
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
   Begin isButtonTest.isButton cmdModificar 
      Height          =   420
      Left            =   3720
      TabIndex        =   14
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCuentas.frx":94C6
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
   Begin isButtonTest.isButton cmdCerrar 
      Height          =   420
      Left            =   3720
      TabIndex        =   15
      Top             =   2100
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCuentas.frx":9DA0
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
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
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label lblCodCuenta 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle"
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
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
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
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Cta."
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
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbTipoCta_Change()
    cmdGrabar.SetFocus
End Sub

Private Sub cmdBuscar1_Click()
    X = InputBox("Ingrese la Cuenta a Buscar", "Buscar Cuenta")
    With rsCuentas
        If .BOF Or .EOF Then Exit Sub
        .Requery
        .MoveFirst
        .Find "Cuenta='" & X & "'"
        If .EOF Or .BOF Then MsgBox "La Cuenta no es válida", vbOKOnly + vbInformation, "Cuentas": .MoveFirst:  Exit Sub
        lblCodCuenta.Caption = !codcuenta
        txtNombreCuenta.Text = !cuenta
        txtDetalle.Text = !Detalle
    End With

End Sub

Private Sub cmdBuscar2_Click()
    X = InputBox("Ingrese el Código de Cuenta a Buscar", "Buscar Cuenta")
    With rsCuentas
        If .BOF Or .EOF Then Exit Sub
        .Requery
        .MoveFirst
        .Find "CodCuenta=" & X
        If .EOF Or .BOF Then MsgBox "El Código de Cuenta no es válido", vbOKOnly + vbInformation, "Cuentas": .MoveFirst: Exit Sub
        lblCodCuenta.Caption = !codcuenta
        txtNombreCuenta.Text = !cuenta
        txtDetalle.Text = !Detalle
    End With
    
End Sub

Private Sub cmdCancelar_Click()
    HabilitarBotones True, False
    HabilitarCuadros True
    Limpiar
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdGrabar_Click()
    If txtNombreCuenta.Text = "" Then MsgBox "Debe ingresar un nombre de cuenta", vbOKOnly + vbInformation, "Cuentas": txtNombreCuenta.SetFocus: Exit Sub
    If txtDetalle.Text = "" Then MsgBox "Debe ingresar un detalle de cuenta", vbOKOnly + vbInformation, "Cuentas": txtDetalle.SetFocus: Exit Sub
    If cmbTipoCta.Text = "" Then MsgBox "Debe elegir un tipo de cuenta", vbCritical + vbOKOnly, "Cuentas": cmbTipoCta.SetFocus: Exit Sub
    
    On Error GoTo LineaError
    
    If Modi = False Then
        Control
        rsControl.MoveFirst
        With rsCuentas
            .Requery
                .AddNew
                !codcuenta = rsControl!codcuenta
                !cuenta = txtNombreCuenta.Text
                !Detalle = txtDetalle.Text
                !tipo = cmbTipoCta.Text
                .Update
                lblCodCuenta.Caption = rsControl!codcuenta
                rsControl!codcuenta = rsControl!codcuenta + 1
                rsControl.UpdateBatch
        End With
    Else
        With rsCuentas
            .Requery
            .Find "codcuenta=" & Int(lblCodCuenta.Caption)
            !Detalle = txtDetalle.Text
            !cuenta = txtNombreCuenta.Text
            !tipo = cmbTipoCta.Text
            .UpdateBatch
        End With
    End If
    HabilitarBotones True, False
    HabilitarCuadros True

LineaError:
    Select Case Err.Number
        Case 3021
            Resume Next
        End Select
End Sub

Private Sub cmdModificar_Click()
    If lblCodCuenta.Caption = "" Then
        MsgBox "Primero debe Buscar una Cuenta", vbOKOnly + vbInformation, "Cuentas"
    Else
        HabilitarBotones False, True
        txtDetalle.Locked = False
        txtDetalle.SetFocus
        Modi = True
    End If
End Sub

Private Sub cmdNuevo_Click()
    HabilitarBotones False, True
    HabilitarCuadros False
    Limpiar
    txtNombreCuenta.SetFocus
    Modi = False
End Sub

Private Sub Form_Load()
    Centrar Me
    Cuentas
    HabilitarBotones True, False
    HabilitarCuadros True
    Limpiar
End Sub

Sub HabilitarBotones(estado1 As Boolean, estado2 As Boolean)
    cmdNuevo.Enabled = estado1
    cmdBuscar1.Enabled = estado1
    cmdBuscar2.Enabled = estado1
    cmdModificar.Enabled = estado1
    cmdCancelar.Enabled = estado2
    cmdGrabar.Enabled = estado2
    cmdCerrar.Enabled = estado1
End Sub

Sub HabilitarCuadros(estado1 As Boolean)
    txtNombreCuenta.Locked = estado1
    txtDetalle.Locked = estado1
    cmbTipoCta.Locked = estado1
End Sub

Sub Limpiar()
    lblCodCuenta.Caption = ""
    txtNombreCuenta.Text = ""
    txtDetalle.Text = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Private Sub txtDetalle_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNombreCuenta_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
