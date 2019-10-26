VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmPP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preparar Presupuesto"
   ClientHeight    =   4080
   ClientLeft      =   5445
   ClientTop       =   2325
   ClientWidth     =   9975
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmPP.frx":324A
   ScaleHeight     =   3697.885
   ScaleMode       =   0  'User
   ScaleWidth      =   9975
   Begin VB.Frame Frame2 
      BackColor       =   &H00884400&
      Caption         =   "Cuentas Actuales"
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
      Height          =   2895
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   4095
      Begin MSDataListLib.DataList dtlCuentas 
         Height          =   2460
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   4339
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
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00662200&
      Caption         =   "Presupuesto Nuevo"
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
      Height          =   2895
      Left            =   4320
      TabIndex        =   10
      Top             =   960
      Width           =   4095
      Begin MSDataListLib.DataList dtlPresupuesto 
         Height          =   2460
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   4339
         _Version        =   393216
         ListField       =   ""
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
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2760
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00884400&
      Caption         =   "Período"
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.ComboBox cmbMes 
         Height          =   360
         ItemData        =   "frmPP.frx":AC67
         Left            =   120
         List            =   "frmPP.frx":AC8F
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpAño 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
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
         CustomFormat    =   "yyyy"
         Format          =   89456643
         CurrentDate     =   43573
      End
      Begin VB.Label Label2 
         BackColor       =   &H00662200&
         BackStyle       =   0  'Transparent
         Caption         =   "AÑO"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00662200&
         BackStyle       =   0  'Transparent
         Caption         =   "MES"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
   End
   Begin isButtonTest.isButton cmdQuitar 
      Height          =   420
      Left            =   8520
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmPP.frx":ACF8
      Style           =   8
      Caption         =   "       Eliminar"
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
   Begin isButtonTest.isButton cmdAgregar 
      Height          =   420
      Left            =   8520
      TabIndex        =   12
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmPP.frx":B5D2
      Style           =   8
      Caption         =   "       Agregar"
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
   Begin isButtonTest.isButton cmdCrear 
      Height          =   420
      Left            =   8520
      TabIndex        =   13
      Top             =   1800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmPP.frx":BEAC
      Style           =   8
      Caption         =   "       Nueva"
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
      BackStyle       =   0  'Transparent
      Caption         =   "Monto"
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbMes_Click()
    With rsPresupuesto
        If .State = 1 Then .Close
        .Open "SELECT cuenta,deuda,saldo,mes,año FROM presupuesto WHERE mes='" & cmbMes.Text & "' and año=" & Year(dtpAño.Value) & " ORDER BY cuenta", Cn, adOpenDynamic, adLockPessimistic
        dtlPresupuesto.ListField = "Cuenta"
        Set dtlPresupuesto.RowSource = rsPresupuesto
    End With

End Sub

Private Sub cmdAgregar_Click()
    If cmbMes.Text = "" Then MsgBox "Elija el mes del presupuesto", vbCritical, "Preparar Presupuesto": cmbMes.SetFocus: Exit Sub
    If txtMonto.Text = "" Then MsgBox "Ingrese el monto", vbCritical, "Preparar Presupuesto": txtMonto.SetFocus: Exit Sub
    
    With rsPresupuesto
        .Requery
        .AddNew
        !cuenta = dtlCuentas.Text
        !deuda = CSng(txtMonto.Text)
        !mes = cmbMes.Text
        !saldo = CSng(txtMonto.Text)
        !año = Year(dtpAño.Value)
        .Update
        .Close
        .Open "SELECT cuenta,deuda,saldo,mes,año FROM presupuesto WHERE mes='" & cmbMes.Text & "' and año=" & Year(dtpAño.Value) & " ORDER BY cuenta", Cn, adOpenDynamic, adLockPessimistic
        dtlPresupuesto.ListField = "Cuenta"
        Set dtlPresupuesto.RowSource = rsPresupuesto
    End With
    
    txtMonto.Text = ""
    
End Sub

Private Sub cmdCrear_Click()
    frmCrearCuenta.Show
    Me.Enabled = False
End Sub

Private Sub cmdQuitar_Click()
    With rsPresupuesto
        .Find "cuenta='" & dtlPresupuesto.Text & "'"
        .Delete
        .Update
        .Close
        .Open "SELECT cuenta,deuda,saldo,mes,año FROM presupuesto WHERE mes='" & cmbMes.Text & "' and año=" & Year(dtpAño.Value) & " ORDER BY cuenta", Cn, adOpenDynamic, adLockPessimistic
        dtlPresupuesto.ListField = "Cuenta"
        Set dtlPresupuesto.RowSource = rsPresupuesto

    End With
End Sub

Private Sub dtlCuentas_Click()
    cmdAgregar.Enabled = True
End Sub

Private Sub dtlPresupuesto_Click()
    cmdQuitar.Enabled = True
End Sub

Private Sub dtpAño_Change()
    With rsPresupuesto
        If .State = 1 Then .Close
        .Open "SELECT cuenta,deuda,saldo,mes,año FROM presupuesto WHERE mes='" & cmbMes.Text & "' and año=" & Year(dtpAño.Value) & " ORDER BY cuenta", Cn, adOpenDynamic, adLockPessimistic
        dtlPresupuesto.ListField = "Cuenta"
        Set dtlPresupuesto.RowSource = rsPresupuesto
    End With

End Sub

Private Sub Form_Load()
    Centrar Me
    With rsCuentasPresupuesto
        If .State = 1 Then .Close
        .Open "SELECT Cuenta FROM CuentasPresupuesto", Cn, adOpenDynamic, adLockPessimistic
        dtlCuentas.ListField = "Cuenta"
        Set dtlCuentas.RowSource = rsCuentasPresupuesto
    End With


End Sub

