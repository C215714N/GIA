VERSION 5.00
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmCrearCuenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Cuenta de Presupuesto"
   ClientHeight    =   885
   ClientLeft      =   8835
   ClientTop       =   2670
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmCrearCuenta.frx":0000
   ScaleHeight     =   885
   ScaleWidth      =   4305
   Begin VB.TextBox txtCuenta 
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
      TabIndex        =   0
      Top             =   360
      Width           =   2715
   End
   Begin isButtonTest.isButton cmdAgregar 
      Height          =   420
      Left            =   2880
      TabIndex        =   2
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCrearCuenta.frx":7A1D
      Style           =   8
      Caption         =   "       Aceptar"
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nueva Cuenta"
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
      TabIndex        =   1
      Top             =   105
      Width           =   1455
   End
End
Attribute VB_Name = "frmCrearCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    If txtCuenta.Text = "" Then MsgBox "Ingrese el nombre de la nueva cuenta", vbCritical, "Agregar Cuenta de Presupuesto": txtCuenta.SetFocus: Exit Sub
    
    With rsCuentasPresupuesto
        If .State = 1 Then .Close
        .Open "SELECT cuenta FROM cuentaspresupuesto WHERE cuenta='" & txtCuenta.Text & "'", Cn, adOpenDynamic, adLockPessimistic
        If .BOF Or .EOF Then
            .AddNew
            !cuenta = txtCuenta.Text
            .Update
            txtCuenta.Text = ""
            txtCuenta.SetFocus
        Else
            MsgBox "La cuenta ya existe", vbCritical, "Cuentas de Presupuesto": txtCuenta.SetFocus
        End If
    End With
End Sub

Private Sub Form_Load()
    Centrar Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmPP.Enabled = True
    With rsCuentasPresupuesto
        .Close
        .Open "SELECT cuenta FROM cuentaspresupuesto", Cn, adOpenDynamic, adLockPessimistic
        frmPP.dtlCuentas.ListField = "Cuenta"
        Set frmPP.dtlCuentas.RowSource = rsCuentasPresupuesto

    End With
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
