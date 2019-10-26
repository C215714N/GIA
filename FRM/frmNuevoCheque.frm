VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmNuevoCheque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo Cheque"
   ClientHeight    =   1485
   ClientLeft      =   5085
   ClientTop       =   2640
   ClientWidth     =   5910
   Icon            =   "frmNuevoCheque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmNuevoCheque.frx":324A
   ScaleHeight     =   1485
   ScaleWidth      =   5910
   Begin VB.ComboBox cmbFirma 
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
      ItemData        =   "frmNuevoCheque.frx":AC67
      Left            =   3000
      List            =   "frmNuevoCheque.frx":AC6E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtMonto 
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
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtNroCheque 
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
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtDestinatario 
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
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   360
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
      Format          =   89260033
      CurrentDate     =   41782
   End
   Begin isButtonTest.isButton cmdAgregar 
      Height          =   420
      Left            =   4440
      TabIndex        =   10
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmNuevoCheque.frx":AC7A
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
   Begin isButtonTest.isButton cmdCancelar 
      Height          =   420
      Left            =   4440
      TabIndex        =   11
      Top             =   900
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmNuevoCheque.frx":B554
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Firma"
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
      Left            =   3000
      TabIndex        =   9
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Monto"
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
      Left            =   1560
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Cheque"
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
      TabIndex        =   7
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Destinatario"
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
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Pago"
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
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmNuevoCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    If txtDestinatario.Text = "" Then MsgBox "Debe agregar el destinatario del cheque", vbCritical, "Cheques": txtDestinatario.SetFocus: Exit Sub
    If txtNroCheque.Text = "" Then MsgBox "Debe agregar el número del cheque", vbCritical, "Cheques": txtNroCheque.SetFocus: Exit Sub
    If txtMonto.Text = "" Then MsgBox "Debe agregar el monto", vbCritical, "Cheques": txtMonto.SetFocus: Exit Sub
    If cmbFirma.Text = "" Then MsgBox "Debe agregar el firmante", vbCritical, "Cheques": cmbFirma.SetFocus: Exit Sub

    With rsCheques
        If .State = 1 Then .Close
        .Open "SELECT * FROM cheques", Cn, adOpenDynamic, adLockPessimistic
        .Requery
        .AddNew
        !fecha = DTPFecha.Value
        !destinatario = txtDestinatario.Text
        !numerocheque = txtNroCheque.Text
        !monto = txtMonto.Text
        !firma = cmbFirma.Text
        !estado = "SIN DEPOSITAR"
        .Update
    End With
    
    If MsgBox("¿Desea ingresar otro cheque?", vbQuestion + vbYesNo, "Cheques") = vbYes Then
        txtDestinatario.Text = ""
        txtMonto.Text = ""
        txtNroCheque.Text = ""
    Else
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub DTPFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Centrar Me
    DTPFecha.Value = Date
End Sub

Private Sub txtDestinatario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtFirma_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNroCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
