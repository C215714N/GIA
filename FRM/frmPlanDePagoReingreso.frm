VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmPlanDePagoReingreso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Plan de Pagos de Reingreso de Alumno"
   ClientHeight    =   1485
   ClientLeft      =   3075
   ClientTop       =   2325
   ClientWidth     =   4515
   Icon            =   "frmPlanDePagoReingreso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmPlanDePagoReingreso.frx":324A
   ScaleHeight     =   1485
   ScaleWidth      =   4515
   Begin VB.TextBox txtCantidadCuotas 
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
      Left            =   1560
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtNroCuota 
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
      Width           =   1335
   End
   Begin VB.TextBox txtMonto 
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
      Top             =   960
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   360
      Left            =   1560
      TabIndex        =   3
      Top             =   960
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
      CurrentDate     =   41353
   End
   Begin isButtonTest.isButton cmdAplicar 
      Height          =   420
      Left            =   3000
      TabIndex        =   8
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmPlanDePagoReingreso.frx":AC67
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
      Left            =   3000
      TabIndex        =   9
      Top             =   900
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmPlanDePagoReingreso.frx":B541
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Cant. Cuotas"
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
      Left            =   1560
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Desde Cuota"
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
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimiento"
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
      Height          =   240
      Left            =   1560
      TabIndex        =   5
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monto $"
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
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   675
   End
End
Attribute VB_Name = "frmPlanDePagoReingreso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAplicar_Click()
    Dim CuotaMax As Integer
    CuotaMax = Int(txtCantidadCuotas.Text) + Int(txtNroCuota.Text) - 1
    ''genera las nuevas cuotas
    With rsPlanDePago
        If .State = 1 Then .Close
        .Open "SELECT * FROM plandepago", Cn, adOpenDynamic, adLockPessimistic
        .Requery
        Do Until Int(txtNroCuota.Text) > CuotaMax
            .AddNew
            !CodAlumno = frmAnalisisDeCuotas.lblCodAlumno.Caption
            !NyA = frmAnalisisDeCuotas.lblNyA.Caption
            !NroCuota = Int(txtNroCuota.Text)
            !deuda = txtMonto.Text
            !totalcobrado = 0
            !DeudaTotal = txtMonto.Text
            !CuotasDebidas = 1
            !fechavto = DTPFecha.Value
            .Update
            txtNroCuota.Text = Int(txtNroCuota.Text) + 1
            If DTPFecha.Month = 12 Then
                DTPFecha.Month = 1
                DTPFecha.Year = DTPFecha.Year + 1
            Else
                DTPFecha.Month = DTPFecha.Month + 1
            End If
            
        Loop
    End With
    '''cambia el estado a reingresado
    With rsVerificaciones
        If .State = 1 Then .Close
        .Open "SELECT codalumno, estado FROM verificaciones WHERE codalumno=" & frmAnalisisDeCuotas.lblCodAlumno.Caption, Cn, adOpenDynamic, adLockPessimistic
        .Requery
        .MoveFirst
        !estado = "Reingresado"
        .UpdateBatch
    End With
    
    rsAnalisisDeCuenta.Requery
    frmAnalisisDeCuotas.formatoGrilla
    Unload Me
    
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Centrar Me
    frmAnalisisDeCuotas.Enabled = True
End Sub

Private Sub txtCantidadCuotas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNroCuota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
