VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmModificarPlanDePago 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar Plan de Pago"
   ClientHeight    =   2055
   ClientLeft      =   1905
   ClientTop       =   645
   ClientWidth     =   3000
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmModificarPlanDePago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3000
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   169345025
      CurrentDate     =   41353
   End
   Begin VB.TextBox txtMonto 
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
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtNroCuota 
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
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin isButtonTest.isButton cmdAplicar 
      Height          =   420
      Left            =   1560
      TabIndex        =   6
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmModificarPlanDePago.frx":10CA
      Style           =   8
      Caption         =   "     Aceptar"
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdCancelar 
      Height          =   420
      Left            =   1560
      TabIndex        =   7
      Top             =   900
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmModificarPlanDePago.frx":19A4
      Style           =   8
      Caption         =   "     Cancelar"
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00662200&
      BackStyle       =   0  'Transparent
      Caption         =   "Monto $"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00662200&
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimiento"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
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
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00662200&
      BackStyle       =   0  'Transparent
      Caption         =   "Desde Cuota"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "frmModificarPlanDePago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAplicar_Click()
    On Error GoTo LineaError
    Dim DeudaTotal As Currency
    
    If txtNroCuota.Text = "" Then MsgBox "Debe Ingresar desde que cuota se modifica el plan de pago", vbOKOnly + vbInformation, "Plan de Pago": txtNroCuota.SetFocus: Exit Sub
    If txtMonto.Text = "" Then MsgBox "Debe Ingresar un monto", vbOKOnly + vbInformation, "Plan de Pago": txtMonto.SetFocus: Exit Sub
      
    '''actualiza los datos
    With rsAnalisisDeCuenta
        If .State = 1 Then .Close
        .Open "SELECT * FROM plandepago WHERE codalumno=" & Int(frmAnalisisDeCuotas.lblCodalumno.Caption) & "and nrocuota>=" & Int(txtNroCuota.Text), Cn, adOpenDynamic, adLockPessimistic
        Do Until .EOF
            !fechavto = dtpFecha.Value
            !deuda = txtMonto.Text
            If !recargoxfecha = True Then !recargoxfecha = False
            If !recargoxmes = True Then !recargoxmes = False
            !DeudaTotal = txtMonto.Text
            .UpdateBatch
            .MoveNext
            If dtpFecha.Month = 12 Then
                dtpFecha.Month = 1
                dtpFecha.Year = dtpFecha.Year + 1
            Else
                dtpFecha.Month = dtpFecha.Month + 1
            End If
        Loop
    End With
    
    ''' muestra en ventana Analisis de cuotas la info actualizada y cierra
    AnalisisDeCuota
    With frmAnalisisDeCuotas
        Set .grilla1.DataSource = rsAnalisisDeCuenta
        .formatoGrilla
    End With
    Unload Me
    
LineaError: ErrCode
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Centrar Me
    dtpFecha.Value = Date
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmAnalisisDeCuotas.Enabled = True
End Sub

