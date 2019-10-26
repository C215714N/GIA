VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCambioEstado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Estado de Orden de Trabajo"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   Icon            =   "frmCambioEstado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4335
   Begin VB.ComboBox cmbEstado 
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
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   840
      Width           =   2655
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
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   1355
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      Left            =   2880
      TabIndex        =   0
      Top             =   720
      Width           =   1355
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
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
      Format          =   220463105
      CurrentDate     =   42264
   End
   Begin VB.Label lblNroOrden 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Orden"
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
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label2 
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
      Left            =   1320
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Estado"
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
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "frmCambioEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGrabar_Click()
    If cmbEstado.Text = "" Then MsgBox "Ingrese el nuevo estado", vbCritical, "Cóndor"
    
    With rsGestionDeOrdenes
        If .State = 1 Then .Close
        .Open "select * from gestiondeordenes", Cn, adOpenDynamic, adLockPessimistic
        .Requery
        .AddNew
        !nroordentrabajo = Int(lblNroOrden.Caption)
        !fecha = dtpFecha.Value
        !gestion = "Cambio de estado a " & cmbEstado.Text
        !Personal = Usuario
        .Update
        .Close
        .Open "select Fecha,Gestion,Personal from gestiondeordenes where nroordentrabajo=" & frmConsultarOrdenes.grilla.Columns(0).Text & " order by fecha,id", Cn, adOpenDynamic, adLockPessimistic
        Set frmConsultarOrdenes.grilla2.DataSource = rsGestionDeOrdenes
        frmConsultarOrdenes.grilla2.Columns(0).Width = 1000
        frmConsultarOrdenes.grilla2.Columns(1).Width = 3000
        frmConsultarOrdenes.grilla2.Columns(2).Width = 1000

    End With
    
    With rsconsultarordenes
    '    If .State = 1 Then .Close
    '    .Open "select nroorden,estado from ordenesdetrabajo where nroorden=" & Int(lblNroOrden.Caption), cn, adOpenDynamic, adLockPessimistic
        .Find "nroorden=" & lblNroOrden.Caption
        !estado = cmbEstado.Text
        .UpdateBatch
        .Close
        .Open "select NroOrden, Cliente,Equipo,Problema,FechaRecibido as Recepción,Estado,FechaEntregado as Entrega from ordenesdetrabajo as o,clientes as c where c.codcliente=o.codcliente and nroorden=" & lblNroOrden.Caption, Cn, adOpenDynamic, adLockPessimistic
    End With
    
    
    Set frmConsultarOrdenes.grilla.DataSource = rsconsultarordenes
    
    frmConsultarOrdenes.grilla.Columns(3).Width = 0
    frmConsultarOrdenes.grilla.Columns(5).Width = 0
    frmConsultarOrdenes.grilla.Columns(6).Width = 0
    frmConsultarOrdenes.lblestado.Caption = frmConsultarOrdenes.grilla.Columns(5).Text

    
    
    Unload Me

End Sub

Private Sub Form_Load()
    Centrar Me
    dtpFecha.Value = Date
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmConsultarOrdenes.Enabled = True
End Sub
