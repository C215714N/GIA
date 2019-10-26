VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAgregaGestion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Gestión a la Orden de Trabajo"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3060
   Icon            =   "frmAgregaGestion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3060
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
      Height          =   420
      Left            =   1560
      TabIndex        =   7
      Top             =   960
      Width           =   1355
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
      Height          =   420
      Left            =   1560
      TabIndex        =   6
      Top             =   360
      Width           =   1355
   End
   Begin VB.TextBox txtGestion 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1800
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
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
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Gestión"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Orden"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmAgregaGestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdGrabar_Click()
    If txtGestion.Text = "" Then MsgBox "Ingrese la gestión", vbCritical, "Cóndor"
    
    With rsGestionDeOrdenes
        If .State = 1 Then .Close
        .Open "select * from gestiondeordenes", Cn, adOpenDynamic, adLockPessimistic
        .Requery
        .AddNew
        !nroordentrabajo = Int(lblNroOrden.Caption)
        !fecha = dtpFecha.Value
        !gestion = txtGestion.Text
        !Personal = Usuario
        .Update
        .Close
        .Open "select Fecha,Gestion,Personal from gestiondeordenes where nroordentrabajo=" & frmConsultarOrdenes.grilla.Columns(0).Text & " order by fecha,id", Cn, adOpenDynamic, adLockPessimistic
        Set frmConsultarOrdenes.grilla2.DataSource = rsGestionDeOrdenes
            frmConsultarOrdenes.grilla2.Columns(0).Width = 1000
            frmConsultarOrdenes.grilla2.Columns(1).Width = 3000
            frmConsultarOrdenes.grilla2.Columns(2).Width = 1000
    End With
    Unload Me
End Sub

Private Sub Form_Load()
    Centrar Me
    dtpFecha.Value = Date
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmConsultarOrdenes.Enabled = True

End Sub

Private Sub lnlNroOrden_Click()

End Sub
