VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmDetalleStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle del Status de la Base"
   ClientHeight    =   5460
   ClientLeft      =   8940
   ClientTop       =   1995
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmDetalleStatus.frx":0000
   ScaleHeight     =   5460
   ScaleWidth      =   9555
   Begin VB.TextBox txtFiltroPDP 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3840
      TabIndex        =   12
      Top             =   120
      Width           =   855
   End
   Begin MSDataGridLib.DataGrid GrillaPlanDePago 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   21
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9
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
   Begin MSDataGridLib.DataGrid GrillaMarcas 
      Height          =   4095
      Left            =   4800
      TabIndex        =   0
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   21
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9
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
   Begin VB.Label lblDeudaMarcas 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   11
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label lblDeudaPDP 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   10
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total Deuda:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   5040
      Width           =   1245
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total Deuda:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   5040
      Width           =   1245
   End
   Begin VB.Label lblTotalPlanDePago 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label lblTotalMarcas 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Total de Alumnos:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   4680
      Width           =   1725
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total de Alumnos:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4680
      Width           =   1725
   End
   Begin VB.Label Label2 
      Caption         =   "Tabla: Marcas"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Tabla: Plan de Pago"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmDetalleStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbFiltroPDP_Change()
    With rsPlanDePago
        If .State = 1 Then .Close
        .Open "SELECT codalumno as Codigo, min(nrocuota) as Cuota,sum(deudatotal) as Deuda, sum(cuotasdebidas)*30-30 as Categoria FROM plandepago WHERE cuotasdebidas=1 and categoria =" & Int(cmbFiltroPDP.Text) & " group by codalumno", Cn, adOpenDynamic, adLockPessimistic
        Set GrillaPlanDePago.DataSource = rsPlanDePago
        
    End With
End Sub

Private Sub Form_Load()
    Centrar Me
    
    With rsMarcar
        If .State = 1 Then .Close
        .Open "SELECT codalumno as Codigo,Cuota,Deuda,Cantidadcuotas *30-30 as Categoria FROM marcas WHERE deuda>1 ORDER BY codalumno", Cn, adOpenDynamic, adLockPessimistic
        Set GrillaMarcas.DataSource = rsMarcar
       
    End With
    
    With rsPlanDePago
        If .State = 1 Then .Close
        .Open "SELECT codalumno as Codigo, min(nrocuota) as Cuota,sum(deudatotal) as Deuda, sum(cuotasdebidas)*30-30 as Categoria FROM plandepago WHERE cuotasdebidas=1 group by codalumno", Cn, adOpenDynamic, adLockPessimistic
        Set GrillaPlanDePago.DataSource = rsPlanDePago
        
    End With
    
     lblTotalMarcas.Caption = rsMarcar.RecordCount
     lblTotalPlanDePago.Caption = rsPlanDePago.RecordCount
     
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmStatus.Enabled = True
End Sub

Private Sub txtFiltroPDP_KeyPress(KeyAscii As Integer)
        Dim c As Integer
        c = Int(txtFiltroPDP.Text)
                If KeyAscii = 13 Then
        With rsPlanDePago
            If .State = 1 Then .Close
            .Open "SELECT codalumno as codigo min(nrocuota) as cuota, sum(deudatotal) as deuda,sum(cuotasdebidas)*30-30 as categoria FROM plandepago WHERE cuotasdebidas like " & c & " group by codalumno, cn, adOpenDynamic, adLockPessimistic"
            Set GrillaPlanDePago.DataSource = rsPlanDePago
            lblTotalPlanDePago.Caption = rsPlanDePago.RecordCount
        End With
    End If
    End Sub
