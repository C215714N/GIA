VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmViaticos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Viáticos"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmViaticos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Palette         =   "frmViaticos.frx":324A
   Picture         =   "frmViaticos.frx":37E4
   ScaleHeight     =   4155
   ScaleWidth      =   4695
   Begin isButtonTest.isButton cmdAgregar 
      Height          =   420
      Left            =   3120
      TabIndex        =   5
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmViaticos.frx":B201
      Style           =   8
      Caption         =   "   Aceptar"
      IconAlign       =   1
      iNonThemeStyle  =   0
      HighlightColor  =   16744576
      FontHighlightColor=   12632256
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
   Begin MSDataGridLib.DataGrid Grilla 
      Height          =   3135
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   5530
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   20
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00884400&
      Caption         =   "Saldo Actual"
      ForeColor       =   &H8000000F&
      Height          =   735
      Left            =   3000
      TabIndex        =   10
      Top             =   120
      Width           =   1575
      Begin VB.Label lblSaldo 
         Alignment       =   2  'Center
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
         Height          =   360
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00884400&
      Caption         =   "Agregar Viático"
      ForeColor       =   &H8000000F&
      Height          =   3135
      Left            =   3000
      TabIndex        =   7
      Top             =   840
      Width           =   1575
      Begin VB.OptionButton optMonto 
         BackColor       =   &H00884400&
         Caption         =   "Rinde"
         ForeColor       =   &H00C0FFC0&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton optMonto 
         BackColor       =   &H00884400&
         Caption         =   "Lleva"
         ForeColor       =   &H00C0C0FF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   4
         Top             =   2040
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   360
         Left            =   120
         TabIndex        =   3
         Top             =   1320
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
         Format          =   89456641
         CurrentDate     =   42277
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00662200&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00662200&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   615
      End
   End
   Begin MSDataListLib.DataCombo dtcAsistente 
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   ""
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
      Caption         =   "Personal"
      ForeColor       =   &H8000000F&
      Height          =   300
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmViaticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    If dtcAsistente.Text = "" Then MsgBox "Elija el Asesor Educativo", vbCritical, "Viáticos": dtcAsistente.SetFocus: Exit Sub
    If txtMonto.Text = "" Then MsgBox "Agregue el monto del viático", vbCritical, "Viáticos": txtMonto.SetFocus: Exit Sub
    If optMonto(0).Value = False And optMonto(1).Value = False Then MsgBox "Elija si el Asesor Educativo lleva o rinde el dinero", vbCritical, "Viáticos": optMonto(0).SetFocus: Exit Sub
    
    With rsViaticos
        If .State = 1 Then .Close
        .Open "SELECT * FROM viaticos", Cn, adOpenDynamic, adLockPessimistic
        .Requery
        .AddNew
        !fecha = DTPFecha.Value
        !asesor = dtcAsistente.Text
        
        If optMonto(0).Value = True Then
            !monto = CSng(txtMonto.Text)
        Else
            !monto = CSng(txtMonto.Text) * -1
        End If
        
        .Update
        .Close
        .Open "SELECT sum(monto) FROM viaticos WHERE asesor='" & dtcAsistente.Text & "'", Cn, adOpenDynamic, adLockPessimistic
        lblSaldo.Caption = FormatCurrency(!expr1000)
        .Close
        .Open "SELECT Fecha,Monto FROM viaticos WHERE asesor='" & dtcAsistente.Text & "' ORDER BY fecha desc,id desc", Cn, adOpenDynamic, adLockPessimistic
    End With
    
    Set grilla.DataSource = rsViaticos
    
    txtMonto.Text = ""
End Sub

Private Sub dtcAsistente_Change()
    With rsViaticos
        If .State = 1 Then .Close
        .Open "SELECT sum(monto) FROM viaticos WHERE asesor='" & dtcAsistente.Text & "'", Cn, adOpenDynamic, adLockPessimistic
        lblSaldo.Caption = Format(!expr1000, "currency")
        .Close
        .Open "SELECT Fecha,Monto FROM viaticos WHERE asesor='" & dtcAsistente.Text & "' ORDER BY fecha desc,id desc", Cn, adOpenDynamic, adLockPessimistic
    End With
    Set grilla.DataSource = rsViaticos
    formatoGrilla
    
End Sub

Private Sub Form_Load()
    Centrar Me
    DTPFecha.Value = Date
    Asistente
    Set dtcAsistente.RowSource = rsPersonal
    dtcAsistente.BoundColumn = "Personal"
    dtcAsistente.ListField = "Personal"
    formatoGrilla
End Sub

Sub formatoGrilla()
    For N = 0 To 1 Step 1
        grilla.Columns(N).Width = 1150 - (N * 250)
        grilla.Columns(N).Alignment = dbgCenter
    Next
End Sub
