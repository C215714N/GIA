VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmPresupuesto 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuesto"
   ClientHeight    =   5745
   ClientLeft      =   3900
   ClientTop       =   1545
   ClientWidth     =   9285
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmPresupuesto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9285
   Begin VB.Frame Frame2 
      BackColor       =   &H00662200&
      Caption         =   "Totales"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   2475
      Left            =   7560
      TabIndex        =   5
      Top             =   960
      Width           =   1606
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label lblPagado 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblDeuda 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Pagado"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Deuda"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSDataGridLib.DataGrid Grilla 
      Height          =   4455
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   21
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00662200&
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3975
      Begin VB.ComboBox cmbMonth 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "frmPresupuesto.frx":10CA
         Left            =   120
         List            =   "frmPresupuesto.frx":10F2
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   420
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpYear 
         Height          =   375
         Left            =   1560
         TabIndex        =   0
         Top             =   420
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
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
         CustomFormat    =   "yyyy"
         Format          =   127270915
         CurrentDate     =   36526
         MaxDate         =   401876
         MinDate         =   36526
      End
      Begin isButtonTest.isButton cmdInforme 
         Height          =   420
         Left            =   2520
         TabIndex        =   12
         Top             =   420
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmPresupuesto.frx":115B
         Style           =   8
         Caption         =   "     Informe"
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "AÑO"
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
         Height          =   255
         Left            =   1560
         TabIndex        =   2
         Top             =   200
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MES"
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
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   200
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ActualizarPresupuesto As Boolean

Private Sub Form_Load()
    Centrar Me
    If Month(Date) = 1 Then
        cmbMonth.Text = "Enero"
    ElseIf Month(Date) = 2 Then
        cmbMonth.Text = "Febrero"
    ElseIf Month(Date) = 3 Then
        cmbMonth.Text = "Marzo"
    ElseIf Month(Date) = 4 Then
        cmbMonth.Text = "Abril"
    ElseIf Month(Date) = 5 Then
        cmbMonth.Text = "Mayo"
    ElseIf Month(Date) = 6 Then
        cmbMonth.Text = "Junio"
    ElseIf Month(Date) = 7 Then
        cmbMonth.Text = "Julio"
    ElseIf Month(Date) = 8 Then
        cmbMonth.Text = "Agosto"
    ElseIf Month(Date) = 9 Then
        cmbMonth.Text = "Septiembre"
    ElseIf Month(Date) = 10 Then
        cmbMonth.Text = "Octubre"
    ElseIf Month(Date) = 11 Then
        cmbMonth.Text = "Noviembre"
    Else
        cmbMonth.Text = "Diciembre"
    End If
    
    dtpYear.Value = Date
    ActualizarPresupuesto = False
End Sub

Private Sub cmbMonth_Click()
    With rsPresupuesto
        If ActualizarPresupuesto = True Then
            If .State = 1 Then .UpdateBatch: .Close: ActualizarPresupuesto = False
        Else
            If .State = 1 Then .Close: ActualizarPresupuesto = False
        End If
        .Open "SELECT sum(deuda) FROM presupuesto WHERE mes='" & cmbMonth.Text & "' AND Año=" & Year(dtpYear.Value), Cn, adOpenDynamic, adLockPessimistic
        lblDeuda.Caption = Format(!expr1000, "currency")
        .Close
        .Open "SELECT sum(pagado) FROM presupuesto WHERE mes='" & cmbMonth.Text & "' AND Año=" & Year(dtpYear.Value), Cn, adOpenDynamic, adLockPessimistic
        lblPagado.Caption = Format(!expr1000, "currency")
        .Close
        .Open "SELECT sum(saldo) FROM presupuesto WHERE mes='" & cmbMonth.Text & "' AND Año=" & Year(dtpYear.Value), Cn, adOpenDynamic, adLockPessimistic
        lblSaldo.Caption = Format(!expr1000, "currency")
        .Close
        .Open "SELECT Cuenta,Deuda,Pagado,Saldo,Observaciones,id FROM presupuesto WHERE mes='" & cmbMonth.Text & "' AND Año=" & Year(dtpYear.Value) & " ORDER BY cuenta", Cn, adOpenDynamic, adLockPessimistic
        Set grilla.DataSource = rsPresupuesto
    End With
    formatoGrilla
    cmdInforme.Enabled = True
End Sub

Private Sub cmdInforme_Click()
    Set dtrPresupuesto.DataSource = rsPresupuesto
    dtrPresupuesto.Caption = "Presupuesto"
    dtrPresupuesto.Sections("Seccion4").Controls("lblmes").Caption = frmPresupuesto.cmbMonth.Text
    dtrPresupuesto.Sections("Seccion4").Controls("lblAño").Caption = frmPresupuesto.dtpYear.Value
    dtrPresupuesto.Sections("Seccion5").Controls("lblDeudaTotal").Caption = lblDeuda.Caption
    dtrPresupuesto.Sections("Seccion5").Controls("lblPagadoTotal").Caption = lblPagado.Caption
    dtrPresupuesto.Sections("Seccion5").Controls("lblSaldoTotal").Caption = lblSaldo.Caption
    dtrPresupuesto.Show
    Me.Enabled = False
End Sub

Private Sub dtpYear_Change()
    With rsPresupuesto
        If ActualizarPresupuesto = True Then
            If .State = 1 Then .UpdateBatch: .Close: ActualizarPresupuesto = False
        Else
            If .State = 1 Then .Close: ActualizarPresupuesto = False
        End If
        .Open "SELECT sum(deuda) FROM presupuesto WHERE mes='" & cmbMonth.Text & "' AND Año=" & Year(dtpYear.Value), Cn, adOpenDynamic, adLockPessimistic
        lblDeuda.Caption = Format(!expr1000, "currency")
        .Close
        .Open "SELECT sum(pagado) FROM presupuesto WHERE mes='" & cmbMonth.Text & "' AND Año=" & Year(dtpYear.Value), Cn, adOpenDynamic, adLockPessimistic
        lblPagado.Caption = Format(!expr1000, "currency")
        .Close
        .Open "SELECT sum(saldo) FROM presupuesto WHERE mes='" & cmbMonth.Text & "' AND Año=" & Year(dtpYear.Value), Cn, adOpenDynamic, adLockPessimistic
        lblSaldo.Caption = Format(!expr1000, "currency")
        .Close
        .Open "SELECT Cuenta,Deuda,Pagado,Saldo,Observaciones,id FROM presupuesto WHERE mes='" & cmbMonth.Text & "' AND Año=" & Year(dtpYear.Value) & " ORDER BY cuenta", Cn, adOpenDynamic, adLockPessimistic
        Set grilla.DataSource = rsPresupuesto
    End With
    formatoGrilla
    cmdInforme.Enabled = True
End Sub

Private Sub grilla_KeyPress(KeyAscii As Integer)
    ActualizarPresupuesto = True
    If KeyAscii = 13 And grilla.Col = 2 Then
        grilla.Columns(3).Text = grilla.Columns(1).Text - grilla.Columns(2).Text
    ElseIf KeyAscii = 13 And grilla.Col = 1 Then
        grilla.Columns(3).Text = grilla.Columns(1).Text - grilla.Columns(2).Text
    End If
End Sub

Private Sub formatoGrilla()
    Dim w As Integer
    For N = 0 To 5 Step 1
        If N > 0 And N < 4 Then
            w = 800
        ElseIf N = 5 Then
            w = 0
        Else:
            w = 2400
        End If
        grilla.Columns(N).Width = w
    Next
End Sub
