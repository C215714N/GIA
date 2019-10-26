VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmComisiones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comisiones"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6405
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComisiones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmComisiones.frx":324A
   ScaleHeight     =   3075
   ScaleMode       =   0  'User
   ScaleWidth      =   6400
   Begin MSDataGridLib.DataGrid Grilla 
      Height          =   1995
      Left            =   120
      TabIndex        =   14
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3519
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
   Begin VB.TextBox txtComision 
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
      Left            =   3960
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtCuota 
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
      Left            =   3000
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   360
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
      Format          =   89260033
      CurrentDate     =   41355
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
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
      Format          =   89260033
      CurrentDate     =   41355
   End
   Begin isButtonTest.isButton cmdConsultar 
      Height          =   420
      Left            =   4920
      TabIndex        =   15
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmComisiones.frx":AC67
      Style           =   8
      Caption         =   "       Consultar"
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
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comisiones"
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
      Height          =   225
      Left            =   4920
      TabIndex        =   13
      Top             =   2200
      Width           =   915
   End
   Begin VB.Label lblTotalComisiones 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4920
      TabIndex        =   12
      Top             =   2450
      Width           =   1335
   End
   Begin VB.Label lblTotalCobrado 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4920
      TabIndex        =   11
      Top             =   1750
      Width           =   1335
   End
   Begin VB.Label lblTotalAlumnos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4920
      TabIndex        =   10
      Top             =   1050
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alumnos"
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
      Height          =   225
      Left            =   4920
      TabIndex        =   9
      Top             =   800
      Width           =   690
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cobrados"
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
      Height          =   225
      Left            =   4920
      TabIndex        =   8
      Top             =   1500
      Width           =   780
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
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
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Comision"
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
      Height          =   195
      Left            =   3960
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuota"
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
      Height          =   195
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta"
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
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   540
   End
End
Attribute VB_Name = "frmComisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Centrar Me
    dtpDesde.Value = Date
    dtpHasta.Value = Date
    txtCuota.Text = ""
    txtComision.Text = ""
End Sub

Private Sub cmdConsultar_Click()
    On Error GoTo Error
    If txtCuota.Text = "" Then MsgBox "Debe elegir una cuota a consultar", vbOKOnly + vbInformation, "Comisiones": txtCuota.SetFocus: Exit Sub
    If txtComision.Text = "" Then MsgBox "Debe ingresar una comisión", vbOKOnly + vbInformation, "Comisiones": txtComision.SetFocus: Exit Sub
    Dim fecha1 As Date
    Dim fecha2 As Date
    
    fecha1 = dtpDesde.Value
    fecha2 = dtpHasta.Value
    
    fecha1 = Format(fecha1, "mm/dd/yyyy")
    fecha2 = Format(fecha2, "mm/dd/yyyy")
             
    With rsPlanDePago
        If .State = 1 Then .Close
        .Open "SELECT sum(totalcobrado), sum((totalcobrado)*" & Int(txtComision.Text) & "/100) FROM plandepago WHERE fechapago>=#" & fecha1 & "# and fechapago<=#" & fecha2 & "# and nrocuota=" & Int(txtCuota.Text), Cn, adOpenDynamic, adLockPessimistic
        lblTotalCobrado.Caption = !expr1000
        lblTotalComisiones.Caption = !expr1001
        lblTotalComisiones.Caption = Format(lblTotalComisiones.Caption, "currency")
        lblTotalCobrado.Caption = Format(lblTotalCobrado.Caption, "Currency")
    End With

    With rsPlanDePago
        If .State = 1 Then .Close
        .Open "SELECT codalumno as [Codigo],nrocuota as [N°],fechapago as [Fecha],totalcobrado as [Monto], ((totalcobrado)*" & Int(txtComision.Text) & "/100) as [Comision] FROM plandepago WHERE fechapago>=#" & fecha1 & "# and fechapago<=#" & fecha2 & "# and nrocuota=" & Int(txtCuota.Text), Cn, adOpenDynamic, adLockPessimistic
    End With

    Set grilla.DataSource = rsPlanDePago
    formatoGrilla
    lblTotalAlumnos.Caption = rsPlanDePago.RecordCount

Error:
    End Sub

Sub formatoGrilla()
    Dim w As Integer
    For N = 0 To 4
        If N = 0 Or N > 2 Then
            w = 800
            If N > 2 Then
                grilla.Columns(N).NumberFormat = "$ #####"
            End If
        ElseIf N = 2 Then
            w = 1150
        Else:
            w = 300
        End If
        grilla.Columns(N).Width = w
        grilla.Columns(N).Alignment = dbgCenter
    Next
End Sub

