VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmInformeVerificados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Verificados"
   ClientHeight    =   3450
   ClientLeft      =   4260
   ClientTop       =   1995
   ClientWidth     =   7500
   Icon            =   "frmInformeVerificados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmInformeVerificados.frx":324A
   ScaleHeight     =   3450
   ScaleWidth      =   7500
   Begin VB.Frame Frame1 
      BackColor       =   &H00662200&
      Caption         =   "Totales"
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
      Height          =   2535
      Left            =   5760
      TabIndex        =   0
      Top             =   720
      Width           =   1605
      Begin VB.Label lblMontoTotal 
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
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   1355
      End
      Begin VB.Label lblVerificados 
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
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1355
      End
      Begin VB.Label lblSuscriptos 
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
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1355
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Suscripciones"
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
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Verificaciones"
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
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingresos"
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
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Width           =   660
      End
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   21
      RowDividerStyle =   0
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
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   360
      Left            =   120
      TabIndex        =   8
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
      Format          =   89456641
      CurrentDate     =   41345
   End
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   360
      Left            =   1560
      TabIndex        =   9
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
      Format          =   89456641
      CurrentDate     =   41345
   End
   Begin isButtonTest.isButton cmdBuscar 
      Height          =   420
      Left            =   3000
      TabIndex        =   12
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmInformeVerificados.frx":AC67
      Style           =   8
      Caption         =   "       Buscar"
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Height          =   225
      Left            =   1560
      TabIndex        =   11
      Top             =   120
      Width           =   450
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   510
   End
End
Attribute VB_Name = "frmInformeVerificados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim fecha1 As Date
    Dim fecha2 As Date

Private Sub cmdBuscar_Click()
    
    
    fecha1 = dtpDesde.Value
    fecha2 = dtpHasta.Value
    fecha1 = Format(fecha1, "mm/dd/yyyy")
    fecha2 = Format(fecha2, "mm/dd/yyyy")
    
    With rsInformeSuscripciones
        If .State = 1 Then .Close
        .Open "SELECT count(*) as Suscripciones,sum(verificado) as Verificados,sum(totalcurso) as Monto FROM informesuscripciones WHERE fechaV>=#" & fecha1 & "# and fechaV<=#" & fecha2 & "# ", Cn, adOpenDynamic, adLockPessimistic
        lblSuscriptos.Caption = !Suscripciones
        If .RecordCount = 1 And lblSuscriptos.Caption = 0 Then lblMontoTotal.Caption = 0: lblVerificados.Caption = 0: Exit Sub
        

        lblVerificados.Caption = !verificados
        lblMontoTotal.Caption = FormatCurrency(!monto)
    
    End With

    With rsInformeSuscripciones
        If .State = 1 Then .Close
        .Open "SELECT Asistente,count(*) as S,sum(verificado) as V,sum(totalcurso) as Monto FROM informesuscripciones WHERE fechaV>=#" & fecha1 & "# and fechaV<=#" & fecha2 & "# group by asistente", Cn, adOpenDynamic, adLockPessimistic
    End With
        
    Set grilla.DataSource = rsInformeSuscripciones
    formatoGrilla

End Sub

Private Sub Form_Load()
    Centrar Me
    dtpDesde.Value = Date
    dtpHasta.Value = Date
End Sub

Sub formatoGrilla()
    grilla.Columns(0).Width = 3200
    grilla.Columns(1).Width = 300
    grilla.Columns(2).Width = 300
    grilla.Columns(3).Width = 1000
    grilla.Columns(3).NumberFormat = "$ #####"
    
End Sub

Private Sub grilla_DblClick()
    With rsAnalisisInforme
        If .State = 1 Then .Close
        .Open "SELECT nya as Alumno, Direccion,Localidad,Tel1 as Telefono1,ptel1 as [Telefono Alumno],tel2 as Telefono2,ptel2 as Celular,Fechasus as Suscripcion,Fechaverif as Verificacion, Totalcurso as [Total Curso] FROM verificaciones WHERE fechaverif>=#" & fecha1 & "# and fechaverif<=#" & fecha2 & "# and asistente='" & grilla.Columns(0).Text & "' ORDER BY fechaverif", Cn, adOpenDynamic, adLockPessimistic
        frmAnalisisInforme.Show
        Set frmAnalisisInforme.grilla.DataSource = rsAnalisisInforme
        frmAnalisisInforme.grilla.Columns(0).Width = 3000
    End With
    frmAnalisisInforme.lblOrigen.Caption = "verificaciones"
    Me.Enabled = False

End Sub

