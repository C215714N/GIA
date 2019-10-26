VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmInformeBajas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Bajas"
   ClientHeight    =   4170
   ClientLeft      =   3285
   ClientTop       =   2565
   ClientWidth     =   7575
   Icon            =   "frmInformeBajas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmInformeBajas.frx":324A
   ScaleHeight     =   4170
   ScaleWidth      =   7575
   Begin VB.Frame Frame2 
      BackColor       =   &H00662200&
      Caption         =   "Cant.Bajas"
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
      Height          =   975
      Left            =   6120
      TabIndex        =   6
      Top             =   0
      Width           =   1335
      Begin VB.Label lblCantidadBajas 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   400
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid Grilla 
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5106
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00884400&
      Caption         =   "Buscar Bajas"
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
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   360
         Left            =   1560
         TabIndex        =   2
         Top             =   480
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
         CurrentDate     =   42108
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   480
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
         CurrentDate     =   42108
      End
      Begin isButtonTest.isButton cmdBuscar 
         Height          =   420
         Left            =   3000
         TabIndex        =   8
         Top             =   400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmInformeBajas.frx":AC67
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
      Begin isButtonTest.isButton cmdImprimir 
         Height          =   420
         Left            =   4440
         TabIndex        =   9
         Top             =   400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmInformeBajas.frx":B541
         Style           =   8
         Caption         =   "       Imprimir"
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmInformeBajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
    If dtpDesde.Value > dtpHasta.Value Then MsgBox "Revise las fechas", vbCritical: dtpDesde.SetFocus: Exit Sub
    
    Dim fecha1 As Date
    Dim fecha2 As Date
    
    fecha1 = Format(dtpDesde.Value, "mm/dd/yyyy")
    fecha2 = Format(dtpHasta.Value, "mm/dd/yyyy")

    With rsBajas
        If .State = 1 Then .Close
        '.Open "SELECT distinct p.CodAlumno,p.Nya as Alumno,capac as Curso,Deuda as [Valor de Cuota],datediff('m',fechapago,now)*-30 as [Sit Cartera],(SELECT min(nrocuota) FROM plandepago as p, bajas as b,verificaciones as v WHERE tipodepago='BAJA' and p.codalumno=b.codalumno and p.codalumno=v.codalumno) as [Nº Cuota],Fecha,Motivo,PagoBaja FROM plandepago as p,verificaciones as v,bajas as b WHERE v.codalumno=p.codalumno and v.codalumno=b.codalumno and fechapago>=#" & fecha1 & "# and fechapago<=#" & fecha2 & "# and tipodepago='BAJA' ORDER BY fecha,p.codalumno", cn, adOpenDynamic, adLockPessimistic
        .Open "SELECT distinct v.CodAlumno,v.nya as Alumno, capac as Curso,Deuda as [Valor de Cuota],sitcartera *-1 as [Sit Cartera], b.NroCuota,Fecha,Motivo, PagoBaja FROM verificaciones as v,bajas as b,plandepago WHERE v.codalumno=b.codalumno and v.codalumno=plandepago.codalumno and fechapago>=#" & fecha1 & "# and fechapago<=#" & fecha2 & "# and tipodepago='BAJA' ORDER BY fecha,v.codalumno", Cn, adOpenDynamic, adLockPessimistic
        Set grilla.DataSource = rsBajas
        If .EOF Or .BOF Then lblCantidadBajas.Caption = "0 Alumnos": cmdImprimir.Enabled = False: Exit Sub
        cmdImprimir.Enabled = True
        lblCantidadBajas.Caption = rsBajas.RecordCount & " Alumnos"
    End With
End Sub

Private Sub cmdImprimir_Click()
    dtrBajas.Show
    dtrBajas.Caption = "Informe de Bajas"
    '''dtrBajas.Orientation = rptOrientLandscape
    dtrBajas.LeftMargin = 1
    dtrBajas.Sections("Sección4").Controls("lbldesde").Caption = dtpDesde.Value
    dtrBajas.Sections("Sección4").Controls("lblhasta").Caption = dtpHasta.Value
    dtrBajas.Sections("Sección4").Controls("lblalumnos").Caption = lblCantidadBajas.Caption

    Set dtrBajas.DataSource = rsBajas
    Me.Enabled = False
End Sub

Private Sub Form_Load()
    Centrar Me
    dtpDesde.Value = Date
    dtpHasta.Value = Date
End Sub
