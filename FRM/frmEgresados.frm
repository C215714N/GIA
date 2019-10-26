VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmEgresados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Egresados"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   Icon            =   "frmEgresados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmEgresados.frx":324A
   ScaleHeight     =   4485
   ScaleWidth      =   4710
   Begin VB.Frame Frame1 
      BackColor       =   &H00662200&
      Caption         =   "Búsqueda"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   4455
      Begin VB.OptionButton optBuscar 
         BackColor       =   &H00662200&
         Caption         =   "Por Nombre"
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
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optBuscar 
         BackColor       =   &H00662200&
         Caption         =   "Por Curso"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   3
         Top             =   840
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   360
         Left            =   1560
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
         CurrentDate     =   41978
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   360
         Left            =   120
         TabIndex        =   0
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
         CurrentDate     =   41978
      End
      Begin isButtonTest.isButton cmdConsultar 
         Height          =   420
         Left            =   3000
         TabIndex        =   9
         Top             =   400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmEgresados.frx":AC67
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
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0 Alumnos"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   2775
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
         TabIndex        =   7
         Top             =   240
         Width           =   495
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
         TabIndex        =   6
         Top             =   240
         Width           =   495
      End
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4683
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
End
Attribute VB_Name = "frmEgresados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConsultar_Click()
    Dim desde As Date
    Dim hasta As Date
    
    desde = Format(dtpDesde.Value, "mm/dd/yyyy")
    hasta = Format(dtpHasta.Value, "mm/dd/yyyy")
    
    With rsEgresados
        If .State = 1 Then .Close
        
        If optBuscar(0).Value = True Then
            .Open "SELECT nya as Alumno,DNI,Fechanac as [Fecha Nacimiento],capac as Curso, Nacionalidad,v.codalumno FROM Verificaciones as V,Egresados as E WHERE v.codalumno=e.codalumno and fecha>=#" & desde & "# and fecha <=#" & hasta & "# ORDER BY nya,fecha", Cn, adOpenDynamic, adLockPessimistic
        Else
            .Open "SELECT nya as Alumno,DNI,Fechanac as [Fecha Nacimiento],capac as Curso,Nacionalidad,v.codalumno FROM Verificaciones as V,Egresados as E WHERE v.codalumno=e.codalumno and fecha>=#" & desde & "# and fecha <=#" & hasta & "# ORDER BY capac,fecha", Cn, adOpenDynamic, adLockPessimistic
        End If
        
        Set grilla.DataSource = rsEgresados
        grilla.Columns(1).Width = 800
        grilla.Columns(2).Width = 1500
        grilla.Columns(0).Width = 3500
        grilla.Columns(3).Width = 2800
        grilla.Columns(5).Width = 0
    End With
    
    Label3.Caption = rsEgresados.RecordCount & " Alumnos"

End Sub

Private Sub Form_Load()
    Centrar Me
    dtpDesde.Day = 1
    dtpDesde.Month = Month(Date)
    dtpDesde.Year = Year(Date)
    dtpHasta.Value = Date
End Sub

Private Sub grilla_DblClick()
    frmExamenes.Show
    frmExamenes.lblOrigen.Caption = "Egresados"
    CodAlumno = grilla.Columns(5).Text
    frmExamenes.txtCodigo.Text = CodAlumno
    With rsVerificaciones
            If .State = 1 Then .Close
            .Open "SELECT nya,capac FROM verificaciones WHERE codalumno=" & CodAlumno, Cn, adOpenDynamic, adLockPessimistic
            frmExamenes.txtAlumno.Text = !NyA
            frmExamenes.txtCurso.Text = !capac
        End With
    
        With rsExamenes
            If .State = 1 Then .Close
            .Open "SELECT Fecha, Modulo as Módulo, teorico as [Examen Teórico], practico as [Examen Práctico], Promedio FROM examenes WHERE codalumno=" & CodAlumno & " ORDER BY fecha,id", Cn, adOpenDynamic, adLockPessimistic
        End With
        
        Set frmExamenes.grilla.DataSource = rsExamenes
        frmExamenes.grilla.Columns(0).Width = 1000
        frmExamenes.grilla.Columns(2).Width = 1400
        frmExamenes.grilla.Columns(3).Width = 1400
        
        frmExamenes.txtCodigo.Enabled = False
        frmExamenes.cmdAgregar.Enabled = False
        Me.Enabled = False
End Sub
