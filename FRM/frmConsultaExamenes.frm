VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmConsultaExamenes 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar Examenes"
   ClientHeight    =   4380
   ClientLeft      =   2640
   ClientTop       =   1995
   ClientWidth     =   6135
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmConsultaExamenes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6135
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "Busqueda"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   5895
      Begin VB.OptionButton optBuscar 
         BackColor       =   &H00662200&
         Caption         =   "Modulo"
         ForeColor       =   &H8000000F&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.OptionButton optBuscar 
         BackColor       =   &H00662200&
         Caption         =   "Nombre"
         ForeColor       =   &H8000000F&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Value           =   -1  'True
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   375
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   108134401
         CurrentDate     =   41978
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   108134401
         CurrentDate     =   41978
      End
      Begin isButtonTest.isButton cmdConsultar 
         Height          =   420
         Left            =   3000
         TabIndex        =   9
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmConsultaExamenes.frx":10CA
         Style           =   8
         Caption         =   "     Consultar"
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
      Begin isButtonTest.isButton cmdImprimir 
         Height          =   420
         Left            =   4440
         TabIndex        =   10
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmConsultaExamenes.frx":19A4
         Style           =   8
         Caption         =   "     Imprimir"
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
      Begin isButtonTest.isButton cmdDiploma 
         Height          =   420
         Left            =   3000
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmConsultaExamenes.frx":227E
         Style           =   8
         Caption         =   "     Diploma"
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
      Begin isButtonTest.isButton cmdExportar 
         Height          =   420
         Left            =   4440
         TabIndex        =   12
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmConsultaExamenes.frx":2778
         Style           =   8
         Caption         =   "     Exportar"
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "0 Alumnos"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1250
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
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
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
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
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmConsultaExamenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConsultar_Click()
    Dim desde As Date
    Dim hasta As Date
    
    desde = Format(dtpDesde.Value, "mm/dd/yyyy")
    hasta = Format(dtpHasta.Value, "mm/dd/yyyy")
    
    With rsExamenes
        If .State = 1 Then .Close
        If optBuscar(0).Value = True Then
            .Open "SELECT V.Codalumno as [Codigo],nya as [Alumno],TipoDoc, DNI as [Documento],Fechanac as [Nacimiento],Nacionalidad, capac as [Curso], Fecha, Modulo, Promedio FROM Verificaciones as V,Examenes as E WHERE v.codalumno=e.codalumno AND fecha BETWEEN #" & desde & "#  AND #" & hasta & "# ORDER BY nya,fecha", Cn, adOpenDynamic, adLockPessimistic
        Else
            .Open "SELECT V.Codalumno as [Codigo],nya as [Alumno],TipoDoc, DNI as [Documento],Fechanac as [Nacimiento],Nacionalidad, capac as [Curso], Fecha, Modulo, Promedio FROM Verificaciones as V,Examenes as E WHERE v.codalumno=e.codalumno AND fecha BETWEEN #" & desde & "#  AND #" & hasta & "# ORDER BY modulo,fecha", Cn, adOpenDynamic, adLockPessimistic
        End If
        Set grilla.DataSource = rsExamenes
        formatoGrilla
    End With
    formatoGrilla
    Label3.Caption = "Alumnos: " & rsExamenes.RecordCount & " Examenes"
    cmdExportar.Enabled = True
    cmdDiploma.Enabled = False
End Sub

Private Sub cmdDiploma_Click()
    With rsDiplomas
        If .State = 1 Then .Close
        .Open "SELECT codalumno AS [Codigo],modulo,fecharetiro AS [Fecha],retiro FROM examenes WHERE codalumno=" & grilla.Columns(0).Text & " AND modulo='" & grilla.Columns(8).Text & "'", Cn, adOpenDynamic, adLockPessimistic
        .Requery
        .MoveFirst
        If !retiro = "" Or retiro = Null Then
            frmRetiroDiploma.Show
            frmRetiroDiploma.lblCodAlumno.Caption = grilla.Columns(0).Text
            frmRetiroDiploma.lblAlumno.Caption = grilla.Columns(1).Text
            frmRetiroDiploma.lblCurso.Caption = grilla.Columns(6).Text
            frmRetiroDiploma.lblModulo.Caption = grilla.Columns(8).Text
            Me.Enabled = False
        Else
            MsgBox "Este diploma ya ha sido retirado el dia " & !fecharetiro, vbCritical, "Diplomas"
        End If
    End With
    

End Sub

Private Sub cmdExportar_Click()
    Call Exportar_Datagrid(grilla.ApproxCount)
End Sub

Private Sub cmdImprimir_Click()
    If Label3.Caption = "0 Alumnos" Then MsgBox "Primero realice la busqueda", vbCritical, "Examenes": Exit Sub
    
    Set dtrNotas.DataSource = rsExamenes
    dtrNotas.LeftMargin = 1
    dtrNotas.Sections("Seccion5").Controls("lblalumnos").Caption = rsExamenes.RecordCount
    '''dtrNotas.Orientation = dtrorientation.landscape
    dtrNotas.Show
    dtrNotas.Caption = "Examenes"
    Me.Enabled = False
End Sub

Private Sub Form_Load()
    Centrar Me
    dtpDesde.Day = 1
    dtpDesde.Month = Month(Date)
    dtpDesde.Year = Year(Date)
    dtpHasta.Value = Date

End Sub

Private Sub Exportar_Datagrid(TotalFilas As Long)
    Me.MousePointer = vbHourglass
    Set obj_excel = CreateObject("Excel.Application")
    Set obj_Libro = obj_excel.workbooks.Open("T:\Examenes.xls")
    Set obj_Hoja = obj_excel.ActiveSheet
       
    Columna = 0
    For X = 0 To grilla.Columns.Count - 1
        If grilla.Columns(X).Visible Then
            Columna = Columna + 1
            obj_Hoja.Cells(1, Columna) = grilla.Columns(X).Caption
            For Y = 0 To TotalFilas - 1
                obj_Hoja.Cells(Y + 2, Columna) = grilla.Columns(X).CellValue(grilla.GetBookmark(Y))
            Next
        End If
    Next
    obj_excel.Visible = True
    With obj_Hoja
        .Columns("A:Z").autofit
    End With
    
    Me.MousePointer = vbDefault
    Set obj_Hoja = Nothing
    Set obj_Libro = Nothing
    Set obj_excel = Nothing
End Sub

Private Sub grilla_Click()
    If Label3.Caption = "0 Alumnos" Then
        cmdDiploma.Enabled = False
    Else
        cmdDiploma.Enabled = True
    End If
End Sub

Sub formatoGrilla()
    For N = 1 To N = 12
        grilla.Columns(N) = 300
    Next
End Sub

