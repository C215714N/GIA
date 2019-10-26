VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmLibroArmado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libros de Aula de Armado"
   ClientHeight    =   4080
   ClientLeft      =   2355
   ClientTop       =   2205
   ClientWidth     =   4515
   Icon            =   "frmLibroArmado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmLibroArmado.frx":324A
   ScaleHeight     =   4080
   ScaleWidth      =   4515
   Begin MSDataGridLib.DataGrid Grilla 
      Height          =   3015
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5318
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
   Begin MSDataListLib.DataCombo dtcCurso 
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   635
      _Version        =   393216
      Style           =   2
      Text            =   "DataCombo1"
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
   Begin isButtonTest.isButton cmdBuscar 
      Height          =   420
      Left            =   3000
      TabIndex        =   4
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmLibroArmado.frx":AC67
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
      Height          =   135
      Left            =   960
      TabIndex        =   3
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dia y Horario"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmLibroArmado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
    '''busca el codigo del curso
    With rsGruposDeArmado
        .Find "curso='" & dtcCurso.Text & "'"
            If .BOF = True Then
                .Find "curso='" & dtcCurso.Text & "'"
            ElseIf .BOF = False Or .EOF = False Then
                Label2.Caption = !ID
                On Error GoTo LineaError
            End If
    End With
    
    '''busca los alumnos del curso
    With rsAlumnosArmado
        If .State = 1 Then .Close
        .Open "SELECT v.codalumno as [Código],v.nya as [Alumnos] FROM verificaciones as v,alumnosdearmado as a WHERE v.codalumno=a.codalumno and grupo=" & rsGruposDeArmado!ID & " ORDER BY nya", Cn, adOpenDynamic, adLockPessimistic
        If .BOF = True And .EOF = True Then
            MsgBox "No se encontraron alumnos en este grupo", vbExclamation, "Gestion Integral del Alumno"
        End If
    End With
    
    Set grilla.DataSource = rsAlumnosArmado
    formatoGrilla
    
LineaError:
    Select Case Err.Number
        Case 3021
            Resume Next
        End Select
End Sub

Private Sub Form_Load()
    Centrar Me
    
    '''consulta los cursos
    With rsGruposDeArmado
        If .State = 1 Then .Close
        .Open "SELECT id,(dia + ' - ' + horario) as curso FROM gruposdearmado", Cn, adOpenDynamic, adLockPessimistic
    End With
    
        ''' carga cursos en DataCombo
    Set dtcCurso.RowSource = rsGruposDeArmado
    dtcCurso.BoundColumn = "curso"
    dtcCurso.ListField = "curso"
    formatoGrilla

End Sub

Private Sub grilla_DblClick()
    frmLibro.Show
    frmLibro.lblFormulario.Caption = Me.Caption
    CodAlumno = frmLibroArmado.grilla.Columns(0).Text
 
    With rsVerificaciones
        If .State = 1 Then .Close
        .Open "SELECT  nya, FechaVerif,cuotas ,capac FROM verificaciones WHERE codalumno=" & CodAlumno, Cn, adOpenDynamic, adLockPessimistic
        frmLibro.lblCodAlumno.Caption = CodAlumno
        frmLibro.lblAlumno.Caption = !NyA
        frmLibro.lblfecha.Caption = !FechaVerif
        frmLibro.lblDuracion.Caption = !cuotas & " Meses"
        frmLibro.lblCapacitacion.Caption = !capac
    End With

    With rsLibro
        If .State = 1 Then .Close
        .Open "SELECT numClase as [N°],Fecha,Tema FROM librodeaula WHERE codalumno=" & CodAlumno & " ORDER BY NumClase", Cn, adOpenDynamic, adLockPessimistic
    End With
    
    Set frmLibro.grilla.DataSource = rsLibro
    frmLibro.formatoGrilla
    Me.Enabled = False
End Sub

Sub formatoGrilla()
    For N = 0 To 1 Step 1
        grilla.Columns(N).Width = 800 + N * 2000
    Next
End Sub

