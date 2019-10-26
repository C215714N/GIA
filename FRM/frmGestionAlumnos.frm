VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmGestionAlumnos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión de Alumnos"
   ClientHeight    =   3075
   ClientLeft      =   5370
   ClientTop       =   2265
   ClientWidth     =   5310
   HasDC           =   0   'False
   Icon            =   "frmGestionAlumnos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmGestionAlumnos.frx":324A
   ScaleHeight     =   3075
   ScaleWidth      =   5310
   Begin MSDataGridLib.DataGrid Grilla 
      Height          =   2655
      Left            =   1560
      TabIndex        =   5
      Top             =   240
      Width           =   3615
      _ExtentX        =   6376
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
   Begin VB.TextBox txtCodAlumno 
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
      Top             =   1440
      Width           =   1355
   End
   Begin isButtonTest.isButton cmdAgregar 
      Height          =   420
      Left            =   120
      TabIndex        =   8
      Top             =   1900
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmGestionAlumnos.frx":AC67
      Style           =   8
      Caption         =   "       Agregar"
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
   Begin isButtonTest.isButton cmdQuitar 
      Height          =   420
      Left            =   120
      TabIndex        =   9
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmGestionAlumnos.frx":B541
      Style           =   8
      Caption         =   "       Eliminar"
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
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Alumno"
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
      TabIndex        =   6
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label lblHorario 
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
      TabIndex        =   4
      Top             =   840
      Width           =   1355
   End
   Begin VB.Label lblDia 
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
      TabIndex        =   3
      Top             =   240
      Width           =   1355
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Horario"
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
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmGestionAlumnos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAgregar_Click()
        
    If txtCodAlumno.Text = "" Then MsgBox "Ingrese el código de alumno a agregar en el curso", vbCritical + vbOKOnly, "Gestión de Alumnos": txtCodAlumno.SetFocus: Exit Sub
    
    '''comprueba q no esté el alumno asignado a un curso
    With rsAlumnosArmado
        If .State = 1 Then .Close
        .Open "SELECT * FROM alumnosdearmado WHERE grupo=" & CodCurso & " and codalumno=" & Int(txtCodAlumno.Text), Cn, adOpenDynamic, adLockPessimistic
        If .BOF Or .EOF Then
            .Requery
            .AddNew
            !CodAlumno = Int(txtCodAlumno.Text)
            !grupo = CodCurso
            .Update
        Else
            MsgBox "El alumno ya tiene asignado un curso", vbCritical + vbOKOnly, "Gestión de Alumnos"
        End If
        
        .Close
        .Open "SELECT v.nya as Alumnos,v.codalumno FROM verificaciones as v,alumnosdearmado as a WHERE v.codalumno=a.codalumno and a.grupo=" & CodCurso, Cn, adOpenDynamic, adLockPessimistic
    End With
    
    Set grilla.DataSource = rsAlumnosArmado
    grilla.Columns(0).Width = 3000
    grilla.Columns(1).Width = 0
    txtCodAlumno.Text = ""
    txtCodAlumno.SetFocus

End Sub

Private Sub cmdQuitar_Click()
    If MsgBox("¿Está seguro que desea quitar al alumno " & grilla.Columns(0).Text & " del grupo?", vbYesNo + vbQuestion, "Gestión de Alumnos") = vbYes Then
        Label4.Caption = grilla.Columns(1).Text
        
        '''carga a los alumnos del curso y los muestra en la grilla
        With rsAlumnosArmado
            If .State = 1 Then .Close
            '.Open "SELECT v.nya,a.grupo,a.codalumno as Alumnos FROM verificaciones as v,alumnosdearmado as a WHERE v.codalumno=a.codalumno and v.nya='" & Label4.Caption & "'", cn, adOpenDynamic, adLockPessimistic
            '.Open "SELECT v.nya as Alumnos FROM verificaciones as v,alumnosdearmado as a WHERE v.codalumno=a.codalumno and v.nya='" & Label4.Caption & "'", cn, adOpenDynamic, adLockPessimistic
            .Open "SELECT * FROM alumnosdearmado WHERE codalumno=" & Label4.Caption, Cn, adOpenDynamic, adLockPessimistic
            .MoveFirst
            .Delete
            .Update
            .Close
            .Open "SELECT v.nya as Alumnos,v.codalumno FROM verificaciones as v,alumnosdearmado as a WHERE v.codalumno=a.codalumno and a.grupo=" & CodCurso, Cn, adOpenDynamic, adLockPessimistic
        End With
    
        Set grilla.DataSource = rsAlumnosArmado
        grilla.Columns(0).Width = 3000
        grilla.Columns(1).Width = 0
    End If
End Sub

Private Sub Form_Load()
    Centrar Me
    '''carga a los alumnos del curso y los muestra en la grilla
    With rsAlumnosArmado
        If .State = 1 Then .Close
        .Open "SELECT v.nya as Alumnos,v.codalumno FROM verificaciones as v,alumnosdearmado as a WHERE v.codalumno=a.codalumno and a.grupo=" & CodCurso & " ORDER BY nya", Cn, adOpenDynamic, adLockPessimistic
    End With
    
    Set grilla.DataSource = rsAlumnosArmado
    grilla.Columns(0).Width = 3000
    grilla.Columns(1).Width = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmGruposArmado.Enabled = True
End Sub
