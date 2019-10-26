VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmLibro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libro de Aula"
   ClientHeight    =   3375
   ClientLeft      =   345
   ClientTop       =   1920
   ClientWidth     =   9615
   Icon            =   "frmLibro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmLibro.frx":324A
   ScaleHeight     =   3403.361
   ScaleMode       =   0  'User
   ScaleWidth      =   9615
   Begin isButtonTest.isButton cmdNuevo 
      Height          =   420
      Left            =   8150
      TabIndex        =   19
      Top             =   238
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmLibro.frx":AC67
      Style           =   8
      Caption         =   "       Nuevo"
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
   Begin isButtonTest.isButton cmdModificar 
      Height          =   420
      Left            =   8145
      TabIndex        =   20
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmLibro.frx":B541
      Style           =   8
      Caption         =   "       Editar"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00884400&
      Caption         =   "Temario"
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
      TabIndex        =   14
      Top             =   2280
      Width           =   3735
      Begin VB.TextBox txtNumClase 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtTema 
         Enabled         =   0   'False
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
         Left            =   1080
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Tema de Clase"
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
         Height          =   300
         Left            =   1080
         TabIndex        =   18
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Clase"
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
         Height          =   300
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   3015
      Left            =   3960
      TabIndex        =   10
      Top             =   240
      Width           =   4125
      _ExtentX        =   7276
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   21
      AllowAddNew     =   -1  'True
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
      Caption         =   "Datos del Alumno"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3735
      Begin VB.Label Label10 
         Height          =   375
         Left            =   3120
         TabIndex        =   17
         Top             =   2400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblCodAlumno 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblCapacitacion 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   6
         Top             =   1700
         Width           =   3495
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Capacitacion"
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
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblAlumno 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   3
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido y Nombre"
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
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo"
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
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblfecha 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicio"
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
         Left            =   1080
         TabIndex        =   9
         Top             =   255
         Width           =   1215
      End
      Begin VB.Label lblDuracion 
         BackColor       =   &H00FFFFFF&
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
         Left            =   2520
         TabIndex        =   5
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Duracion:"
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
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   255
         Width           =   735
      End
   End
   Begin isButtonTest.isButton cmdAgregar 
      Height          =   420
      Left            =   8150
      TabIndex        =   21
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmLibro.frx":BE1B
      Style           =   8
      Caption         =   "       Aceptar"
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
   Begin isButtonTest.isButton cmdEliminar 
      Height          =   420
      Left            =   8150
      TabIndex        =   22
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmLibro.frx":C6F5
      Style           =   8
      Caption         =   "       Eliminar"
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
   Begin isButtonTest.isButton cmdImprimir 
      Height          =   420
      Left            =   8150
      TabIndex        =   23
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmLibro.frx":CFCF
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
   Begin VB.Label lblFormulario 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   16
      Top             =   3600
      Visible         =   0   'False
      Width           =   3375
   End
End
Attribute VB_Name = "frmLibro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fecha As Date

Private Sub cmdAgregar_Click()

'''guarda la modificacion al libro de aula
If ModiLibro = False Then
    With rsLibro
        If .State = 1 Then .Close
            .Open "SELECT * FROM LibroDeAula", Cn, adOpenDynamic, adLockPessimistic
            .AddNew
            !CodAlumno = Int(lblCodAlumno.Caption)
            !numclase = Int(txtNumClase.Text)
            !Tema = txtTema.Text
            !fecha = Date
            .Update
            .Close
            .Open "SELECT numClase as [N°],Fecha,Tema FROM librodeaula WHERE codalumno=" & CodAlumno & " ORDER BY NumClase", Cn, adOpenDynamic, adLockPessimistic
    End With
Else
    With rsLibro
        fecha = Format(Label10.Caption, "mm/dd/yyyy")
        .Close
        .Open "SELECT numClase as [N°],Fecha,Tema FROM librodeaula WHERE codalumno=" & CodAlumno & " and fecha=#" & fecha & "#", Cn, adOpenDynamic, adLockPessimistic
        .MoveFirst
            !N° = Int(txtNumClase.Text)
            !Tema = txtTema.Text
            .UpdateBatch
            .Close
            .Open "SELECT numClase as [N°],Fecha,Tema FROM librodeaula WHERE codalumno=" & CodAlumno & " ORDER BY NumClase", Cn, adOpenDynamic, adLockPessimistic
    End With
End If

'''muestra grilla actualizada
Set grilla.DataSource = rsLibro
    formatoGrilla

botones True, False
    
LineaError:
    Select Case Err.Number
        Case 3021
            Resume Next
        End Select
End Sub

Private Sub cmdEliminar_Click()
    On Error GoTo LineaError
    If txtTema.Text = "" Then MsgBox "Primero debe elegir una clase", vbCritical + vbOKOnly, "Libro de Aula": Exit Sub
    fecha = Format(Label10.Caption, "mm/dd/yyyy")
    With rsLibro
        .Close
        .Open "SELECT numClase as [N°],Fecha,Tema FROM librodeaula WHERE codalumno=" & CodAlumno & " and numclase=" & Int(txtNumClase.Text) & " and fecha=#" & fecha & "#", Cn, adOpenDynamic, adLockPessimistic
        .MoveFirst
        .Delete
        .Requery
        .Close
        .Open "SELECT numClase as [N°],Fecha,Tema FROM librodeaula WHERE codalumno=" & CodAlumno & " ORDER BY NumClase", Cn, adOpenDynamic, adLockPessimistic
    End With
    
    Set grilla.DataSource = rsLibro
    formatoGrilla
    txtNumClase.Text = ""
    txtTema.Text = ""
    botones True, False

LineaError:
    Select Case Err.Number
        Case 3021
            Resume Next
        End Select
End Sub

Private Sub cmdImprimir_Click()
    Set dtrLibro.DataSource = rsLibro
    dtrLibro.Sections("Sección4").Controls("lbldesde").Caption = lblfecha.Caption
    dtrLibro.Sections("Sección4").Controls("lblhasta").Caption = lblDuracion.Caption
    dtrLibro.Sections("Sección4").Controls("etiqueta13").Caption = lblCapacitacion.Caption
    dtrLibro.Sections("Sección4").Controls("lblinforme").Caption = "Libro de Aula de " & lblAlumno.Caption
    dtrLibro.Show
    dtrLibro.Caption = "Libro de Aula del alumno " & lblAlumno.Caption
    Me.Enabled = False

End Sub

Private Sub cmdModificar_Click()
    If txtTema.Text = "" Then MsgBox "Primero debe elegir una clase", vbCritical + vbOKOnly, "Libro de Aula": Exit Sub
    ModiLibro = True
    botones False, True
End Sub

Private Sub cmdNuevo_Click()
    botones False, True
    txtTema.Text = ""
    txtNumClase.SetFocus
    ModiLibro = False
End Sub

Private Sub Form_Load()
    Centrar Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lblFormulario.Caption = "Libros de Aula de las Reservas" Then
        frmLibroOperador.Enabled = True
    Else
        frmLibroArmado.Enabled = True
    End If
End Sub

Private Sub botones(estado1 As Boolean, estado2 As Boolean)
    cmdNuevo.Enabled = estado1
    cmdModificar.Enabled = estado1
    cmdEliminar.Enabled = estado1
    cmdAgregar.Enabled = estado2
    txtTema.Enabled = estado2
    txtNumClase.Enabled = estado2
End Sub

Private Sub grilla_Click()
    botones True, False
    txtTema.Text = grilla.Columns(2).Text
    txtNumClase.Text = grilla.Columns(0).Text
    Label10.Caption = grilla.Columns(1).Text
End Sub

Sub formatoGrilla()
On Error GoTo LineaError
    Dim w As Integer
    For N = 0 To 2 Step 1
        w = 300 + (N * 850)
        If N < 2 Then
            grilla.Columns(N).Alignment = dbgCenter
        End If
        grilla.Columns(N).Width = w
    Next
        rsLibro.MoveLast

LineaError:
    Select Case Err.Number
        Case 3021
            Resume Next
        End Select
End Sub
