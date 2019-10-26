VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmLibroOperador 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libros de Aula de las Reservas"
   ClientHeight    =   5250
   ClientLeft      =   7365
   ClientTop       =   2175
   ClientWidth     =   6885
   Icon            =   "frmLibroOperador.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmLibroOperador.frx":324A
   ScaleHeight     =   5250
   ScaleWidth      =   6885
   Begin VB.Frame lblasistencia 
      BackColor       =   &H00662200&
      Caption         =   "Presentismo"
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
      Left            =   4680
      TabIndex        =   7
      Top             =   0
      Width           =   2055
      Begin VB.TextBox txtAsistencia 
         DataField       =   "PA"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   420
      End
      Begin isButtonTest.isButton cmdAsistencia 
         Height          =   420
         Left            =   600
         TabIndex        =   9
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmLibroOperador.frx":AC67
         Style           =   8
         Caption         =   "       Asistencia"
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
   End
   Begin MSDataGridLib.DataGrid Grilla 
      Height          =   3975
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7011
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00884400&
      Caption         =   "Elija Turno"
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
      TabIndex        =   4
      Top             =   0
      Width           =   4455
      Begin VB.ComboBox cmbHora 
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
         ItemData        =   "frmLibroOperador.frx":B541
         Left            =   1560
         List            =   "frmLibroOperador.frx":B55D
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtpFecha 
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
         CurrentDate     =   41580
      End
      Begin isButtonTest.isButton cmdBuscar 
         Height          =   420
         Left            =   3000
         TabIndex        =   8
         Top             =   400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmLibroOperador.frx":B5B1
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
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hora"
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
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmLibroOperador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fecha As Date

Private Sub cmdAsistencia_Click()
    ''' graba el presente
    With rsAsistencia
        If .State = 1 Then .Close
        .Open "SELECT * FROM reservas WHERE codalumno=" & Int(grilla.Columns(0).Text) & "and hora='" & cmbHora.Text & "' and fecha=#" & fecha & "#", Cn, adOpenDynamic, adLockPessimistic
        .MoveFirst
        !pa = txtAsistencia.Text
        .UpdateBatch
        cmdAsistencia.Enabled = False
    End With
        '''refresca la grilla
        rsReservas.Requery
        formatoGrilla
End Sub

Private Sub cmdBuscar_Click()
    '''control de error si no eligio hora de reserva
    If cmbHora.Text = "" Then MsgBox "Primero debe elegir un horario de turno", vbOKOnly + vbCritical, "Libro de Aula de las Reservas": cmbHora.SetFocus: Exit Sub
    
    ''' asigna formato fecha a la variable para la busqueda
    fecha = Format(DTPFecha.Value, "mm/dd/yyyy")
    
    '''consulta de reservas
    With rsReservas
        If .State = 1 Then .Close
        .Open "SELECT codalumno as Código,nya as [Apellido y Nombre], pa as [P / A] FROM Reservas WHERE fecha=#" & fecha & "# and hora='" & cmbHora.Text & "' ORDER BY nya", Cn, adOpenDynamic, adLockPessimistic
    End With
    
    '''muestra consulta en grilla
    Set grilla.DataSource = rsReservas
    formatoGrilla
    cmdAsistencia.Enabled = False
End Sub

Private Sub Form_Load()
    Centrar Me
    DTPFecha.Value = Date
End Sub

Private Sub grilla_Click()
    txtAsistencia.Enabled = True
    txtAsistencia.Text = grilla.Columns(2).Text
    txtAsistencia.Visible = True
    txtAsistencia.SetFocus
    cmdAsistencia.Enabled = True
    lblAsistencia.Visible = True
End Sub

Private Sub grilla_DblClick()
    frmLibro.Show
    frmLibro.lblFormulario.Caption = Me.Caption
    
    CodAlumno = frmLibroOperador.grilla.Columns(0).Text
 
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
    Exit Sub

End Sub

Private Sub txtAsistencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Sub formatoGrilla()
    Dim w As Integer
    For N = 0 To 2
        If N = 1 Then
            w = 3800
        Else:
            w = 400 * (N + 2)
        End If
        grilla.Columns(N).Width = w
    Next
End Sub
