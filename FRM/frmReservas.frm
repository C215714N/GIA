VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmReservas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reservas"
   ClientHeight    =   4575
   ClientLeft      =   3180
   ClientTop       =   2100
   ClientWidth     =   12525
   Icon            =   "frmReservas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmReservas.frx":324A
   ScaleHeight     =   4575
   ScaleWidth      =   12525
   Begin VB.TextBox txtAsistencia 
      DataField       =   "PA"
      DataSource      =   "Data1"
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
      Left            =   6960
      TabIndex        =   12
      Top             =   4080
      Width           =   3975
   End
   Begin VB.Frame frameHorarios 
      BackColor       =   &H00884400&
      Caption         =   "Horarios"
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
      Height          =   1440
      Left            =   100
      TabIndex        =   3
      ToolTipText     =   "Seleccione el Horario"
      Top             =   3000
      Width           =   2700
      Begin VB.OptionButton rbt4 
         BackColor       =   &H00884400&
         Caption         =   "12:30 Hs"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1000
         Width           =   975
      End
      Begin VB.OptionButton rbt1 
         BackColor       =   &H00884400&
         Caption         =   "08:00 Hs"
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
         Left            =   240
         TabIndex        =   10
         Top             =   250
         Width           =   975
      End
      Begin VB.OptionButton rbt2 
         BackColor       =   &H00884400&
         Caption         =   "09:30 Hs"
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
         Left            =   240
         TabIndex        =   9
         Top             =   500
         Width           =   975
      End
      Begin VB.OptionButton rbt3 
         BackColor       =   &H00884400&
         Caption         =   "11:00 Hs"
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
         Left            =   240
         TabIndex        =   8
         Top             =   750
         Width           =   975
      End
      Begin VB.OptionButton rbt5 
         BackColor       =   &H00884400&
         Caption         =   "14:00 Hs"
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
         Left            =   1440
         TabIndex        =   7
         Top             =   250
         Width           =   975
      End
      Begin VB.OptionButton rbt6 
         BackColor       =   &H00884400&
         Caption         =   "15:30 Hs"
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
         Left            =   1440
         TabIndex        =   6
         Top             =   500
         Width           =   975
      End
      Begin VB.OptionButton rbt7 
         BackColor       =   &H00884400&
         Caption         =   "17:00 Hs"
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
         Left            =   1440
         TabIndex        =   5
         Top             =   750
         Width           =   975
      End
      Begin VB.OptionButton rbt8 
         BackColor       =   &H00884400&
         Caption         =   "18:30 Hs"
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
         Left            =   1440
         TabIndex        =   4
         Top             =   1000
         Width           =   975
      End
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   3420
      Left            =   2880
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6033
      _Version        =   393216
      AllowUpdate     =   0   'False
      ForeColor       =   0
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
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2670
      Left            =   105
      TabIndex        =   1
      ToolTipText     =   "Seleccione la Fecha"
      Top             =   360
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4710
      _Version        =   393216
      ForeColor       =   8930304
      BackColor       =   8930304
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   16777215
      StartOfWeek     =   41680898
      TitleBackColor  =   8930304
      TitleForeColor  =   16777215
      TrailingForeColor=   14737632
      CurrentDate     =   40179
      MinDate         =   36161
   End
   Begin MSDataListLib.DataCombo dtcAlumno 
      Height          =   360
      Left            =   2880
      TabIndex        =   13
      Top             =   4080
      Width           =   3975
      _ExtentX        =   7011
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
      Left            =   11040
      TabIndex        =   18
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmReservas.frx":AC67
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
   Begin isButtonTest.isButton btnAceptar 
      Height          =   420
      Left            =   11040
      TabIndex        =   19
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmReservas.frx":B541
      Style           =   8
      Caption         =   "       Aceptar"
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
   Begin isButtonTest.isButton cmdCancelar 
      Height          =   420
      Left            =   11040
      TabIndex        =   20
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmReservas.frx":BE1B
      Style           =   8
      Caption         =   "       Cancelar"
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
   Begin isButtonTest.isButton btnAgregar 
      Height          =   420
      Left            =   11040
      TabIndex        =   21
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmReservas.frx":C6F5
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
   Begin isButtonTest.isButton btnEliminar 
      Height          =   420
      Left            =   11040
      TabIndex        =   22
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmReservas.frx":CFCF
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
   Begin isButtonTest.isButton btnAsistencia 
      Height          =   420
      Left            =   11040
      TabIndex        =   23
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmReservas.frx":D8A9
      Style           =   8
      Caption         =   "       Asistencia"
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
   Begin VB.Label lblreservas 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Height          =   240
      Left            =   2280
      TabIndex        =   16
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblNyA 
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
      Left            =   2880
      TabIndex        =   15
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label lblAsistencia 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Asistencia"
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
      Left            =   6960
      TabIndex        =   14
      Top             =   3840
      Width           =   825
   End
   Begin VB.Label lblCodalumno 
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
      Left            =   2880
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reservas para este Turno :"
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
      TabIndex        =   17
      Top             =   120
      Width           =   2100
   End
End
Attribute VB_Name = "frmReservas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Index As String
Dim BuscarAlumno As Boolean
Option Compare Text

Private Sub btnAgregar_Click()
    dtcAlumno.Visible = True
    btnAceptar.Enabled = True
    lblNyA.Visible = True
    btnEliminar.Enabled = False
    btnAgregar.Enabled = False
    btnAsistencia.Enabled = False
    dtcAlumno.Enabled = True
    dtcAlumno.SetFocus
    cmdCancelar.Enabled = True
    cmdBuscar.Enabled = False
    BuscarAlumno = False
    

End Sub

Private Sub btnAsistencia_Click()
    dtcAlumno.Visible = True
    dtcAlumno.Enabled = False
    lblNyA.Visible = True
    txtAsistencia.Visible = True
    lblAsistencia.Visible = True
    btnAceptar.Enabled = False
    txtAsistencia.SetFocus
End Sub

Private Sub btnEliminar_Click()
        a = MsgBox("¿Esta seguro que desea eliminar al alumno " & grilla.Columns(0).Text & "?", vbYesNo, "Eliminar Selección")
        If a = vbYes Then
            With rsReservas
                Dim fecha As Date
                fecha = Format(MonthView1.Value, "mm/dd/yyyy")
                If .State = 1 Then .Close
                .Open "SELECT * FROM reservas WHERE nya='" & dtcAlumno.Text & "' and fecha=#" & fecha & "# and hora='" & Index & "'", Cn, adOpenDynamic, adLockPessimistic
                .MoveFirst
                .Delete
                .Update
                Reservas
            End With
       End If
End Sub


Private Sub btnMañana_Click()
    'Activa las opciones de horario del turno mañana
    'rbtNadaTarde.Value = True
    'rbtNadaMañana.Value = True
    frameMañana.Visible = True
    frameTarde.Visible = False
    'Desactivar los botones de modificar y eliminar
    btnEliminar.Enabled = False
End Sub

Private Sub Reservas()
    dtcAlumno.Visible = False
    lblNyA.Visible = False
    btnAceptar.Enabled = False
    btnEliminar.Enabled = False
    btnAsistencia.Enabled = True

Dim fecha As Date
fecha = Format(MonthView1.Value, "mm/dd/yyyy")

    With rsVerificaciones
        If .State = 1 Then .Close
        .Open "SELECT codalumno,(nya + ' - ' + capac) as alumno FROM verificaciones WHERE capac='Operador de PC' or capac='Programación' or capac='Diseño Web' or capac='Diseño Gráfico' or capac='Programación + Access' or capac = 'Redes Sociales' ORDER BY nya", Cn, adOpenDynamic, adLockPessimistic
    End With
    
        ''' carga alumnos en DataCombo
    Set dtcAlumno.RowSource = rsVerificaciones
    dtcAlumno.BoundColumn = "alumno"
    dtcAlumno.ListField = "alumno"
    
    
    With rsReservas
        If .State = 1 Then .Close
        .Open "SELECT nya as [Apellido y Nombre], pa as [P/A],Fecha, hora as Horario FROM Reservas WHERE fecha=#" & fecha & "# and hora ='" & Index & "' ORDER BY nya", Cn, adOpenDynamic, adLockPessimistic
        If .BOF Or .EOF Then lblreservas.Caption = 0: lblAsistencia.Visible = False: txtAsistencia.Text = "": txtAsistencia.Visible = False: btnAgregar.Enabled = True: btnEliminar.Enabled = False: btnAsistencia.Enabled = False:  Exit Sub
    End With
    
    Set grilla.DataSource = rsReservas
    formatoGrilla
    lblreservas.Caption = rsReservas.RecordCount
    Equipos
    
    If Int(lblreservas.Caption) >= rsEquipos!Equipos Then
        btnAgregar.Enabled = False
    Else
        btnAgregar.Enabled = True
    End If
    txtAsistencia.Visible = False
    txtAsistencia.Text = ""
    lblAsistencia.Visible = False
End Sub

Private Sub btnTarde_Click()
    'Activa las opciones de horario del turno tarde
    frameMañana.Visible = False
    frameTarde.Visible = True

     'Desactivar los botones de modificar y eliminar
    btnEliminar.Enabled = False
   
End Sub




Private Sub cmdBuscar_Click()
    BuscarAlumno = True
    dtcAlumno.Visible = True
    btnAceptar.Enabled = True
    lblNyA.Visible = True
    btnEliminar.Enabled = False
    btnAgregar.Enabled = True
    btnAsistencia.Enabled = False
    dtcAlumno.Enabled = True
    dtcAlumno.SetFocus
    cmdCancelar.Enabled = False
    cmdBuscar.Enabled = True

End Sub

Private Sub cmdCancelar_Click()
    dtcAlumno.Visible = False
    btnAceptar.Enabled = False
    lblNyA.Visible = False
    btnEliminar.Enabled = False
    btnAgregar.Enabled = True
    btnAsistencia.Enabled = False
    dtcAlumno.Enabled = False
    cmdCancelar.Enabled = False
    cmdBuscar.Enabled = True
End Sub





Private Sub Form_Load()
    Centrar Me
    MonthView1.Value = Date
    ''' consulta alumnos
    With rsVerificaciones
        If .State = 1 Then .Close
        .Open "SELECT max(codalumno),nya,(nya + ' - ' + capac) as Alumno,capac FROM verificaciones WHERE capac='Operador de PC' or capac='Programación' or capac='Diseño Web' or capac='Diseño Gráfico' or capac='Programación + Access' or capac='Redes Sociales' group by nya,capac ORDER BY nya", Cn, adOpenDynamic, adLockPessimistic
    End With

        ''' carga alumnos en DataCombo
    Set dtcAlumno.RowSource = rsVerificaciones
    dtcAlumno.BoundColumn = "alumno"
    dtcAlumno.ListField = "alumno"

End Sub

Private Sub grilla_Click()
    lblNyA.Visible = True
    dtcAlumno.Text = grilla.Columns(0).Text
    dtcAlumno.Visible = True
    btnEliminar.Enabled = True
    btnAsistencia.Enabled = True
    dtcAlumno.Enabled = False

End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    Reservas
End Sub

Private Sub rbt1_Click()
    Index = rbt1.Caption
    Reservas ' Modulo que busca la reserva actual
End Sub

Private Sub rbt2_Click()
    Index = rbt2.Caption
    Reservas ' Modulo que busca la reserva actual
End Sub

Private Sub rbt3_Click()
    Index = rbt3.Caption
    Reservas ' Modulo que busca la reserva actual
End Sub

Private Sub rbt4_Click()
    Index = rbt4.Caption
    Reservas ' Modulo que busca la reserva actual
End Sub

Private Sub rbt5_Click()
    Index = rbt5.Caption
    Reservas ' Modulo que busca la reserva actual
End Sub

Private Sub rbt6_Click()
    Index = rbt6.Caption
    Reservas ' Modulo que busca la reserva actual
End Sub

Private Sub rbt7_Click()
    Index = rbt7.Caption
    Reservas ' Modulo que busca la reserva actual
End Sub

Private Sub rbt8_Click()
    Index = rbt8.Caption
    Reservas ' Modulo que busca la reserva actual
End Sub

Private Sub rbt9_Click()
    Index = rbt9.Caption
    Reservas ' Modulo que busca la reserva actual
End Sub

Private Sub btnAceptar_Click()
    '''control de errores
On Error GoTo Error
    'agrega la reserva
If BuscarAlumno = False Then
    With rsVerificaciones
        If .State = 1 Then .Close
        .Open "SELECT codalumno,(nya + ' - ' + capac) as alumno FROM verificaciones WHERE capac='Operador de PC' or capac='Programación' or capac='Diseño Web' or capac='Diseño Gráfico' or capac='Programación + Access' or capac='Redes Sociales' ORDER BY nya", Cn, adOpenDynamic, adLockPessimistic
        .Find "alumno='" & dtcAlumno.Text & "'"
        lblCodAlumno.Caption = !CodAlumno
    End With

    With rsReservas
        If .State = 1 Then .Close
        .Open "SELECT * FROM reservas", Cn, adOpenDynamic, adLockPessimistic
        .Requery
            .AddNew
            !NyA = dtcAlumno.Text
            !fecha = MonthView1.Value
           
            If rbt1.Value = True Then
                !hora = rbt1.Caption
            ElseIf rbt2.Value = True Then
                !hora = rbt2.Caption
            ElseIf rbt3.Value = True Then
                !hora = rbt3.Caption
            ElseIf rbt4.Value = True Then
                !hora = rbt4.Caption
            ElseIf rbt5.Value = True Then
                !hora = rbt5.Caption
            ElseIf rbt6.Value = True Then
                !hora = rbt6.Caption
            ElseIf rbt7.Value = True Then
                !hora = rbt7.Caption
            ElseIf rbt8.Value = True Then
                !hora = rbt8.Caption
            ElseIf rbt9.Value = True Then
                !hora = rbt9.Caption
            End If
            !pa = ""
            !CodAlumno = Int(lblCodAlumno.Caption)
            .Update
            .Close
    End With
    
    'Determina el estado de los botones
    btnAceptar.Enabled = False
    dtcAlumno.Visible = False
    lblNyA.Visible = False
    btnAsistencia.Enabled = True
    cmdCancelar.Enabled = False
    cmdBuscar.Enabled = True
    Reservas
    Exit Sub
Error:
    MsgBox "El alumno " & dtcAlumno.Text & " ya tiene reserva en este turno", vbOKOnly + vbCritical, "Reservas"
Else
    With rsVerificaciones
        If .State = 1 Then .Close
        .Open "SELECT codalumno,(nya + ' - ' + capac) as alumno FROM verificaciones WHERE capac='Operador de PC' or capac='Programación' or capac='Diseño Web' or capac='Diseño Gráfico' or capac='Programación + Access' or capac='Redes Sociales' ORDER BY nya", Cn, adOpenDynamic, adLockPessimistic
        .Find "alumno='" & dtcAlumno.Text & "'"
        lblCodAlumno.Caption = !CodAlumno
    End With

    
    With rsReservas
        If .State = 1 Then .Close
            .Open "SELECT nya as [Apellido y Nombre], pa as [P/A],Fecha, hora as Horario FROM Reservas WHERE codalumno=" & Int(lblCodAlumno.Caption) & " ORDER BY  fecha desc,hora", Cn, adOpenDynamic, adLockPessimistic
            Set grilla.DataSource = rsReservas
            lblreservas.Caption = rsReservas.RecordCount
    End With
   
    formatoGrilla

    With rsVerificaciones
        If .State = 1 Then .Close
        .Open "SELECT codalumno,(nya + ' - ' + capac) as alumno FROM verificaciones WHERE capac='Operador de PC' or capac='Programación' or capac='Diseño Web' or capac='Diseño Gráfico' or capac='Programación + Access' ORDER BY nya", Cn, adOpenDynamic, adLockPessimistic
    End With
    
        ''' carga alumnos en DataCombo
    Set dtcAlumno.RowSource = rsVerificaciones
    dtcAlumno.BoundColumn = "alumno"
    dtcAlumno.ListField = "alumno"

End If
End Sub

Private Sub txtAsistencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rsReservas
            Dim fecha As Date
            fecha = Format(MonthView1.Value, "mm/dd/yyyy")
            If .State = 1 Then .Close
            .Open "SELECT * FROM reservas WHERE nya='" & dtcAlumno.Text & "' and fecha=#" & fecha & "# and hora='" & Index & "'", Cn, adOpenDynamic, adLockPessimistic
            .MoveFirst
            !pa = txtAsistencia.Text
            .UpdateBatch
            Reservas
        End With
    End If
End Sub

Sub formatoGrilla()
    Dim w As Integer
    For N = 0 To 3 Step 1
        If N = 1 Or N = 3 Then
            w = 800
        Else:
            w = 4800 - (N * 1825)
        End If
        grilla.Columns(N).Width = w
    Next
    grilla.Columns(0).Width = 4800
    grilla.Columns(1).Width = 800
    grilla.Columns(2).Width = 1200
    grilla.Columns(3).Width = 800
End Sub

