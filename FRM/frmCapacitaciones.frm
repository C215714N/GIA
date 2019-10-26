VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmCapacitaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Capacitaciones"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   Icon            =   "frmCapacitaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmCapacitaciones.frx":324A
   ScaleHeight     =   4365
   ScaleMode       =   0  'User
   ScaleWidth      =   5100.001
   Begin VB.TextBox txtCurso 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   0   'False
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
      Caption         =   "Capacitaciones"
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
   Begin isButtonTest.isButton cmdGrabar 
      Height          =   420
      Left            =   3579
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCapacitaciones.frx":AC67
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
   Begin isButtonTest.isButton cmdCancelar 
      Height          =   420
      Left            =   3579
      TabIndex        =   5
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCapacitaciones.frx":B541
      Style           =   8
      Caption         =   "       Cancelar"
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
   Begin isButtonTest.isButton cmdNuevo 
      Height          =   420
      Left            =   3579
      TabIndex        =   8
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCapacitaciones.frx":BE1B
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
      Left            =   3579
      TabIndex        =   6
      Top             =   960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCapacitaciones.frx":C6F5
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
   Begin isButtonTest.isButton cmdSalir 
      Height          =   420
      Left            =   3579
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCapacitaciones.frx":CFCF
      Style           =   8
      Caption         =   "       Volver"
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
      Left            =   3579
      TabIndex        =   9
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmCapacitaciones.frx":D8A9
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre del Curso:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label lblID 
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
      Left            =   3600
      TabIndex        =   2
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmCapacitaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    HabilitarBotones True, False
    txtCurso.Locked = True
    grilla.Enabled = True
End Sub

Private Sub cmdEliminar_Click()
    If txtCurso.Text = "" Then
        MsgBox "Primero debe elejir una capacitación", vbOKOnly + vbInformation, "Capacitación"
    Else
        a = MsgBox("¿Está seguro que desea eliminar esa capacitación?", vbYesNo + vbQuestion, "Capacitaciones")
        If a = vbYes Then
            With rsCapacitaciones
                .Requery
                .Find "capacitacion='" & lblID.Caption & "'"
                .Delete
                .Update
            End With
            grilla.Refresh
            txtCurso.Text = ""
        End If
    End If
End Sub

Private Sub cmdGrabar_Click()
If txtCurso.Text = "" Then MsgBox "Primero debe agregar el nombre del curso", vbOKOnly + vbInformation, "Capacitaciones": txtCurso.SetFocus: Exit Sub
On Error GoTo LineaError

If Modi = True Then
    With rsCapacitaciones
        .Requery
        .Find "capacitacion='" & lblID.Caption & "'"
        !capacitacion = txtCurso.Text
        .UpdateBatch
    End With
    HabilitarBotones True, False
    grilla.Enabled = True
    txtCurso.Locked = True
    txtCurso.Text = ""
Else
    With rsCapacitaciones
        .Requery
        .AddNew
        !capacitacion = txtCurso.Text
        .Update
    End With
    HabilitarBotones True, False
    grilla.Enabled = True
    txtCurso.Locked = True
    txtCurso.Text = ""
End If

LineaError:
    Select Case Err.Number
        Case 3021
            Resume Next
        End Select
End Sub

Private Sub cmdModificar_Click()
    If txtCurso.Text = "" Then
        MsgBox "Primero debe elejir una capacitación", vbOKOnly + vbInformation, "Capacitación"
    Else
        txtCurso.Locked = False
        txtCurso.SetFocus
        HabilitarBotones False, True
        grilla.Enabled = False
        Modi = True
    End If
End Sub

Private Sub cmdNuevo_Click()
    txtCurso.Locked = False
    HabilitarBotones False, True
    txtCurso.SetFocus
    grilla.Enabled = False
    txtCurso.Text = ""
    Modi = False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Centrar Me
    Capacitaciones
    Set grilla.DataSource = rsCapacitaciones
    grilla.Columns(0).Width = 3200
    txtCurso.Locked = True
    txtCurso.Text = ""
    HabilitarBotones True, False
End Sub

Private Sub grilla_Click()
    lblID.Caption = grilla.Text
    txtCurso.Text = grilla.Text
End Sub

Sub HabilitarBotones(estado1 As Boolean, estado2 As Boolean)
    cmdNuevo.Enabled = estado1
    cmdModificar.Enabled = estado1
    cmdSalir.Enabled = estado1
    cmdEliminar.Enabled = estado1
    cmdGrabar.Enabled = estado2
    cmdCancelar.Enabled = estado2
End Sub

Private Sub txtCurso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
