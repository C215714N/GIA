VERSION 5.00
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmGruposArmado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupos de Armado"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3315
   Icon            =   "frmGruposArmado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmGruposArmado.frx":324A
   ScaleHeight     =   1980
   ScaleWidth      =   3315
   Begin VB.Frame Frame1 
      BackColor       =   &H00884400&
      Caption         =   "Cursos"
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
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3050
      Begin VB.ComboBox cmbDia 
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
         ItemData        =   "frmGruposArmado.frx":AC67
         Left            =   120
         List            =   "frmGruposArmado.frx":AC7D
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox cmbHorario 
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
         ItemData        =   "frmGruposArmado.frx":ACB4
         Left            =   120
         List            =   "frmGruposArmado.frx":ACD6
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1080
         Width           =   1335
      End
      Begin isButtonTest.isButton cmdAlumnos 
         Height          =   420
         Left            =   1560
         TabIndex        =   5
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmGruposArmado.frx":AD3E
         Style           =   8
         Caption         =   "       Alumnos"
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
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmGruposArmado.frx":B618
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
      Begin isButtonTest.isButton cmdEliminar 
         Height          =   420
         Left            =   1560
         TabIndex        =   7
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmGruposArmado.frx":BEF2
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
         Caption         =   "Horario"
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
         TabIndex        =   4
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Día"
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
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmGruposArmado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAlumnos_Click()
    ''' control de curso
    If cmbDia.Text = "" Then MsgBox "Primero elija un día de cursada", vbCritical + vbOKOnly, "Administración de Grupos de Armado": cmbDia.SetFocus: Exit Sub
    If cmbHorario.Text = "" Then MsgBox "Debe elegir un horario de cursada", vbOKOnly + vbCritical, "Administración de Grupos de Armado": cmbHorario.SetFocus: Exit Sub
    
    ''' busca el curso
    With rsGruposDeArmado
        If .State = 1 Then .Close
        .Open "SELECT * FROM gruposdearmado WHERE dia='" & cmbDia.Text & "' and horario='" & cmbHorario.Text & "'", Cn, adOpenDynamic, adLockPessimistic
        If .BOF Or .EOF Then MsgBox "No hay curso abierto el día " & cmbDia.Text & " a las " & cmbHorario.Text, vbCritical + vbOKOnly, "Administración de Grupos de Armado": cmbDia.SetFocus: Exit Sub
        .MoveFirst
        CodCurso = !ID
    End With
    
    '''muestra formulario de gestion de alumnos del curso
    frmGestionAlumnos.Show
    frmGestionAlumnos.lblDia.Caption = cmbDia.Text
    frmGestionAlumnos.lblHorario.Caption = cmbHorario.Text
    Me.Enabled = False
End Sub

Private Sub cmdEliminar_Click()
''' comprobacion de informacion
    If cmbDia.Text = "" Then MsgBox "Debe elegir un día", vbOKOnly + vbCritical, "Grupos de Armado": cmbDia.SetFocus: Exit Sub
    If cmbHorario.Text = "" Then MsgBox "Debe elegir un horario", vbOKOnly + vbCritical, "Grupos de Armado": cmbHorario.SetFocus: Exit Sub
    
    '''agrega grupo
    If MsgBox("Elminar el grupo del día " & cmbDia.Text & " a las " & cmbHorario.Text & "?", vbYesNo + vbQuestion, "Grupos de Armado") = vbYes Then
        With rsGruposDeArmado
            If .State = 1 Then .Close
            .Open "SELECT * FROM Gruposdearmado WHERE dia='" & cmbDia.Text & "' and horario='" & cmbHorario.Text & "'", Cn, adOpenDynamic, adLockPessimistic
            If .RecordCount < 1 Then MsgBox "El grupo no existe", vbOKOnly + vbCritical, "Grupos de Armado": cmbDia.SetFocus: Exit Sub
            .Requery
            .Delete
            .Update
        End With
    End If
    MsgBox "El grupo fue eliminado exitosamente", , "Grupos de Armado"

End Sub

Private Sub cmdNuevo_Click()
''' comprobacion de informacion
    If cmbDia.Text = "" Then MsgBox "Debe elegir un día", vbOKOnly + vbCritical, "Grupos de Armado": cmbDia.SetFocus: Exit Sub
    If cmbHorario.Text = "" Then MsgBox "Debe elegir un horario", vbOKOnly + vbCritical, "Grupos de Armado": cmbHorario.SetFocus: Exit Sub
    
    '''agrega grupo
    If MsgBox("¿Crear un grupo el día " & cmbDia.Text & " a las " & cmbHorario.Text & "?", vbYesNo + vbQuestion, "Grupos de Armado") = vbYes Then
        With rsGruposDeArmado
            If .State = 1 Then .Close
            .Open "SELECT * FROM Gruposdearmado WHERE dia='" & cmbDia.Text & "' and horario='" & cmbHorario.Text & "'", Cn, adOpenDynamic, adLockPessimistic
            If .RecordCount > 0 Then MsgBox "El grupo ya existe", vbOKOnly + vbCritical, "Grupos de Armado": cmbDia.SetFocus: Exit Sub
            .Requery
            .AddNew
            !dia = cmbDia.Text
            !horario = cmbHorario.Text
            .Update
        End With
        MsgBox "El grupo fue creado exitosamente", , "Grupos de Armado"
    End If
End Sub

Private Sub Form_Load()
    Centrar Me
End Sub
