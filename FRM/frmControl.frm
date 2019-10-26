VERSION 5.00
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmControl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control"
   ClientHeight    =   3165
   ClientLeft      =   5160
   ClientTop       =   3645
   ClientWidth     =   5565
   Icon            =   "frmControl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmControl.frx":324A
   ScaleHeight     =   3165
   ScaleWidth      =   5565
   Begin VB.Frame Frame1 
      BackColor       =   &H00884400&
      Caption         =   "Control"
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
      Height          =   3015
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   3855
      Begin VB.TextBox txtDerechoExamen 
         Alignment       =   1  'Right Justify
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00884400&
         Caption         =   "Recargos"
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
         Left            =   2520
         TabIndex        =   13
         Top             =   0
         Width           =   1335
         Begin VB.TextBox txtRecargoXMes 
            Alignment       =   1  'Right Justify
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
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtRecargoXFecha 
            Alignment       =   1  'Right Justify
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
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fuera de Mes"
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
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fuera Fecha"
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
            TabIndex        =   15
            Top             =   360
            Width           =   1020
         End
      End
      Begin VB.TextBox txtMatricula 
         Alignment       =   1  'Right Justify
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
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtSucursal 
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
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtEmpresa 
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
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1920
         Width           =   3615
      End
      Begin VB.TextBox txtUltimaFecha 
         Alignment       =   2  'Center
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtNroAsiento 
         Alignment       =   1  'Right Justify
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtCodAlumno 
         Alignment       =   1  'Right Justify
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
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Der.Examen"
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
         Left            =   1320
         TabIndex        =   19
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Matrícula"
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
         TabIndex        =   12
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursal"
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
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   960
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa"
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
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro Asiento"
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
         Height          =   195
         Left            =   2640
         TabIndex        =   9
         Top             =   2280
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Última Fecha"
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
         Left            =   1320
         TabIndex        =   8
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cod. Alumno"
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
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1200
      End
   End
   Begin isButtonTest.isButton cmdGrabar 
      Height          =   420
      Left            =   4080
      TabIndex        =   20
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmControl.frx":AC67
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
      Left            =   4080
      TabIndex        =   21
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmControl.frx":B541
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
   Begin isButtonTest.isButton cmdModificar 
      Height          =   420
      Left            =   4080
      TabIndex        =   22
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmControl.frx":BE1B
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
   Begin isButtonTest.isButton cmdCerrar 
      Height          =   420
      Left            =   4080
      TabIndex        =   23
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmControl.frx":C6F5
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
End
Attribute VB_Name = "frmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    HabilitarBotones True, False
    HabilitarCuadros True
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdGrabar_Click()
    ''' declara variable para comprobar fechas
    Dim fecha As Date
    fecha = Format(rsControl!ultimafecha, "dd/mm/yyyy")

    ''' actualiza tabla control
    With rsControl
        .Requery
        .MoveFirst
        !CodAlumno = txtCodAlumno.Text
        !recargopormes = txtRecargoXMes.Text
        !recargoporfecha = txtRecargoXFecha.Text
        !nroasiento = txtNroAsiento.Text
        !ultimafecha = txtUltimaFecha.Text
        !empresa = txtEmpresa.Text
        !sucursal = txtSucursal.Text
        !matricula = Int(txtMatricula.Text)
        !derechoExamen = txtDerechoExamen.Text
        .UpdateBatch
        
        ''' actualiza barra de tìtulo del mdi
        MDI.Caption = "Gestión Integral del Alumno - " & !empresa & " - " & !sucursal
        
    End With
    HabilitarCuadros True
    HabilitarBotones True, False
    
    ''' si se cambio por una fecha anterior, actualizar plan de pago y marcas
    If fecha > rsControl!ultimafecha Then
    
        fecha = Format(rsControl!ultimafecha, "mm/dd/yyyy")
        '''actualiza a cuota quitando recargos
        With rsRestaurarPlanDePago
            If .State = 1 Then .Close
            .Open "SELECT * FROM plandepago WHERE fechavto>#" & fecha & "# ORDER BY codalumno", Cn, adOpenDynamic, adLockPessimistic
            .MoveFirst
            Do Until .EOF
                If !recargoxfecha = True Then
                    !recargoxfecha = False
                End If
                If !recargoxmes = True Then
                    !recargoxmes = False
                End If
                !DeudaTotal = !deuda - !totalcobrado
                .UpdateBatch
                .MoveNext
            Loop
        End With
        
        ''' elimina situaciones de cartera archivadas
        With rsSituacionesDeCartera
            If .State = 1 Then .Close
            .Open "SELECT * FROM situacionesdecartera WHERE fecha>=#" & fecha & "#", Cn, adOpenDynamic, adLockPessimistic
            .Requery
            .MoveFirst
            Do Until .EOF
                .Delete
                .Update
                .MoveFirst
            Loop
        End With
        
        ''' elimina totales de situaciones de cartera archivadas
        With rsTotalesSituaciones
            If .State = 1 Then .Close
            .Open "SELECT * FROM TotalesSituaciones WHERE fecha>=#" & fecha & "#", Cn, adOpenDynamic, adLockPessimistic
            .Requery
            .MoveFirst
            Do Until .EOF
                .Delete
                .Update
                .MoveFirst
            Loop
        End With
    End If
End Sub

Private Sub cmdModificar_Click()
    HabilitarCuadros False
    HabilitarBotones False, True
    txtCodAlumno.SetFocus
End Sub

Private Sub Form_Load()
    Centrar Me
    Control
    With rsControl
        .Requery
        .MoveFirst
        txtCodAlumno.Text = !CodAlumno
        txtRecargoXFecha.Text = !recargoporfecha
        txtRecargoXMes.Text = !recargopormes
        txtUltimaFecha.Text = !ultimafecha
        txtNroAsiento.Text = !nroasiento
        txtEmpresa.Text = !empresa
        txtSucursal.Text = !sucursal
        txtMatricula.Text = !matricula
        txtDerechoExamen.Text = !derechoExamen
    End With
End Sub

Sub HabilitarBotones(estado1 As Boolean, estado2 As Boolean)
    cmdModificar.Enabled = estado1
    cmdCerrar.Enabled = estado1
    cmdGrabar.Enabled = estado2
    cmdCancelar.Enabled = estado2
End Sub

Sub HabilitarCuadros(estado1 As Boolean)
    txtCodAlumno.Locked = estado1
    txtRecargoXFecha.Locked = estado1
    txtRecargoXMes.Locked = estado1
    txtUltimaFecha.Locked = estado1
    txtNroAsiento.Locked = estado1
    txtEmpresa.Locked = estado1
    txtSucursal.Locked = estado1
    txtMatricula.Locked = estado1
    txtDerechoExamen.Locked = estado1
End Sub

