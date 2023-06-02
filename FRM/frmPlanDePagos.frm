VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmPlanDePagos 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plan de Pago"
   ClientHeight    =   5175
   ClientLeft      =   6855
   ClientTop       =   2160
   ClientWidth     =   4485
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmPlanDePagos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5129.738
   ScaleMode       =   0  'User
   ScaleWidth      =   4485
   Begin VB.TextBox txtNroCuota 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   14
      Text            =   "0"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtCuotaDos 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   11
      Text            =   "1"
      Top             =   1275
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid grilla 
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   4250
      _ExtentX        =   7488
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtTotalCuotas 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3480
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtDeuda 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin MSComCtl2.DTPicker DTPFechaVto 
      Height          =   360
      Left            =   1080
      TabIndex        =   0
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
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
      Format          =   130875393
      CurrentDate     =   41323
   End
   Begin MSComCtl2.DTPicker dtpVtoDos 
      Height          =   360
      Left            =   1080
      TabIndex        =   10
      Top             =   1275
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   635
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
      Format          =   130875393
      CurrentDate     =   41323
   End
   Begin isButtonTest.isButton cmdCrearPlan 
      Height          =   420
      Left            =   1560
      TabIndex        =   12
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmPlanDePagos.frx":10CA
      Style           =   8
      Caption         =   "     Asignar"
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
   Begin isButtonTest.isButton cmdSalir 
      Height          =   420
      Left            =   3000
      TabIndex        =   13
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmPlanDePagos.frx":19A4
      Style           =   8
      Caption         =   "     Volver"
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
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cuotas"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   240
      Left            =   3480
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblNya 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   7
      Top             =   150
      Width           =   3255
   End
   Begin VB.Label lblCodAlumno 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   6
      Top             =   150
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deuda $"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   240
      Left            =   2520
      TabIndex        =   5
      Top             =   600
      Width           =   720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimiento"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cuota"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
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
      Top             =   600
      Width           =   630
   End
End
Attribute VB_Name = "frmPlanDePagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCrearPlan_Click()
    tabla
    CrearPlan
    cmdCrearPlan.Enabled = False
    txtTotalCuotas.Enabled = False
    
    '''Agrega alumno a alumnos del mes
    With rsAlumnosDelMes
        If .State = 1 Then .Close
        .Open "SELECT * FROM alumnosdelmes", Cn, adOpenDynamic, adLockPessimistic
        .Requery
        .AddNew
        !CodAlumno = lblCodAlumno.Caption
        !totalcuotas = Int(txtTotalCuotas.Text)
        .Update
    End With

'    Marcar
    If Int(txtDeuda.Text) > "1" Then
        frmComisionCuota.Show
        frmComisionCuota.lblTotalCurso.Caption = Format(Int(txtDeuda.Text) * Int(txtTotalCuotas.Text), "currency")
        frmComisionCuota.lblTotalCuota1.Caption = Format(txtDeuda.Text, "currency")
        Me.Enabled = False
    End If

End Sub

Private Sub cmdSalir_Click()
    frmVerificaciones.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()
    Centrar Me
    PlanDePago
    lblCodAlumno.Caption = frmVerificaciones.lblCodAlumno.Caption
    lblNya.Caption = frmVerificaciones.txtNya.Text
    txtTotalCuotas.Text = frmVerificaciones.txtTotalCuotas.Text
    txtDeuda.Text = Val(frmVerificaciones.txtTotalCurso.Text) / Val(txtTotalCuotas.Text)
    DTPFechaVto.Value = Date
    dtpVtoDos.Value = Date
    cmdCrearPlan.Enabled = True
    txtTotalCuotas.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmVerificaciones.Enabled = True
    Unload Me
End Sub

Sub tabla()
    Dim NroCuota As Integer
    NroCuota = 0
    grilla.Rows = 2
    grilla.Col = 0
    grilla.Row = 0
    grilla.Text = "N° Cuota"
    grilla.Col = 1
    grilla.Text = "Fecha Vto"
    grilla.Col = 2
    grilla.Text = "Deuda"
    grilla.Col = 0
    grilla.Row = 1
    grilla.ColWidth(1) = 2000
End Sub

Public Sub revisar()
    If DTPFechaVto.Value < dtpVtoDos.Value Then MsgBox ("la cuota siguiente no puede ser menor al primer vencimiento")
End Sub

Sub CrearPlan()
    With rsPlanDePago
        Dim FechaLimite As Integer
        FechaLimite = 25
        
        Do Until NroCuota > Int(txtTotalCuotas.Text)
            .Requery
            .AddNew
            !CodAlumno = lblCodAlumno.Caption
            !NyA = lblNya.Caption
            !NroCuota = NroCuota
        ''' Vencimiento Cuotas y calculo fecha limite
            If NroCuota >= 1 Then
                !fechavto = dtpVtoDos.Value
                DTPFechaVto.Value = dtpVtoDos.Value
                DTPFechaVto.Day = FechaLimite
                !FechaLimite = DTPFechaVto.Value
            Else: !fechavto = DTPFechaVto.Value
            End If
            
            !deuda = txtDeuda.Text
            !totalcobrado = 0
        ''' Alumnos Becados
            If Int(txtTotalCuotas.Text) = 1 And Int(txtDeuda.Text) = 1 Then
                !DeudaTotal = 0
                !CuotasDebidas = 0
        ''' Plan de Pago regular
            Else
                !DeudaTotal = txtDeuda.Text
                !CuotasDebidas = 1
            End If
            .Update
            grilla.Text = NroCuota
            grilla.Col = 1
            
            If NroCuota >= 1 Then
                grilla.Text = dtpVtoDos.Value
            Else
                grilla.Text = DTPFechaVto.Value
            End If
            
            grilla.Col = 2
            grilla.Text = txtDeuda.Text
            grilla.Rows = grilla.Rows + 1
            grilla.Col = 0
            grilla.Row = grilla.Row + 1
            NroCuota = NroCuota + 1
            
        ''' Cambio de Año
            If NroCuota > txtCuotaDos.Text Then
                If dtpVtoDos.Month = 12 Then
                    dtpVtoDos.Month = 1
                    dtpVtoDos.Year = dtpVtoDos.Year + 1
                Else: dtpVtoDos.Month = dtpVtoDos.Month + 1
                End If
            End If
        Loop
    End With
End Sub
