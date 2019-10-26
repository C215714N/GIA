VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmAuditoria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auditoría"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3225
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAuditoria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmAuditoria.frx":324A
   ScaleHeight     =   5265
   ScaleWidth      =   3225
   Begin VB.Frame Frame2 
      BackColor       =   &H00884400&
      Caption         =   "Restantes"
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
      Height          =   3615
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3015
      Begin VB.ListBox List1 
         Height          =   2460
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin VB.ListBox List2 
         Height          =   2460
         Left            =   1560
         TabIndex        =   11
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblMarcas 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblPlanDePago 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Marcas"
         ForeColor       =   &H8000000F&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "P.Pago"
         ForeColor       =   &H8000000F&
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   1005
      End
   End
   Begin isButtonTest.isButton cmdAgregar 
      Height          =   420
      Left            =   1680
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmAuditoria.frx":AC67
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00884400&
      Caption         =   "Actualzar"
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
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.TextBox txtCodigo 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin isButtonTest.isButton btnActualizar 
         Height          =   420
         Left            =   1560
         TabIndex        =   5
         Top             =   400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmAuditoria.frx":B541
         Style           =   8
         Caption         =   "       Actualizar"
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde Código"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComCtl2.DTPicker dtpFechaFutura 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   4680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      CurrentDate     =   42492
   End
End
Attribute VB_Name = "frmAuditoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdActualizar_Click()
        Dim fechafutura As Date
        fechafutura = Format(dtpFechaFutura.Value, "mm/dd/yyyy")
        
           '''actualiza la info de situacion de cartera en marcas
            ''' consulta cuotas debidas a la fecha
            With rsPlanDePago
                If .State = 1 Then .Close
                .Open "SELECT p.codalumno,min(p.nrocuota) as Cuota, sum(p.Deudatotal) as Deuda,sum(p.CuotasDebidas) as CuotasDebidas,  DateDiff('m',Min(p.fechavto),#" & fechafutura & "#) AS Meses, Max(p.NroCuota) AS MaxCuota,max(V.Cuotas) AS UltimaCuota FROM plandepago as p, verificaciones as v WHERE v.codalumno=p.codalumno and p.cuotasdebidas > 0 and p.fechavto<#" & fechafutura & "# and p.codalumno>=" & Int(txtCodigo.Text) & " group by p.codalumno ORDER BY p.codalumno", Cn, adOpenDynamic, adLockPessimistic
                .MoveFirst
            End With
            
            With rsMarcar
                If .State = 1 Then .Close
                .Open "SELECT * FROM marcas WHERE codalumno>=" & txtCodigo.Text & " ORDER BY codalumno", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                '''actualiza el alumno
                Do Until .EOF
                    If rsPlanDePago.EOF Then
                        !cuota = 0
                        !deuda = 0
                        !cantidadcuotas = 0
                        !cobrado = 0
                        !pago = 0
                        .Update
                        .MoveNext
                    ElseIf !CodAlumno = rsPlanDePago!CodAlumno Then
                        !cuota = rsPlanDePago!cuota
                        !deuda = rsPlanDePago!deuda
                        If rsPlanDePago!maxcuota = rsPlanDePago!ultimacuota And rsPlanDePago!Meses > rsPlanDePago!CuotasDebidas Then
                            !cantidadcuotas = rsPlanDePago!Meses
                        Else
                            !cantidadcuotas = rsPlanDePago!CuotasDebidas
                        End If
                        !cobrado = 0
                        !pago = 0
                        .UpdateBatch
                        .MoveNext
                        rsPlanDePago.MoveNext
                    Else
                        !cuota = 0
                        !deuda = 0
                        !cantidadcuotas = 0
                        !cobrado = 0
                        !pago = 0
                        .Update
                        .MoveNext
                    End If
                Loop
                
                MsgBox "La Base de Datos fue actualizada exitosamente", , "Auditoría"
            End With


End Sub

Private Sub cmdAgregar_Click()
    If MsgBox("¿Está seguro que desea agregar estos códigos a la tabla MARCAS?", vbQuestion + vbYesNo, "GIA") = vbYes Then
            With rsActualizarMarcas
                If .State = 1 Then .Close
                .Open "SELECT * FROM marcas", Cn, adOpenDynamic, adLockPessimistic
        
                Do Until List1.ListCount = 0
            
                    .Requery
                    .AddNew
                    !CodAlumno = Int(List1.List(0))
                    !cuota = 0
                    !deuda = 0
                    !pago = 0
                    !cantidadcuotas = 0
                    !cobrado = 0
                    !LPA = ""
                    .Update
                    List1.RemoveItem (0)
                
                Loop
            End With
    End If
End Sub

Private Sub Form_Load()
    Centrar Me
    Dim contador As Integer
    contador = 1
    With rsMarcas
        If .State = 1 Then .Close
        .Open "SELECT codalumno FROM marcas ORDER BY codalumno", Cn, adOpenDynamic, adLockPessimistic
        .MoveFirst
        Do Until .EOF
            If contador = !CodAlumno Then
                .MoveNext
                contador = contador + 1
            Else
                List1.AddItem (contador)
                contador = contador + 1
            End If
        Loop
    End With
    lblMarcas.Caption = List1.ListCount
    contador = 1
    With rsPlanDePago
        If .State = 1 Then .Close
        .Open "SELECT distinct codalumno FROM plandepago WHERE codalumno>0 ORDER BY codalumno", Cn, adOpenDynamic, adLockPessimistic
        .MoveFirst
        Do Until .EOF
            If contador = !CodAlumno Then
                .MoveNext
                contador = contador + 1
            Else
                List2.AddItem (contador)
                contador = contador + 1
            End If
        Loop
    End With
    lblPlanDePago.Caption = List2.ListCount

    dtpFechaFutura.Value = Date
    If dtpFechaFutura.Month = 12 Then
        dtpFechaFutura.Month = 1
        dtpFechaFutura.Year = dtpFechaFutura.Year + 1
    Else
        dtpFechaFutura.Day = 25
        dtpFechaFutura.Month = dtpFechaFutura.Month + 1
    End If
    
End Sub

