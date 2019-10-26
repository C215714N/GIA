VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConsultarOrdenes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar Órdenes de Trabajo"
   ClientHeight    =   5895
   ClientLeft      =   3120
   ClientTop       =   1695
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10575
   Begin VB.CommandButton cmdEstado 
      Caption         =   "Cambiar Estado"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   25
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdAgregarGestion 
      Caption         =   "Agregar Gestión"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   22
      Top             =   1560
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid grilla2 
      Height          =   3375
      Left            =   5040
      TabIndex        =   15
      Top             =   2400
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin VB.Frame Frame2 
      Caption         =   "Ingreso"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5040
      TabIndex        =   12
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtPersonal 
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
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtFecha 
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
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtFalla 
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
         Height          =   975
         Left            =   1920
         TabIndex        =   13
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label lblestado 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Shape Shape1 
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label Label8 
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   "Personal"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Falla / Diagnóstico"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   14
         Top             =   240
         Width           =   2415
      End
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   3375
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      Caption         =   "Consulta de Órdenes de Trabajo"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2760
         TabIndex        =   28
         Top             =   1715
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkEntregados 
         Caption         =   "Entregados"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         Top             =   1715
         Width           =   1215
      End
      Begin VB.CheckBox chkPendientes 
         Caption         =   "Pendientes"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1715
         Width           =   1215
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   10
         Top             =   1560
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo dtcCliente 
         Height          =   360
         Left            =   1680
         TabIndex        =   9
         Top             =   1125
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.TextBox txtNroOrden 
         Alignment       =   1  'Right Justify
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
         Left            =   1680
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   300
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
         Format          =   220463105
         CurrentDate     =   42206
      End
      Begin VB.OptionButton optBuscar 
         Caption         =   "Por Cliente"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.OptionButton optBuscar 
         Caption         =   "Por Nº Orden"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   780
         Width           =   1695
      End
      Begin VB.OptionButton optBuscar 
         Caption         =   "Por Fecha"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   375
         Left            =   3360
         TabIndex        =   5
         Top             =   300
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
         Format          =   220463105
         CurrentDate     =   42206
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   7
         Top             =   360
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "De"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   225
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Gestión de Ordenes"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   21
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Label Label6 
      Caption         =   "Órdenes de Trabajo"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2160
      Width           =   2535
   End
End
Attribute VB_Name = "frmConsultarOrdenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkEntregados_Click()
    If chkEntregados.Value = 1 And chkPendientes.Value = 1 Then
        chkTodos.Value = 1
    Else
        chkTodos.Value = 0
    End If
End Sub

Private Sub chkPendientes_Click()
    If chkEntregados.Value = 1 And chkPendientes.Value = 1 Then
        chkTodos.Value = 1
    Else
        chkTodos.Value = 0
    End If

End Sub

Private Sub chkTodos_Click()
    If chkTodos.Value = 1 Then
        chkPendientes.Value = 1
        chkEntregados.Value = 1
    Else
        chkPendientes.Value = 0
        chkEntregados.Value = 0
    End If
End Sub

Private Sub cmdAgregarGestion_Click()
    frmAgregaGestion.Show
    frmAgregaGestion.lblNroOrden.Caption = grilla.Columns(0).Text
    Me.Enabled = False
End Sub

Private Sub cmdBuscar_Click()
    
    If chkTodos.Value = 0 And chkPendientes.Value = 0 And chkEntregados.Value = 0 Then MsgBox "Elija el tipo de resultado", vbCritical, "Cóndor": chkTodos.SetFocus: Exit Sub
        
    If optBuscar(0).Value = True Then
        Dim fecha1 As Date
        Dim fecha2 As Date
        fecha1 = Format(dtpDesde.Value, "mm/dd/yyyy")
        fecha2 = Format(dtpHasta.Value, "mm/dd/yyyy")
        
        If chkTodos.Value = 1 Then
            With rsconsultarordenes
                If .State = 1 Then .Close
                .Open "select NroOrden, Cliente,Equipo,Problema,FechaRecibido as Recepción,Estado,FechaEntregado as Entrega from ordenesdetrabajo as o,clientes as c where c.codcliente=o.codcliente and fecharecibido>=#" & fecha1 & "# and fecharecibido <=#" & fecha2 & "#", Cn, adOpenDynamic, adLockPessimistic
            End With
        ElseIf chkPendientes.Value = 1 Then
            With rsconsultarordenes
                If .State = 1 Then .Close
                .Open "select NroOrden, Cliente,Equipo,Problema,FechaRecibido as Recepción,Estado,FechaEntregado as Entrega from ordenesdetrabajo as o,clientes as c where c.codcliente=o.codcliente and fecharecibido>=#" & fecha1 & "# and fecharecibido <=#" & fecha2 & "# and estado<>'ENTREGADO'", Cn, adOpenDynamic, adLockPessimistic
            End With
        ElseIf chkEntregados.Value = 1 Then
            With rsconsultarordenes
                If .State = 1 Then .Close
                .Open "select NroOrden, Cliente,Equipo,Problema,FechaRecibido as Recepción,Estado,FechaEntregado as Entrega from ordenesdetrabajo as o,clientes as c where c.codcliente=o.codcliente and fecharecibido>=#" & fecha1 & "# and fecharecibido <=#" & fecha2 & "# and estado='ENTREGADO'", Cn, adOpenDynamic, adLockPessimistic
            End With
        End If
        
    ElseIf optBuscar(1).Value = True Then
            With rsconsultarordenes
                If .State = 1 Then .Close
                .Open "select NroOrden, Cliente,Equipo,Problema,FechaRecibido as Recepción,Estado,FechaEntregado as Entrega from ordenesdetrabajo as o,clientes as c where c.codcliente=o.codcliente and nroorden=" & txtNroOrden.Text, Cn, adOpenDynamic, adLockPessimistic
            End With
    Else
        With rsBuscarClientes
            If .State = 1 Then .Close
            .Open "select codcliente from clientes where cliente='" & dtcCliente.Text & "'", Cn, adOpenDynamic, adLockPessimistic
        End With
          
        If chkTodos.Value = 1 Then
            With rsconsultarordenes
                If .State = 1 Then .Close
                .Open "select NroOrden, Cliente,Equipo,Problema,FechaRecibido as Recepción,Estado,FechaEntregado as Entrega from ordenesdetrabajo as o,clientes as c where c.codcliente=o.codcliente and o.codcliente=" & rsBuscarClientes!codcliente, Cn, adOpenDynamic, adLockPessimistic
            End With
        ElseIf chkEntregados.Value = 1 Then
            With rsconsultarordenes
                If .State = 1 Then .Close
                .Open "select NroOrden, Cliente,Equipo,Problema,FechaRecibido as Recepción,Estado,FechaEntregado as Entrega from ordenesdetrabajo as o,clientes as c where estado='ENTREGADO' and c.codcliente=o.codcliente and o.codcliente=" & rsBuscarClientes!codcliente, Cn, adOpenDynamic, adLockPessimistic
            End With
        ElseIf chkPendientes.Value = 1 Then
            With rsconsultarordenes
                If .State = 1 Then .Close
                .Open "select NroOrden, Cliente,Equipo,Problema,FechaRecibido as Recepción,Estado,FechaEntregado as Entrega from ordenesdetrabajo as o,clientes as c where estado<>'ENTREGADO' and c.codcliente=o.codcliente and o.codcliente=" & rsBuscarClientes!codcliente, Cn, adOpenDynamic, adLockPessimistic
            End With
        End If
        
    End If
    
    Set grilla.DataSource = rsconsultarordenes
    
    grilla.Columns(3).Width = 0
    grilla.Columns(5).Width = 0
    grilla.Columns(6).Width = 0
    lblestado.Caption = ""
    Shape1.FillColor = vbBlack
End Sub

Private Sub cmdEstado_Click()
    frmCambioEstado.Show
    frmCambioEstado.lblNroOrden.Caption = grilla.Columns(0).Text
    Me.Enabled = False
End Sub

Private Sub Form_Load()
    Centrar Me
    dtpDesde.Value = Date
    dtpHasta.Value = Date
End Sub

Private Sub grilla_Click()
    If rsconsultarordenes.RecordCount > 0 Then
        txtFalla.Text = grilla.Columns(3).Text
        txtFecha.Text = grilla.Columns(4).Text
        
        With rsGestionDeOrdenes
            If .State = 1 Then .Close
            .Open "select Fecha,Gestion,Personal from gestiondeordenes where nroordentrabajo=" & grilla.Columns(0).Text & " order by fecha,id", Cn, adOpenDynamic, adLockPessimistic
            Set grilla2.DataSource = rsGestionDeOrdenes
            grilla2.Columns(0).Width = 1000
            grilla2.Columns(1).Width = 3000
            grilla2.Columns(2).Width = 1000
            
            txtPersonal.Text = grilla2.Columns(2).Text
        End With
        
        If grilla.Columns(5).Text = "A REVISAR" Then Shape1.FillColor = vbRed
        If grilla.Columns(5).Text = "PRESUPUESTADO" Then Shape1.FillColor = vbYellow
        If grilla.Columns(5).Text = "REALIZADO" Then Shape1.FillColor = vbGreen
        If grilla.Columns(5).Text = "ENTREGADO" Then Shape1.FillColor = vbBlue
        lblestado.Caption = grilla.Columns(5).Text
    End If
End Sub

Private Sub optBuscar_Click(Index As Integer)
    Select Case Index
        Case 0
            dtpDesde.Enabled = True
            dtpHasta.Enabled = True
            txtNroOrden.Enabled = False
            dtcCliente.Enabled = False
            dtpDesde.SetFocus
        Case 1
            dtpDesde.Enabled = False
            dtpHasta.Enabled = False
            txtNroOrden.Enabled = True
            dtcCliente.Enabled = False
            txtNroOrden.SetFocus
        Case 2
            dtpDesde.Enabled = False
            dtpHasta.Enabled = False
            txtNroOrden.Enabled = False
            dtcCliente.Enabled = True
            dtcCliente.SetFocus
    End Select
End Sub
