VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmConsultarCheques 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar Cheques"
   ClientHeight    =   4680
   ClientLeft      =   3360
   ClientTop       =   1620
   ClientWidth     =   8925
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmConsultarCheques.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8925
   Begin MSFlexGridLib.MSFlexGrid grilla 
      Height          =   2805
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   4948
      _Version        =   393216
      Cols            =   6
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00662200&
      Caption         =   "Consultar Cheques"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5775
      Begin VB.Frame Frame2 
         BackColor       =   &H00662200&
         Caption         =   "Filtrar por"
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
         Height          =   1215
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1215
         Begin VB.OptionButton optBuscar 
            BackColor       =   &H00662200&
            Caption         =   "Firma"
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
            Height          =   300
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   540
            Width           =   900
         End
         Begin VB.OptionButton optBuscar 
            BackColor       =   &H00662200&
            Caption         =   "Fecha"
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
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   900
         End
         Begin VB.OptionButton optBuscar 
            BackColor       =   &H00662200&
            Caption         =   "Ambos"
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
            Height          =   300
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   1020
         End
      End
      Begin VB.ComboBox cmbFirma 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "frmConsultarCheques.frx":10CA
         Left            =   1440
         List            =   "frmConsultarCheques.frx":10D1
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   375
         Left            =   2880
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   131203073
         CurrentDate     =   41782
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   375
         Left            =   1440
         TabIndex        =   0
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   131203073
         CurrentDate     =   41782
      End
      Begin isButtonTest.isButton cmdBuscar 
         Height          =   420
         Left            =   4320
         TabIndex        =   15
         Top             =   405
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmConsultarCheques.frx":10DD
         Style           =   8
         Caption         =   "     Buscar"
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheques"
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
         Height          =   360
         Left            =   4320
         TabIndex        =   19
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblCantidad 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   4335
         TabIndex        =   18
         Top             =   1095
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
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
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
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
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin isButtonTest.isButton cmdEliminar 
      Height          =   420
      Left            =   7440
      TabIndex        =   16
      Top             =   420
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmConsultarCheques.frx":19B7
      Style           =   8
      Caption         =   "     Eliminar"
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
   Begin isButtonTest.isButton cmdDepositar 
      Height          =   420
      Left            =   7440
      TabIndex        =   17
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmConsultarCheques.frx":2291
      Style           =   8
      Caption         =   "     Deposito"
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
   Begin VB.Label lblTotalADepositar 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   6000
      TabIndex        =   8
      Top             =   435
      Width           =   1335
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   6000
      TabIndex        =   6
      Top             =   1095
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "A Depositar"
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
      Height          =   360
      Left            =   6000
      TabIndex        =   9
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Total $"
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
      Height          =   360
      Left            =   6015
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmConsultarCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBuscar_Click()
    Busqueda
End Sub

Private Sub cmddepositar_Click()
    If MsgBox("�Esta Seguro que se ha depositado este cheque?", vbQuestion + vbYesNo, "Cheques") = vbYes Then
        Grilla.Col = 2
        With rsCheques
            .Close
            .Open "SELECT * FROM cheques WHERE numerocheque='" & Grilla.Text & "'", Cn, adOpenDynamic, adLockPessimistic
            !estado = "DEPOSITADO"
            .UpdateBatch
        End With
        Busqueda
    End If
End Sub

Private Sub cmdEliminar_Click()
    Grilla.Col = 2
    If MsgBox("�Esta seguro que desea eliminar el cheque N� " & Grilla.Text & "?", vbQuestion + vbYesNo, "Consultar Cheques") = vbYes Then
        With rsCheques
            .Close
            .Open "SELECT * FROM cheques WHERE numerocheque='" & Grilla.Text & "'", Cn, adOpenDynamic, adLockPessimistic
            .Requery
            .MoveFirst
            .Delete
            .Close
        End With
        
        Busqueda
        cmdEliminar.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    Centrar Me
    dtpDesde.Value = Date
    dtpHasta.Value = Date
End Sub

Private Sub grilla_Click()
    cmdEliminar.Enabled = True
    cmdDepositar.Enabled = True
End Sub

Private Sub optBuscar_Click(Index As Integer)
    
    Select Case Index
        Case 0
            dtpDesde.Enabled = True
            dtpHasta.Enabled = True
            cmbFirma.Enabled = False
        Case 1
            dtpDesde.Enabled = False
            dtpHasta.Enabled = False
            cmbFirma.Enabled = True
        Case 2
            dtpDesde.Enabled = True
            dtpHasta.Enabled = True
            cmbFirma.Enabled = True
    End Select
End Sub

Private Sub Busqueda()
    '''declara variables para consultar las fechas
    Dim fecha1 As Date
    Dim fecha2 As Date
    
    ''asigna las fechas de busqueda a las variables con formato para sql
    fecha1 = Format(dtpDesde.Value, "mm/dd/yyyy")
    fecha2 = Format(dtpHasta.Value, "mm/dd/yyyy")
    
    If optBuscar(0).Value = True Then
        If dtpHasta.Value < dtpDesde.Value Then MsgBox "Fechas invalidas", vbCritical, "Consultar Cheques": dtpDesde.SetFocus: Exit Sub
        
    '''Consulta por FECHA
        With rsCheques
            If .State = 1 Then .Close
            
            .Open "SELECT count(*) FROM cheques WHERE fecha>=#" & fecha1 & "# and fecha <=#" & fecha2 & "#", Cn, adOpenDynamic, adLockPessimistic
            lblCantidad.Caption = !expr1000
            .Close
            
            .Open "SELECT sum(monto) FROM cheques WHERE fecha>=#" & fecha1 & "# and fecha <=#" & fecha2 & "#", Cn, adOpenDynamic, adLockPessimistic
            lblTotal.Caption = Format(!expr1000, "currency")
            .Close
            
            .Open "SELECT sum(monto) FROM cheques WHERE fecha>=#" & fecha1 & "# and fecha <=#" & fecha2 & "# and estado<>'DEPOSITADO'", Cn, adOpenDynamic, adLockPessimistic
            lblTotalADepositar.Caption = Format(!expr1000, "currency")
            .Close
            
            .Open "SELECT Fecha,destinatario as Beneficiario, NumeroCheque,Monto as Importe,Firma,Estado FROM cheques WHERE fecha>=#" & fecha1 & "# and fecha <=#" & fecha2 & "# ORDER BY fecha", Cn, adOpenDynamic, adLockPessimistic
        End With
        
    ElseIf optBuscar(1).Value = True Then
        If cmbFirma.Text = "" Then MsgBox "Seleccione la firma a consultar", vbCritical, "Consultar Cheques": cmbFirma.SetFocus: Exit Sub
    
    '''Consulta por FIRMA
        With rsCheques
            If .State = 1 Then .Close
            .Open "SELECT count(*) FROM cheques WHERE firma='" & cmbFirma.Text & "'", Cn, adOpenDynamic, adLockPessimistic
            lblCantidad.Caption = !expr1000
            .Close
            
            .Open "SELECT sum(monto) FROM cheques WHERE firma='" & cmbFirma.Text & "'", Cn, adOpenDynamic, adLockPessimistic
            lblTotal.Caption = Format(!expr1000, "currency")
            .Close
            
            .Open "SELECT sum(monto) FROM cheques WHERE firma='" & cmbFirma.Text & "' and estado<>'DEPOSITADO'", Cn, adOpenDynamic, adLockPessimistic
            lblTotalADepositar.Caption = Format(!expr1000, "currency")
            .Close
            
            .Open "SELECT Fecha,destinatario as Beneficiario, NumeroCheque,Monto as Importe,Firma,Estado FROM cheques WHERE firma='" & cmbFirma.Text & "' ORDER BY fecha", Cn, adOpenDynamic, adLockPessimistic
        
        End With
    
    ElseIf optBuscar(2).Value = True Then
        If dtpHasta.Value < dtpDesde.Value Then MsgBox "Fechas invalidas", vbCritical, "Consultar Cheques": dtpDesde.SetFocus: Exit Sub
        If cmbFirma.Text = "" Then MsgBox "Seleccione la firma a consultar", vbCritical, "Consultar Cheques": cmbFirma.SetFocus: Exit Sub
    
    '''Consulta por FECHA & por FIRMA
        With rsCheques
            If .State = 1 Then .Close
            .Open "SELECT count(*) FROM cheques WHERE fecha>=#" & fecha1 & "# and fecha <=#" & fecha2 & "# and firma='" & cmbFirma.Text & "'", Cn, adOpenDynamic, adLockPessimistic
            lblCantidad.Caption = !expr1000
            .Close
            
            .Open "SELECT sum(monto) FROM cheques WHERE fecha>=#" & fecha1 & "# and fecha <=#" & fecha2 & "# and firma='" & cmbFirma.Text & "'", Cn, adOpenDynamic, adLockPessimistic
            lblTotal.Caption = Format(!expr1000, "currency")
            .Close
            
            .Open "SELECT sum(monto) FROM cheques WHERE fecha>=#" & fecha1 & "# and fecha <=#" & fecha2 & "# and firma='" & cmbFirma.Text & "' and estado<>'DEPOSITADO'", Cn, adOpenDynamic, adLockPessimistic
            lblTotalADepositar.Caption = Format(!expr1000, "currency")
            .Close
            
            .Open "SELECT Fecha,destinatario as Beneficiario, NumeroCheque,Monto as Importe,Firma,Estado FROM cheques WHERE fecha>=#" & fecha1 & "# and fecha <=#" & fecha2 & "# and firma='" & cmbFirma.Text & "' ORDER BY fecha", Cn, adOpenDynamic, adLockPessimistic
        End With

    Else
        MsgBox "Elija parametros de busqueda", vbCritical, "Consultar cheques"
    End If
    
    Grilla.Clear
    If rsCheques.BOF Or rsCheques.EOF Then Exit Sub
        Grilla.Rows = Int(lblCantidad.Caption) + 2
        Grilla.Col = 0
        Grilla.Row = 0
        Grilla.Text = "Fecha"
        Grilla.Col = 1
        Grilla.Text = "Beneficiario"
        Grilla.Col = 2
        Grilla.Text = "N� Cheque"
        Grilla.Col = 3
        Grilla.Text = "Importe"
        Grilla.Col = 4
        Grilla.Text = "Firma"
        Grilla.Col = 5
        Grilla.Text = "Estado"
        Grilla.Col = 0
        Grilla.Row = Grilla.Row + 1
    
    rsCheques.MoveFirst
    Do Until rsCheques.EOF
        Grilla.Text = rsCheques!fecha
        Grilla.Col = 1
        Grilla.Text = rsCheques!beneficiario
        Grilla.Col = 2
        Grilla.Text = rsCheques!numerocheque
        Grilla.Col = 3
        Grilla.Text = rsCheques!importe
        Grilla.Col = 4
        Grilla.Text = rsCheques!firma
        Grilla.Col = 5
        Grilla.Text = rsCheques!estado
        If Grilla.Text = "DEPOSITADO" Then
            Grilla.CellForeColor = vbGreen
            Grilla.CellFontBold = True
        Else
            Grilla.CellForeColor = vbRed
            Grilla.CellFontBold = True
        End If
        Grilla.Col = 0
        Grilla.Row = Grilla.Row + 1
        rsCheques.MoveNext
    Loop
    formatoGrilla
End Sub

Sub formatoGrilla()
    Dim w As Integer
    For N = 0 To 5 Step 1
        If N = 1 Then
            w = 2000
        ElseIf N = 0 Or N = 5 Then
            w = 1150
        Else:
            w = 900
        End If
        Grilla.ColWidth(N) = w
    Next
End Sub
