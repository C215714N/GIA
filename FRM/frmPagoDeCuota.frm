VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmPagoDeCuota 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago de Cuota"
   ClientHeight    =   3030
   ClientLeft      =   13545
   ClientTop       =   2865
   ClientWidth     =   7380
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmPagoDeCuota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   7380
   Begin VB.TextBox txtComision 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Left            =   5640
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc 
      Height          =   330
      Left            =   0
      Top             =   4800
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   2415
      Left            =   120
      TabIndex        =   13
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   21
      RowDividerStyle =   0
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9.75
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
   Begin VB.TextBox txtResta 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Left            =   5640
      TabIndex        =   5
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtTotalPago 
      Alignment       =   1  'Right Justify
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
      Height          =   360
      Left            =   3960
      TabIndex        =   4
      Text            =   "0"
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtNroFactura 
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
      Left            =   3960
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtMonto 
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
      Left            =   5640
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox cmbTipoPago 
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
      ItemData        =   "frmPagoDeCuota.frx":10CA
      Left            =   3960
      List            =   "frmPagoDeCuota.frx":10DA
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin isButtonTest.isButton cmdCobrar 
      Height          =   420
      Left            =   3960
      TabIndex        =   6
      Top             =   2400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   741
      Icon            =   "frmPagoDeCuota.frx":1106
      Style           =   8
      Caption         =   "     Cobrar"
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
      Left            =   5640
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   741
      Icon            =   "frmPagoDeCuota.frx":19E0
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
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comision $"
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
      Left            =   5640
      TabIndex        =   15
      Top             =   120
      Width           =   1230
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Plan de Pago"
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
      Height          =   225
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Height          =   225
      Left            =   3960
      TabIndex        =   12
      Top             =   1560
      Width           =   1350
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resta $"
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
      Height          =   225
      Left            =   5640
      TabIndex        =   11
      Top             =   1560
      Width           =   1350
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Factura"
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
      Height          =   225
      Left            =   3960
      TabIndex        =   10
      Top             =   870
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Monto $"
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
      Height          =   225
      Left            =   5640
      TabIndex        =   9
      Top             =   840
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Pago"
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
      Height          =   225
      Left            =   3960
      TabIndex        =   8
      Top             =   120
      Width           =   1350
   End
End
Attribute VB_Name = "frmPagoDeCuota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Centrar Me
    ContabilidadTemp
    txtResta.Text = frmCobranza.txtAdeuda.Text
    txtNroFactura.Text = ""
    txtMonto.Text = ""
    Adodc.CursorLocation = adUseClient
    Adodc.ConnectionString = DbCon
    Adodc.RecordSource = "SELECT * FROM contabilidadtemp"
    Adodc.Refresh
    Set grilla.DataSource = Adodc
    formatoGrilla
    Dim tipoPago As String
End Sub
Private Sub cmbTipoPago_KeyPress(keyAscii As Integer)
    If keyAscii = 13 And cmbTipoPago.Text = "Comision" Then
        txtComision.Enabled = True
        txtComision.SetFocus
    ElseIf keyAscii = 13 And txtNroFactura.Enabled = True Then
        txtNroFactura.SetFocus
    ElseIf keyAscii = 13 And txtNroFactura.Enabled = False Then
        txtMonto.SetFocus
    Else: txtComision.Enabled = False
    End If
End Sub

Private Sub cmdCobrar_Click()
    If cmbTipoPago.Text = "" Then MsgBox "Debe elegir un Tipo de Pago", vbOKOnly + vbInformation, "Pago de Cuota": cmbTipoPago.SetFocus: Exit Sub
    If txtNroFactura.Text = "" Then MsgBox "Debe agregar un numero de factura", vbOKOnly + vbInformation, "Pago de Cuota": txtNroFactura.SetFocus: Exit Sub
    If txtMonto.Text = "" Then MsgBox "Debe agregar un monto a pagar", vbOKOnly + vbInformation, "Pago de Cuota": txtMonto.SetFocus: Exit Sub

    On Error GoTo continuar
    Cobranza
    With rsCobranza
        .Find "nrocuota=" & Int(frmCobranza.txtNroCuota.Text)
        !DeudaTotal = Val(txtResta.Text)
        !totalcobrado = !totalcobrado + Val(txtTotalPago.Text)
        !fechapago = Date
        !recibo = txtNroFactura.Text
        If Val(txtResta.Text) = 0 Then
            !CuotasDebidas = 0
            !tipodepago = "PAG"
        Else
            !tipodepago = "Par"
        End If
        .UpdateBatch
    End With
    Contabilidad
    With rsContabilidadTemp
        .MoveFirst
        Do Until .EOF
            rsContabilidad.Requery
            rsContabilidad.AddNew
            rsContabilidad!asiento = !asiento
            rsContabilidad!fecha = !fecha
            rsContabilidad!cuenta = !cuenta
            rsContabilidad!Detalle = !Detalle
            rsContabilidad!Debe = !Debe
            rsContabilidad!Haber = !Haber
            rsContabilidad!nrofactura = !nrofactura
            rsContabilidad!NroCuota = !NroCuota
            rsContabilidad!CodAlumno = !CodAlumno
            rsContabilidad.Update
            .Delete
            .Update
            .MoveFirst
        Loop
    End With
    '''si paga cuota futura no lo baja a marcas
    If Month(Format(frmCobranza.txtFechaVto.Text, "dd/mm/yyyy")) > Month(Date) And Year(Format(frmCobranza.txtFechaVto.Text, "mm/dd/yyyy")) >= Year(Date) Then
        GoTo continuar
    End If
    ''' si paga cuota 1 en adelante lo baja a marcas
    If Int(frmCobranza.txtNroCuota.Text) > 0 Then
        Marcar
        With rsMarcar
            .Find "Codalumno=" & Int(frmCobranza.lblCodAlumno.Caption)
            !cobrado = !cobrado + CSng(txtTotalPago.Text)
            If !cobrado >= !deuda Then
                !pago = 1
            End If
        .UpdateBatch
        End With
    ElseIf Int(frmCobranza.txtNroCuota.Text) = 0 And DateDiff("m", Date, frmCobranza.txtFechaVto.Text) < 0 Then
        Marcar
        With rsMarcar
            .Find "Codalumno=" & Int(frmCobranza.lblCodAlumno.Caption)
            !cobrado = !cobrado + CSng(txtTotalPago.Text)
            If txtResta.Text = "0" Then
                !pago = !pago + 1
            End If
        .UpdateBatch
        End With
    End If
continuar:
    planPago
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    With rsContabilidadTemp
        .Requery
        If .BOF Or .EOF Then frmCobranza.Enabled = True: Unload Me: Exit Sub
        .MoveFirst
        Do Until .EOF
            .Delete
            .Update
            .MoveFirst
        Loop
    End With
    cmdCobrar.Enabled = False
    Unload Me
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    With rsContabilidadTemp
        .Requery
        If .BOF Or .EOF Then frmCobranza.Enabled = True: Unload Me: Exit Sub
        .MoveFirst
        Do Until .EOF
            .Delete
            .Update
            .MoveFirst
        Loop
    End With
    planPago
End Sub
Private Sub txtMonto_KeyPress(keyAscii As Integer)
If keyAscii = 13 Then
    revision
    cmdCobrar.Enabled = True
    txtTotalPago.Text = CCur(txtTotalPago.Text) + CCur(txtMonto.Text)
    txtResta.Text = FormatCurrency(txtResta.Text) - FormatCurrency(txtMonto.Text)
    
    If Val(txtResta.Text) = 0 Then
        cmbTipoPago.Enabled = False
    ElseIf Val(txtResta.Text) < 0 Then
        MsgBox "El monto es superior a la deuda", vbOKOnly + vbInformation, "Pago de Cuota":
        txtResta.Text = Val(txtResta.Text) + Val(txtTotalPago.Text)
        txtTotalPago.Text = Val(txtTotalPago.Text) - Val(txtMonto.Text)
        txtMonto.SetFocus
        Exit Sub
    End If
    
    tipoPago = cmbTipoPago.Text
    If tipoPago = "Comision" Then
        bajarAsiento tipoPago
        tipoPago = "Salida"
        bajarAsiento tipoPago
        tipoPago = "Efectivo"
    End If
    bajarAsiento tipoPago

    txtNroFactura.Enabled = False
    If Val(txtResta.Text) = 0 Then
        cmdCobrar.SetFocus
    Else
        cmbTipoPago.SetFocus
    End If
    txtMonto.Text = 0
    formatoGrilla
End If
End Sub
Sub revision()
    If cmbTipoPago.Text = "" Then MsgBox "Debe elegir un Tipo de Pago", vbOKOnly + vbInformation, "Pago de Cuota": cmbTipoPago.SetFocus
    If txtNroFactura.Text = "" Then MsgBox "Debe agregar un numero de factura", vbOKOnly + vbInformation, "Pago de Cuota": txtNroFactura.SetFocus
    If txtMonto.Text = "" Then MsgBox "Debe agregar un monto a pagar", vbOKOnly + vbInformation, "Pago de Cuota": txtMonto.SetFocus
End Sub
Sub planPago()
    frmCobranza.Enabled = True
    frmCobranza.Adodc.Refresh
    rsPlanDePago.Requery
    frmCobranza.formatoGrilla
    frmCobranza.txtAdeuda.Text = txtResta.Text
    frmCobranza.cmdPagar.Enabled = False
    cmdCobrar.Enabled = False
End Sub

Sub formatoGrilla()
    grilla.Columns(0).Width = 0
    grilla.Columns(1).Width = 0
    grilla.Columns(4).Width = 0
    grilla.Columns(2).Width = 800
    grilla.Columns(6).Width = 0

End Sub
Private Sub txtNroFactura_KeyPress(keyAscii As Integer)
    continue keyAscii
End Sub
Sub bajarAsiento(tipoPago)
    With rsContabilidadTemp
        .Requery
        .AddNew
        !nrofactura = txtNroFactura.Text
        ''' Egreso por comision
        If tipoPago = "Salida" Then
            !Debe = Null
            !Haber = txtComision.Text
        Else
            !Debe = txtMonto.Text
            !Haber = Null
        End If
        !fecha = Date
        !asiento = Null
        !CodAlumno = Val(frmCobranza.lblCodAlumno.Caption)
        !NroCuota = Val(frmCobranza.txtNroCuota.Text)
        ''' Seleccion de cuenta
        If tipoPago = "Efectivo" Or tipoPago = "Salida" Then
            !cuenta = "CAJA ADMINISTRACION"
        ElseIf tipoPago = "Comision" Then
            !cuenta = "HONORARIOS ASESORES"
        ElseIf tipoPago = "Tarjeta" Then
            !cuenta = "DEBITO TARJETA CREDITO"
        Else: !cuenta = "Descuento"
        End If
        
        If tipoPago = "Comision" Or tipoPago = "Salida" Then
            !Detalle = "COMISION CUOTA " & frmCobranza.txtNroCuota.Text & " ALUMNO " & frmCobranza.lblCodAlumno.Caption
        Else:
            !Detalle = "ALUMNO " & frmCobranza.lblCodAlumno.Caption & " CUOTA " & frmCobranza.txtNroCuota.Text
        End If
        .Update
    End With
    Adodc.Refresh
End Sub

