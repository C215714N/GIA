VERSION 5.00
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmComisionCuota 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comisión de Primera Cuota"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5475
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComisionCuota.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmComisionCuota.frx":324A
   ScaleHeight     =   3060
   ScaleWidth      =   5475
   Begin VB.Frame Frame6 
      BackColor       =   &H00662200&
      Caption         =   "Totales"
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
      Height          =   2895
      Left            =   3720
      TabIndex        =   15
      Top             =   0
      Width           =   1600
      Begin isButtonTest.isButton cmdAceptar 
         Height          =   420
         Left            =   120
         TabIndex        =   4
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmComisionCuota.frx":AC67
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
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Total $"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Comisiones"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label lblTotalComisiones 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   120
         TabIndex        =   18
         Top             =   1150
         Width           =   1335
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Curso"
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   250
         Width           =   975
      End
      Begin VB.Label lblTotalCurso 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   360
         Left            =   120
         TabIndex        =   16
         Top             =   500
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00884400&
      Caption         =   "Comisiones"
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
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   3555
      Begin VB.Frame Frame4 
         BackColor       =   &H00884400&
         Caption         =   "N° Factura"
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
         Height          =   855
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   1600
         Begin VB.TextBox txtNroFactura 
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00884400&
         Caption         =   "Pago 1° Cuota"
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
         Height          =   1695
         Left            =   1800
         TabIndex        =   12
         Top             =   1080
         Width           =   1600
         Begin VB.TextBox txtPagoParcial 
            Alignment       =   1  'Right Justify
            Height          =   360
            Left            =   120
            TabIndex        =   3
            Text            =   "0"
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Cuota $"
            ForeColor       =   &H8000000F&
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lblTotalCuota1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   360
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Parcial Cuota"
            ForeColor       =   &H8000000F&
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00884400&
         Caption         =   "Coordinador"
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
         Height          =   855
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1600
         Begin VB.TextBox txtCoordinador 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   120
            TabIndex        =   0
            Top             =   300
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00884400&
         Caption         =   "Asesor"
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
         Height          =   1695
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1600
         Begin VB.TextBox txtPorcentajeAsesor 
            Height          =   375
            Left            =   120
            TabIndex        =   1
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Porcentaje"
            ForeColor       =   &H8000000F&
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Comision"
            ForeColor       =   &H8000000F&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1080
            TabIndex        =   8
            Top             =   480
            Width           =   375
         End
         Begin VB.Label lblComisionAsesor 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Height          =   360
            Left            =   120
            TabIndex        =   7
            Top             =   1200
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmComisionCuota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aceptar As Boolean

Private Sub Form_Load()
    Centrar Me
        lblTotalCurso.Caption = Format(0, "currency")
        lblComisionAsesor.Caption = Format(0, "currency")
        lblTotalComisiones.Caption = Format(0, "currency")
    aceptar = False
End Sub

Private Sub cmdAceptar_Click()
    If txtCoordinador.Text = "" Then MsgBox "ingrese el valor de la comision del Coordinador", vbCritical, "Comision de Primera Cuota": txtCoordinador.SetFocus: Exit Sub
    If txtPorcentajeAsesor.Text = "" Then MsgBox "Ingrese el porcentaje de comisión del Asesor", vbCritical, "Comisión de Primera Cuota": txtPorcentajeAsesor.SetFocus: Exit Sub
    If txtPagoParcial.Text = "" Then MsgBox "Ingrese el pago parcial de la primera cuota", vbCritical, "Comisión de Primera Cuota": txtPagoParcial.SetFocus: Exit Sub
    If txtNroFactura.Text = "" Then MsgBox "Ingrese el número de factura", vbCritical, "Comisión de Primera Cuota": txtNroFactura.SetFocus: Exit Sub

    With rsContabilidad
        If .State = 1 Then .Close
        .Open "SELECT * FROM contabilidad", Cn, adOpenDynamic, adLockPessimistic
    ''' Movimientos de Caja
        '''CAJA ADMINISTRACION (Parcial 1° Cuota - DEBE)
            .AddNew
            !fecha = Date
            !cuenta = "CAJA ADMINISTRACION"
            !Detalle = "Parcial de 1ª Cuota de " & frmPlanDePagos.lblCodAlumno.Caption
            !Debe = Int(lblTotalCuota1.Caption) - Int(txtPagoParcial.Text)
            !Haber = Null
            !nrofactura = txtNroFactura.Text
            !CodAlumno = Int(frmPlanDePagos.lblCodAlumno.Caption)
            !NroCuota = 1
            .Update
        '''DESCUENTO (Parcial 1° Cuota - HABER)
            .Requery
            .AddNew
            !fecha = Date
            !cuenta = "Descuento"
            !Detalle = "Parcial de 1ª Cuota de " & frmPlanDePagos.lblCodAlumno.Caption
            !Debe = Int(txtPagoParcial.Text)
            !nrofactura = txtNroFactura.Text
            !Haber = Null
            !CodAlumno = Int(frmPlanDePagos.lblCodAlumno.Caption)
            !NroCuota = 1
            .Update
        '''COMISION COORDINADOR (Comision 1° Cuota - DEBE)
            .AddNew
            !fecha = Date
            !cuenta = "COMISIONES VARIAS"
            !Detalle = "Comisión Coord. 1ª Cuota de " & frmPlanDePagos.lblCodAlumno.Caption
            !Debe = Int(txtCoordinador.Text)
            !Haber = Null
            !CodAlumno = Null
            !NroCuota = Null
            .Update
        '''CAJA ADMINISTRACION (Comision 1° Cuota - HABER)
            .Requery
            .AddNew
            !fecha = Date
            !cuenta = "CAJA ADMINISTRACION"
            !Detalle = "Comisión Coord. 1ª Cuota de " & frmPlanDePagos.lblCodAlumno.Caption
            !Debe = Null
            !Haber = Int(txtCoordinador.Text)
            !CodAlumno = Null
            !NroCuota = Null
            .Update
        '''COMISION ASESOR (Comision 1° Cuota - DEBE)
            .Requery
            .AddNew
            !fecha = Date
            !cuenta = "HONORARIOS ASESORES"
            !Detalle = "Comisión de 1ª Cuota de " & frmPlanDePagos.lblCodAlumno.Caption
            !Debe = Int(lblComisionAsesor.Caption)
            !Haber = Null
            !CodAlumno = Null
            !NroCuota = Null
            .Update
        '''CAJA ADMINISTRACION (Comision 1° Cuota - HABER)
            .Requery
            .AddNew
            !fecha = Date
            !cuenta = "CAJA ADMINISTRACION"
            !Detalle = "Comisión de 1ª Cuota de " & frmPlanDePagos.lblCodAlumno.Caption
            !Debe = Null
            !Haber = Int(lblComisionAsesor.Caption)
            !CodAlumno = Null
            !NroCuota = Null
            .Update
    End With
    
    With rsPlanDePago
        If .State = 1 Then .Close
        .Open "SELECT * FROM plandepago WHERE codalumno=" & Int(frmPlanDePagos.lblCodAlumno.Caption), Cn, adOpenDynamic, adLockPessimistic
    '''Plan de Pago (Liquidacion de Cuota)
        .Requery
        !tipodepago = "PAG"
        !fechapago = Date
        !DeudaTotal = 0
        !recibo = txtNroFactura.Text
        !totalcobrado = Int(lblTotalCuota1.Caption)
        !CuotasDebidas = 0
        .UpdateBatch
    End With
    aceptar = True
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If aceptar = True Then
        frmPlanDePagos.Enabled = True
    Else
        MsgBox "No se han cargado las comisiones en el sistema", vbCritical, "Comisión de Primera Cuota"
        Cancel = True
    End If
End Sub

Private Sub txtCoordinador_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FormatoNumeros
        CalcularComision
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtPorcentajeAsesor_keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FormatoNumeros
        CalcularComision
        SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNroFactura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtPagoParcial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FormatoNumeros
        CalcularComision
        SendKeys "{TAB}"
    End If
End Sub

Sub FormatoNumeros()
    If txtCoordinador.Text = "" Then
            txtCoordinador.Text = "0"
        ElseIf txtPorcentajeAsesor.Text = "" Then
            txtPorcentajeAsesor.Text = "0"
        ElseIf lblComisionAsesor.Caption = "" Then
            lblComisionAsesor.Caption = "0"
        ElseIf txtPagoParcial.Text = "" Then
            txtPagoParcial.Text = "0"
    End If
End Sub

Sub CalcularComision()
    lblComisionAsesor.Caption = FormatCurrency((Int(txtPorcentajeAsesor.Text) * Int(lblTotalCurso.Caption)) / 100)
    lblTotalComisiones.Caption = FormatCurrency(Int(txtCoordinador.Text) + Int(lblComisionAsesor.Caption))
    lblTotal.Caption = FormatCurrency(Int(lblTotalCuota1.Caption) - Int(lblTotalComisiones.Caption) - Int(txtPagoParcial.Text))
End Sub

