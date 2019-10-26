VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmContabilidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad"
   ClientHeight    =   4380
   ClientLeft      =   3330
   ClientTop       =   1605
   ClientWidth     =   7110
   Icon            =   "frmContabilidad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmContabilidad.frx":324A
   ScaleHeight     =   4380
   ScaleWidth      =   7110
   Begin VB.Frame Frame1 
      BackColor       =   &H00662200&
      Caption         =   "Bajar"
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
      Left            =   5400
      TabIndex        =   18
      Top             =   120
      Width           =   1575
      Begin isButtonTest.isButton cmdBajarCuenta 
         Height          =   420
         Left            =   120
         TabIndex        =   5
         Top             =   800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmContabilidad.frx":AC67
         Style           =   8
         Caption         =   "       Cuenta"
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
      Begin isButtonTest.isButton cmdBajarAsiento 
         Height          =   420
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmContabilidad.frx":B541
         Style           =   8
         Caption         =   "       Asiento"
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
         Left            =   120
         TabIndex        =   7
         Top             =   1300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmContabilidad.frx":BE1B
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
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   2175
      Left            =   120
      TabIndex        =   17
      Top             =   2040
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin VB.TextBox txtNroFactura 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3000
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox txtHaber 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4200
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtDebe 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3000
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtDetalle 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   5175
   End
   Begin MSDataListLib.DataCombo dtcCuenta 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nro Factura"
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
      Left            =   3000
      TabIndex        =   16
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Haber"
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
      Left            =   4200
      TabIndex        =   15
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Debe"
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
      Left            =   3000
      TabIndex        =   14
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Detalle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Detalle"
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
      TabIndex        =   13
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
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
      TabIndex        =   12
      Top             =   720
      Width           =   585
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      TabIndex        =   11
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      TabIndex        =   10
      Top             =   120
      Width           =   510
   End
   Begin VB.Label lblNroAsiento 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N° Asiento"
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
      TabIndex        =   8
      Top             =   120
      Width           =   840
   End
End
Attribute VB_Name = "frmContabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBajarAsiento_Click()
    
If Debe = Haber Then
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
    
    Dim nroasiento As Long
    With rsControl
        .Requery
        .MoveFirst
        nroasiento = !nroasiento
        nroasiento = nroasiento + 1
        !nroasiento = nroasiento
        .UpdateBatch
    End With
    lblNroAsiento.Caption = nroasiento
    txtNroFactura.Text = ""
    txtDetalle.Text = ""
    txtNroFactura.Locked = False
    txtNroFactura.SetFocus
    cmdBajarAsiento.Enabled = False
Else
    MsgBox "No coinciden Debe y Haber", vbOKOnly + vbInformation, "Contabilidad"
End If
End Sub

Private Sub cmdBajarCuenta_Click()
    If txtNroFactura.Text = "" Then MsgBox "Debe ingresar un número de factura", vbOKOnly + vbInformation, "Contabilidad": txtNroFactura.SetFocus: Exit Sub
    If txtDetalle.Text = "" Then MsgBox "Debe ingresar un detalle", vbOKOnly + vbInformation, "Contabilidad": txtDetalle.SetFocus: Exit Sub
    If dtcCuenta.Text = "" Then MsgBox "Debe ingresar una cuenta", vbOKOnly + vbInformation, "Contabilidad": dtcCuenta.SetFocus: Exit Sub
    If txtHaber.Text = "" And txtDebe.Text = "" Then MsgBox "Debe ingresar un monto a la cuenta", vbOKOnly + vbInformation, "Contabilidad": txtDebe.SetFocus: Exit Sub

    Debe = Val(txtDebe.Text) + Debe
    Haber = Haber + Val(txtHaber.Text)
    
    If Debe > 0 And Debe = Haber Then cmdBajarAsiento.Enabled = True
    
    With rsContabilidadTemp
        .Requery
        .AddNew
        !asiento = lblNroAsiento.Caption
        !fecha = CDate(lblfecha.Caption)
        !cuenta = dtcCuenta.Text
        !Detalle = txtDetalle.Text
        !NroCuota = Null
        !CodAlumno = Null
        If txtDebe.Text = "" Then
            !Debe = Null
        Else
            !Debe = txtDebe.Text
        End If
        If txtHaber.Text = "" Then
            !Haber = Null
        Else
            !Haber = txtHaber.Text
        End If
        !nrofactura = txtNroFactura.Text
        .Update
    End With
    formatoGrilla
    Limpiar
End Sub

Private Sub cmdCancelar_Click()
    With rsContabilidadTemp
        .Requery
        If .BOF Or .EOF Then Unload Me: Exit Sub
        .MoveFirst
        Do Until .EOF
            .Delete
            .Update
            .MoveFirst
        Loop
    End With
    Unload Me
End Sub


Private Sub dtcCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Centrar Me
    Cuentas
    Control
    Contabilidad
    ContabilidadTemp
    lblNroAsiento.Caption = rsControl!nroasiento
    lblfecha.Caption = Date
    Set dtcCuenta.RowSource = rsCuentas
    dtcCuenta.BoundColumn = "cuenta"
    dtcCuenta.ListField = "cuenta"
    Set grilla.DataSource = rsContabilidadTemp
    formatoGrilla
    txtNroFactura.Locked = False
    Haber = 0
    Debe = 0
End Sub

Sub formatoGrilla()
    grilla.Columns(0).Width = 0
    grilla.Columns(1).Width = 0
    grilla.Columns(2).Width = 0
    grilla.Columns(7).Width = 0
End Sub

Sub Limpiar()
    txtNroFactura.Locked = True
    dtcCuenta.Text = ""
    'txtDetalle.Text = ""
    dtcCuenta.SetFocus
    txtDebe.Text = ""
    txtHaber.Text = ""
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    With rsContabilidadTemp
        .Requery
        If .BOF Or .EOF Then Unload Me: Exit Sub
        .MoveFirst
        Do Until .EOF
            .Delete
            .Update
            .MoveFirst
        Loop
    End With

    Unload Me
End Sub

Private Sub txtDebe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtDetalle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtHaber_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub txtNroFactura_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
