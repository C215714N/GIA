VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmPersonal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal"
   ClientHeight    =   4260
   ClientLeft      =   4530
   ClientTop       =   1725
   ClientWidth     =   6375
   Icon            =   "frmPersonal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmPersonal.frx":324A
   ScaleHeight     =   4260
   ScaleWidth      =   6375
   Begin VB.Frame Frame4 
      BackColor       =   &H00662200&
      Caption         =   "Personal"
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
      Height          =   4000
      Left            =   4680
      TabIndex        =   23
      Top             =   120
      Width           =   1575
      Begin isButtonTest.isButton cmdBuscar 
         Height          =   420
         Left            =   120
         TabIndex        =   25
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmPersonal.frx":AC67
         Style           =   8
         Caption         =   "       Buscar"
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
      Begin isButtonTest.isButton cmdGrabar 
         Height          =   420
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmPersonal.frx":B541
         Style           =   8
         Caption         =   "       Guardar"
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
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmPersonal.frx":BE1B
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
      Begin isButtonTest.isButton cmdNuevo 
         Height          =   420
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmPersonal.frx":C6F5
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
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmPersonal.frx":CFCF
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
         Left            =   120
         TabIndex        =   28
         Top             =   3360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmPersonal.frx":D8A9
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00884400&
      Caption         =   "Datos Personales"
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
      Height          =   4000
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   4455
      Begin MSDataListLib.DataCombo dtcCargo 
         Height          =   360
         Left            =   1560
         TabIndex        =   9
         Top             =   3480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   635
         _Version        =   393216
         Text            =   ""
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
      Begin MSComCtl2.DTPicker dtpFechaNacimiento 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
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
         Format          =   89260033
         CurrentDate     =   41319
      End
      Begin VB.TextBox txtNya 
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
         TabIndex        =   0
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox txtTelCel 
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
         Left            =   2280
         TabIndex        =   7
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox txtTelCasa 
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
         TabIndex        =   6
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox txtLocalidad 
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
         TabIndex        =   5
         Top             =   2280
         Width           =   4215
      End
      Begin VB.TextBox txtDireccion 
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
         TabIndex        =   4
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox txtDNI 
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
         Left            =   2760
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
      Begin VB.ComboBox cmbTipoDoc 
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
         ItemData        =   "frmPersonal.frx":E183
         Left            =   1560
         List            =   "frmPersonal.frx":E190
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dtpFechaIngreso 
         Height          =   360
         Left            =   120
         TabIndex        =   8
         Top             =   3480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   635
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
         Format          =   89260033
         CurrentDate     =   41319
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Ingreso"
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
         Height          =   300
         Left            =   120
         TabIndex        =   22
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Cargo"
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
         Height          =   300
         Left            =   1560
         TabIndex        =   21
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Nac."
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
         Height          =   300
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono Celular"
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
         Height          =   300
         Left            =   2280
         TabIndex        =   19
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono Casa"
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
         Height          =   300
         Left            =   120
         TabIndex        =   18
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre y Apellido"
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
         Height          =   300
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Documento"
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
         Height          =   300
         Left            =   2760
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Doc."
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
         Height          =   300
         Left            =   1560
         TabIndex        =   15
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección"
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
         Height          =   300
         Left            =   120
         TabIndex        =   14
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Localidad"
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
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   2040
         Width           =   855
      End
   End
   Begin VB.Label lblID 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Left            =   1560
      TabIndex        =   24
      Top             =   4920
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "frmPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Private Sub cmbTipoDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDNI.SetFocus
End Sub

Private Sub cmdBuscar_Click()
    frmBuscarPersonal.Show
    Me.Enabled = False
    
End Sub

Private Sub cmdCancelar_Click()
    HabilitarCuadros True, False
    HabilitarBotones True, False
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdGrabar_Click()
    If txtNya.Text = "" Then MsgBox "Debe ingresar un Nombre de Empleado", vbOKOnly + vbInformation, "Personal": txtNya.SetFocus: Exit Sub
    If cmbTipoDoc.Text = "" Then MsgBox "Debe ingresar un Tipo de Documento", vbOKOnly + vbInformation, "Personal": cmbTipoDoc.SetFocus: Exit Sub
    If txtDNI.Text = "" Then MsgBox "Debe ingresar un Número de Documento", vbOKOnly + vbInformation, "Personal": txtDNI.SetFocus: Exit Sub
    If txtDireccion.Text = "" Then MsgBox "Debe ingresar una Dirección", vbOKOnly + vbInformation, "Personal": txtDireccion.SetFocus: Exit Sub
    If txtLocalidad.Text = "" Then MsgBox "Debe ingresar una Localidad", vbOKOnly + vbInformation, "Personal": txtLocalidad.SetFocus: Exit Sub
    If txtTelCasa.Text = "" Then MsgBox "Debe ingresar un Teléfono", vbOKOnly + vbInformation, "Personal": txtTelCasa.SetFocus: Exit Sub
    If txtTelCel.Text = "" Then MsgBox "Debe ingresar un Teléfono", vbOKOnly + vbInformation, "Personal": txtTelCel.SetFocus: Exit Sub
    If dtcCargo.Text = "" Then MsgBox "Debe ingresar un Cargo", vbOKOnly + vbInformation, "Personal": dtcCargo.SetFocus: Exit Sub
    
    On Error GoTo LineaError
    
    If Modi = False Then
        With rsPersonal
            .Requery
            .AddNew
            !NyA = txtNya.Text
            !tipodoc = cmbTipoDoc.Text
            !dni = txtDNI.Text
            !direccion = txtDireccion.Text
            !localidad = txtLocalidad.Text
            !Fechanacimiento = dtpFechaNacimiento.Value
            !cargo = dtcCargo.Text
            !telcasa = txtTelCasa.Text
            !telcel = txtTelCel.Text
            !fechaingreso = dtpFechaIngreso.Value
            .Update
            .Requery
        End With
    Else
        With rsPersonal
            .Requery
            .Find "ID='" & lblID.Caption & "'"
            !NyA = txtNya.Text
            !tipodoc = cmbTipoDoc.Text
            !dni = txtDNI.Text
            !direccion = txtDireccion.Text
            !localidad = txtLocalidad.Text
            !Fechanacimiento = dtpFechaNacimiento.Value
            !cargo = dtcCargo.Text
            !telcasa = txtTelCasa.Text
            !telcel = txtTelCel.Text
            !fechaingreso = dtpFechaIngreso.Value
            .UpdateBatch
            .Requery
        End With
    End If
    HabilitarBotones True, False
    HabilitarCuadros True, False
    Limpiar

LineaError:
    Select Case Err.Number
        Case 3021
            Resume Next
        End Select
End Sub

Private Sub cmdModificar_Click()
    If txtNya.Text = "" Then
        MsgBox "Primero debe Buscar un Empleado", vbOKOnly, "Personal"
        cmdBuscar.SetFocus
    Else
        HabilitarBotones False, True
        HabilitarCuadros False, True
        txtNya.SetFocus
        Modi = True
    End If
End Sub

Private Sub cmdNuevo_Click()
    HabilitarBotones False, True
    HabilitarCuadros False, True
    Limpiar
    txtNya.SetFocus
    Modi = False
End Sub

Private Sub dtcCargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpFechaIngreso.SetFocus
End Sub

Private Sub dtpFechaIngreso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdGrabar.SetFocus
End Sub

Private Sub dtpFechaNacimiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtcCargo.SetFocus
End Sub

Private Sub Form_Load()
    Centrar Me
    Personal
    Cargos
    HabilitarBotones True, False
    HabilitarCuadros True, False
    Limpiar
    Set dtcCargo.RowSource = rsCargos
    dtcCargo.BoundColumn = "cargo"
    dtcCargo.ListField = "cargo"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub

Sub HabilitarBotones(estado1 As Boolean, estado2 As Boolean)
    cmdNuevo.Enabled = estado1
    cmdModificar.Enabled = estado1
    cmdBuscar.Enabled = estado1
    cmdGrabar.Enabled = estado2
    cmdCancelar.Enabled = estado2
    cmdCerrar.Enabled = estado1
End Sub

Sub HabilitarCuadros(estado1 As Boolean, estado2 As Boolean)
    txtNya.Locked = estado1
    cmbTipoDoc.Locked = estado1
    txtDireccion.Locked = estado1
    txtLocalidad.Locked = estado1
    txtTelCasa.Locked = estado1
    txtTelCel.Locked = estado1
    txtDNI.Locked = estado1
    dtpFechaNacimiento.Enabled = estado2
    dtpFechaIngreso.Enabled = estado2
    dtcCargo.Locked = estado1
End Sub

Sub Limpiar()
    txtNya.Text = ""
    txtDireccion.Text = ""
    txtLocalidad.Text = ""
    txtTelCasa.Text = ""
    txtTelCel.Text = ""
    txtDNI.Text = ""
    dtpFechaNacimiento.Value = Date
    dtpFechaIngreso.Value = Date
    dtcCargo.Text = ""
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtLocalidad.SetFocus
End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDireccion.SetFocus
    
    If KeyAscii = 46 Then KeyAscii = 0
End Sub

Private Sub txtLocalidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTelCasa.SetFocus
End Sub

Private Sub txtNya_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmbTipoDoc.SetFocus
End Sub

Private Sub txtTelCasa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTelCel.SetFocus
End Sub

Private Sub txtTelCel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpFechaNacimiento.SetFocus
End Sub
