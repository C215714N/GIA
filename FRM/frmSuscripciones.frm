VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmSuscripciones 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion Integral del Alumno - Suscripciones"
   ClientHeight    =   5085
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   11415
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmSuscripciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   11415
   Begin VB.Frame Frame5 
      BackColor       =   &H00662200&
      Caption         =   "Observaciones"
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
      Height          =   2000
      Left            =   4680
      TabIndex        =   53
      Top             =   3000
      Width           =   4935
      Begin RichTextLib.RichTextBox txtObservaciones 
         Height          =   1575
         Left            =   120
         TabIndex        =   58
         Top             =   300
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   2778
         _Version        =   393217
         Enabled         =   -1  'True
         MaxLength       =   1000
         Appearance      =   0
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmSuscripciones.frx":10CA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00662200&
      Caption         =   "Suscripciones"
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
      Height          =   4850
      Left            =   9720
      TabIndex        =   34
      Top             =   120
      Width           =   1600
      Begin isButtonTest.isButton cmdNuevo 
         Height          =   420
         Left            =   120
         TabIndex        =   54
         Top             =   600
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   741
         Icon            =   "frmSuscripciones.frx":1150
         Style           =   8
         Caption         =   "     Nuevo"
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
      Begin isButtonTest.isButton cmdModificar 
         Height          =   420
         Left            =   120
         TabIndex        =   55
         Top             =   1200
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   741
         Icon            =   "frmSuscripciones.frx":1A2A
         Style           =   8
         Caption         =   "     Editar"
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
      Begin isButtonTest.isButton cmdBuscar 
         Height          =   420
         Left            =   120
         TabIndex        =   56
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmSuscripciones.frx":2304
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
      Begin isButtonTest.isButton cmdGrabar 
         Height          =   420
         Left            =   120
         TabIndex        =   26
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmSuscripciones.frx":2BDE
         Style           =   8
         Caption         =   "     Aceptar"
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
      Begin isButtonTest.isButton cmdCancelar 
         Height          =   420
         Left            =   120
         TabIndex        =   27
         Top             =   3000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmSuscripciones.frx":34B8
         Style           =   8
         Caption         =   "     Cancelar"
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
      Begin isButtonTest.isButton cmdCerrar 
         Height          =   420
         Left            =   120
         TabIndex        =   57
         Top             =   3600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmSuscripciones.frx":3D92
         Style           =   8
         Caption         =   "     Volver"
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
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00662200&
      Caption         =   "Curso"
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
      Height          =   2895
      Left            =   6720
      TabIndex        =   29
      Top             =   120
      Width           =   2895
      Begin VB.TextBox txtTotalMatricula 
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
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1200
         Width           =   1215
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
         Left            =   1440
         TabIndex        =   20
         Top             =   1200
         Width           =   1335
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
         Height          =   360
         ItemData        =   "frmSuscripciones.frx":466C
         Left            =   120
         List            =   "frmSuscripciones.frx":4679
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CheckBox chkManuales 
         BackColor       =   &H00662200&
         Caption         =   "Manuales"
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
         Left            =   1440
         TabIndex        =   25
         Top             =   2520
         Width           =   1260
      End
      Begin VB.CheckBox chkExamenes 
         BackColor       =   &H00662200&
         Caption         =   "Examenes"
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
         Left            =   1440
         TabIndex        =   24
         Top             =   2280
         Width           =   1260
      End
      Begin VB.TextBox txtGastoAdm 
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
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1800
         Width           =   1215
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtTotalCurso 
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
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpFechaSuscripcion 
         Height          =   375
         Left            =   1440
         TabIndex        =   22
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
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
         Format          =   127074305
         CurrentDate     =   41308
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Factura"
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
         Left            =   1440
         TabIndex        =   38
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Forma Pago"
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
         TabIndex        =   37
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Matricula"
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
         TabIndex        =   36
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Susc."
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
         Left            =   1440
         TabIndex        =   33
         Top             =   1560
         Width           =   960
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Gasto Adm."
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
         TabIndex        =   32
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label cuotas 
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
         Left            =   1440
         TabIndex        =   31
         Top             =   360
         Width           =   585
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Curso"
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
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Width           =   885
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00662200&
      Caption         =   "Telefonos"
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
      Height          =   2000
      Left            =   120
      TabIndex        =   28
      Top             =   3000
      Width           =   4455
      Begin VB.TextBox txtPT4 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1500
         Width           =   2055
      End
      Begin VB.TextBox txtPT3 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1100
         Width           =   2055
      End
      Begin VB.TextBox txtPT2 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   700
         Width           =   2055
      End
      Begin VB.TextBox txtPT1 
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   300
         Width           =   2055
      End
      Begin VB.TextBox txtTel4 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1500
         Width           =   2055
      End
      Begin VB.TextBox txtTel3 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1100
         Width           =   2055
      End
      Begin VB.TextBox txtTel2 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   700
         Width           =   2055
      End
      Begin VB.TextBox txtTel1 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00662200&
      Caption         =   "Datos Personales"
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
      Height          =   2895
      Left            =   120
      TabIndex        =   39
      Top             =   120
      Width           =   6495
      Begin VB.TextBox txtDocumento 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4800
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtEdad 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtNacionalidad 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtCP 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   1800
         Width           =   975
      End
      Begin VB.ComboBox cmbTipoDoc 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmSuscripciones.frx":469B
         Left            =   3720
         List            =   "frmSuscripciones.frx":46A8
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtDireccion 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtNya 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   600
         Width           =   3495
      End
      Begin MSDataListLib.DataCombo dtcLocalidad 
         Height          =   360
         Left            =   3720
         TabIndex        =   4
         Top             =   1200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpFechaNacimiento 
         Height          =   375
         Left            =   3720
         TabIndex        =   6
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
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
         Format          =   127139841
         CurrentDate     =   41308
      End
      Begin MSDataListLib.DataCombo dtcAsistente 
         Height          =   360
         Left            =   3720
         TabIndex        =   8
         Top             =   2400
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo dtcCapacitacion 
         Height          =   360
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   661
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century Gothic"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Asistente"
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
         Left            =   3720
         TabIndex        =   52
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacitacion"
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
         TabIndex        =   51
         Top             =   2160
         Width           =   1365
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edad"
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
         Left            =   2760
         TabIndex        =   50
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nacionalidad"
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
         Left            =   120
         TabIndex        =   49
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Doc"
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
         Left            =   3720
         TabIndex        =   48
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "N° Documento"
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
         Left            =   4800
         TabIndex        =   47
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Localidad"
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
         Left            =   3720
         TabIndex        =   46
         Top             =   960
         Width           =   2625
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Direccion"
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
         TabIndex        =   45
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Nacimiento"
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
         Left            =   3720
         TabIndex        =   44
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido y Nombres"
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
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   2475
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.P."
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
         Left            =   5400
         TabIndex        =   42
         Top             =   1560
         Width           =   330
      End
   End
   Begin VB.Label lblID 
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Left            =   6120
      TabIndex        =   35
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmSuscripciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkEscuela_Click()
    If txtTotalCurso.text = "1" And chkEscuela.Value = 1 Then
        lblTotalMatricula.Caption = Format(rsControl!matriculaescuela, "currency")
    Else
            lblTotalMatricula.Caption = Format(rsControl!matricula, "currency")
    End If
End Sub
Private Sub cmdBuscar_Click()
    frmBuscarSuscripcion.Show
    Verificar = False
    Me.Enabled = False
End Sub

Private Sub cmdCancelar_Click()
    HabilitarBotones True, False
    HabilitarCuadros True, False
    Limpiar
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdGrabar_Click()
    If txtNya.text = "" Then MsgBox "Debe ingresar un Nombre de Alumno", vbOKOnly + vbInformation, "Suscripciones": txtNya.SetFocus: Exit Sub
    If cmbTipoDoc.text = "" Then MsgBox "Debe ingresar un Tipo de Documento", vbOKOnly + vbInformation, "Suscripciones": cmbTipoDoc.SetFocus: Exit Sub
    If txtDocumento.text = "" Then MsgBox "Debe ingresar un Numero de Documento", vbOKOnly + vbInformation, "Suscripciones": txtDocumento.SetFocus: Exit Sub
    If txtDireccion.text = "" Then MsgBox "Debe ingresar una Direccion", vbOKOnly + vbInformation, "Suscripciones": txtDireccion.SetFocus: Exit Sub
    If txtCP.text = "" Then MsgBox "Debe ingresar un Codigo Postal", vbOKOnly + vbInformation, "Suscripciones": txtCP.SetFocus: Exit Sub
    If dtcLocalidad.text = "" Then MsgBox "Debe ingresar una Localidad", vbOKOnly + vbInformation, "Suscripciones": txtLocalidad.SetFocus: Exit Sub
    If txtNacionalidad.text = "" Then MsgBox "Debe ingresar una Nacionalidad del Alumno", vbOKOnly + vbInformation, "Suscripciones": txtNacionalidad.SetFocus: Exit Sub
    If dtcCapacitacion.text = "" Then MsgBox "Debe ingresar una Capacitacion ", vbOKOnly + vbInformation, "Suscripciones": dtcCapacitacion.SetFocus: Exit Sub
    If dtcAsistente.text = "" Then MsgBox "Debe ingresar un Asistente", vbOKOnly + vbInformation, "Suscripciones": dtcAsistente.SetFocus: Exit Sub
    If txtPT1.text = "" Then MsgBox "Debe ingresar al menos un Telefono", vbOKOnly + vbInformation, "Suscripciones": txtPT1.SetFocus: Exit Sub
    If txtTel1.text = "" Then MsgBox "Debe ingresar al menos un Telefono", vbOKOnly + vbInformation, "Suscripciones": txtTel1.SetFocus: Exit Sub
    If txtTotalCurso.text = "" Or txtTotalCurso.text = "0" Then MsgBox "Debe ingresar el Precio del Curso." & vbNewLine & "El mismo debe ser superior a Cero", vbOKOnly + vbInformation, "Suscripciones": txtTotalCurso.SetFocus: Exit Sub
    If txtTotalCuotas.text = "" Or txtTotalCuotas.text = "0" Then MsgBox "Debe ingresar la Cantidad de Cuotas." & vbNewLine & "La misma debe ser superior a Cero", vbOKOnly + vbInformation, "Suscripciones": txtTotalCuotas.SetFocus: Exit Sub
    If txtGastoAdm.text = "" Then MsgBox "Debe ingresar el Gasto Administrativo", vbOKOnly + vbInformation, "Suscripciones": txtGastoAdm.SetFocus: Exit Sub
    If cmbTipoPago.text = "" Then MsgBox "Ingrese el tipo de pago de la Matricula", vbOKOnly + vbInformation, "Suscripciones": cmbTipoPago.SetFocus: Exit Sub
    If txtNroFactura.text = "" Then MsgBox "Ingrese el Numero de Factura del pago de la Matricula", vbInformation, "Suscripciones": txtNroFactura.SetFocus: Exit Sub
    If txtTotalMatricula.text = "" Then MsgBox "Debe ingresar el valor de la Matricula", vbOKOnly + vbInformation, "Suscripciones": txtTotalMatricula.SetFocus: Exit Sub
    
    On Error GoTo LineaError
    
    If Modi = False Then
        With rsSuscripciones
            .Requery
            .AddNew
            !NyA = txtNya.text
            !tipodoc = cmbTipoDoc.text
            !dni = txtDocumento.text
            !direccion = txtDireccion.text
            !cp = txtCP.text
            !localidad = dtcLocalidad.text
            !nacionalidad = txtNacionalidad.text
            !fechanac = dtpFechaNacimiento.Value
            !capac = dtcCapacitacion.text
            !Asistente = dtcAsistente.text
            !edad = txtEdad.text
            !ptel1 = txtPT1.text
            !ptel2 = txtPT2.text
            !ptel3 = txtPT3.text
            !ptel4 = txtPT4.text
            !tel1 = txtTel1.text
            !tel2 = txtTel2.text
            !tel3 = txtTel3.text
            !tel4 = txtTel4.text
            !totalcurso = Int(txtTotalCurso.text)
            !cuotas = Int(txtTotalCuotas.text)
            !gastoadm = Int(txtGastoAdm.text)
            !fechasus = dtpFechaSuscripcion.Value
            !observaciones = txtObservaciones.text
            !manuales = chkManuales.Value
            !dchoexamen = chkExamenes.Value
            !totalmatricula = Int(txtTotalMatricula.text)
            !nrofactura = txtNroFactura.text
            .Update
            lblID.Caption = !ID
            .Requery
        End With
        
        With rsContabilidad
            If .State = 1 Then .Close
            .Open "SELECT * FROM contabilidad", Cn, adOpenDynamic, adLockPessimistic
            .Requery
            .AddNew
            !fecha = Date
            !asiento = Null
            !NroCuota = Null
            !CodAlumno = Null
            !Cuenta = "MATRICULA DE CURSO"
            !Detalle = "Matricula del Alumno " & txtNya.text
            !nrofactura = txtNroFactura.text
            !Haber = CSng(txtGastoAdm.text)
            !Debe = Null
            .Update
            .Requery
            .AddNew
            !fecha = dtpFechaSuscripcion.Value
            
            If cmbTipoPago.text = "Efectivo" Then
                !Cuenta = "CAJA ADMINISTRACION"
            ElseIf cmbTipoPago.text = "Descuento" Then
                !Cuenta = "Descuento"
            Else
                !Cuenta = "DEBITO TARJETA CREDITO"
            End If
            
            !Detalle = "Matricula del Alumno " & txtNya.text
            !nrofactura = txtNroFactura.text
            !Debe = CSng(txtGastoAdm.text)
            !asiento = Null
            !NroCuota = Null
            !CodAlumno = Null
            !Haber = Null
            .Update
        End With
        
        With rsMatriculas
            If .State = 1 Then .Close
            .Open "SELECT * FROM matriculas", Cn, adOpenDynamic, adLockPessimistic
            .Requery
            .AddNew
            !ID = Int(lblID.Caption)
            !totalmatricula = CSng(txtTotalMatricula.text)
            !abonado = CSng(txtGastoAdm.text)
            !Debe = CSng(txtTotalMatricula.text) - CSng(txtGastoAdm.text)
            .Update
        End With
        
        With rsInformeSuscripciones
            If .State = 1 Then .Close
            .Open "SELECT * FROM informesuscripciones", Cn, adOpenDynamic, adLockPessimistic
            .Requery
            .AddNew
            !fechaS = dtpFechaSuscripcion.Value
            !fechaV = Null
            !Asistente = dtcAsistente.text
            !curso = dtcCapacitacion.text
            !totalcurso = FormatCurrency(txtTotalCurso.text)
            !verificado = 0
            .Update
            .Requery
        End With
        
        ''''si el alumno es 100% lo agrega en la tabla correspondiente
        If Int(txtTotalCurso.text) = 1 And Int(txtTotalCuotas.text) = 1 Then
            
            With rsAlumnosBecados
                If .State = 1 Then .Close
                .Open "SELECT * FROM alumnosbecados", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .AddNew
                !idreferencial = Int(lblID.Caption)
                !matricula = Int(txtGastoAdm.text)
                !Debe = Int(txtTotalMatricula.text) - Int(txtGastoAdm.text)
                !cancelacion = Date + 1
                .Update
            End With
        End If
        
        HabilitarBotones True, False
        HabilitarCuadros True, False
        Limpiar
    Else
        With rsSuscripciones
            .Requery
            .Find "ID='" & lblID.Caption & "'"
            !NyA = txtNya.text
            !tipodoc = cmbTipoDoc.text
            !dni = txtDocumento.text
            !direccion = txtDireccion.text
            !cp = txtCP.text
            !localidad = dtcLocalidad.text
            !nacionalidad = txtNacionalidad.text
            !fechanac = dtpFechaNacimiento.Value
            !capac = dtcCapacitacion.text
            !Asistente = dtcAsistente.text
            !edad = txtEdad.text
            !ptel1 = txtPT1.text
            !ptel2 = txtPT2.text
            !ptel3 = txtPT3.text
            !ptel4 = txtPT4.text
            !tel1 = txtTel1.text
            !tel2 = txtTel2.text
            !tel3 = txtTel3.text
            !tel4 = txtTel4.text
            !totalcurso = Int(txtTotalCurso.text)
            !cuotas = Int(txtTotalCuotas.text)
            !gastoadm = Int(txtGastoAdm.text)
            !fechasus = dtpFechaSuscripcion.Value
            !observaciones = txtObservaciones.text
            !manuales = chkManuales.Value
            !dchoexamen = chkExamenes.Value
            !totalmatricula = Int(txtTotalMatricula.text)
            !nrofactura = txtNroFactura.text
            .UpdateBatch
            .Requery
        End With
        HabilitarBotones True, False
        HabilitarCuadros True, False
        Limpiar
    End If

LineaError: ErrCode
End Sub

Private Sub cmdModificar_Click()
    If txtNya.text = "" Then
        MsgBox "Primero debe realizar una Busqueda", vbOKOnly + vbInformation, "Gestion Integral del Alumno"
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

Private Sub dtcLocalidad_KeyPress(keyAscii As Integer)
    If keyAscii = 13 Then
        txtNacionalidad.SetFocus
        With rsLocalidades
            .Find "localidad='" & dtcLocalidad.text & "'"
             txtCP.text = !cp
        End With
    End If
End Sub

Private Sub dtpFechaNacimiento_Change()
    If (Month(dtpFechaNacimiento.Value) < Month(Date)) Or (Day(dtpFechaNacimiento.Value) <= Day(Date) And Month(dtpFechaNacimiento.Value) = Month(Date)) Then
        txtEdad.text = DateDiff("yyyy", dtpFechaNacimiento.Value, Date)
    Else
        txtEdad.text = DateDiff("yyyy", dtpFechaNacimiento.Value, Date) - 1
    End If
End Sub
Private Sub Form_Load()
    Centrar Me
    Suscripciones
    Capacitaciones
    Asistente
    Localidades
    HabilitarBotones True, False
    HabilitarCuadros True, False
    Limpiar
    
    Set dtcLocalidad.RowSource = rsLocalidades
    dtcLocalidad.BoundColumn = "localidad"
    dtcLocalidad.ListField = "localidad"
    
    Set dtcCapacitacion.RowSource = rsCapacitaciones
    dtcCapacitacion.BoundColumn = "capacitacion"
    dtcCapacitacion.ListField = "capacitacion"
    
    Set dtcAsistente.RowSource = rsPersonal
    dtcAsistente.BoundColumn = "Personal"
    dtcAsistente.ListField = "Personal"
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
    dtcLocalidad.Locked = estado1
    txtNacionalidad.Locked = estado1
    txtPT1.Locked = estado1
    txtPT2.Locked = estado1
    txtPT3.Locked = estado1
    cmbTipoPago.Locked = estado1
    txtPT4.Locked = estado1
    txtTel1.Locked = estado1
    txtTel2.Locked = estado1
    txtTel3.Locked = estado1
    txtTel4.Locked = estado1
    txtNroFactura.Locked = estado1
    txtObservaciones.Locked = estado1
    txtGastoAdm.Locked = estado1
    txtTotalMatricula.Locked = estado1
    txtTotalCuotas.Locked = estado1
    txtTotalCurso.Locked = estado1
    txtDocumento.Locked = estado1
    dtpFechaNacimiento.Enabled = estado2
    dtpFechaSuscripcion.Enabled = estado2
    txtNroFactura.Locked = estado1
    chkExamenes.Enabled = estado2
    chkManuales.Enabled = estado2
End Sub

Sub Limpiar()
    txtNya.text = ""
    txtDireccion.text = ""
    txtEdad.text = ""
    txtCP.text = ""
    dtcLocalidad.text = ""
    txtNacionalidad.text = ""
    txtPT1.text = ""
    txtPT2.text = ""
    txtPT3.text = ""
    txtPT4.text = ""
    txtTel1.text = ""
    txtTel2.text = ""
    txtTel3.text = ""
    txtTel4.text = ""
    txtNroFactura.text = ""
    txtObservaciones.text = ""
    txtGastoAdm.text = ""
    txtTotalCuotas.text = ""
    txtTotalCurso.text = ""
    txtDocumento.text = ""
    txtTotalMatricula.text = ""
    dtpFechaNacimiento.Value = Date
    dtpFechaSuscripcion.Value = Date
    chkExamenes.Value = 0
    chkManuales.Value = 0
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
End Sub
Private Sub txtCP_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub txtDireccion_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub txtDocumento_KeyPress(keyAscii As Integer)
    Continue keyAscii
    If keyAscii = 46 Then keyAscii = 0
End Sub
Private Sub txtGastoAdm_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub txtLocalidad_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub txtNacionalidad_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub txtNroFactura_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub txtNya_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub txtPT1_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub txtPT2_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub txtPT3_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub txtPT4_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub txtTel1_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub txtTel2_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub txtTel3_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub txtTel4_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub txtTotalCuotas_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub txtTotalCurso_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub txtTotalMatricula_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub chkExamenes_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub chkManuales_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub cmbTipoDoc_KeyPress(keyAscii As Integer)
    Continue keyAscii
    keyAscii = 0
End Sub
Private Sub cmbTipoPago_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub dtcAsistente_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub dtcCapacitacion_KeyPress(keyAscii As Integer)
    Continue keyAscii
End Sub
Private Sub dtpFechaSuscripcion_KeyPress(keyAscii As Integer)
        If KeyPress = 13 Then cmbTipoPago.SetFocus
End Sub
