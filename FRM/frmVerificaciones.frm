VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmVerificaciones 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión Integral del Alumno - Verificaciones"
   ClientHeight    =   5085
   ClientLeft      =   3945
   ClientTop       =   1605
   ClientWidth     =   11415
   Icon            =   "frmVerificaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmVerificaciones.frx":324A
   ScaleHeight     =   5085
   ScaleWidth      =   11415
   Begin VB.Frame Frame6 
      BackColor       =   &H00662200&
      Caption         =   "Observaciones"
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
      Height          =   2000
      Left            =   4680
      TabIndex        =   52
      Top             =   3000
      Width           =   4935
      Begin VB.TextBox txtObservaciones 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1515
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   300
         Width           =   4695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00884400&
      Caption         =   "Teléfonos"
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
      Height          =   2000
      Left            =   120
      TabIndex        =   35
      Top             =   3000
      Width           =   4455
      Begin VB.TextBox txtPT1 
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
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   300
         Width           =   2055
      End
      Begin VB.TextBox txtPT2 
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
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   700
         Width           =   2055
      End
      Begin VB.TextBox txtPT3 
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
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1100
         Width           =   2055
      End
      Begin VB.TextBox txtPT4 
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
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1500
         Width           =   2055
      End
      Begin VB.TextBox txtTel1 
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
         TabIndex        =   9
         Top             =   300
         Width           =   2055
      End
      Begin VB.TextBox txtTel2 
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
         TabIndex        =   11
         Top             =   700
         Width           =   2055
      End
      Begin VB.TextBox txtTel3 
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
         TabIndex        =   13
         Top             =   1100
         Width           =   2055
      End
      Begin VB.TextBox txtTel4 
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
         TabIndex        =   15
         Top             =   1500
         Width           =   2055
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
      Height          =   2895
      Left            =   120
      TabIndex        =   36
      Top             =   120
      Width           =   6495
      Begin MSDataListLib.DataCombo dtcLocalidad 
         Height          =   360
         Left            =   3720
         TabIndex        =   4
         Top             =   1200
         Width           =   2655
         _ExtentX        =   4683
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
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   600
         Width           =   3495
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
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   3495
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
         ItemData        =   "frmVerificaciones.frx":AC67
         Left            =   3720
         List            =   "frmVerificaciones.frx":AC74
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtCP 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox txtNacionalidad 
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
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtEdad 
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
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtDocumento 
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
         Left            =   4800
         TabIndex        =   2
         Top             =   600
         Width           =   1575
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   89194497
         CurrentDate     =   41308
      End
      Begin MSDataListLib.DataCombo dtcAsistente 
         Height          =   360
         Left            =   3720
         TabIndex        =   8
         Top             =   2400
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
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
      Begin MSDataListLib.DataCombo dtcCapacitacion 
         Height          =   360
         Left            =   120
         TabIndex        =   7
         Top             =   2400
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.P."
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
         Height          =   240
         Left            =   5400
         TabIndex        =   49
         Top             =   1560
         Width           =   330
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Apellido y Nombres"
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
         Height          =   240
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Nacimiento"
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
         Height          =   255
         Left            =   3720
         TabIndex        =   45
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   120
         TabIndex        =   44
         Top             =   960
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   240
         Left            =   3720
         TabIndex        =   43
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Documento"
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
         Height          =   255
         Left            =   4800
         TabIndex        =   42
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Doc"
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
         Height          =   255
         Left            =   3720
         TabIndex        =   41
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nacionalidad"
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
         Height          =   240
         Left            =   120
         TabIndex        =   40
         Top             =   1560
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edad"
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
         Height          =   240
         Left            =   2760
         TabIndex        =   39
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capacitación"
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
         Height          =   240
         Left            =   120
         TabIndex        =   38
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Asistente"
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
         Height          =   255
         Left            =   3720
         TabIndex        =   37
         Top             =   2160
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00662200&
      Caption         =   "Curso"
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
      Height          =   2175
      Left            =   6720
      TabIndex        =   30
      Top             =   840
      Width           =   2895
      Begin VB.CheckBox chkManuales 
         BackColor       =   &H00662200&
         Caption         =   "Manuales"
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
         Height          =   255
         Left            =   1440
         TabIndex        =   23
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CheckBox chkExamenes 
         BackColor       =   &H00662200&
         Caption         =   "Exámenes"
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
         Height          =   315
         Left            =   1440
         TabIndex        =   22
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtTotalCurso 
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
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtTotalCuotas 
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
         TabIndex        =   18
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtGastoAdm 
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
         TabIndex        =   19
         Top             =   1680
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpFechaSuscripcion 
         Height          =   360
         Left            =   1440
         TabIndex        =   20
         Top             =   480
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
         Format          =   89194497
         CurrentDate     =   41308
      End
      Begin MSComCtl2.DTPicker DTPFechaVerificacion 
         Height          =   360
         Left            =   1440
         TabIndex        =   21
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
         Format          =   89194497
         CurrentDate     =   41308
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Verif."
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
         Height          =   480
         Left            =   1440
         TabIndex        =   47
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Curso"
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
         Height          =   240
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label cuotas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cuotas"
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
         Height          =   240
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Width           =   705
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Gasto Adm."
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
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Susc."
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
         Height          =   240
         Left            =   1440
         TabIndex        =   31
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00662200&
      Caption         =   "Código"
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
      Height          =   735
      Left            =   6720
      TabIndex        =   50
      Top             =   120
      Width           =   2895
      Begin VB.Label lblCodAlumno 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1440
         TabIndex        =   51
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00552233&
      Caption         =   "Verificaciones"
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
      Height          =   4845
      Left            =   9720
      TabIndex        =   29
      Top             =   120
      Width           =   1575
      Begin isButtonTest.isButton cmdVerificar 
         Height          =   420
         Left            =   120
         TabIndex        =   53
         Top             =   600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmVerificaciones.frx":AC85
         Style           =   8
         Caption         =   "       Verificar"
         IconSize        =   18
         IconAlign       =   1
         CaptionAlign    =   1
         iNonThemeStyle  =   7
         ShowFocus       =   -1  'True
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
         TabIndex        =   54
         Top             =   1200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmVerificaciones.frx":B55F
         Style           =   8
         Caption         =   "       Editar"
         IconSize        =   18
         IconAlign       =   1
         CaptionAlign    =   1
         iNonThemeStyle  =   7
         ShowFocus       =   -1  'True
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
      Begin isButtonTest.isButton cmdBuscar 
         Height          =   420
         Left            =   120
         TabIndex        =   55
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmVerificaciones.frx":BE39
         Style           =   8
         Caption         =   "       Buscar"
         IconSize        =   18
         IconAlign       =   1
         CaptionAlign    =   1
         iNonThemeStyle  =   7
         ShowFocus       =   -1  'True
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
      Begin isButtonTest.isButton cmdPlanDePago 
         Height          =   420
         Left            =   120
         TabIndex        =   56
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmVerificaciones.frx":C713
         Style           =   8
         Caption         =   "       Plan Pago"
         IconSize        =   18
         IconAlign       =   1
         CaptionAlign    =   1
         iNonThemeStyle  =   7
         ShowFocus       =   -1  'True
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
      Begin isButtonTest.isButton cmdGrabar 
         Height          =   420
         Left            =   120
         TabIndex        =   25
         Top             =   3000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmVerificaciones.frx":CFED
         Style           =   8
         Caption         =   "       Aceptar"
         IconSize        =   18
         IconAlign       =   1
         CaptionAlign    =   1
         iNonThemeStyle  =   7
         ShowFocus       =   -1  'True
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
         TabIndex        =   26
         Top             =   3600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmVerificaciones.frx":D8C7
         Style           =   8
         Caption         =   "       Cancelar"
         IconSize        =   18
         IconAlign       =   1
         CaptionAlign    =   1
         iNonThemeStyle  =   7
         ShowFocus       =   -1  'True
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
         TabIndex        =   57
         Top             =   4200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmVerificaciones.frx":E1A1
         Style           =   8
         Caption         =   "       Volver"
         IconSize        =   18
         IconAlign       =   1
         CaptionAlign    =   1
         iNonThemeStyle  =   7
         ShowFocus       =   -1  'True
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
   Begin VB.Label Label20 
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
      Height          =   375
      Left            =   4920
      TabIndex        =   48
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmVerificaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
    frmBuscarVerificacion.Show
    Me.Enabled = False
    Analisis = False
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
    Dim codigo As Long

    If txtNya.Text = "" Then MsgBox "Debe ingresar un Nombre de Alumno", vbOKOnly + vbInformation, "Suscripciones": txtNya.SetFocus: Exit Sub
    If cmbTipoDoc.Text = "" Then MsgBox "Debe ingresar un Tipo de Documento", vbOKOnly + vbInformation, "Suscripciones": cmbTipoDoc.SetFocus: Exit Sub
    If txtDocumento.Text = "" Then MsgBox "Debe ingresar un Número de Documento", vbOKOnly + vbInformation, "Suscripciones": txtDocumento.SetFocus: Exit Sub
    If txtDireccion.Text = "" Then MsgBox "Debe ingresar una Dirección", vbOKOnly + vbInformation, "Suscripciones": txtDireccion.SetFocus: Exit Sub
    If txtCP.Text = "" Then MsgBox "Debe ingresar un Código Postal", vbOKOnly + vbInformation, "Suscripciones": txtCP.SetFocus: Exit Sub
    If dtcLocalidad.Text = "" Then MsgBox "Debe ingresar una Localidad", vbOKOnly + vbInformation, "Suscripciones": txtLocalidad.SetFocus: Exit Sub
    If txtNacionalidad.Text = "" Then MsgBox "Debe ingresar una Nacionalidad del Alumno", vbOKOnly + vbInformation, "Suscripciones": txtNacionalidad.SetFocus: Exit Sub
    If txtEdad.Text = "" Then MsgBox "Debe ingresar la Edad del Alumno", vbOKOnly + vbInformation, "Suscripciones": txtEdad.SetFocus: Exit Sub
    If dtcCapacitacion.Text = "" Then MsgBox "Debe ingresar una Capacitación ", vbOKOnly + vbInformation, "Suscripciones": dtcCapacitacion.SetFocus: Exit Sub
    If dtcAsistente.Text = "" Then MsgBox "Debe ingresar un Asistente", vbOKOnly + vbInformation, "Suscripciones": dtcAsistente.SetFocus: Exit Sub
    If txtPT1.Text = "" Then MsgBox "Debe ingresar al menos un Teléfono", vbOKOnly + vbInformation, "Suscripciones": txtPT1.SetFocus: Exit Sub
    If txtTel1.Text = "" Then MsgBox "Debe ingresar al menos un Teléfono", vbOKOnly + vbInformation, "Suscripciones": txtTel1.SetFocus: Exit Sub
    If txtTotalCurso.Text = "" Or txtTotalCurso.Text = "0" Then MsgBox "Debe ingresar el Precio del Curso." & vbNewLine & "El mismo debe ser superior a Cero", vbOKOnly + vbInformation, "Suscripciones": txtTotalCurso.SetFocus: Exit Sub
    If txtTotalCuotas.Text = "" Or txtTotalCuotas.Text = "0" Then MsgBox "Debe ingresar la Cantidad de Cuotas." & vbNewLine & "La misma debe ser superior a Cero", vbOKOnly + vbInformation, "Suscripciones": txtTotalCuotas.SetFocus: Exit Sub
    If txtGastoAdm.Text = "" Then MsgBox "Debe ingresar el Gasto Administrativo", vbOKOnly + vbInformation, "Suscripciones": txtGastoAdm.SetFocus: Exit Sub

    On Error GoTo LineaError

    If Modi = False Then
        With rsControl
            .Find "ID=1"
            codigo = !CodAlumno
            lblCodAlumno.Caption = codigo
            !CodAlumno = codigo + 1
            .UpdateBatch
        End With
        
        With rsVerificaciones
            .Requery
            .AddNew
            !CodAlumno = lblCodAlumno.Caption
            !NyA = txtNya.Text
            !tipodoc = cmbTipoDoc.Text
            !dni = txtDocumento.Text
            !direccion = txtDireccion.Text
            !cp = txtCP.Text
            !localidad = dtcLocalidad.Text
            !nacionalidad = txtNacionalidad.Text
            !edad = txtEdad.Text
            !fechanac = dtpFechaNacimiento.Value
            !capac = dtcCapacitacion.Text
            !Asistente = dtcAsistente.Text
            !ptel1 = txtPT1.Text
            !ptel2 = txtPT2.Text
            !ptel3 = txtPT3.Text
            !ptel4 = txtPT4.Text
            !tel1 = txtTel1.Text
            !tel2 = txtTel2.Text
            !tel3 = txtTel3.Text
            !tel4 = txtTel4.Text
            !totalcurso = Int(txtTotalCurso.Text)
            !cuotas = Int(txtTotalCuotas.Text)
            !gastoadm = Int(txtGastoAdm.Text)
            !fechasus = dtpFechaSuscripcion.Value
            !FechaVerif = DTPFechaVerificacion.Value
            !observaciones = txtObservaciones.Text
            !estado = "Activo"
            If chkManuales.Value = 1 Then
                !manuales = True
            Else
                !manuales = False
            End If
            If chkExamenes.Value = 1 Then
                !dchoexamen = True
            Else
                !dchoexamen = False
            End If
            .Update
            .Requery
        End With
       
       With rsInformeSuscripciones
            If .State = 1 Then .Close
            .Open "SELECT * FROM informesuscripciones WHERE curso='" & dtcCapacitacion.Text & "' and asistente='" & dtcAsistente.Text & "' and totalcurso=" & txtTotalCurso.Text & " and verificado=0", Cn, adOpenDynamic, adLockPessimistic
            If .BOF Or .EOF Then MsgBox "El Alumno fue verificado con Exito" & vbNewLine & vbNewLine & "Recuerde que para una Correcta Gestion Administrativa debera asignarle un plan de pago, incluso si la capacitacion estuviese Completamente Becada", vbOKOnly + vbInformation: GoSub continuar
            .MoveFirst
            !fechaV = DTPFechaVerificacion.Value
            !verificado = 1
            .UpdateBatch

        End With

    Else
        With rsVerificaciones
            .Requery
            .Find "Codalumno=" & lblCodAlumno.Caption
            .Find "CodAlumno='" & lblCodAlumno.Caption & "'"
            !NyA = txtNya.Text
            !tipodoc = cmbTipoDoc.Text
            !dni = txtDocumento.Text
            !direccion = txtDireccion.Text
            !cp = txtCP.Text
            !localidad = dtcLocalidad.Text
            !nacionalidad = txtNacionalidad.Text
            !edad = txtEdad.Text
            !fechanac = dtpFechaNacimiento.Value
            !capac = dtcCapacitacion.Text
            !Asistente = dtcAsistente.Text
            !ptel1 = txtPT1.Text
            !ptel2 = txtPT2.Text
            !ptel3 = txtPT3.Text
            !ptel4 = txtPT4.Text
            !tel1 = txtTel1.Text
            !tel2 = txtTel2.Text
            !tel3 = txtTel3.Text
            !tel4 = txtTel4.Text
            !totalcurso = Int(txtTotalCurso.Text)
            !cuotas = Int(txtTotalCuotas.Text)
            !gastoadm = Int(txtGastoAdm.Text)
            !fechasus = dtpFechaSuscripcion.Value
            !observaciones = txtObservaciones.Text
            !FechaVerif = DTPFechaVerificacion.Value
            !manuales = chkManuales.Value
            !dchoexamen = chkExamenes.Value
            .UpdateBatch
            .Requery
            
            '''actualiza nombre en tabla de plan de pago
            With rsAnalisisDeCuenta
                If .State = 1 Then .Close
                .Open "SELECT * FROM PlanDePago WHERE codalumno=" & Int(lblCodAlumno.Caption), Cn, adOpenDynamic, adLockPessimistic
                If .BOF Or .EOF Then MsgBox "Los datos fueron actualizados con Exito" & vbNewLine & vbNewLine & "Recuerde que para una Correcta Gestion Administrativa debera asignarle un plan de pago, incluso si la capacitacion estuviese Completamente Becada", vbOKOnly + vbInformation: GoSub continuar
                .MoveFirst
                !NyA = txtNya.Text
                .UpdateBatch
                .Requery
            End With
        End With
    End If

continuar:
    If Trim(Len(lblCodAlumno.Caption)) = 1 Then lblCodAlumno.Caption = Format(lblCodAlumno.Caption, "0000#")
    If Trim(Len(lblCodAlumno.Caption)) = 2 Then lblCodAlumno.Caption = Format(lblCodAlumno.Caption, "000##")
    If Trim(Len(lblCodAlumno.Caption)) = 3 Then lblCodAlumno.Caption = Format(lblCodAlumno.Caption, "00###")
    If Trim(Len(lblCodAlumno.Caption)) = 4 Then lblCodAlumno.Caption = Format(lblCodAlumno.Caption, "0####")

    HabilitarBotones True, False
    HabilitarCuadros True, False
    
LineaError:
    Select Case Err.Number
        Case 3021
            Resume Next
        End Select
End Sub

Private Sub cmdModificar_Click()
    If lblCodAlumno.Caption = "" Then
        MsgBox "Primero debe Buscar a un Alumno Verificado", vbInformation + vbOKOnly, "Gestion Integral del Alumno"
    Else
        HabilitarCuadros False, True
        HabilitarBotones False, True
        txtNya.SetFocus
        Modi = True
    End If
End Sub

Private Sub cmdPlanDePago_Click()
    If lblCodAlumno.Caption = "" Then
        MsgBox "Primero debe Buscar un Alumno", vbOKOnly + vbInformation, "Verificaciones"
    Else
        PlanDePago
        With rsPlanDePago
            .Find "Codalumno='" & Val(lblCodAlumno.Caption) & "'"
            If .EOF Then
                frmPlanDePagos.Show
                If txtTotalCuotas.Text = 1 Then
                    frmPlanDePagos.txtCuotaDos.Visible = False
                    frmPlanDePagos.dtpVtoDos.Visible = False
                Else
                    frmPlanDePagos.txtCuotaDos.Visible = True
                    frmPlanDePagos.dtpVtoDos.Visible = True
                End If
                Me.Enabled = False
            Else
                MsgBox "El Alumno ya tiene asignado un plan de pago", vbOKOnly + vbInformation, "Verificaciones"
            End If
        End With
    End If
End Sub

Private Sub cmdVerificar_Click()
    frmBuscarSuscripcion.Show
    Verificar = True
    Me.Enabled = False
    Modi = False
End Sub
Private Sub dtcAsistente_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub dtcCapacitacion_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub dtcLocalidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNacionalidad.SetFocus
        With rsLocalidades
            .Find "localidad='" & dtcLocalidad.Text & "'"
            txtCP.Text = !cp
        End With
    End If
End Sub

Private Sub dtpFechaNacimiento_Change()
    If Month(dtpFechaNacimiento.Value) < Month(Date) Then
        txtEdad.Text = DateDiff("yyyy", dtpFechaNacimiento.Value, Date)
    ElseIf Day(dtpFechaNacimiento.Value) <= Day(Date) And Month(dtpFechaNacimiento.Value) = Month(Date) Then
        txtEdad.Text = DateDiff("yyyy", dtpFechaNacimiento.Value, Date)
    ElseIf Day(dtpFechaNacimiento.Value) > Day(Date) And Month(dtpFechaNacimiento.Value) >= Month(Date) Then
        txtEdad.Text = DateDiff("yyyy", dtpFechaNacimiento.Value, Date) - 1
    Else
      txtEdad.Text = DateDiff("yyyy", dtpFechaNacimiento.Value, Date) - 1
    End If

End Sub

Private Sub Form_Load()
    Centrar Me
    Verificaciones
    Localidades
    Control
    Capacitaciones
    Asistente
    Set dtcLocalidad.RowSource = rsLocalidades
    dtcLocalidad.BoundColumn = "localidad"
    dtcLocalidad.ListField = "localidad"
    
    Set dtcCapacitacion.RowSource = rsCapacitaciones
    dtcCapacitacion.BoundColumn = "capacitacion"
    dtcCapacitacion.ListField = "capacitacion"
    Set dtcAsistente.RowSource = rsPersonal
    dtcAsistente.BoundColumn = "Personal"
    dtcAsistente.ListField = "Personal"
    HabilitarBotones True, False
    HabilitarCuadros True, False
    Limpiar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Label20.Caption = "frmCuotasXFecha" Then
        frmCuotasXFecha.Enabled = True
    ElseIf Label20.Caption = "frmAnalisisDeCuotas" Then
        frmAnalisisDeCuotas.Enabled = True
        
        '''refresca datos en analisis de cuotas
        AnalisisDeCuota
        Set frmAnalisisDeCuotas.grilla1.DataSource = rsAnalisisDeCuenta
        frmAnalisisDeCuotas.formatoGrilla

    ElseIf Label20.Caption = "frmAnalisisSituacion" Then
        frmAnalisisSituacion.Enabled = True
    ElseIf Label20.Caption = "frmMarcas" Then
        frmMarcas.Enabled = True
    End If
End Sub

Sub HabilitarBotones(estado1 As Boolean, estado2 As Boolean)
    cmdVerificar.Enabled = estado1
    cmdModificar.Enabled = estado1
    cmdBuscar.Enabled = estado1
    cmdGrabar.Enabled = estado2
    cmdCancelar.Enabled = estado2
    cmdCerrar.Enabled = estado1
    cmdPlanDePago.Enabled = estado1
End Sub

Sub HabilitarCuadros(estado1 As Boolean, estado2 As Boolean)
    txtNya.Locked = estado1
    cmbTipoDoc.Locked = estado1
    txtDireccion.Locked = estado1
    txtEdad.Locked = estado1
    txtCP.Locked = estado1
    dtcLocalidad.Locked = estado1
    txtNacionalidad.Locked = estado1
    txtPT1.Locked = estado1
    txtPT2.Locked = estado1
    txtPT3.Locked = estado1
    txtPT4.Locked = estado1
    txtTel1.Locked = estado1
    txtTel2.Locked = estado1
    txtTel3.Locked = estado1
    txtTel4.Locked = estado1
    txtObservaciones.Locked = estado1
    txtGastoAdm.Locked = estado1
    txtTotalCuotas.Locked = estado1
    txtTotalCurso.Locked = estado1
    txtDocumento.Locked = estado1
    dtpFechaNacimiento.Enabled = estado2
    dtpFechaSuscripcion.Enabled = estado2
    DTPFechaVerificacion.Enabled = estado2
    dtcAsistente.Locked = estado1
    dtcCapacitacion.Locked = estado1
    chkManuales.Enabled = estado2
    chkExamenes.Enabled = estado2
End Sub

Sub Limpiar()
    lblCodAlumno.Caption = ""
    txtNya.Text = ""
    txtDireccion.Text = ""
    txtEdad.Text = ""
    txtCP.Text = ""
    
    txtNacionalidad.Text = ""
    txtPT1.Text = ""
    txtPT2.Text = ""
    txtPT3.Text = ""
    txtPT4.Text = ""
    txtTel1.Text = ""
    txtTel2.Text = ""
    txtTel3.Text = ""
    txtTel4.Text = ""
    txtObservaciones.Text = ""
    txtGastoAdm.Text = ""
    txtTotalCuotas.Text = ""
    txtTotalCurso.Text = ""
    txtDocumento.Text = ""
    dtpFechaNacimiento.Value = Date
    dtpFechaSuscripcion.Value = Date
    chkManuales.Value = 0
    chkExamenes.Value = 0
End Sub


Private Sub txtCP_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtDocumento_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
    If KeyAscii = 46 Then KeyAscii = 0
End Sub
Private Sub txtEdad_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtGastoAdm_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtLocalidad_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtNacionalidad_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtNya_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtPT1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtPT2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtPT3_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtPT4_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtTel1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtTel2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtTel3_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtTel4_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtTotalCuotas_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub txtTotalCurso_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chkExamenes_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub chkManuales_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
Private Sub cmbTipoDoc_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Public Sub Nex()
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
