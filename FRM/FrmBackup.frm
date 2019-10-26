VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copia de Seguridad"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5430
   FillColor       =   &H80000005&
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmBackup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "FrmBackup.frx":9ED32
   ScaleHeight     =   3630
   ScaleWidth      =   5430
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   420
      Left            =   3600
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdNuevaCarpeta 
      Caption         =   "Crear Carpeta"
      Height          =   420
      Left            =   3600
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton cmdRestaurar 
      Caption         =   "Restaurar"
      Height          =   420
      Left            =   3600
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "Guardar"
      Height          =   420
      Left            =   3600
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   2520
      Left            =   120
      MouseIcon       =   "FrmBackup.frx":A674F
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Seleccionar Directorio de Destino"
      Top             =   960
      Width           =   3375
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H80000014&
      ForeColor       =   &H80000007&
      Height          =   360
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Seleccionar Unidad de Destino"
      Top             =   360
      Width           =   3375
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5520
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBackup.frx":A6A59
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBackup.frx":A6E16
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBackup.frx":A6F70
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBackup.frx":A70CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBackup.frx":A781B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBackup.frx":A7975
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBackup.frx":A7F0F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBackup.frx":A8069
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBackup.frx":A81C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBackup.frx":A82D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBackup.frx":A872A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tblBotones 
      Align           =   4  'Align Right
      Height          =   3630
      Left            =   5040
      TabIndex        =   5
      Top             =   0
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   6403
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copiar"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nueva Carpeta"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Auditoria"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Carpeta"
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
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   750
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad"
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
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2000
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccionar Disco de Destino"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccionar Directorio"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   2175
   End
End
Attribute VB_Name = "FrmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text

Private Sub Form_Load()
    Centrar Me
End Sub

Private Sub Drive1_Change()
    On Local Error GoTo LineaError
        Dir1.Path = Drive1.Drive
    Exit Sub

LineaError:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Sub cmdBackup_Click()
On Local Error GoTo LineaError
    Dim NuevaRuta As String
    NuevaRuta = IIf(Right(Dir1.Path, 1) = "\", Dir1.Path, Dir1.Path + "\") + "GIA-Backup-" & Date & ".mdb"
    Screen.MousePointer = vbHourglass
    'a = CopyFile("t:\Base.mdb", NuevaRuta, 0)
    Screen.MousePointer = vbDefault
    
    If a <> 0 Then
        MsgBox "La Copia de Seguridad Se Realizó Correctamente!" + Chr(13) + NuevaRuta, 64
    End If
  Exit Sub
  
LineaError:
      Screen.MousePointer = vbDefault
      MsgBox Err.Description, vbCritical
End Sub

Sub cmdNuevaCarpeta_Click()
On Local Error GoTo L
   Dim N As String
    N = InputBox("Ingrese el nombre de la Carpeta...!", "Crear Nueva Carpeta", , 2000, 1000)
    If InStr(1, N, "\") > 0 Then MsgBox "Prohibido Ingresar==>   \", 64: N = "": Exit Sub
    If N <> "" Then
      MkDir IIf(Right(Dir1.Path, 1) = "\", Dir1.Path, Dir1.Path & "\") + N
      Dir1.Refresh
    End If
  Exit Sub
L:
      MsgBox Err.Description, vbCritical
End Sub

Sub cmdSalir_Click()
    Unload Me
End Sub
