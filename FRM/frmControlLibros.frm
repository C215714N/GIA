VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmControlLibros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Manuales"
   ClientHeight    =   3075
   ClientLeft      =   6030
   ClientTop       =   2235
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmControlLibros.frx":0000
   ScaleHeight     =   3075
   ScaleWidth      =   6405
   Begin VB.Frame Frame1 
      BackColor       =   &H00662200&
      Caption         =   "Agregar"
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
      Height          =   2775
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   1600
      Begin VB.TextBox txtPrecio 
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
         Width           =   1355
      End
      Begin VB.TextBox txtStock 
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
         TabIndex        =   3
         Top             =   1080
         Width           =   1355
      End
      Begin VB.TextBox txtManual 
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
         TabIndex        =   2
         Top             =   480
         Width           =   1355
      End
      Begin isButtonTest.isButton cmdAgregar 
         Height          =   420
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmControlLibros.frx":7A1D
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Precio"
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
         TabIndex        =   7
         Top             =   1440
         Width           =   1350
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Stock"
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
         TabIndex        =   6
         Top             =   840
         Width           =   1355
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Manual"
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
         TabIndex        =   5
         Top             =   240
         Width           =   1355
      End
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   21
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
End
Attribute VB_Name = "frmControlLibros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Centrar Me
    With rsManuales
        If .State = 1 Then .Close
        .Open "SELECT Manual, Stock, Precio FROM manuales ORDER BY manual", Cn, adOpenDynamic, adLockPessimistic
    End With

    Set grilla.DataSource = rsManuales
    formatoGrilla
End Sub

Private Sub cmdAgregar_Click()
    ''' controla que esten todos los datos
    If txtManual.Text = "" Then MsgBox "Ingrese nombre del manual", vbCritical, "Control de Manuales": txtManual.SetFocus: Exit Sub
    If txtStock.Text = "" Then MsgBox "Ingrese cantidad de manuales", vbCritical, "Control de Manuales": txtCantidad.SetFocus: Exit Sub
    If txtPrecio.Text = "" Then MsgBox "Ingrese precio del manual", vbCritical, "Control de Manuales": txtPrecio.SetFocus: Exit Sub
    
    With rsManuales
        .Requery
        .Find "manual='" & txtManual.Text & "'"
        If .EOF Or .BOF Then
            .AddNew
            !manual = txtManual.Text
            !stock = Int(txtStock.Text)
            !precio = txtPrecio.Text
            .Update
        Else
            !stock = Int(txtStock.Text)
            !precio = txtPrecio.Text
            .UpdateBatch
        End If
        .Close
        .Open "SELECT Manual, Stock, Precio FROM manuales ORDER BY manual", Cn, adOpenDynamic, adLockPessimistic
        Set grilla.DataSource = rsManuales
        formatoGrilla
    End With
    '''Restablece los Datos
    txtManual.Text = ""
    txtStock.Text = ""
    txtPrecio.Text = ""
End Sub

Private Sub grilla_Click()
    txtManual.Text = grilla.Columns(0).Text
    txtStock.Text = grilla.Columns(1).Text
    txtPrecio.Text = grilla.Columns(2).Text
End Sub

Sub formatoGrilla()
    Dim w As Integer
    For N = 0 To 2
        If N = 0 Then
            w = 2100
        Else:
            w = 800
            grilla.Columns(N).Alignment = dbgCenter
            If N = 2 Then
                grilla.Columns(N).NumberFormat = "$ #####"
            End If
        End If
        grilla.Columns(N).Width = w
    Next
End Sub

