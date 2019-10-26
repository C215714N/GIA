VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmUltimasCuotas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Últimas Cuotas"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   Icon            =   "frmUltimasCuotas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmUltimasCuotas.frx":324A
   ScaleHeight     =   4215
   ScaleWidth      =   5790
   Begin isButtonTest.isButton cmdBuscar 
      Height          =   420
      Left            =   1560
      TabIndex        =   3
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmUltimasCuotas.frx":AC67
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
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   20
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
   Begin VB.Label lblDeudaTotal 
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
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Deuda Total"
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
      TabIndex        =   0
      Top             =   120
      Width           =   960
   End
End
Attribute VB_Name = "frmUltimasCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
    ''' suma el total de ultimas cuotas y lo carga en lbl
    With rsUltimasCuotas
        If .State = 1 Then .Close
        .Open "SELECT sum(Deuda) FROM verificaciones as v,marcas as m WHERE m.codalumno=v.codalumno and v.cuotas > 1 and v.cuotas=m.cuota and m.cantidadcuotas=1", Cn, adOpenDynamic, adLockPessimistic
    End With
    
    lblDeudaTotal.Caption = Format(rsUltimasCuotas!expr1000, "currency")
       
    ''' consulta ultimas cuotas dentro de un periodo determinado
    With rsUltimasCuotas
        If .State = 1 Then .Close
        .Open "SELECT m.codalumno as Codigo, v.nya as Alumno,m.deuda as Deuda FROM verificaciones as v,marcas as m WHERE m.codalumno=v.codalumno and v.cuotas > 1 and v.cuotas=m.cuota and m.cantidadcuotas=1 ORDER BY m.codalumno", Cn, adOpenDynamic, adLockPessimistic
    End With
       
    '''carga consulta en la grilla
    Set grilla.DataSource = rsUltimasCuotas
    formatoGrilla
End Sub

Sub formatoGrilla()
    Dim w As Integer
    For N = 0 To 2 Step 1
        If N = 1 Then
            w = 3400
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

Private Sub Form_Load()
    Centrar Me
End Sub
