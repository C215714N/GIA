VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmSituacionDeCartera 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Situación de Cartera"
   ClientHeight    =   5445
   ClientLeft      =   345
   ClientTop       =   2505
   ClientWidth     =   8310
   Icon            =   "frmSituacionDeCartera.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmSituacionDeCartera.frx":324A
   ScaleHeight     =   5445
   ScaleWidth      =   8310
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   360
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
      CurrentDate     =   41624
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00552233&
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
      Height          =   4935
      Left            =   6600
      TabIndex        =   7
      Top             =   360
      Width           =   1575
      Begin VB.TextBox txtAlumnos 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   480
         Width           =   1355
      End
      Begin VB.TextBox txtDeuda 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1080
         Width           =   1355
      End
      Begin VB.TextBox txtCobranza 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1680
         Width           =   1355
      End
      Begin VB.TextBox txtResto 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   3480
         Width           =   1355
      End
      Begin VB.TextBox txtPorcentaje 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2880
         Width           =   1355
      End
      Begin VB.TextBox txtCobrado 
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2280
         Width           =   1355
      End
      Begin isButtonTest.isButton cmdAnalizar 
         Height          =   420
         Left            =   120
         TabIndex        =   3
         Top             =   3900
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmSituacionDeCartera.frx":AC67
         Style           =   8
         Caption         =   "       Analisis"
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
      Begin isButtonTest.isButton cmdInforme 
         Height          =   420
         Left            =   120
         TabIndex        =   4
         Top             =   4400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   741
         Icon            =   "frmSituacionDeCartera.frx":B541
         Style           =   8
         Caption         =   "       Informe"
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
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Resto"
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
         TabIndex        =   19
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje"
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
         TabIndex        =   18
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Cobrado"
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
         TabIndex        =   17
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cobranza"
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
         TabIndex        =   16
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Deuda"
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
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Alumnos"
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
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid grilla 
      Height          =   4455
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7858
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
   Begin isButtonTest.isButton cmdBuscar 
      Height          =   420
      Left            =   1560
      TabIndex        =   1
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmSituacionDeCartera.frx":BE1B
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
   Begin isButtonTest.isButton cmdCancelar 
      Height          =   420
      Left            =   3000
      TabIndex        =   2
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmSituacionDeCartera.frx":C6F5
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Situacion al Dia"
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
      TabIndex        =   5
      Top             =   120
      Width           =   1455
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSituacionDeCartera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim alumnos As Long
    Dim Cobranza As Single
    Dim resto As Single
    Dim totalcobrado As Single
    Dim deuda As Single


Private Sub cmdAnalizar_Click()
    Situacion = grilla.Columns(0).Text
    Me.Enabled = False
    frmAnalisisSituacion.Show
    frmAnalisisSituacion.txtResta.Text = grilla.Columns(6).Text
    frmAnalisisSituacion.txtResta.Text = Format(frmAnalisisSituacion.txtResta.Text, "currency")
End Sub


Private Sub cmdBuscar_Click()
    
''' analiza la situacion de cartera al dia de hoy
If DTPFecha.Value = Date Then
    
    alumnos = 0
    deuda = 0
    Cobranza = 0
    totalcobrado = 0
    resto = 0
    
    With rsSituacionDeCartera
        If .State = 1 Then .Close
        .Open "SELECT cantidadcuotas * 30 -30 , count(codalumno), sum(deuda), sum(cobrado), sum(pago), Round(sum(cobrado) * 100 / sum(deuda)), sum(deuda)-sum(cobrado) FROM marcas WHERE cantidadcuotas > 0 group by cantidadcuotas", Cn, adOpenDynamic, adLockPessimistic
        .MoveFirst
        
        Do Until .EOF
            alumnos = alumnos + !expr1001
            deuda = deuda + !expr1002
            Cobranza = Cobranza + !expr1003
            totalcobrado = totalcobrado + !expr1004
            resto = resto + !expr1006
            .MoveNext
        Loop
        txtAlumnos.Text = alumnos
        txtDeuda.Text = FormatCurrency(deuda)
        txtCobranza.Text = FormatCurrency(Cobranza)
        txtCobrado.Text = totalcobrado
        txtResto.Text = FormatCurrency(resto)
        txtPorcentaje.Text = FormatCurrency(txtCobranza.Text) * 100 / FormatCurrency(txtDeuda.Text) & "%"
        
        If txtPorcentaje.Text = "100%" Then
            txtPorcentaje.Text = "100%"
        Else
            txtPorcentaje.Text = Format(txtPorcentaje.Text, "##.##%")
        End If
    End With
    
    With rsSituacionDeCartera
        If .State = 1 Then .Close
        .Open "SELECT cantidadcuotas * 30 -30 as Dias, count(codalumno) as [Alumnos], sum(deuda) as Deuda, sum(cobrado) as [Cobranza], sum(pago) as [Cant], Round(sum(cobrado) * 100 / sum(deuda)) as [%], sum(deuda)-sum(cobrado) as [Resta] FROM marcas WHERE cantidadcuotas > 0 group by cantidadcuotas", Cn, adOpenDynamic, adLockPessimistic
    End With
    
    cmdAnalizar.Enabled = True
'''busca la situacion de cartera anterior
ElseIf DTPFecha.Value < Date Then
    With rsSituacionDeCartera
        If .State = 1 Then .Close
        .Open "SELECT Dias,Alumnos as [Alumnos],Deuda,Cobranza,Cobrado as [Cobrado],Porcentaje as [%],Resto as [Resta] FROM SituacionesDeCartera WHERE fecha=#" & Format(DTPFecha.Value, "mm/dd/yyyy") & "#", Cn, adOpenDynamic, adLockPessimistic
        If .BOF Or .EOF Then MsgBox "Ese día no se inició sesión", vbOKOnly, "Situación de Cartera - GIA": Exit Sub
    End With
    '''carga los totales
    With rsTotalesSituaciones
        If .State = 1 Then .Close
        .Open "SELECT * FROM totalessituaciones WHERE fecha=#" & Format(DTPFecha.Value, "mm/dd/yyyy") & "#", Cn, adOpenDynamic, adLockPessimistic
        txtAlumnos.Text = !alumnos
        txtDeuda.Text = !deuda
        txtCobranza.Text = !Cobranza
        txtCobrado.Text = !cobrado
        txtPorcentaje.Text = !porcentaje
        txtResto.Text = !resto
    End With
    cmdAnalizar.Enabled = False
End If
    
    '''muestra la situacion de cartera en la grilla
    Set grilla.DataSource = rsSituacionDeCartera
    formatoGrilla
End Sub

Private Sub formatoGrilla()
    Dim N As Integer
    For N = 0 To 6 Step 1
        grilla.Columns(N).Alignment = dbgCenter
        If N = 2 Or N = 3 Or N = 6 Then
                    grilla.Columns(N).Width = 1200
                    grilla.Columns(N).Alignment = dbgLeft
                    grilla.Columns(N).NumberFormat = "$   ######"
            ElseIf N = 0 Or N = 4 Or N = 5 Then
                    grilla.Columns(N).Width = 500
            Else:   grilla.Columns(1).Width = 800
        End If
    Next
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdInforme_Click()
    Set dtrSituacion.DataSource = rsSituacionDeCartera
    dtrSituacion.Show
    dtrSituacion.Caption = "Situación de Cartera"
    dtrSituacion.Sections("Sección5").Controls("etiqueta21").Caption = txtAlumnos.Text
    dtrSituacion.Sections("Sección5").Controls("etiqueta11").Caption = txtDeuda.Text
    dtrSituacion.Sections("Sección5").Controls("etiqueta12").Caption = txtCobranza.Text
    dtrSituacion.Sections("Sección5").Controls("etiqueta13").Caption = txtCobrado.Text
    dtrSituacion.Sections("Sección5").Controls("etiqueta14").Caption = txtPorcentaje.Text
    dtrSituacion.Sections("Sección5").Controls("etiqueta15").Caption = txtResto.Text
    dtrSituacion.Sections("Sección4").Controls("etiqueta25").Caption = DTPFecha.Value

End Sub

Private Sub DTPFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Centrar Me
    '''asigna valores 0 a las variables de los totales
    alumnos = 0
    deuda = 0
    Cobranza = 0
    totalcobrado = 0
    resto = 0
    DTPFecha.Value = Date
End Sub

