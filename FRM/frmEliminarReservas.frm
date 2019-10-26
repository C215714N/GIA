VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmEliminarReservas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eliminar Reservas"
   ClientHeight    =   1155
   ClientLeft      =   5685
   ClientTop       =   4275
   ClientWidth     =   3015
   Icon            =   "frmEliminarReservas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmEliminarReservas.frx":324A
   ScaleHeight     =   1155
   ScaleWidth      =   3015
   Begin MSComCtl2.DTPicker DTP 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
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
      Format          =   89260033
      CurrentDate     =   41037
   End
   Begin isButtonTest.isButton cmdEliminar 
      Height          =   420
      Left            =   1560
      TabIndex        =   2
      Top             =   100
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmEliminarReservas.frx":AC67
      Style           =   8
      Caption         =   "       Eliminar"
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
   Begin isButtonTest.isButton cmdSalir 
      Height          =   420
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmEliminarReservas.frx":B541
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Anteriores a:"
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
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1125
   End
End
Attribute VB_Name = "frmEliminarReservas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    'Elimina los datos anteriores

    Dim feha As Date
    fecha = Format(DTP.Value, "mm/dd/yyyy")

    If MsgBox("¿Desea eliminar la base de datos anterior al " & DTP.Value & "?", vbYesNo, "Reservas") = vbYes Then
        With rsEliminar
            If .State = 1 Then .Close
            .Open "SELECT * FROM reservas WHERE Fecha <#" & fecha & "#", Cn, adOpenDynamic, adLockPessimistic
            If .BOF Or .EOF Then Exit Sub
            .MoveFirst
            Do Until .EOF
                .Delete
                .UpdateBatch
                .MoveNext
            Loop
        End With
        MsgBox ("Se han Eliminado las reservas anteriores a " & DTP.Value)
    End If
End Sub
Private Sub Form_Load()
    Centrar Me
    DTP.Value = Date
End Sub

