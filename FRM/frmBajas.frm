VERSION 5.00
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmBajas 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Bajas"
   ClientHeight    =   1440
   ClientLeft      =   5475
   ClientTop       =   3645
   ClientWidth     =   3990
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmBajas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmBajas.frx":324A
   ScaleHeight     =   1440
   ScaleWidth      =   3990
   Begin VB.ComboBox cmbPagoBaja 
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
      ItemData        =   "frmBajas.frx":11DFF
      Left            =   2500
      List            =   "frmBajas.frx":11E09
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtmotivo 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin isButtonTest.isButton cmdConfirmar 
      Height          =   420
      Left            =   2500
      TabIndex        =   3
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmBajas.frx":11E15
      Style           =   8
      Caption         =   "       Dar Baja"
      IconAlign       =   1
      CaptionAlign    =   1
      iNonThemeStyle  =   0
      HighlightColor  =   16744576
      FontHighlightColor=   12632256
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
      Caption         =   "Pago de Baja"
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
      Left            =   2500
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmBajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConfirmar_Click()
If txtmotivo.Text = "" Then MsgBox "Debera escribir el motivo de la baja": txtmotivo.SetFocus: Exit Sub
If cmbPagoBaja.Text = "" Then MsgBox "Defina si el alumno pagó la baja": cmbPagoBaja.SetFocus: Exit Sub

With rsMarcar
    If .State = 1 Then .Close
    .Open "SELECT * FROM marcas WHERE codalumno=" & frmAnalisisDeCuotas.lblCodAlumno.Caption, Cn, adOpenDynamic, adLockPessimistic
    .Requery
    .MoveFirst
End With

With rsBajas
    If .State = 1 Then .Close
    .Open "SELECT * FROM bajas", Cn, adOpenDynamic, adLockPessimistic
    .AddNew
    !motivo = txtmotivo.Text
    !fecha = Date
    !CodAlumno = frmAnalisisDeCuotas.lblCodAlumno.Caption
    !pagobaja = cmbPagoBaja.Text
    !sitcartera = rsMarcar!cantidadcuotas * 30 - 30
    !NroCuota = rsMarcar!cuota
    .Update
End With

With rsPlanDePago
            If .State = 1 Then .Close
            .Open "SELECT * FROM plandepago WHERE codalumno=" & CodAlumno, Cn, adOpenDynamic, adLockPessimistic
            .MoveFirst
            Do Until .EOF
                If !tipodepago = "PAG" Then
                    .MoveNext
                ElseIf !tipodepago = "Par" Then
                    .MoveNext
                Else
                    !tipodepago = "BAJA"
                    !fechapago = Date
                    !DeudaTotal = 0
                    !CuotasDebidas = 0
                    .UpdateBatch
                    .MoveNext
                End If
            Loop
        End With
        
        With rsVerificaciones
            If .State = 1 Then .Close
            .Open "SELECT codalumno, estado FROM verificaciones WHERE codalumno=" & frmAnalisisDeCuotas.lblCodAlumno.Caption, Cn, adOpenDynamic, adLockPessimistic
            .Requery
            .MoveFirst
            !estado = "Baja"
            .UpdateBatch
        End With

        frmBajas.Hide
        frmAnalisisDeCuotas.Enabled = True
        
End Sub

Private Sub Form_Load()
    Centrar Me
End Sub
