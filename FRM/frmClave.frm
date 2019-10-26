VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmClave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión Integral del Alumno"
   ClientHeight    =   2610
   ClientLeft      =   840
   ClientTop       =   3330
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmClave.frx":324A
   ScaleHeight     =   2610
   ScaleMode       =   0  'User
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker dtpFechaFutura 
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   85524481
      CurrentDate     =   42125
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   120
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
      Format          =   85524481
      CurrentDate     =   41327
   End
   Begin VB.TextBox txtClave 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
   Begin isButtonTest.isButton cmdIngresar 
      Height          =   420
      Left            =   2400
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmClave.frx":182AA
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
   Begin isButtonTest.isButton cmdSalir 
      Height          =   420
      Left            =   3840
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmClave.frx":18B84
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTRASEÑA"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2475
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2475
      TabIndex        =   5
      Top             =   600
      Width           =   840
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USUARIO"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2430
      TabIndex        =   9
      Top             =   615
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTRASEÑA"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2430
      TabIndex        =   8
      Top             =   1335
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   -240
      Picture         =   "frmClave.frx":1945E
      Stretch         =   -1  'True
      Top             =   -360
      Width           =   3000
   End
End
Attribute VB_Name = "frmClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Compare Text
Private Sub cmdIngresar_Click()
    
    ''' ingreso con usuario administrador y veo todo
    If txtUsuario.Text = "Admin" And txtClave.Text = "C215714N" Then
        Usuario = txtUsuario.Text
        Clave = txtClave.Text
        MDI.Caption = frmClave.Caption
        MDI.mnuConfig.Visible = True
        MDI.subEquipos.Visible = True
        MDI.subEliminarReservas.Visible = True
        MDI.subManuales.Visible = True
        '''MDI.mnuserviciotecnico.Visible = True
        MDI.g84.Visible = True
        Me.Hide
        MDI.Show
        MDI.StatusBar1.Panels(5).Text = "Usuario: " & "Administrador"
        Exit Sub
        
    ElseIf txtUsuario.Text = "Operador" And txtClave.Text = "AulaPC" Then
        Usuario = txtUsuario.Text
        MDI.Caption = frmClave.Caption
        MDI.StatusBar1.Panels(5).Text = "Usuario: " & txtUsuario.Text
        MDI.mnuAlumnos.Visible = False
        MDI.subDerechosExamenes.Visible = False
        MDI.mnuGestion.Visible = False
        MDI.mnuEmpleados.Visible = False
        MDI.mnuAdm.Visible = False
        MDI.mnuConfig.Visible = False
        MDI.subDiplomas.Visible = False
        MDI.subEquipos.Visible = False
        MDI.subEliminarReservas.Visible = False
        MDI.g56.Visible = False
        MDI.subGrupoArmado.Visible = False
        MDI.subBuscarExamenes.Visible = False
        MDI.subExamenes.Visible = False
        MDI.g84.Visible = False
        MDI.Show
        Me.Hide
        Exit Sub
        
    ElseIf txtUsuario.Text = "Armado" And txtClave.Text = "TecnicoPC" Then
        Usuario = txtUsuario.Text
        MDI.Caption = frmClave.Caption
        MDI.StatusBar1.Panels(5).Text = "Usuario: " & txtUsuario.Text
        MDI.mnuAlumnos.Visible = False
        MDI.subDerechosExamenes.Visible = False
        MDI.mnuGestion.Visible = False
        MDI.mnuEmpleados.Visible = False
        MDI.mnuAdm.Visible = False
        MDI.mnuConfig.Visible = False
        MDI.subDiplomas.Visible = False
        MDI.subEquipos.Visible = False
        MDI.subEliminarReservas.Visible = False
        MDI.g56.Visible = False
        MDI.subGrupoArmado.Visible = True
        MDI.subBuscarExamenes.Visible = False
        MDI.subExamenes.Visible = False
        MDI.g84.Visible = False
        MDI.mnuReservas.Visible = False
        MDI.subLibroDeAula.Visible = True
        MDI.mnuLibro.Visible = False
        MDI.Show
        Me.Hide
        Exit Sub
    End If
        
    '''declaro variables para las fechas
    Dim fecha As Date
    Dim fechafutura As Date
    fecha = Format(DTPFecha.Value, "dd/mm/yyyy")
    
    Control
    ''' controla si la fecha esta muy adelantada
    If DateDiff("d", rsControl!ultimafecha, fecha) > 10 Then
        MsgBox "Está intentando ingresar con una fecha muy tardía. Póngase en contacto con el soporte técnico", vbCritical + vbOKOnly, "Gestión Integral del Alumno"
        Exit Sub
    End If

    
    '''asigno valores a variable fecha y configuro fecha futura para mes siguiente
    
    dtpFechaFutura.Day = 1
    dtpFechaFutura.Year = Year(Date)
    dtpFechaFutura.Month = Month(Date)
    
    If dtpFechaFutura.Month = 12 Then
        dtpFechaFutura.Month = 1
        dtpFechaFutura.Year = dtpFechaFutura.Year + 1
    Else
        dtpFechaFutura.Month = dtpFechaFutura.Month + 1
        dtpFechaFutura.Year = DTPFecha.Year
    End If
    
    '''aplico formato a las variables de fechas
    fechafutura = Format(dtpFechaFutura.Value, "mm/dd/yyyy")
    fecha = Format(DTPFecha.Value, "mm/dd/yyyy")

    '''controla que no se ingrese con fecha anterior a la ya ingresada
    If DTPFecha.Value < rsControl!ultimafecha Then
        MsgBox "No puede ingresar al sistema con esa fecha", vbOKOnly + vbInformation, "Gestión Integral del Alumno"
    Else
        '''ingreso con usuario de administracion
        If txtUsuario.Text = "adm" And txtClave.Text = "1950" Then
            Usuario = txtUsuario.Text
            Clave = txtClave.Text
            MDI.mnuConfig.Visible = False
            MDI.subAdmGrupos.Visible = False
            MDI.subEquipos.Visible = False
            MDI.subEliminarReservas.Visible = False
            MDI.g84.Visible = False
            MDI.subManuales.Visible = False
            MDI.mnuReservas.Visible = True
            MDI.g56.Visible = True
            MDI.subCopiarPresupuesto.Visible = False
            MDI.subGrupoArmado.Visible = True
            MDI.g4.Visible = False
            MDI.subPP.Visible = False
            MDI.SubPresupuesto.Visible = False
            
        '''ingreso con usuario general
        ElseIf txtUsuario.Text = "adm" And txtClave.Text = "SuperV" Then
            Usuario = txtUsuario.Text
            Clave = txtClave.Text
            MDI.mnuConfig.Visible = False
            MDI.subEquipos.Visible = False
            MDI.subEliminarReservas.Visible = False
            MDI.g84.Visible = False
            MDI.mnuLibro.Visible = True
            MDI.subGrupoArmado.Visible = True
            MDI.g56.Visible = True
            
        '''Usuario Cobranzas
        ElseIf txtUsuario.Text = "cobranza" And txtClave.Text = "llamados" Then
            Usuario = txtUsuario.Text
            Clave = txtClave.Text
            MDI.mnuAdm.Visible = False
            MDI.mnuControlAlumnos.Visible = False
            MDI.mnuEmpleados.Visible = False
            '''MDI.mnuserviciotecnico.Visible = False
            MDI.mnuAlumnos.Visible = False
            MDI.subInformes.Visible = False
            MDI.subCuotas.Visible = False
            MDI.subComisiones.Visible = False
            MDI.mnuLibro.Visible = False
            MDI.mnuConfig.Visible = False
            MDI.subCopiarPresupuesto.Visible = False
            MDI.subAdmGrupos.Visible = False
            MDI.subEquipos.Visible = False
            MDI.subEliminarReservas.Visible = False
            MDI.g84.Visible = False
            MDI.subManuales.Visible = False
            MDI.mnuReservas.Visible = True
            MDI.g56.Visible = True
            MDI.subGrupoArmado.Visible = True
            MDI.g4.Visible = False
            MDI.subPP.Visible = False
            MDI.SubPresupuesto.Visible = False
            MDI.subContabilidad.Visible = False
            MDI.g2.Visible = False
       
       '''error de ingreso
        Else
            MsgBox "Usuario o clave incorrecta." & vbNewLine & "Ingrese un usuario y contraseña válidos", vbOKOnly + vbInformation, "Gestión Integral del Alumno": txtUsuario.SetFocus: Exit Sub
            Exit Sub
        End If

        '''guarda situacion de cartera del ultimo dia ingresado
        If DTPFecha.Value > rsControl!ultimafecha Then
            With rsSituacionDeCartera
                If .State = 1 Then .Close
  
                '''consulta la situacion actual
                '.Open "SELECT cantidadcuotas * 30 -30 as Dias, count(codalumno) as [Total de Alumnos], sum(deuda) as Deuda, sum(cobrado) as Cobranza, sum(pago) as [Total Cobrado], round(sum(cobrado) * 100 / sum(deuda),2) as [Porcentaje Cobrado], sum(deuda)-sum(cobrado) as [Resto a Cobrar] FROM marcas WHERE cantidadcuotas > 0 group by cantidadcuotas", cn, adOpenDynamic, adLockPessimistic
                .Open "SELECT cantidadcuotas * 30 -30 as Dias, count(codalumno) as [Total de Alumnos], sum(deuda) as Deuda, sum(cobrado) as Cobranza, sum(pago) as [Total Cobrado], sum(cobrado) * 100 / sum(deuda) as [Porcentaje Cobrado], sum(deuda)-sum(cobrado) as [Resto a Cobrar] FROM marcas WHERE cantidadcuotas > 0 group by cantidadcuotas", Cn, adOpenDynamic, adLockPessimistic
                
                .MoveFirst
                
                '''agrega la info en situaciones de cartera
                With rsSituacionesDeCartera
                    If .State = 1 Then .Close
                    .Open "SELECT * FROM SituacionesDeCartera", Cn, adOpenDynamic, adLockPessimistic
                    Do Until rsSituacionDeCartera.EOF
                        .Requery
                        .AddNew
                        !fecha = rsControl!ultimafecha
                        !dias = rsSituacionDeCartera!dias
                        !deuda = rsSituacionDeCartera!deuda
                        !cobrado = rsSituacionDeCartera![Total Cobrado]
                        !alumnos = rsSituacionDeCartera![Total de Alumnos]
                        !Cobranza = rsSituacionDeCartera!Cobranza
                        !porcentaje = rsSituacionDeCartera![Porcentaje Cobrado]
                        !resto = rsSituacionDeCartera![Resto a Cobrar]
                        .Update
                        rsSituacionDeCartera.MoveNext
                    Loop
                End With
                
                ''' declara variables para los totales
                Dim alumnos As Long
                Dim Cobranza As Single
                Dim resto As Single
                Dim totalcobrado As Single
                Dim deuda As Single
                alumnos = 0
                deuda = 0
                Cobranza = 0
                totalcobrado = 0
                resto = 0

                '''calcula totales a la ultima fecha
                .Close
                .Open "SELECT cantidadcuotas * 30 -30 , COUNT(codalumno), SUM(deuda), SUM(cobrado), SUM(pago), SUM(cobrado) * 100 / SUM(deuda), SUM(deuda)-sum(cobrado) FROM marcas WHERE cantidadcuotas > 0 group by cantidadcuotas", Cn, adOpenDynamic, adLockPessimistic
                .MoveFirst
        
                Do Until .EOF
                    alumnos = alumnos + !expr1001
                    deuda = deuda + !expr1002
                    Cobranza = Cobranza + !expr1003
                    totalcobrado = totalcobrado + !expr1004
                    resto = resto + !expr1006
                    .MoveNext
                Loop
                
                '''agrega totales a la ultima fecha
                With rsTotalesSituaciones
                    If .State = 1 Then .Close
                    .Open "SELECT * FROM TotalesSituaciones", Cn, adOpenDynamic, adLockPessimistic
                    .Requery
                    .AddNew
                    !fecha = rsControl!ultimafecha
                    !alumnos = alumnos
                    !deuda = deuda
                    !Cobranza = Cobranza
                    !resto = resto
                    !cobrado = totalcobrado
                    !porcentaje = Cobranza * 100 / deuda
                    .Update
                End With
            End With
        End If
    
    '''CONTROL DE FECHA
            With rsControl
               .Close
               .Open "SELECT month(ultimafecha) FROM control", Cn, adOpenDynamic, adLockPessimistic
            End With
        If DTPFecha.Month <> rsControl!expr1000 Then
            Control
            Marcar
        '''Agrega alumnos del mes a situacion de cartera cuando inicia mes
            With rsAlumnosDelMes
                If .State = 1 Then .Close
                .Open "SELECT * FROM alumnosdelmes ORDER BY codalumno", Cn, adOpenDynamic, adLockPessimistic
                If .BOF Or .EOF Then GoTo continuar '''si no hay alumnos que agregar continua con la proxima accion
                .MoveFirst
                Do Until .EOF
                    rsMarcar.Requery
                    rsMarcar.AddNew
                    rsMarcar!CodAlumno = !CodAlumno
                    rsMarcar.Update
                    .Delete
                    .Requery
                Loop
            End With

continuar:
'''RECARGO FUERA DE MES
    Control
    rsControl.MoveFirst
    With rsPlanDePago
        If .State = 1 Then .Close
        .Open "SELECT * FROM plandepago WHERE fechavto<#" & fecha & "# and cuotasdebidas>0", Cn, adOpenDynamic, adLockPessimistic
        If .BOF Or .EOF Then GoTo SituacionDeCartera
        .MoveFirst
        Do Until .EOF
                !recargoxmes = True
                !DeudaTotal = !DeudaTotal + rsControl!recargopormes
                .UpdateBatch
                .MoveNext
        Loop
    End With
            
SituacionDeCartera:

            With rsControl
                .Close
                .Open "SELECT month(ultimafecha) FROM control", Cn, adOpenDynamic, adLockPessimistic
            End With

            '''actualiza la info de situacion de cartera en marcas
            ''' consulta cuotas debidas a la fecha
            With rsPlanDePago
                If .State = 1 Then .Close
                .Open "SELECT p.codalumno, MIN(p.nrocuota) as Cuota, SUM(p.Deudatotal) as Deuda, SUM(p.CuotasDebidas) as CuotasDebidas,  DATEDIFF('m',Min(p.fechavto),#" & fechafutura & "#) AS Meses, MAX(p.NroCuota) AS MaxCuota,max(V.Cuotas) AS UltimaCuota FROM plandepago as p, verificaciones as v WHERE v.codalumno=p.codalumno and p.cuotasdebidas > 0 and p.fechavto<#" & fechafutura & "# GROUP BY p.codalumno ORDER BY p.codalumno", Cn, adOpenDynamic, adLockPessimistic
                .MoveFirst
            End With
            
            With rsMarcar
                If .State = 1 Then .Close
                .Open "SELECT * FROM marcas ORDER BY codalumno", Cn, adOpenDynamic, adLockPessimistic
                .Requery
                .MoveFirst
                '''actualiza el alumno
                Do Until .EOF
                    If rsPlanDePago.EOF Then
                        !cuota = 0
                        !deuda = 0
                        !cantidadcuotas = 0
                        !cobrado = 0
                        !pago = 0
                        .Update
                        .MoveNext
                    ElseIf !CodAlumno = rsPlanDePago!CodAlumno Then
                        !cuota = rsPlanDePago!cuota
                        !deuda = rsPlanDePago!deuda
                        If rsPlanDePago!maxcuota = rsPlanDePago!ultimacuota And rsPlanDePago!Meses > rsPlanDePago!CuotasDebidas Then
                            !cantidadcuotas = rsPlanDePago!Meses
                        Else
                            !cantidadcuotas = rsPlanDePago!CuotasDebidas
                        End If
                        !cobrado = 0
                        !pago = 0
                        .UpdateBatch
                        .MoveNext
                        rsPlanDePago.MoveNext
                    Else
                        !cuota = 0
                        !deuda = 0
                        !cantidadcuotas = 0
                        !cobrado = 0
                        !pago = 0
                        .Update
                        .MoveNext
                    End If
                Loop
            End With

        Control
        rsControl.MoveFirst
        rsControl!bachiller = 0
        rsControl.UpdateBatch

        End If
          
Recargo:
        ''' aplicar recargo x fuera de fecha
        Control
        rsControl.MoveFirst
        
        With rsPlanDePago
            If .State = 1 Then .Close
            .Open "SELECT * FROM plandepago WHERE fechavto<#" & fecha & "# and cuotasdebidas>0 and recargoxfecha=false", Cn, adOpenDynamic, adLockPessimistic
            If .BOF Or .EOF Then GoTo bachi
            .MoveFirst
            Do Until .EOF
                !recargoxfecha = True
                !DeudaTotal = !DeudaTotal + rsControl!recargoporfecha
                .UpdateBatch
                .MoveNext
            Loop
        End With
        
        
bachi:
        ''' regargo para bachiller
      Control
          If rsControl!bachiller = 1 Then GoTo fecha
        
        If Day(Date) > 15 Then
            With rsPlanDePago
                If .State = 1 Then .Close
                .Open "SELECT * FROM plandepago WHERE cuotasdebidas=1 and day(fechavto)=8 and month(fechavto)=" & Month(Date) & " and year(fechavto)=" & Year(Date), Cn, adOpenDynamic, adLockPessimistic
                If .BOF Or .EOF Then GoTo fecha
                .MoveFirst
                Do Until .EOF
                    !DeudaTotal = !DeudaTotal + rsControl!recargoporfecha
                    .UpdateBatch
                    .MoveNext
                Loop
                rsControl!bachiller = 1
                rsControl.UpdateBatch
            End With
        End If
fecha:
        '''modifica la ultima fecha en la tabla control
        Control
        With rsControl
            !ultimafecha = DTPFecha.Value
            .UpdateBatch
        End With
        
'''ACCESO FORMULARIO PRINCIPAL
        MDI.Caption = frmClave.Caption
        Me.Hide
        MDI.Show
        MDI.StatusBar1.Panels(5).Text = "Usuario: " & txtUsuario.Text
    End If
End Sub

Private Sub cmdIngresar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdIngresar_Click
End Sub

Private Sub cmdSalir_Click()
    a = MsgBox("¿Está seguro que desea Salir?", vbYesNo + vbQuestion, "Gestion Integral del Alumno")
    If a = vbYes Then
        End
    End If
End Sub

Private Sub DTPFecha_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
    Centrar Me
    Control
    DTPFecha.Value = Date
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    a = MsgBox("¿Está seguro que desea Salir?", vbYesNo + vbQuestion, "Gestion Integral del Alumno")
    If a = vbNo Then
        Cancel = True
    End If
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then cmdIngresar_Click
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub
