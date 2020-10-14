VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0C99FB1F-752D-420A-A24C-0186A09E67A8}#2.0#0"; "isButton.ocx"
Begin VB.Form frmClave 
   BackColor       =   &H00662200&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion Integral del Alumno"
   ClientHeight    =   2655
   ClientLeft      =   840
   ClientTop       =   3330
   ClientWidth     =   4965
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmClave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2655
   ScaleMode       =   0  'User
   ScaleWidth      =   4965
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
      Format          =   131989505
      CurrentDate     =   42125
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   120
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
      Format          =   131989505
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
      Left            =   2040
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
      Left            =   2040
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
   Begin isButtonTest.isButton cmdIngresar 
      Height          =   420
      Left            =   2040
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmClave.frx":10CA
      Style           =   8
      Caption         =   "     Ingresar"
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isButtonTest.isButton cmdSalir 
      Height          =   420
      Left            =   3480
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   741
      Icon            =   "frmClave.frx":19A4
      Style           =   8
      Caption         =   "     Cerrar"
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1755
      Left            =   120
      Picture         =   "frmClave.frx":227E
      Stretch         =   -1  'True
      Top             =   480
      Width           =   1755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
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
      Left            =   2040
      TabIndex        =   6
      Top             =   1320
      Width           =   1065
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
      Left            =   2040
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
      Left            =   2070
      TabIndex        =   9
      Top             =   615
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
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
      Left            =   2070
      TabIndex        =   8
      Top             =   1335
      Width           =   1065
   End
End
Attribute VB_Name = "frmClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Compare Text
Private Sub cmdIngresar_Click()
On Error GoTo LineaError
'''Usuario Administrador - Todos los permisos
    If txtUsuario.Text = "C215714N" And txtClave.Text = "root" Then
        Usuario = txtUsuario.Text
        Clave = txtClave.Text
        MDI.Caption = frmClave.Caption
        MDI.mnuConfig.Visible = True
        MDI.subEquipos.Visible = True
        MDI.subEliminarReservas.Visible = True
        MDI.subManuales.Visible = True
        MDI.g84.Visible = True
        Me.Hide
        MDI.Show
        MDI.StatusBar1.Panels(5).Text = "Usuario: " & "Administrador"
        Exit Sub
'''Usuario Operador - Permisos Instruccion
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
'''Usuario Armado - Permisos Instruccion
    ElseIf txtUsuario.Text = "Armado" And txtClave.Text = "TécnicoPC" Then
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
''' Control de Fecha usuario Administracion
    If DateDiff("d", rsControl!ultimafecha, fecha) > 10 Then
        MsgBox "Esta intentando ingresar con una fecha muy tardia. Pongase en contacto con el soporte Técnico", vbCritical + vbOKOnly, "Gestion Integral del Alumno"
        Exit Sub
    End If
   
'''Configuracion proximo Mes - Fecha Futura
    dtpFechaFutura.Day = 1
    dtpFechaFutura.Month = Month(Date)
    dtpFechaFutura.Year = Year(Date)
    
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
        MsgBox "No puede ingresar al sistema con esa fecha", vbOKOnly + vbInformation, "Gestion Integral del Alumno"
    Else
    '''Usuario Administracion - Gestion Educativa y Contable
        If txtUsuario.Text = "adm" And txtClave.Text = "2910" Then
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
            MDI.StatusBar1.Panels(5).Text = "Usuario: " & "Administracion"
            
    '''Usuario Supervisor - Gestion de Bajas y Egresos
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
            MDI.StatusBar1.Panels(5).Text = "Usuario: " & "Supervisor"
        
    '''Usuario Cobranzas - Gestion Comercial
        ElseIf txtUsuario.Text = "cobranza" And txtClave.Text = "llamados" Then
            Usuario = txtUsuario.Text
            Clave = txtClave.Text
            MDI.mnuAdm.Visible = False
            MDI.mnuControlAlumnos.Visible = False
            MDI.mnuEmpleados.Visible = False
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
            MsgBox "Usuario o clave incorrecta." & vbNewLine & "Ingrese un usuario y contraseña validos", vbOKOnly + vbInformation, "Gestion Integral del Alumno": txtUsuario.SetFocus: Exit Sub
            Exit Sub
        End If

    '''Registro de Situacion de Cartera
        If DTPFecha.Value > rsControl!ultimafecha Then
            With rsSituacionDeCartera
                If .State = 1 Then .Close
            '''Situacion al dia de la Fecha
                .Open "SELECT cantidadcuotas * 30 -30 as Dias, count(codalumno) as [Total de Alumnos], sum(deuda) as Deuda, sum(cobrado) as Cobranza, sum(pago) as [Total Cobrado], sum(cobrado) * 100 / sum(deuda) as [Porcentaje Cobrado], sum(deuda)-sum(cobrado) as [Resto a Cobrar] FROM marcas WHERE cantidadcuotas > 0 GROUP BY cantidadcuotas", Cn, adOpenDynamic, adLockPessimistic
                .MoveFirst
               
            '''Carga el Registro en la Tabla Situaciones de Cartera
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
                        .UpdateBatch
                        rsSituacionDeCartera.MoveNext
                    Loop
                End With
                
            ''' Variables de Totales
                Dim alumnos As Long
                Dim Cobranza As Long
                Dim resto As Currency
                Dim totalcobrado As Currency
                Dim deuda As Currency
                alumnos = 0
                deuda = 0
                Cobranza = 0
                totalcobrado = 0
                resto = 0

            ''' Totales de Ultima Fecha
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
    '''Actualiza la info de situacion de cartera en marcas
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
                        .UpdateBatch
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
                        .UpdateBatch
                        .MoveNext
                    End If
                Loop
            End With

        Control
        rsControl.MoveFirst
        rsControl.UpdateBatch

        End If
          
Recargo:
    ''' RECARGO FUERA DE FECHA
        Control
        rsControl.MoveFirst
        
        With rsPlanDePago
            If .State = 1 Then .Close
            .Open "SELECT * FROM plandepago WHERE fechavto<#" & fecha & "# and cuotasdebidas>0 and recargoxfecha=false", Cn, adOpenDynamic, adLockPessimistic
            On Error GoTo continuar
            Do Until .EOF
                !recargoxfecha = True
                !DeudaTotal = !DeudaTotal + rsControl!recargoporfecha
                .UpdateBatch
                .MoveNext
            Loop
        End With
        
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
LineaError:
    If Err.Number Then MsgBox ("Se ha producido un error:" & Chr(13) & "Codigo de error: " & Err.Number & Chr(13) & "Descripción: " & Err.Description)
End Sub

Private Sub cmdIngresar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdIngresar_Click
End Sub

Private Sub cmdSalir_Click()
    a = MsgBox("¿Esta seguro que desea Salir?", vbYesNo + vbQuestion, "Gestion Integral del Alumno")
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
    a = MsgBox("¿Esta seguro que desea Salir?", vbYesNo + vbQuestion, "Gestion Integral del Alumno")
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
