VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.MDIForm MDI 
   BackColor       =   &H00662200&
   Caption         =   "Gestion Integral del Alumno"
   ClientHeight    =   8130
   ClientLeft      =   90
   ClientTop       =   495
   ClientWidth     =   12585
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDI.frx":324A
   WindowState     =   2  'Maximized
   Begin HookMenu.XpMenu XpMenu2 
      Left            =   240
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   50
      CheckBorderColor=   7021576
      SelMenuBorder   =   7021576
      SelMenuBackColor=   14073525
      SelMenuForeColor=   16646297
      SelCheckBackColor=   14134407
      MenuBorderColor =   6956042
      SeparatorColor  =   -2147483632
      MenuBackColor   =   14609903
      MenuForeColor   =   0
      CheckBackColor  =   15326939
      CheckForeColor  =   10027263
      DisabledMenuBorderColor=   -2147483632
      DisabledMenuBackColor=   15660791
      DisabledMenuForeColor=   -2147483631
      MenuBarBackColor=   15790320
      MenuPopupBackColor=   16777215
      ShortCutNormalColor=   0
      ShortCutSelectColor=   16646297
      ArrowNormalColor=   10027263
      ArrowSelectColor=   12484864
      ShadowColor     =   0
      Bmp:1           =   "MDI.frx":F8856
      Key:1           =   "#mnuAlumnos"
      Bmp:2           =   "MDI.frx":F95BE
      Key:2           =   "#subSuscripciones"
      Bmp:3           =   "MDI.frx":FA326
      Key:3           =   "#subVerificaciones"
      Bmp:4           =   "MDI.frx":FB08E
      Key:4           =   "#subCobranza"
      Bmp:5           =   "MDI.frx":FBDF6
      Key:5           =   "#subGestion"
      Bmp:6           =   "MDI.frx":FCB5E
      Key:6           =   "#SubSituacion"
      Bmp:7           =   "MDI.frx":FD8C6
      Key:7           =   "#subCuotasXFecha"
      Bmp:8           =   "MDI.frx":FE62E
      Key:8           =   "#SubMarcas"
      Bmp:9           =   "MDI.frx":FF396
      Key:9           =   "#subCuotas"
      Bmp:10          =   "MDI.frx":1000FE
      Key:10          =   "#subUltimasCuotas"
      Bmp:11          =   "MDI.frx":100E66
      Key:11          =   "#subInformes"
      Bmp:12          =   "MDI.frx":101BCE
      Key:12          =   "#subInformeSuscripciones"
      Bmp:13          =   "MDI.frx":102936
      Key:13          =   "#subInformesVerificaciones"
      Bmp:14          =   "MDI.frx":10369E
      Key:14          =   "#subComisiones"
      Bmp:15          =   "MDI.frx":104406
      Key:15          =   "#subBecaTotal"
      Bmp:16          =   "MDI.frx":10516E
      Key:16          =   "#subMatriculas"
      Bmp:17          =   "MDI.frx":105ED6
      Key:17          =   "#subEgresados"
      Bmp:18          =   "MDI.frx":106C3E
      Key:18          =   "#subInformeBajas"
      Bmp:19          =   "MDI.frx":1079A6
      Key:19          =   "#mnuLibro"
      Bmp:20          =   "MDI.frx":10870E
      Key:20          =   "#subGrupoArmado"
      Bmp:21          =   "MDI.frx":109476
      Key:21          =   "#subAdmGrupos"
      Bmp:22          =   "MDI.frx":10A1DE
      Key:22          =   "#subLibroDeAula"
      Bmp:23          =   "MDI.frx":10AF46
      Key:23          =   "#subCapacitacion"
      Bmp:24          =   "MDI.frx":10BCAE
      Key:24          =   "#subDerechosExamenes"
      Bmp:25          =   "MDI.frx":10CA16
      Key:25          =   "#subExamenes"
      Bmp:26          =   "MDI.frx":10D77E
      Key:26          =   "#subBuscarExamenes"
      Bmp:27          =   "MDI.frx":10E4E6
      Key:27          =   "#subDiplomas"
      Bmp:28          =   "MDI.frx":10F24E
      Key:28          =   "#subViaticos"
      Bmp:29          =   "MDI.frx":10FFB6
      Key:29          =   "#subContabilidad"
      Bmp:30          =   "MDI.frx":110D1E
      Key:30          =   "#subConsultarCtas"
      Bmp:31          =   "MDI.frx":111A86
      Key:31          =   "#subCuentas"
      Bmp:32          =   "MDI.frx":1127EE
      Key:32          =   "#subNuevoCheque"
      Bmp:33          =   "MDI.frx":113556
      Key:33          =   "#subConsultarCheques"
      Bmp:34          =   "MDI.frx":1142BE
      Key:34          =   "#subRestaurar"
      Bmp:35          =   "MDI.frx":115026
      Mask:35         =   1
      Key:35          =   "#subBackUp"
      Bmp:36          =   "MDI.frx":115878
      Key:36          =   "#subControl"
      Bmp:37          =   "MDI.frx":1165E0
      Key:37          =   "#subReingresos"
      Bmp:38          =   "MDI.frx":117348
      Key:38          =   "#subPersonal"
      Bmp:39          =   "MDI.frx":1180B0
      Key:39          =   "#subEquipos"
      Bmp:40          =   "MDI.frx":118E18
      Key:40          =   "#subPP"
      Bmp:41          =   "MDI.frx":119B80
      Key:41          =   "#SubPresupuesto"
      Bmp:42          =   "MDI.frx":11A8E8
      Key:42          =   "#subCopiarPresupuesto"
      Bmp:43          =   "MDI.frx":11B650
      Key:43          =   "#subManuales"
      Bmp:44          =   "MDI.frx":11C3B8
      Key:44          =   "#subCargos"
      Bmp:45          =   "MDI.frx":11D120
      Key:45          =   "#subStatus"
      Bmp:46          =   "MDI.frx":11DE88
      Key:46          =   "#subAuditoria"
      Bmp:47          =   "MDI.frx":11EBF0
      Key:47          =   "#subCopias"
      Bmp:48          =   "MDI.frx":11F958
      Key:48          =   "#subVentaManual"
      Bmp:49          =   "MDI.frx":1206C0
      Key:49          =   "#subReservas"
      Bmp:50          =   "MDI.frx":121428
      Key:50          =   "#subEliminarReservas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7755
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "29/6/2024"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "01:32"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuAlumnos 
      Caption         =   "Alumnos"
      Begin VB.Menu subSuscripciones 
         Caption         =   "Suscripciones"
         Shortcut        =   {F1}
      End
      Begin VB.Menu subVerificaciones 
         Caption         =   "Verificaciones"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuGestion 
      Caption         =   "Gestion Educativa"
      Begin VB.Menu subCobranza 
         Caption         =   "Co&branza"
         Shortcut        =   ^B
      End
      Begin VB.Menu g1 
         Caption         =   "-"
      End
      Begin VB.Menu subGestion 
         Caption         =   "Gestion"
         Begin VB.Menu SubSituacion 
            Caption         =   "&Situacion de Cartera"
            Shortcut        =   ^S
         End
         Begin VB.Menu subCuotasXFecha 
            Caption         =   "Cuotas Por &Fecha"
            Shortcut        =   ^F
         End
         Begin VB.Menu SubMarcas 
            Caption         =   "Marcas"
         End
         Begin VB.Menu subCuotas 
            Caption         =   "Cuo&tas"
            Shortcut        =   ^T
         End
         Begin VB.Menu subUltimasCuotas 
            Caption         =   "Ultimas Cuotas"
            Shortcut        =   ^U
         End
      End
      Begin VB.Menu subInformes 
         Caption         =   "Informes"
         Begin VB.Menu subInformeSuscripciones 
            Caption         =   "Suscripciones"
         End
         Begin VB.Menu subInformesVerificaciones 
            Caption         =   "Verificaciones"
         End
         Begin VB.Menu subComisiones 
            Caption         =   "Comisiones"
            Shortcut        =   ^P
         End
         Begin VB.Menu subBecaTotal 
            Caption         =   "Alumnos 100%"
         End
         Begin VB.Menu subMatriculas 
            Caption         =   "Matriculas"
         End
         Begin VB.Menu subEgresados 
            Caption         =   "Egresados"
            Shortcut        =   {F3}
         End
         Begin VB.Menu subInformeBajas 
            Caption         =   "Bajas"
            Shortcut        =   {F4}
         End
      End
   End
   Begin VB.Menu mnuControlAlumnos 
      Caption         =   "Control Estudiantil"
      Begin VB.Menu mnuLibro 
         Caption         =   "Libros de Aula de &Operador"
         Shortcut        =   ^O
      End
      Begin VB.Menu subGrupoArmado 
         Caption         =   "Grupos de Armado"
         Begin VB.Menu subAdmGrupos 
            Caption         =   "Administrar Grupos"
         End
         Begin VB.Menu subLibroDeAula 
            Caption         =   "Libros de Aula de &Armado"
            Shortcut        =   ^A
         End
      End
      Begin VB.Menu g56 
         Caption         =   "-"
      End
      Begin VB.Menu subCapacitacion 
         Caption         =   "Capacitacio&nes"
         Shortcut        =   ^N
      End
      Begin VB.Menu subDerechosExamenes 
         Caption         =   "&Derechos de Examen"
         Shortcut        =   ^D
      End
      Begin VB.Menu subExamenes 
         Caption         =   "&Examenes"
         Shortcut        =   ^E
      End
      Begin VB.Menu subBuscarExamenes 
         Caption         =   "Buscar Examenes"
      End
      Begin VB.Menu subDiplomas 
         Caption         =   "Diplomas Entregados"
      End
   End
   Begin VB.Menu mnuAdm 
      Caption         =   "Gestion Comercial"
      Begin VB.Menu subViaticos 
         Caption         =   "Viaticos"
         Shortcut        =   {F8}
      End
      Begin VB.Menu subContabilidad 
         Caption         =   "Contabilidad"
         Shortcut        =   {F9}
      End
      Begin VB.Menu g2 
         Caption         =   "-"
      End
      Begin VB.Menu subConsultarCtas 
         Caption         =   "Consultar Cuentas"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu subCuentas 
         Caption         =   "Mantenimiento de Cuentas"
         Shortcut        =   +{F3}
      End
      Begin VB.Menu g3 
         Caption         =   "-"
      End
      Begin VB.Menu subNuevoCheque 
         Caption         =   "Agregar Cheques"
      End
      Begin VB.Menu subConsultarCheques 
         Caption         =   "Consultar Che&ques"
         Shortcut        =   ^Q
      End
      Begin VB.Menu g4 
         Caption         =   "-"
      End
      Begin VB.Menu subPP 
         Caption         =   "Preparar Presupuesto"
      End
      Begin VB.Menu SubPresupuesto 
         Caption         =   "Presupuesto"
      End
      Begin VB.Menu subCopiarPresupuesto 
         Caption         =   "Copiar Presupuesto"
      End
      Begin VB.Menu g85 
         Caption         =   "-"
      End
      Begin VB.Menu subManuales 
         Caption         =   "Control de Manuales"
      End
      Begin VB.Menu subVentaManual 
         Caption         =   "Venta de &Manuales"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mnuReservas 
      Caption         =   "Turnos"
      Begin VB.Menu subReservas 
         Caption         =   "&Reservas"
         Shortcut        =   ^R
      End
      Begin VB.Menu g84 
         Caption         =   "-"
      End
      Begin VB.Menu subEliminarReservas 
         Caption         =   "Eliminar Reservas"
         Shortcut        =   +{DEL}
      End
      Begin VB.Menu subEquipos 
         Caption         =   "Equipos"
      End
   End
   Begin VB.Menu mnuEmpleados 
      Caption         =   "Empleados"
      Begin VB.Menu subPersonal 
         Caption         =   "Persona&l"
         Shortcut        =   ^L
      End
      Begin VB.Menu subCargos 
         Caption         =   "Cargos"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuraciones"
      Begin VB.Menu subControl 
         Caption         =   "Control"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu subReingresos 
         Caption         =   "Reingresos"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu subStatus 
         Caption         =   "Status de la Base"
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu subAuditoria 
         Caption         =   "Auditoria"
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu g8 
         Caption         =   "-"
      End
      Begin VB.Menu subCopias 
         Caption         =   "Copias de Seguridad"
         Begin VB.Menu subBackUp 
            Caption         =   "Realizar Copia de Seguridad"
            Shortcut        =   {F11}
         End
         Begin VB.Menu subRestaurar 
            Caption         =   "Restaurar Copia de Seguridad"
            Shortcut        =   {F12}
         End
      End
   End
   Begin VB.Menu mnuSesion 
      Caption         =   "Cerrar Sesion"
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "Salir"
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub MDIForm_Load()
    Centrar Me
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmClave.Show
    frmClave.Caption = MDI.Caption
    frmClave.txtClave = ""
    frmClave.txtUsuario.Text = ""
    frmClave.txtUsuario.SetFocus
End Sub

Private Sub mnuLibro_Click()
    frmLibroOperador.Show
End Sub

Private Sub mnuSalir_Click()
    a = MsgBox("¿Esta seguro que desea Salir?", vbYesNo + vbQuestion, "Gestion Integral del Alumno")
    If a = vbYes Then
        End
    End If
End Sub

Private Sub mnuSesion_Click()
    Unload Me
End Sub

Private Sub subAdmGrupos_Click()
    frmGruposArmado.Show
End Sub

Private Sub subAuditoria_Click()
    frmAuditoria.Show
End Sub

Private Sub subBackUp_Click()
'''    FrmBackup.Show
    Dim Origen As String
    Dim Destino As String
    Origen = "" & DB & ""
    Destino = "T:\CopiaBase.mdb"
    If MsgBox("¿Realizar Copia de Seguridad?", vbQuestion + vbYesNo, "Gestion Integral del Alumno") = vbYes Then
            Set Fs = CreateObject("Scripting.FileSystemObject")
            Fs.CopyFile Origen, Destino
            MsgBox "La Copia de Respaldo se Realizo Correctamente", vbInformation + vbOKOnly, "Gestion Integral del Alumno"
    End If
End Sub

Private Sub subBecaTotal_Click()
    frmBecaTotal.Show
End Sub

Private Sub subBuscarExamenes_Click()
    frmConsultaExamenes.Show
End Sub

Private Sub subCapacitacion_Click()
    frmCapacitaciones.Show
End Sub

Private Sub subCargos_Click()
    frmCargos.Show
End Sub

Private Sub subclientes_Click()
    frmClientes.Show
End Sub

Private Sub subCobranza_Click()
    frmBuscarCobranza.Show
End Sub

Private Sub subComisiones_Click()
    frmComisiones.Show
End Sub

Private Sub subConsultarCheques_Click()
    frmConsultarCheques.Show
End Sub

Private Sub subConsultarCtas_Click()
    frmConsultarCuentas.Show
End Sub

Private Sub subContabilidad_Click()
    frmContabilidad.Show
End Sub

Private Sub subControl_Click()
    frmControl.Show
End Sub

Private Sub subCopiarPresupuesto_Click()
    If MsgBox("¿Copiar el presupuesto actual para el mes Praximo?", vbQuestion + vbYesNo, "Copiar Presupuesto") = vbYes Then
          
    '''Carga el presupuesto del mes en curso
        With rsCuentasPresupuesto
            If .State = 1 Then .Close
            .Open "SELECT * FROM presupuesto WHERE a¿o=" & Year(Date) & "and mes='" & MonthName(Month(Date)) & "'", Cn, adOpenDynamic, adLockPessimistic
            .Requery
            .MoveFirst
        End With
        
    '''Abre la tabla presupuesto para agregar las cuentas del mes Praximo
        With rsPresupuesto
            If .State = 1 Then .Close
            .Open "SELECT * FROM presupuesto", Cn, adOpenDynamic, adLockPessimistic
            .Requery
            Do Until rsCuentasPresupuesto.EOF
                .AddNew
                !Cuenta = rsCuentasPresupuesto!Cuenta
                !deuda = rsCuentasPresupuesto!deuda
                !pagado = 0
                !saldo = rsCuentasPresupuesto!deuda
                !observaciones = ""
                
                If Month(Date) = 12 Then
                    !mes = "Enero"
                    !a¿o = Year(Date) + 1
                Else
                    !a¿o = Year(Date)
                    !mes = MonthName(Month(Date) + 1)
                End If
                .UpdateBatch
                rsCuentasPresupuesto.MoveNext
            Loop
        End With
    End If
    
    MsgBox "El presupuesto se ha creado correctamente", , "Copia de Presupuesto"
End Sub

Private Sub subCuentas_Click()
    frmCuentas.Show
End Sub

Private Sub subCuotas_Click()
    frmBuscarVerificacion.Show
    Analisis = True
End Sub

Private Sub subCuotasXFecha_Click()
    frmCuotasXFecha.Show
End Sub

Private Sub subDerechosExamenes_Click()
    frmDerechosExamenes.Show
End Sub

Private Sub subDiplomas_Click()
    frmDiplomasEntregados.Show
End Sub

Private Sub subEgresados_Click()
    frmEgresados.Show
End Sub

Private Sub subEliminarReservas_Click()
    frmEliminarReservas.Show
End Sub

Private Sub subEquipos_Click()
    frmEquipos.Show
End Sub

Private Sub subExamenes_Click()
    frmExamenes.Show
End Sub

Private Sub subInformeBajas_Click()
    frmInformeBajas.Show
End Sub

Private Sub subInformeSuscripciones_Click()
    frmInformeSuscripciones.Show
End Sub

Private Sub subInformesVerificaciones_Click()
    frmInformeVerificados.Show
End Sub

Private Sub subLibroDeAula_Click()
    frmLibroArmado.Show
End Sub

Private Sub subManuales_Click()
    frmControlLibros.Show
End Sub

Private Sub SubMarcas_Click()
    frmMarcas.Show
End Sub

Private Sub subMatriculas_Click()
    frmMatriculas.Show
End Sub

Private Sub subnuevaorden_Click()
    frmNuevaOrden.Show
End Sub

Private Sub subNuevoCheque_Click()
    frmNuevoCheque.Show
End Sub

Private Sub subordenes_Click()
    frmConsultarOrdenes.Show
End Sub

Private Sub subPersonal_Click()
    frmPersonal.Show
End Sub

Private Sub subPP_Click()
    frmPP.Show
End Sub

Private Sub SubPresupuesto_Click()
    frmPresupuesto.Show
End Sub

Private Sub subReingresos_Click()
    frmReingresos.Show
End Sub

Private Sub subReservas_Click()
    frmReservas.Show
End Sub

Private Sub subRestaurar_Click()
    Dim Origen As String
    Dim Destino As String
    Origen = "" & DB & ""
    Destino = "T:\CopiaBase.mdb"
    If MsgBox("¿Restaurar Copia de Seguridad?", vbQuestion + vbYesNo, "Gestion Integral del Alumno") = vbYes Then
            Set Fs = CreateObject("Scripting.FileSystemObject")
            Fs.CopyFile Destino, Origen
            MsgBox "La Restauracion se Realizo Correctamente", vbInformation + vbOKOnly, "Gestion Integral del Alumno"
    End If

End Sub

Private Sub SubSituacion_Click()
    frmSituacionDeCartera.Show
End Sub

Private Sub subStatus_Click()
    frmStatus.Show
End Sub

Private Sub subSuscripciones_Click()
    frmSuscripciones.Show
End Sub

Private Sub subUltimasCuotas_Click()
    frmUltimasCuotas.Show
End Sub

Private Sub subVentaManual_Click()
    frmVentaManuales.Show
End Sub

Private Sub subVerificaciones_Click()
    frmVerificaciones.Show
End Sub

Private Sub subViaticos_Click()
    frmViaticos.Show
End Sub
