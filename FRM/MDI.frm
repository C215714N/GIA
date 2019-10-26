VERSION 5.00
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.MDIForm MDI 
   BackColor       =   &H8000000C&
   Caption         =   "PROARTEC - Gestión Integral del Alumno"
   ClientHeight    =   8130
   ClientLeft      =   90
   ClientTop       =   495
   ClientWidth     =   12585
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDI.frx":324A
   WindowState     =   2  'Maximized
   Begin HookMenu.XpMenu XpMenu2 
      Left            =   960
      Top             =   5160
      _ExtentX        =   900
      _ExtentY        =   900
      BitmapSize      =   28
      BmpCount        =   52
      CheckBorderColor=   7021576
      SelMenuBorder   =   7021576
      SelMenuBackColor=   14073525
      SelMenuForeColor=   16646297
      SelCheckBackColor=   14791828
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
      Key:1           =   "#subSuscripciones"
      Bmp:2           =   "MDI.frx":F95BE
      Key:2           =   "#subVerificaciones"
      Bmp:3           =   "MDI.frx":FA326
      Key:3           =   "#subCobranza"
      Bmp:4           =   "MDI.frx":FB08E
      Key:4           =   "#subCuotasXFecha"
      Bmp:5           =   "MDI.frx":FBDF6
      Key:5           =   "#mnuLibro"
      Bmp:6           =   "MDI.frx":FCB5E
      Key:6           =   "#subGrupoArmado"
      Bmp:7           =   "MDI.frx":FD8C6
      Key:7           =   "#subInformeBajas"
      Bmp:8           =   "MDI.frx":FE62E
      Mask:8          =   -256
      Key:8           =   "#subRestaurar"
      Bmp:9           =   "MDI.frx":FEE80
      Mask:9          =   1
      Key:9           =   "#subBackUp"
      Bmp:10          =   "MDI.frx":FF6D2
      Mask:10         =   15466495
      Key:10          =   "#subEquipos"
      Bmp:11          =   "MDI.frx":FFF24
      Key:11          =   "#subCopias"
      Bmp:12          =   "MDI.frx":100776
      Key:12          =   "#subReingresos"
      Bmp:13          =   "MDI.frx":1014DE
      Key:13          =   "#mnuEmpleados"
      Bmp:14          =   "MDI.frx":102246
      Key:14          =   "#subPersonal"
      Bmp:15          =   "MDI.frx":102FAE
      Key:15          =   "#subCargos"
      Bmp:16          =   "MDI.frx":103D16
      Mask:16         =   1
      Key:16          =   "#subComisiones"
      Bmp:17          =   "MDI.frx":104568
      Key:17          =   "#subCapacitacion"
      Bmp:18          =   "MDI.frx":1052D0
      Key:18          =   "#subManuales"
      Bmp:19          =   "MDI.frx":106038
      Key:19          =   "#SubSituacion"
      Bmp:20          =   "MDI.frx":106DA0
      Key:20          =   "#SubMarcas"
      Bmp:21          =   "MDI.frx":107B08
      Key:21          =   "#subInformes"
      Bmp:22          =   "MDI.frx":108870
      Mask:22         =   14745599
      Key:22          =   "#subExamenes"
      Bmp:23          =   "MDI.frx":1090C2
      Key:23          =   "#subBuscarExamenes"
      Bmp:24          =   "MDI.frx":109914
      Key:24          =   "#subDiplomas"
      Bmp:25          =   "MDI.frx":10A67C
      Key:25          =   "#subConsultarCtas"
      Bmp:26          =   "MDI.frx":10B3E4
      Key:26          =   "#subCuentas"
      Bmp:27          =   "MDI.frx":10C14C
      Key:27          =   "#subGestion"
      Bmp:28          =   "MDI.frx":10CEB4
      Key:28          =   "#subAdmGrupos"
      Bmp:29          =   "MDI.frx":10DC1C
      Key:29          =   "#subLibroDeAula"
      Bmp:30          =   "MDI.frx":10E984
      Key:30          =   "#subDerechosExamenes"
      Bmp:31          =   "MDI.frx":10F6EC
      Key:31          =   "#subViaticos"
      Bmp:32          =   "MDI.frx":110454
      Mask:32         =   1
      Key:32          =   "#subVentaManual"
      Bmp:33          =   "MDI.frx":110CA6
      Key:33          =   "#subAuditoria"
      Bmp:34          =   "MDI.frx":1114F8
      Key:34          =   "#subStatus"
      Bmp:35          =   "MDI.frx":111D4A
      Key:35          =   "#subEliminarReservas"
      Bmp:36          =   "MDI.frx":112AB2
      Key:36          =   "#subReservas"
      Bmp:37          =   "MDI.frx":11381A
      Key:37          =   "#subEgresados"
      Bmp:38          =   "MDI.frx":114582
      Key:38          =   "#subContabilidad"
      Bmp:39          =   "MDI.frx":1152EA
      Key:39          =   "#subCuotas"
      Bmp:40          =   "MDI.frx":116052
      Key:40          =   "#subUltimasCuotas"
      Bmp:41          =   "MDI.frx":116DBA
      Key:41          =   "#subMatriculas"
      Bmp:42          =   "MDI.frx":117B22
      Mask:42         =   1
      Key:42          =   "#subBecaTotal"
      Bmp:43          =   "MDI.frx":118374
      Key:43          =   "#subNuevoCheque"
      Bmp:44          =   "MDI.frx":1190DC
      Key:44          =   "#subConsultarCheques"
      Bmp:45          =   "MDI.frx":119E44
      Key:45          =   "#subPP"
      Bmp:46          =   "MDI.frx":11ABAC
      Key:46          =   "#SubPresupuesto"
      Bmp:47          =   "MDI.frx":11B914
      Mask:47         =   14024703
      Key:47          =   "#subCopiarPresupuesto"
      Bmp:48          =   "MDI.frx":11C166
      Key:48          =   "#subnuevaorden"
      Bmp:49          =   "MDI.frx":11CECE
      Key:49          =   "#subordenes"
      Bmp:50          =   "MDI.frx":11DC36
      Key:50          =   "#subclientes"
      Bmp:51          =   "MDI.frx":11E99E
      Mask:51         =   15794175
      Key:51          =   "#subControl"
      Mask:52         =   16711935
      Key:52          =   "#mnuserviciotecnico"
      UseSystemFont   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Century Gothic"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
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
            TextSave        =   "05/10/2019"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "01:58 p.m."
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
         Weight          =   400
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
      Caption         =   "Gestión Educativa"
      Begin VB.Menu subCobranza 
         Caption         =   "Co&branza"
         Shortcut        =   ^B
      End
      Begin VB.Menu g1 
         Caption         =   "-"
      End
      Begin VB.Menu subGestion 
         Caption         =   "Gestión"
         Begin VB.Menu SubSituacion 
            Caption         =   "&Situación de Cartera"
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
            Caption         =   "&Cuotas"
            Shortcut        =   ^C
         End
         Begin VB.Menu subUltimasCuotas 
            Caption         =   "&Últimas Cuotas"
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
            Caption         =   "Matrículas"
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
         Caption         =   "&Derechos de Exámenes"
         Shortcut        =   ^D
      End
      Begin VB.Menu subExamenes 
         Caption         =   "&Exámenes"
         Shortcut        =   ^E
      End
      Begin VB.Menu subBuscarExamenes 
         Caption         =   "Buscar Exámenes"
      End
      Begin VB.Menu subDiplomas 
         Caption         =   "Diplomas Entregados"
      End
   End
   Begin VB.Menu mnuAdm 
      Caption         =   "Gestión Comercial"
      Begin VB.Menu subViaticos 
         Caption         =   "Viáticos"
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
      End
      Begin VB.Menu subCuentas 
         Caption         =   "Mantenimiento de Cuentas"
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
         Shortcut        =   ^X
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
      End
      Begin VB.Menu subReingresos 
         Caption         =   "Reingresos"
      End
      Begin VB.Menu subStatus 
         Caption         =   "Status de la Base"
      End
      Begin VB.Menu subAuditoria 
         Caption         =   "Auditoría"
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
      Caption         =   "Cerrar Sesión"
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
    a = MsgBox("¿Está seguro que desea Salir?", vbYesNo + vbQuestion, "Gestion Integral del Alumno")
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
    Origen = "T:\base.mdb"
    Destino = "T:\CopiaBase.mdb"
    If MsgBox("¿Realizar Copia de Seguridad?", vbQuestion + vbYesNo, "Gestión Integral del Alumno") = vbYes Then
            Set Fs = CreateObject("Scripting.FileSystemObject")
            Fs.CopyFile Origen, Destino
            MsgBox "La Copia de Respaldo se Realizó Correctamente", vbInformation + vbOKOnly, "Gestión Integral del Alumno"
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
    If MsgBox("¿Copiar el presupuesto actual para el mes próximo?", vbQuestion + vbYesNo, "Copiar Presupuesto") = vbYes Then
          
    '''Carga el presupuesto del mes en curso
        With rsCuentasPresupuesto
            If .State = 1 Then .Close
            .Open "SELECT * FROM presupuesto WHERE año=" & Year(Date) & "and mes='" & MonthName(Month(Date)) & "'", Cn, adOpenDynamic, adLockPessimistic
            .Requery
            .MoveFirst
        End With
        
    '''Abre la tabla presupuesto para agregar las cuentas del mes pròximo
        With rsPresupuesto
            If .State = 1 Then .Close
            .Open "SELECT * FROM presupuesto", Cn, adOpenDynamic, adLockPessimistic
            .Requery
            Do Until rsCuentasPresupuesto.EOF
                .AddNew
                !cuenta = rsCuentasPresupuesto!cuenta
                !deuda = rsCuentasPresupuesto!deuda
                !pagado = 0
                !saldo = rsCuentasPresupuesto!deuda
                !observaciones = ""
                
                If Month(Date) = 12 Then
                    !mes = "Enero"
                    !año = Year(Date) + 1
                Else
                    !año = Year(Date)
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
    Origen = "T:\base.mdb"
    Destino = "T:\CopiaBase.mdb"
    If MsgBox("¿Restaurar Copia de Seguridad?", vbQuestion + vbYesNo, "Gestión Integral del Alumno") = vbYes Then
            Set Fs = CreateObject("Scripting.FileSystemObject")
            Fs.CopyFile Destino, Origen
            MsgBox "La Restauración se Realizó Correctamente", vbInformation + vbOKOnly, "Gestión Integral del Alumno"
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
