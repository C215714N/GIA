Attribute VB_Name = "Conexiones"
Sub Main()
''' Conexion con la Base de Datos
    Cn.CursorLocation = adUseClient
    Cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=T:\Base.mdb;Persist Security Info=False;Jet OLEDB:Database Password=ascir"
''' Formulario de Inicio de Sesion
    frmClave.Show
''' Carga los Datos de la Empresa - Control
    Control
    With rsControl
        .MoveFirst
        frmClave.Caption = frmClave.Caption & " - " & !empresa & " - " & !sucursal
    End With
End Sub

Sub Control()
    With rsControl
        If .State = 1 Then .Close
        .Open "SELECT * FROM Control", Cn, adOpenDynamic, adLockPessimistic
    End With
End Sub

Sub Suscripciones()
    With rsSuscripciones
        If .State = 1 Then .Close
        .Open "SELECT * FROM Suscripciones", Cn, adOpenDynamic, adLockPessimistic
    End With
End Sub

Sub Verificaciones()
    With rsVerificaciones
        If .State = 1 Then .Close
        .Open "SELECT * FROM Verificaciones", Cn, adOpenDynamic, adLockPessimistic
    End With
End Sub

Sub Capacitaciones()
    With rsCapacitaciones
        If .State = 1 Then .Close
        .Open "SELECT capacitacion FROM Capacitacion ORDER BY capacitacion", Cn, adOpenDynamic, adLockPessimistic
    End With
End Sub

Public Sub Equipos()
    With rsEquipos
        If .State = 1 Then .Close
        .Open "SELECT * FROM equipos", Cn, adOpenDynamic, adLockPessimistic
        .MoveFirst
    End With
End Sub

Sub Personal()
    With rsPersonal
        If .State = 1 Then .Close
        .Open "SELECT * FROM Personal", Cn, adOpenDynamic, adLockPessimistic
    End With
End Sub

Sub Asistente()
    With rsPersonal
        If .State = 1 Then .Close
        .Open "SELECT NyA as [Personal], Cargo FROM Personal WHERE Cargo='Asistente' ORDER BY nya", Cn, adOpenDynamic, adLockPessimistic
    End With
End Sub

Sub Cargos()
    With rsCargos
        If .State = 1 Then .Close
        .Open "SELECT * FROM Cargos", Cn, adOpenDynamic, adLockPessimistic
    End With
End Sub

Sub PlanDePago()
    With rsPlanDePago
        If .State = 1 Then .Close
        .Open "SELECT * FROM PlanDePago", Cn, adOpenDynamic, adLockPessimistic
    End With
End Sub

Sub Cuentas()
    With rsCuentas
        If .State = 1 Then .Close
        .Open "SELECT * FROM cuentas ORDER BY cuenta", Cn, adOpenDynamic, adLockPessimistic
    End With
End Sub

Sub Contabilidad()
    With rsContabilidad
        If .State = 1 Then .Close
        .Open "SELECT * FROM contabilidad", Cn, adOpenDynamic, adLockPessimistic
    End With
End Sub

Sub ContabilidadTemp()
    With rsContabilidadTemp
        If .State = 1 Then .Close
        .Open "SELECT * FROM ContabilidadTemp", Cn, adOpenDynamic, adLockPessimistic
    End With
End Sub

Sub AnalisisDeCuota()
    With rsAnalisisDeCuenta
        If .State = 1 Then .Close
        .Open "SELECT nrocuota as [N°],fechavto as [Vencimiento],tipodepago as [Pago],deuda as [Deuda],observaciones,codalumno as [Codigo],nya as [Alumno] FROM PlanDePago WHERE codalumno=" & CodAlumno & " ORDER BY nrocuota", Cn, adOpenDynamic, adLockPessimistic
    End With
End Sub

Sub Historico()
    With rsHistorico
        If .State = 1 Then .Close
        .Open "SELECT Nrocuota as [N°],nrofactura as [Recibo],fecha AS [Fecha],debe as [Monto] FROM contabilidad WHERE codalumno=" & CodAlumno & " ORDER BY Nrocuota, fecha", Cn, adOpenDynamic, adLockPessimistic
    End With
End Sub

Sub Marcar()
    With rsMarcar
        If .State = 1 Then .Close
        .Open "SELECT * FROM marcas ORDER BY codalumno", Cn, adOpenDynamic, adLockPessimistic
        .MoveFirst
    End With
End Sub

Sub Cobranza()
    With rsCobranza
        If .State = 1 Then .Close
        .Open "SELECT * FROM PlanDePago WHERE codalumno=" & CodAlumno, Cn, adOpenDynamic, adLockPessimistic
    End With
End Sub

Public Sub Localidades()
    With rsLocalidades
        If .State = 1 Then .Close
        .Open "SELECT * FROM Localidad ORDER BY localidad", Cn, adOpenDynamic, adLockPessimistic
    End With
End Sub

Public Sub Centrar(frm As Form)
'''Centra el Formulario en Pantalla
    frm.Top = (MDI.Height - frm.Height) \ 4 - 500
    frm.Left = (MDI.Width - frm.Width) \ 2
End Sub
