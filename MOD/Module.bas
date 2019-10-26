Attribute VB_Name = "Declaraciones"
''' Conexion Base de Datos
    Global Cn                           As New ADODB.Connection

''' Control General
    Global rsControl                    As New ADODB.Recordset
    Global rsStatus                     As New ADODB.Recordset
    Global rsLocalidades                As New ADODB.Recordset

''' Control de Empleados
    Global rsPersonal                   As New ADODB.Recordset
    Global rsCargos                     As New ADODB.Recordset
    Global rsViaticos                   As New ADODB.Recordset

''' Control de Alumnos
    Global rsSuscripciones              As New ADODB.Recordset
    Global rsInformeSuscripciones       As New ADODB.Recordset
    Global rsVerificaciones             As New ADODB.Recordset
    Global rsMatriculas                 As New ADODB.Recordset
    
    Global rsAlumnosDelMes              As New ADODB.Recordset
    Global rsAlumnosBecados             As New ADODB.Recordset
    Global rsBajas                      As New ADODB.Recordset
    
''' Gestion Comercial
    Global rsCobranza                   As New ADODB.Recordset
    Global rsContabilidad               As New ADODB.Recordset
    Global rsContabilidadTemp           As New ADODB.Recordset
    Global rsCuentas                    As New ADODB.Recordset
    Global rsPresupuesto                As New ADODB.Recordset
    Global rsCuentasPresupuesto         As New ADODB.Recordset
    Global rsCheques                    As New ADODB.Recordset
    
    Global rsSituacionDeCartera         As New ADODB.Recordset
    Global rsSituacionesDeCartera       As New ADODB.Recordset
    Global rsTotalesSituaciones         As New ADODB.Recordset
    Global rsAnalisisDeCuenta           As New ADODB.Recordset
    Global rsAnalisisSituacionDeDeuda   As New ADODB.Recordset
      
    Global rsPlanDePago                 As New ADODB.Recordset
    Global rsRestaurarPlanDePago        As New ADODB.Recordset
    Global rsControlDeP                 As New ADODB.Recordset
    Global rsCuotasXFecha               As New ADODB.Recordset
    Global rsUltimasCuotas              As New ADODB.Recordset
    Global rsHistorico                  As New ADODB.Recordset
    Global rsMarcar                     As New ADODB.Recordset
    Global rsMarcas                     As New ADODB.Recordset
    
''' Gestion Estudiantil
    Global rsCapacitaciones             As New ADODB.Recordset
    Global rsLibro                      As New ADODB.Recordset
    Global rsAsistencia                 As New ADODB.Recordset
    
    Global rsGruposDeArmado             As New ADODB.Recordset
    Global rsAlumnosArmado              As New ADODB.Recordset
    Global rsEquipos                    As New ADODB.Recordset
    Global rsReservas                   As New ADODB.Recordset
    Global rsEliminar                   As New ADODB.Recordset
    
    Global rsManuales                   As New ADODB.Recordset
    Global rsVentaManuales              As New ADODB.Recordset
    Global rsDerechosExamenes           As New ADODB.Recordset
    Global rsExamenes                   As New ADODB.Recordset
    Global rsEgresados                  As New ADODB.Recordset
    Global rsDiplomas                   As New ADODB.Recordset
    
''' Gestion Servicio Tecnico
    Global rsclientes                   As New ADODB.Recordset
    Global rsAnalisisInforme            As New ADODB.Recordset
    Global rsconsultarordenes           As New ADODB.Recordset

'''DECLARACION DE VARIABLES
    ''' Variables Control de Acceso
        Global Usuario          As String
        Global Clave            As String

    ''' Variables Gestion Educativa
        Global Analisis         As Boolean
        Global Verificar        As Boolean
        Global CuotasXFecha     As Boolean
        Global Modi             As Boolean
        Global ModiReservas     As Boolean
        Global ModiLibro        As Boolean

    ''' Variables Gestion Comercial
        Global CodCurso         As Long
        Global CodAlumno        As Long
        Global Debe             As Single
        Global Haber            As Single
        Global Situacion        As Integer
