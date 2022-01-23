Attribute VB_Name = "Conexiones"
Public dbBase As Database
Public rstListadoPacientes As Recordset
Public rstCargaPaciente As Recordset
Public rstComboLocalidad As Recordset
Public rstComboObraSocial As Recordset
Public rstExisteLocalidad As Recordset
Public rstExisteObraSocial As Recordset
Public rstDuplicadoDNI As Recordset
Public rstCargaAfiliaciones As Recordset
Public rstEliminarPaciente As Recordset
Public rstEditarPaciente As Recordset
Public rstEditarAfiliacion As Recordset
Public rstListadoLocalidades As Recordset
Public rstListadoObrasSociales As Recordset
Public rstListadoDiagnosticos As Recordset
Public rstEliminarDiagnostico As Recordset
Public rstEliminarLocalidad As Recordset
Public rstEliminarObraSocial As Recordset
Public rstDuplicadoDiagnostico As Recordset
Public rstCargaDiagnostico As Recordset
Public rstEditarDiagnostico As Recordset
Public rstEditarLocalidad As Recordset
Public rstCargaLocalidad As Recordset
Public rstDuplicadoLocalidad As Recordset
Public rstEditarObraSocial As Recordset
Public rstCargaObraSocial As Recordset
Public rstDuplicadoObraSocial As Recordset
Public rstListadoConsultas As Recordset
Public rstConsultasObservaciones As Recordset
Public rstDatosPacienteConsulta As Recordset
Public rstDatosObrasSocialesConsulta As Recordset
Public rstComboDiagnostico As Recordset
Public rstCargaConsulta As Recordset
Public rstEditarConsulta As Recordset
Public rstExisteDiagnostico As Recordset
Public rstExisteConsulta As Recordset
Public rstEliminarConsulta As Recordset
Public rstEditarTipoHistorial As Recordset
Public rstComboJerarquiaHistorial As Recordset
Public rstCargaTipoHistorial As Recordset
Public rstAsignarNumero As Recordset
Public rstListadoHistorialPrincipal As Recordset
Public rstListadoHistorialAccesorio As Recordset
Public rstDatosConsultaHistorial As Recordset
Public rstCargaHistorial As Recordset
Public rstEliminarTipoHistorial As Recordset
Public rstComprobarEliminacionTipoHistorial As Recordset
Public rstDatosCargaHistorial As Recordset
Public rstBuscarPaciente As Recordset
Public rstReporteGeneralPacientes As Recordset

Public Function BaseDeDatos() As String
    BaseDeDatos = "Fichas.mdb"
End Function

Public Sub Conexion(NombreBase As Database, Direccion As String)
    Set NombreBase = OpenDatabase(Direccion)
End Sub

Public Sub SetRecordset(NombreRecordset As Recordset, SQL As String)
    Set NombreRecordset = dbBase.OpenRecordset(SQL)
End Sub

Public Sub DesconectarRecordset()
    Set rstListadoPacientes = Nothing
    Set rstCargaPaciente = Nothing
    Set rstComboLocalidad = Nothing
    Set rstComboObraSocial = Nothing
    Set rstExisteLocalidad = Nothing
    Set rstExisteObraSocial = Nothing
    Set rstDuplicadoDNI = Nothing
    Set rstCargaAfiliaciones = Nothing
    Set rstEliminarPaciente = Nothing
    Set rstEditarPaciente = Nothing
    Set rstEditarAfiliacion = Nothing
    Set rstListadoLocalidades = Nothing
    Set rstListadoObrasSociales = Nothing
    Set rstListadoDiagnosticos = Nothing
    Set rstEliminarDiagnostico = Nothing
    Set rstEliminarLocalidad = Nothing
    Set rstEliminarObraSocial = Nothing
    Set rstDuplicadoDiagnostico = Nothing
    Set rstCargaDiagnostico = Nothing
    Set rstEditarDiagnostico = Nothing
    Set rstEditarLocalidad = Nothing
    Set rstCargaLocalidad = Nothing
    Set rstDuplicadoLocalidad = Nothing
    Set rstEditarObraSocial = Nothing
    Set rstCargaObraSocial = Nothing
    Set rstDuplicadoObraSocial = Nothing
    Set rstListadoConsultas = Nothing
    Set rstConsultasObservaciones = Nothing
    Set rstDatosPacienteConsulta = Nothing
    Set rstDatosObrasSocialesConsulta = Nothing
    Set rstComboDiagnostico = Nothing
    Set rstCargaConsulta = Nothing
    Set rstEditarPaciente = Nothing
    Set rstExisteDiagnostico = Nothing
    Set rstExisteConsulta = Nothing
    Set rstEliminarConsulta = Nothing
    Set rstEditarTipoHistorial = Nothing
    Set rstComboJerarquiaHistorial = Nothing
    Set rstCargaTipoHistorial = Nothing
    Set rstAsignarNumero = Nothing
    Set rstListadoHistorialPrincipal = Nothing
    Set rstListadoHistorialAccesorio = Nothing
    Set rstDatosConsultaHistorial = Nothing
    Set rstCargaHistorial = Nothing
    Set rstEliminarTipoHistorial = Nothing
    Set rstComprobarEliminacionTipoHistorial = Nothing
    Set rstDatosCargaHistorial = Nothing
    Set rstBuscarPaciente = Nothing
    Set rstReporteGeneralPacientes = Nothing
End Sub
