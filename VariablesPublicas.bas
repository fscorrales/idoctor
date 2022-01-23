Attribute VB_Name = "VariablesPublicas"
Public bolEditandoPaciente As Boolean
Public bolEditandoConsulta As Boolean
Public bolEditandoDiagnostico As Boolean
Public bolEditandoLocalidad As Boolean
Public bolEditandoObraSocial As Boolean
Public bolCargaLocalidadDesdePacientes As Boolean
Public strPasajeAPacientes As String
Public bolCargaObraSocialPrimariaDesdePacientes As Boolean
Public bolCargaObraSocialSecundariaDesdePacientes As Boolean
Public bolCargaDiagnosticoDesdeConsultas As Boolean
Public strPasajeAConsultas As String
Public AgregarPacientesDesdeListadoConsultas As Boolean
Public bolEditandoPacienteDesdeConsulta As Boolean
Public bolEditandoTipoHistorial As Boolean
Public strBuscarPaciente As String

Public Function CalcularEdad(FechaDeNacimiento As Date) As String
    Select Case Month(Date)
        Case Is > Month(FechaDeNacimiento)
            CalcularEdad = Year(Date) - Year(FechaDeNacimiento)
        Case Is = Month(FechaDeNacimiento)
            If Day(FechaDeNacimiento) <= Day(Date) Then
                CalcularEdad = Year(Date) - Year(FechaDeNacimiento)
            ElseIf Day(FechaDeNacimiento) > Day(Date) Then
                CalcularEdad = (Year(Date) - Year(FechaDeNacimiento)) - 1
            End If
        Case Is < Month(FechaDeNacimiento)
            CalcularEdad = (Year(Date) - Year(FechaDeNacimiento)) - 1
    End Select
End Function

Public Sub Buscar(ValorBuscado As String, Recordset As Recordset, SQL As String, Indice As String)
    Call SetRecordset(Recordset, SQL)
    With Recordset
        .Index = Indice
        .Seek "=", ValorBuscado
    End With
End Sub

Public Sub Encontrar(ValorBuscado As String, Recordset As Recordset, SQL As String, CampoBuscado As String)
    Call SetRecordset(Recordset, SQL)
    ValorBuscado = Format(ValorBuscado, "'&&&&&&&&&&&&&&&&&&&&&&&" _
    & "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&" _
    & "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&" _
    & "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&'")
    Recordset.FindFirst CampoBuscado & "=" & ValorBuscado
End Sub

Public Sub BlanquearVariables()
    bolEditandoPaciente = False
    bolEditandoDiagnostico = False
    bolEditandoLocalidad = False
    bolCargaLocalidadDesdePacientes = False
    strPasajeAPacientes = ""
    bolCargaObraSocialPrimariaDesdePacientes = False
    bolCargaObraSocialSecundariaDesdePacientes = False
    bolEditandoConsulta = False
    bolCargaDiagnosticoDesdeConsultas = False
    strPasajeAConsultas = ""
    AgregarPacientesDesdeListadoConsultas = False
    bolEditandoPacienteDesdeConsulta = False
    bolEditandoTipoHistorial = False
    strBuscarPaciente = ""
End Sub
