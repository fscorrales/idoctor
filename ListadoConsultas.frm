VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoConsultas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de Consultas"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Acciones Posibles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   6480
      Width           =   9615
      Begin VB.CommandButton cmdTratamientoProfilaxis 
         Caption         =   "Cargar TRATAMIENTO - PROFILAXIS"
         Height          =   375
         Left            =   3360
         TabIndex        =   13
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton cmdVacunasOtros 
         Caption         =   "Cargar VACUNAS - OTROS"
         Height          =   375
         Left            =   6480
         TabIndex        =   12
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton cmdLaboratorio 
         Caption         =   "Cargar LABORATORIO"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton cmdAgregarConsulta 
         Caption         =   "Agregar Nueva Consulta"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   2895
      End
      Begin VB.CommandButton cmdEditarConsulta 
         Caption         =   "Editar Datos de la Consulta"
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   360
         Width           =   2895
      End
      Begin VB.CommandButton cmdEliminarConsulta 
         Caption         =   "Eliminar Consulta Definitivamente"
         Height          =   375
         Left            =   6480
         TabIndex        =   7
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.TextBox txtObservaciones 
      Height          =   1215
      Left            =   5040
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   5040
      Width           =   4575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Detalle de la Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   4920
      TabIndex        =   4
      Top             =   4680
      Width           =   4815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Consultas por Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   4680
      Width           =   4815
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgConsultas 
         Height          =   1215
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   2143
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pacientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.ComboBox txtCategoriaBuscar 
         Height          =   315
         Left            =   1320
         TabIndex        =   16
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtBuscar 
         Height          =   285
         Left            =   3120
         TabIndex        =   15
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Ejecutar Función BUSCAR"
         Height          =   375
         Left            =   6600
         TabIndex        =   14
         Top             =   360
         Width           =   2895
      End
      Begin VB.CommandButton cmdAgregarPaciente 
         Caption         =   "Agregar Nuevo Paciente"
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Top             =   3960
         Width           =   2895
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgPacientes 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   5530
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label2 
         Caption         =   "BUSCADOR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "ListadoConsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAgregarConsulta_Click()
    Dim i As Integer
    i = dgPacientes.Row
    Call SetRecordset(rstDatosPacienteConsulta, "Select * From PACIENTES Where DNI = " & "'" & dgPacientes.TextMatrix(i, 1) & "'")
    Call SetRecordset(rstDatosObrasSocialesConsulta, "Select * From AFILIACIONES Where DNI = " & "'" & dgPacientes.TextMatrix(i, 1) & "' ORDER BY TipoAfiliacion")
    i = 0
    Unload ListadoConsultas
    Consultas.Show
End Sub

Private Sub cmdAgregarPaciente_Click()
    Unload ListadoConsultas
    Pacientes.Show
    AgregarPacientesDesdeListadoConsultas = True
End Sub

Private Sub cmdBuscar_Click()
    Call BuscarPaciente(txtCategoriaBuscar.Text, txtBuscar.Text)
End Sub

Private Sub cmdEditarConsulta_Click()
    Dim i As String
    Dim X As String
    i = dgPacientes.Row
    X = dgConsultas.Row
    Call SetRecordset(rstDatosPacienteConsulta, "Select * From PACIENTES Where DNI = " & "'" & dgPacientes.TextMatrix(i, 1) & "'")
    Call SetRecordset(rstDatosObrasSocialesConsulta, "Select * From AFILIACIONES Where DNI = " & "'" & dgPacientes.TextMatrix(i, 1) & "' ORDER BY TipoAfiliacion")
    Call SetRecordset(rstEditarConsulta, "Select * from INGRESOS Where NumeroIngreso = " & dgConsultas.TextMatrix(X, 0))
    Unload ListadoConsultas
    bolEditandoConsulta = True
    Consultas.Show
    i = ""
    o = ""

End Sub

Private Sub cmdLaboratorio_Click()
    Dim X As String
    X = dgConsultas.Row
    Call SetRecordset(rstDatosConsultaHistorial, "Select INGRESOS.*, PACIENTES.Apellido, PACIENTES.Nombre from PACIENTES INNER JOIN INGRESOS ON PACIENTES.dni = INGRESOS.dni Where NumeroIngreso = " & dgConsultas.TextMatrix(X, 0))
    Call SetRecordset(rstDatosCargaHistorial, "Select * From HISTORIAL Where NumeroIngreso = " & dgConsultas.TextMatrix(X, 0) & "And CodigoHistorial LIKE 'L###' Order by CodigoHistorial")
    Unload ListadoConsultas
    Laboratorio.Show
    X = ""
End Sub

Private Sub cmdTratamientoProfilaxis_Click()
    Dim X As String
    X = dgConsultas.Row
    Call SetRecordset(rstDatosConsultaHistorial, "Select INGRESOS.*, PACIENTES.Apellido, PACIENTES.Nombre from PACIENTES INNER JOIN INGRESOS ON PACIENTES.dni = INGRESOS.dni Where NumeroIngreso = " & dgConsultas.TextMatrix(X, 0))
    Call SetRecordset(rstDatosCargaHistorial, "Select * From HISTORIAL Where NumeroIngreso = " & dgConsultas.TextMatrix(X, 0) & "And CodigoHistorial LIKE '[T,P]###' Order by CodigoHistorial")
    Unload ListadoConsultas
    TratamientoYProfilaxis.Show
    X = ""
End Sub

Private Sub cmdVacunasOtros_Click()
    Dim X As String
    X = dgConsultas.Row
    Call SetRecordset(rstDatosConsultaHistorial, "Select INGRESOS.*, PACIENTES.Apellido, PACIENTES.Nombre from PACIENTES INNER JOIN INGRESOS ON PACIENTES.dni = INGRESOS.dni Where NumeroIngreso = " & dgConsultas.TextMatrix(X, 0))
    Call SetRecordset(rstDatosCargaHistorial, "Select * From HISTORIAL Where NumeroIngreso = " & dgConsultas.TextMatrix(X, 0) & "And CodigoHistorial LIKE '[V,O]###' Order by CodigoHistorial")
    Unload ListadoConsultas
    VacunasYOtros.Show
    X = ""
End Sub

Private Sub dgConsultas_RowColChange()
    Dim i As Integer
    i = dgConsultas.Row
    txtObservaciones.Text = dgConsultas.TextMatrix(i, 3)
    i = 0
End Sub

Private Sub dgPacientes_RowColChange()
    Dim i As Integer
    i = dgPacientes.Row
    ConfigurarConsultas
    CargarConsultas (dgPacientes.TextMatrix(i, 1))
    txtObservaciones.Text = dgConsultas.TextMatrix(1, 3)
    i = 0
    If dgConsultas.TextMatrix(1, 0) = "" Then
        cmdEliminarConsulta.Enabled = False
        cmdEditarConsulta.Enabled = False
        cmdLaboratorio.Enabled = False
        cmdTratamientoProfilaxis.Enabled = False
        cmdVacunasOtros.Enabled = False
    Else
        cmdEliminarConsulta.Enabled = True
        cmdEditarConsulta.Enabled = True
        cmdLaboratorio.Enabled = True
        cmdTratamientoProfilaxis.Enabled = True
        cmdVacunasOtros.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Call CenterMe(ListadoConsultas)
    With ListadoConsultas
        .Width = 10000
        .Height = 8300
    End With
    ConfigurarPacientes
    CargarPacientes
    ConfigurarConsultas
    CargarConsultas (dgPacientes.TextMatrix(1, 1))
    txtObservaciones.Text = dgConsultas.TextMatrix(1, 3)
    If dgConsultas.TextMatrix(1, 0) = "" Then
        cmdEliminarConsulta.Enabled = False
        cmdEditarConsulta.Enabled = False
        cmdLaboratorio.Enabled = False
        cmdTratamientoProfilaxis.Enabled = False
        cmdVacunasOtros.Enabled = False
    Else
        cmdEliminarConsulta.Enabled = True
        cmdEditarConsulta.Enabled = True
        cmdLaboratorio.Enabled = True
        cmdTratamientoProfilaxis.Enabled = True
        cmdVacunasOtros.Enabled = True
    End If
    txtCategoriaBuscar.AddItem "DNI"
    txtCategoriaBuscar.AddItem "Apellido"
    txtCategoriaBuscar.AddItem "Apellido, Nombre"
    txtCategoriaBuscar.Text = "Apellido"
    If Len(strBuscarPaciente) <> 0 Then
        Call BuscarPaciente("DNI", strBuscarPaciente)
        strBuscarPaciente = ""
    End If
End Sub

Sub ConfigurarPacientes()
    With dgPacientes
        .Clear
        .Cols = 5
        .Rows = 2
        .TextMatrix(0, 0) = "Apellido y Nombre"
        .TextMatrix(0, 1) = "D.N.I."
        .TextMatrix(0, 2) = "Edad"
        .TextMatrix(0, 3) = "Domicilio"
        .TextMatrix(0, 4) = "Localidad"
        .ColWidth(0) = 3000
        .ColWidth(1) = 1000
        .ColWidth(2) = 900
        .ColWidth(3) = 2500
        .ColWidth(4) = 1500
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 7
        .ColAlignment(2) = 4
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
    End With
End Sub

Sub CargarPacientes()
    Dim i As Integer
    i = 0
    dgPacientes.Rows = 2
    Call SetRecordset(rstListadoPacientes, "Select * From PACIENTES Order by Apellido")
    If rstListadoPacientes.BOF = False Then
        With rstListadoPacientes
            .MoveFirst
            While .EOF = False
                i = i + 1
                dgPacientes.RowHeight(i) = 300
                dgPacientes.TextMatrix(i, 0) = .Fields("Apellido") & ", " & .Fields("Nombre")
                If Len(.Fields("DNI")) <> "0" Then
                    dgPacientes.TextMatrix(i, 1) = .Fields("DNI")
                Else
                    dgPacientes.TextMatrix(i, 1) = "FALTA DNI"
                End If
                If Len(.Fields("FechaNacimiento")) <> "0" Then
                    dgPacientes.TextMatrix(i, 2) = CalcularEdad(.Fields("FechaNacimiento")) & " años"
                Else
                    dgPacientes.TextMatrix(i, 2) = "SIN FECHA"
                End If
                If Len(.Fields("Domicilio")) <> "0" Then
                    dgPacientes.TextMatrix(i, 3) = .Fields("Domicilio")
                Else
                    dgPacientes.TextMatrix(i, 3) = "FALTA DOMICILIO"
                End If
                If Len(.Fields("Localidad")) <> "0" Then
                    dgPacientes.TextMatrix(i, 4) = .Fields("Localidad")
                Else
                    dgPacientes.TextMatrix(i, 4) = "FALTA LOCALIDAD"
                End If
                .MoveNext
                dgPacientes.Rows = dgPacientes.Rows + 1
            Wend
        End With
        dgPacientes.Rows = dgPacientes.Rows - 1
    End If
    Set rstListadoPacientes = Nothing
End Sub

Sub ConfigurarConsultas()
    With dgConsultas
        .Clear
        .Cols = 4
        .Rows = 2
        .TextMatrix(0, 0) = "Numero Ingreso"
        .TextMatrix(0, 1) = "Fecha"
        .TextMatrix(0, 2) = "Diagnóstico"
        .TextMatrix(0, 3) = "Observaciones"
        .ColWidth(0) = 1
        .ColWidth(1) = 1000
        .ColWidth(2) = 3000
        .ColWidth(3) = 1
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
    End With
End Sub

Sub CargarConsultas(Numero As String)
    Dim i As Integer
    i = 0
    dgConsultas.Rows = 2
    Call SetRecordset(rstListadoConsultas, "Select Fecha,Diagnostico,Observaciones,NumeroIngreso FROM INGRESOS Where DNI = " & "'" & Numero & "' Order By Fecha")
    If rstListadoConsultas.BOF = False Then
        With rstListadoConsultas
            .MoveFirst
            While .EOF = False
                i = i + 1
                dgConsultas.RowHeight(i) = 300
                dgConsultas.TextMatrix(i, 0) = .Fields("NumeroIngreso")
                dgConsultas.TextMatrix(i, 1) = .Fields("Fecha")
                dgConsultas.TextMatrix(i, 2) = .Fields("Diagnostico")
                If Len(.Fields("Observaciones")) <> 0 Then
                    dgConsultas.TextMatrix(i, 3) = .Fields("Observaciones")
                Else
                    dgConsultas.TextMatrix(i, 3) = ""
                End If
                .MoveNext
                dgConsultas.Rows = dgConsultas.Rows + 1
            Wend
        End With
        dgConsultas.Rows = dgConsultas.Rows - 1
    End If

End Sub

Private Sub cmdEliminarConsulta_Click()
    Dim X As String
    Dim Y As String
    Dim Borrar As Integer
    X = dgPacientes.Row
    Y = dgConsultas.Row
    Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la CONSULTA del PACIENTE: " & dgPacientes.TextMatrix(X, 0) & " de fecha: " & dgConsultas.TextMatrix(Y, 1) & "?", vbQuestion + vbYesNo, "BORRANDO CONSULTA")
    If Borrar = 6 Then
        Call SetRecordset(rstEliminarConsulta, "Select * From INGRESOS Where NumeroIngreso = " & dgConsultas.TextMatrix(Y, 0))
        rstEliminarConsulta.Delete
        Set rstEliminarConsulta = Nothing
    ConfigurarPacientes
    CargarPacientes
    ConfigurarConsultas
    CargarConsultas (dgPacientes.TextMatrix(1, 1))
    txtObservaciones.Text = dgConsultas.TextMatrix(1, 3)
    End If
    Borrar = 0
    X = ""
    Y = ""
End Sub

Public Sub BuscarPaciente(Categoria As String, ValorBuscado As String)
        
    Dim strApellido As String
    Dim strNombre As String
    Dim i As Integer
    
    If EsIgualTextoEspecificado(Categoria, "TIPO DE DATO A BUSCAR", "DNI", "Apellido", "Apellido, Nombre") = False Then
        Exit Sub
    ElseIf Trim(ValorBuscado) = "" Then
        MsgBox "Debe especificar el DATO a buscar", vbInformation + vbOKOnly, "DATO A BUSCAR NO INGRESADO"
    Else
        Select Case Categoria
        Dim X As Integer
        Case Is = "DNI"
            Call Buscar(ValorBuscado, rstBuscarPaciente, "PACIENTES", "pkPaciente")
            If rstBuscarPaciente.NoMatch = True Then
                MsgBox "El DNI buscado no existe." & vbCrLf & "Verifique que haya ingresado correctamente el dato", vbInformation + vbOKOnly, "BÚSQUEDA SIN ÉXITO"
            Else
                With dgPacientes
                    For X = 1 To (.Rows - 1)
                        If .TextMatrix(X, 1) = ValorBuscado Then
                            .Row = X
                            If strBuscarPaciente = "" Then
                                .SetFocus
                            End If
                            .TopRow = X
                            CargarConsultas (dgPacientes.TextMatrix(X, 1))
                            txtObservaciones.Text = dgConsultas.TextMatrix(1, 3)
                            If dgConsultas.TextMatrix(1, 0) = "" Then
                                cmdEliminarConsulta.Enabled = False
                                cmdEditarConsulta.Enabled = False
                                cmdLaboratorio.Enabled = False
                                cmdTratamientoProfilaxis.Enabled = False
                                cmdVacunasOtros.Enabled = False
                            Else
                                cmdEliminarConsulta.Enabled = True
                                cmdEditarConsulta.Enabled = True
                                cmdLaboratorio.Enabled = True
                                cmdTratamientoProfilaxis.Enabled = True
                                cmdVacunasOtros.Enabled = True
                            End If
                            Exit For
                        End If
                    Next X
                End With
            End If
        Case Is = "Apellido"
            Call SetRecordset(rstBuscarPaciente, "Select DNI, Apellido From PACIENTES Where Apellido LIKE " & "'" & ValorBuscado & "*' Order By Apellido")
            If rstBuscarPaciente.EOF = True Then
                MsgBox "El Apellido buscado no existe." & vbCrLf & "Verifique que haya ingresado correctamente el dato", vbInformation + vbOKOnly, "BÚSQUEDA SIN ÉXITO"
            Else
                With dgPacientes
                rstBuscarPaciente.MoveFirst
                    For X = 1 To (.Rows - 1)
                        If .TextMatrix(X, 1) = rstBuscarPaciente.Fields("DNI") Then
                            .Row = X
                            .SetFocus
                            .TopRow = X
                            CargarConsultas (dgPacientes.TextMatrix(X, 1))
                            txtObservaciones.Text = dgConsultas.TextMatrix(1, 3)
                            If dgConsultas.TextMatrix(1, 0) = "" Then
                                cmdEliminarConsulta.Enabled = False
                                cmdEditarConsulta.Enabled = False
                                cmdLaboratorio.Enabled = False
                                cmdTratamientoProfilaxis.Enabled = False
                                cmdVacunasOtros.Enabled = False
                            Else
                                cmdEliminarConsulta.Enabled = True
                                cmdEditarConsulta.Enabled = True
                                cmdLaboratorio.Enabled = True
                                cmdTratamientoProfilaxis.Enabled = True
                                cmdVacunasOtros.Enabled = True
                            End If
                            Exit For
                        End If
                    Next X
                End With
            End If
        Case Is = "Apellido, Nombre"
            i = InStr(1, ValorBuscado, ",", vbBinaryCompare)
            strApellido = Trim(Left(ValorBuscado, i - 1))
            i = Len(ValorBuscado) - i - 1
            strNombre = Trim(Right(ValorBuscado, i + 1))
            Call SetRecordset(rstBuscarPaciente, "Select DNI, Apellido, Nombre From PACIENTES" _
            & " Where Apellido Like " & " '" & strApellido _
            & "*' And Nombre Like " & " '" & strNombre _
            & "*' Order By Apellido, Nombre")
            If rstBuscarPaciente.EOF = True Then
                MsgBox "El Apellido, Nombre buscado no existe." & vbCrLf & "Verifique que haya ingresado correctamente el dato", vbInformation + vbOKOnly, "BÚSQUEDA SIN ÉXITO"
            Else
                With dgPacientes
                rstBuscarPaciente.MoveFirst
                    For X = 1 To (.Rows - 1)
                        If .TextMatrix(X, 1) = rstBuscarPaciente.Fields("DNI") Then
                            .Row = X
                            .SetFocus
                            .TopRow = X
                            CargarConsultas (dgPacientes.TextMatrix(X, 1))
                            txtObservaciones.Text = dgConsultas.TextMatrix(1, 3)
                            If dgConsultas.TextMatrix(1, 0) = "" Then
                                cmdEliminarConsulta.Enabled = False
                                cmdEditarConsulta.Enabled = False
                                cmdLaboratorio.Enabled = False
                                cmdTratamientoProfilaxis.Enabled = False
                                cmdVacunasOtros.Enabled = False
                            Else
                                cmdEliminarConsulta.Enabled = True
                                cmdEditarConsulta.Enabled = True
                                cmdLaboratorio.Enabled = True
                                cmdTratamientoProfilaxis.Enabled = True
                                cmdVacunasOtros.Enabled = True
                            End If
                            Exit For
                        End If
                    Next X
                End With
            End If
        End Select
        X = 0
        Set rstBuscarPaciente = Nothing
    End If
End Sub
