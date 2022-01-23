VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoPacientes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de Pacientes"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
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
      TabIndex        =   2
      Top             =   5160
      Width           =   9615
      Begin VB.CommandButton cmdEliminarPaciente 
         Caption         =   "Eliminar Paciente Definitivamente"
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         Top             =   840
         Width           =   4095
      End
      Begin VB.CommandButton cmdEditarPaciente 
         Caption         =   "Editar Datos del Paciente"
         Height          =   375
         Left            =   5040
         TabIndex        =   5
         Top             =   360
         Width           =   4095
      End
      Begin VB.CommandButton cmdIngresarConsulta 
         Caption         =   "Ingresar Nueva Consulta"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   4095
      End
      Begin VB.CommandButton cmdAgregarPaciente 
         Caption         =   "Agregar Nuevo Paciente"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   4095
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
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgPacientes 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ListadoPacientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregarPaciente_Click()
    Unload ListadoPacientes
    Pacientes.Show
End Sub

Private Sub cmdEditarPaciente_Click()
    Dim i As String
    i = dgPacientes.Row
    Call SetRecordset(rstEditarPaciente, "Select * from PACIENTES Where DNI = " & "'" & dgPacientes.TextMatrix(i, 1) & "'")
    Call SetRecordset(rstEditarAfiliacion, "Select * from AFILIACIONES Where DNI = " & "'" & dgPacientes.TextMatrix(i, 1) & "'" & " Order by TipoAfiliacion Asc")
    Unload ListadoPacientes
    With Pacientes
        .Show
        .txtApellido.Text = rstEditarPaciente.Fields("Apellido")
        .txtNombre.Text = rstEditarPaciente.Fields("Nombre")
        .txtDNI.Text = rstEditarPaciente.Fields("DNI")
        .txtNacimiento.Text = rstEditarPaciente.Fields("FechaNacimiento")
        If Len(rstEditarPaciente.Fields("Domicilio")) <> "0" Then
            .txtDireccion.Text = rstEditarPaciente.Fields("Domicilio")
        Else
            .txtDireccion.Text = ""
        End If
        If Len(rstEditarPaciente.Fields("Domicilio")) <> "0" Then
            .txtDireccion.Text = rstEditarPaciente.Fields("Domicilio")
        Else
            .txtDireccion.Text = ""
        End If
        .txtLocalidad.Text = rstEditarPaciente.Fields("Localidad")
        If Len(rstEditarPaciente.Fields("Telefono1")) <> "0" Then
            .txtTelefono1.Text = rstEditarPaciente.Fields("Telefono1")
        Else
            .txtTelefono1.Text = ""
        End If
        If Len(rstEditarPaciente.Fields("Telefono2")) <> "0" Then
            .txtTelefono2.Text = rstEditarPaciente.Fields("Telefono2")
        Else
            .txtTelefono2.Text = ""
        End If
        .txtObraSocialPrimaria.Text = rstEditarAfiliacion.Fields("ObraSocial")
        If Len(rstEditarAfiliacion.Fields("NumeroAfiliado")) <> "0" Then
            .txtAfiliadoPrimario.Text = rstEditarAfiliacion.Fields("NumeroAfiliado")
        Else
            .txtAfiliadoPrimario.Text = ""
        End If
        If rstEditarAfiliacion.RecordCount > "1" Then
            rstEditarAfiliacion.MoveNext
            .txtObraSocialSecundaria.Text = rstEditarAfiliacion.Fields("ObraSocial")
            If Len(rstEditarAfiliacion.Fields("NumeroAfiliado")) <> "0" Then
                .txtAfiliadoSecundario.Text = rstEditarAfiliacion.Fields("NumeroAfiliado")
            Else
                .txtAfiliadoSecundario.Text = ""
            End If
            rstEditarAfiliacion.MovePrevious
        End If
    End With
    i = ""
    bolEditandoPaciente = True
End Sub

Private Sub cmdEliminarPaciente_Click()
    Dim i As String
    Dim Borrar As Integer
    i = dgPacientes.Row
    Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE el PACIENTE: " & dgPacientes.TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que las Consultas Asociadas al Paciente también serán ELIMINADAS", vbQuestion + vbYesNo, "BORRANDO PACIENTE")
    If Borrar = 6 Then
        Call Buscar(dgPacientes.TextMatrix(i, 1), rstEliminarPaciente, "PACIENTES", "pkPaciente")
        rstEliminarPaciente.Delete
        Set rstEliminarPaciente = Nothing
        ConfigurarPacientes
        CargarPacientes
    End If
    Borrar = 0
    i = ""
End Sub

Private Sub cmdIngresarConsulta_Click()
    Dim i As Integer
    i = dgPacientes.Row
    Call SetRecordset(rstDatosPacienteConsulta, "Select * From PACIENTES Where DNI = " & "'" & dgPacientes.TextMatrix(i, 1) & "'")
    Call SetRecordset(rstDatosObrasSocialesConsulta, "Select * From AFILIACIONES Where DNI = " & "'" & dgPacientes.TextMatrix(i, 1) & "'")
    i = 0
    Unload ListadoPacientes
    Consultas.Show
End Sub

Private Sub Form_Load()
    Call CenterMe(ListadoPacientes)
    With ListadoPacientes
        .Width = 10000
        .Height = 7000
    End With
    ConfigurarPacientes
    CargarPacientes
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
