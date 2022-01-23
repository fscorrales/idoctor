VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ReporteGeneralPaciente 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reporte General Por Paciente"
   ClientHeight    =   10035
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frLaboratorio 
      Caption         =   "Tratamiento y Profilaxis"
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
      Height          =   2895
      Left            =   120
      TabIndex        =   8
      Top             =   7080
      Width           =   9615
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgTratamientoYProfilaxis 
         Height          =   2400
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   4233
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Vacunas y Otros"
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
      Height          =   2895
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   9615
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgVacunasYOtros 
         Height          =   2400
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   4233
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selección del Paciente a Consultar"
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
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9615
      Begin VB.CommandButton cmdEjecutar 
         Caption         =   "Ejecutar Filtro de DATOS"
         Height          =   375
         Left            =   6600
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
      Begin VB.ComboBox txtPacientes 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   5055
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
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Laboratorio"
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
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   9615
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgLaboratorio 
         Height          =   2400
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9360
         _ExtentX        =   16510
         _ExtentY        =   4233
         _Version        =   393216
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ReporteGeneralPaciente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DNI As String
Dim NumeroColumnas As Integer
Dim i As Integer

Private Sub cmdEjecutar_Click()
    DNI = InStrRev(txtPacientes.Text, "(", , vbTextCompare)
    If DNI <> "0" Then
        DNI = Mid(txtPacientes.Text, DNI + 1)
        DNI = Left(DNI, Len(DNI) - 1)
    End If
    ConfigurarLaboratorio
    CargarLaboratorio
    ConfigurarVacunasYotros
    CargarVacunas
    CargarOtros
    ConfigurarTratamientoYProfilaxis
    CargarTratamiento
    CargarProfilaxis
End Sub

Private Sub Form_Load()
    With ReporteGeneralPaciente
        .Width = 10000
        .Height = 10400
    End With
    Call CenterMe(ReporteGeneralPaciente)
    CargarPacientes
    ConfigurarLaboratorio
    CargarLaboratorio
    ConfigurarVacunasYotros
    CargarVacunas
    CargarOtros
    ConfigurarTratamientoYProfilaxis
    CargarTratamiento
    CargarProfilaxis
End Sub

Sub CargarPacientes()
    Call SetRecordset(rstReporteGeneralPacientes, "Select Apellido, Nombre, DNI From PACIENTES Order By Apellido ASC")
    If rstReporteGeneralPacientes.BOF = False Then
        With rstReporteGeneralPacientes
        .MoveFirst
        While .EOF = False
            txtPacientes.AddItem .Fields("Apellido") & ", " & .Fields("Nombre") & " (" & .Fields("DNI") & ")"
            .MoveNext
        Wend
        End With
    End If
    Set rstReporteGeneralPacientes = Nothing
End Sub

Sub ConfigurarLaboratorio()
    Dim X As Integer
    Call SetRecordset(rstReporteGeneralPacientes, "Select FechaDesde From Ingresos Inner Join Historial On Ingresos.NumeroIngreso = Historial.NumeroIngreso Where Ingresos.DNI = " & "'" & DNI & "' And CodigoHistorial LIKE 'L###' Group by FechaDesde")
    NumeroColumnas = rstReporteGeneralPacientes.RecordCount
    With dgLaboratorio
        .Clear
        .Cols = (NumeroColumnas + 2)
        .Rows = 2
        .TextMatrix(0, 0) = "Código Historial"
        .TextMatrix(0, 1) = "Descripción"
        .ColWidth(0) = 1
        .ColWidth(1) = 2500
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .AllowUserResizing = flexResizeColumns
        If rstReporteGeneralPacientes.BOF = False Then
            rstReporteGeneralPacientes.MoveFirst
            .FixedCols = 2
            X = 2
            While rstReporteGeneralPacientes.EOF = False
                .TextMatrix(0, X) = rstReporteGeneralPacientes.Fields("FechaDesde")
                .ColWidth(X) = 1000
                .ColAlignment(X) = 7
                X = X + 1
                rstReporteGeneralPacientes.MoveNext
            Wend
        End If
    End With
    Set rstReporteGeneralPacientes = Nothing
    X = 0
End Sub

Sub CargarLaboratorio()
    Dim i As Integer
    i = 1
    dgLaboratorio.Rows = 2
    dgLaboratorio.RowHeight(i) = 300
    dgLaboratorio.TextMatrix(i, 0) = "SECCION"
    dgLaboratorio.TextMatrix(i, 1) = "LABORATORIO"
    dgLaboratorio.Rows = dgLaboratorio.Rows + 1
    Call SetRecordset(rstListadoHistorialPrincipal, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'L###' AND Jerarquia = 'Principal' ORDER BY CodigoHistorial")
    If rstListadoHistorialPrincipal.BOF = False Then
        With rstListadoHistorialPrincipal
            .MoveFirst
            While .EOF = False
                i = i + 1
                dgLaboratorio.RowHeight(i) = 300
                dgLaboratorio.TextMatrix(i, 0) = .Fields("CodigoHistorial")
                dgLaboratorio.TextMatrix(i, 1) = " + " & .Fields("Descripcion")
                dgLaboratorio.Rows = dgLaboratorio.Rows + 1
                Call SetRecordset(rstListadoHistorialAccesorio, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'L###' AND Jerarquia = " & "'" & rstListadoHistorialPrincipal.Fields("Descripcion") & "'")
                If rstListadoHistorialAccesorio.BOF = False Then
                    With rstListadoHistorialAccesorio
                        .MoveFirst
                        While .EOF = False
                            i = i + 1
                            dgLaboratorio.RowHeight(i) = 300
                            dgLaboratorio.TextMatrix(i, 0) = .Fields("CodigoHistorial")
                            dgLaboratorio.TextMatrix(i, 1) = "     - " & .Fields("Descripcion")
                            dgLaboratorio.Rows = dgLaboratorio.Rows + 1
                            .MoveNext
                        Wend
                    End With
                End If
                .MoveNext
            Wend
        End With
    End If
    dgLaboratorio.Rows = dgLaboratorio.Rows - 1
    Set rstListadoHistorialPrincipal = Nothing
    Set rstListadoHistorialAccesorio = Nothing
    If NumeroColumnas > 0 Then
        Dim z As Integer
        For z = 2 To (NumeroColumnas + 1)
            With dgLaboratorio
                Call SetRecordset(rstReporteGeneralPacientes, "Select HISTORIAL.* From Ingresos Inner Join Historial On Ingresos.NumeroIngreso = Historial.NumeroIngreso Where Ingresos.DNI = " & "'" & DNI & "' And CodigoHistorial LIKE 'L###' And Year(FechaDesde) = " & Year(.TextMatrix(0, z)) & "And Month(FechaDesde) = " & Month(.TextMatrix(0, z)) & "And Day(FechaDesde) = " & Day(.TextMatrix(0, z)))
                If rstReporteGeneralPacientes.BOF = False Then
                    With rstReporteGeneralPacientes
                        .MoveFirst
                        While .EOF = False
                        Dim X As Integer
                            For X = 1 To (dgLaboratorio.Rows - 1)
                                If dgLaboratorio.TextMatrix(X, 0) = .Fields("CodigoHistorial") Then
                                    dgLaboratorio.TextMatrix(X, z) = .Fields("Dato")
                                    Exit For
                                End If
                            Next X
                            X = 0
                            .MoveNext
                        Wend
                    End With
                End If
            End With
        Next z
        z = 0
        Set rstReporteGeneralPacientes = Nothing
    End If
    NumeroColumnas = 0
End Sub

Sub ConfigurarVacunasYotros()
    Dim X As Integer
    Call SetRecordset(rstReporteGeneralPacientes, "Select FechaDesde From Ingresos Inner Join Historial On Ingresos.NumeroIngreso = Historial.NumeroIngreso Where Ingresos.DNI = " & "'" & DNI & "' And CodigoHistorial LIKE '[V,O]###' Group by FechaDesde")
    NumeroColumnas = rstReporteGeneralPacientes.RecordCount
    With dgVacunasYOtros
        .Clear
        .Cols = (NumeroColumnas + 2)
        .Rows = 2
        .TextMatrix(0, 0) = "Código Historial"
        .TextMatrix(0, 1) = "Descripción"
        .ColWidth(0) = 1
        .ColWidth(1) = 2500
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .AllowUserResizing = flexResizeColumns
        If rstReporteGeneralPacientes.BOF = False Then
            rstReporteGeneralPacientes.MoveFirst
            .FixedCols = 2
            X = 2
            While rstReporteGeneralPacientes.EOF = False
                .TextMatrix(0, X) = rstReporteGeneralPacientes.Fields("FechaDesde")
                .ColWidth(X) = 1000
                .ColAlignment(X) = 7
                X = X + 1
                rstReporteGeneralPacientes.MoveNext
            Wend
        End If
    End With
    Set rstReporteGeneralPacientes = Nothing
    X = 0
End Sub

Sub CargarVacunas()
    i = 1
    dgVacunasYOtros.Rows = 2
    dgVacunasYOtros.RowHeight(i) = 300
    dgVacunasYOtros.TextMatrix(i, 0) = "SECCION"
    dgVacunasYOtros.TextMatrix(i, 1) = "VACUNAS"
    dgVacunasYOtros.Rows = dgVacunasYOtros.Rows + 1
    Call SetRecordset(rstListadoHistorialPrincipal, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'V###' AND Jerarquia = 'Principal' ORDER BY CodigoHistorial")
    If rstListadoHistorialPrincipal.BOF = False Then
        With rstListadoHistorialPrincipal
            .MoveFirst
            While .EOF = False
                i = i + 1
                dgVacunasYOtros.RowHeight(i) = 300
                dgVacunasYOtros.TextMatrix(i, 0) = .Fields("CodigoHistorial")
                dgVacunasYOtros.TextMatrix(i, 1) = " + " & .Fields("Descripcion")
                dgVacunasYOtros.Rows = dgVacunasYOtros.Rows + 1
                Call SetRecordset(rstListadoHistorialAccesorio, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'V###' AND Jerarquia = " & "'" & rstListadoHistorialPrincipal.Fields("Descripcion") & "'")
                If rstListadoHistorialAccesorio.BOF = False Then
                    With rstListadoHistorialAccesorio
                        .MoveFirst
                        While .EOF = False
                            i = i + 1
                            dgVacunasYOtros.RowHeight(i) = 300
                            dgVacunasYOtros.TextMatrix(i, 0) = .Fields("CodigoHistorial")
                            dgVacunasYOtros.TextMatrix(i, 1) = "     - " & .Fields("Descripcion")
                            dgVacunasYOtros.Rows = dgVacunasYOtros.Rows + 1
                            .MoveNext
                        Wend
                    End With
                End If
                .MoveNext
            Wend
        End With
    End If
    dgVacunasYOtros.Rows = dgVacunasYOtros.Rows - 1
    Set rstListadoHistorialPrincipal = Nothing
    Set rstListadoHistorialAccesorio = Nothing
    If NumeroColumnas > 0 Then
        Dim z As Integer
        For z = 2 To (NumeroColumnas + 1)
            With dgVacunasYOtros
                Call SetRecordset(rstReporteGeneralPacientes, "Select HISTORIAL.* From Ingresos Inner Join Historial On Ingresos.NumeroIngreso = Historial.NumeroIngreso Where Ingresos.DNI = " & "'" & DNI & "' And CodigoHistorial LIKE '[V,O]###' And Year(FechaDesde) = " & Year(.TextMatrix(0, z)) & "And Month(FechaDesde) = " & Month(.TextMatrix(0, z)) & "And Day(FechaDesde) = " & Day(.TextMatrix(0, z)))
                If rstReporteGeneralPacientes.BOF = False Then
                    With rstReporteGeneralPacientes
                        .MoveFirst
                        While .EOF = False
                        Dim X As Integer
                            For X = 1 To (dgVacunasYOtros.Rows - 1)
                                If Left(.Fields("CodigoHistorial"), 1) = "V" Then
                                    If dgVacunasYOtros.TextMatrix(X, 0) = .Fields("CodigoHistorial") Then
                                        dgVacunasYOtros.TextMatrix(X, z) = .Fields("Dato")
                                        Exit For
                                    End If
                                End If
                            Next X
                            X = 0
                            .MoveNext
                        Wend
                    End With
                End If
            End With
        Next z
        z = 0
        Set rstReporteGeneralPacientes = Nothing
    End If
End Sub
Sub CargarOtros()
    dgVacunasYOtros.Rows = dgVacunasYOtros.Rows + 1
    i = i + 1
    dgVacunasYOtros.RowHeight(i) = 300
    dgVacunasYOtros.TextMatrix(i, 0) = "SECCION"
    dgVacunasYOtros.TextMatrix(i, 1) = "OTROS"
    dgVacunasYOtros.Rows = dgVacunasYOtros.Rows + 1
    Call SetRecordset(rstListadoHistorialPrincipal, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'O###' AND Jerarquia = 'Principal' ORDER BY CodigoHistorial")
    If rstListadoHistorialPrincipal.BOF = False Then
        With rstListadoHistorialPrincipal
            .MoveFirst
            While .EOF = False
                i = i + 1
                dgVacunasYOtros.RowHeight(i) = 300
                dgVacunasYOtros.TextMatrix(i, 0) = .Fields("CodigoHistorial")
                dgVacunasYOtros.TextMatrix(i, 1) = " + " & .Fields("Descripcion")
                dgVacunasYOtros.Rows = dgVacunasYOtros.Rows + 1
                Call SetRecordset(rstListadoHistorialAccesorio, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'O###' AND Jerarquia = " & "'" & rstListadoHistorialPrincipal.Fields("Descripcion") & "'")
                If rstListadoHistorialAccesorio.BOF = False Then
                    With rstListadoHistorialAccesorio
                        .MoveFirst
                        While .EOF = False
                            i = i + 1
                            dgVacunasYOtros.RowHeight(i) = 300
                            dgVacunasYOtros.TextMatrix(i, 0) = .Fields("CodigoHistorial")
                            dgVacunasYOtros.TextMatrix(i, 1) = "     - " & .Fields("Descripcion")
                            dgVacunasYOtros.Rows = dgVacunasYOtros.Rows + 1
                            .MoveNext
                        Wend
                    End With
                End If
                .MoveNext
            Wend
        End With
    End If
    dgVacunasYOtros.Rows = dgVacunasYOtros.Rows - 1
    Set rstListadoHistorialPrincipal = Nothing
    Set rstListadoHistorialAccesorio = Nothing
    If NumeroColumnas > 0 Then
        Dim z As Integer
        For z = 2 To (NumeroColumnas + 1)
            With dgVacunasYOtros
                Call SetRecordset(rstReporteGeneralPacientes, "Select HISTORIAL.* From Ingresos Inner Join Historial On Ingresos.NumeroIngreso = Historial.NumeroIngreso Where Ingresos.DNI = " & "'" & DNI & "' And CodigoHistorial LIKE '[V,O]###' And Year(FechaDesde) = " & Year(.TextMatrix(0, z)) & "And Month(FechaDesde) = " & Month(.TextMatrix(0, z)) & "And Day(FechaDesde) = " & Day(.TextMatrix(0, z)))
                If rstReporteGeneralPacientes.BOF = False Then
                    With rstReporteGeneralPacientes
                        .MoveFirst
                        While .EOF = False
                        Dim X As Integer
                            For X = 1 To (dgVacunasYOtros.Rows - 1)
                                If Left(.Fields("CodigoHistorial"), 1) = "O" Then
                                    If dgVacunasYOtros.TextMatrix(X, 0) = .Fields("CodigoHistorial") Then
                                        dgVacunasYOtros.TextMatrix(X, z) = .Fields("Dato")
                                        Exit For
                                    End If
                                End If
                            Next X
                            X = 0
                            .MoveNext
                        Wend
                    End With
                End If
            End With
        Next z
        z = 0
        Set rstReporteGeneralPacientes = Nothing
    End If
    NumeroColumnas = 0
End Sub

Sub ConfigurarTratamientoYProfilaxis()
    With dgTratamientoYProfilaxis
        .Clear
        .Cols = 4
        .Rows = 2
        .TextMatrix(0, 0) = "Descripción"
        .TextMatrix(0, 1) = "Desde"
        .TextMatrix(0, 2) = "Hasta"
        .TextMatrix(0, 3) = "Datos"
        .ColWidth(0) = 5100
        .ColWidth(1) = 1250
        .ColWidth(2) = 1250
        .ColWidth(3) = 1450
        .FixedCols = 1
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
    End With
End Sub

Sub CargarTratamiento()
    i = 1
    dgTratamientoYProfilaxis.Rows = 2
    dgTratamientoYProfilaxis.RowHeight(i) = 300
    dgTratamientoYProfilaxis.TextMatrix(i, 0) = "TRATAMIENTO - Fecha Consulta"
    dgTratamientoYProfilaxis.Rows = dgTratamientoYProfilaxis.Rows + 1
    Call SetRecordset(rstReporteGeneralPacientes, "Select  INGRESOS.Fecha, HISTORIAL.FechaDesde, HISTORIAL.FechaHasta, HISTORIAL.Dato, TIPOHISTORIAL.Descripcion, TIPOHISTORIAL.Jerarquia From INGRESOS INNER JOIN (TIPOHISTORIAL INNER JOIN HISTORIAL ON TIPOHISTORIAL.CodigoHistorial = HISTORIAL.CodigoHistorial) ON INGRESOS.NumeroIngreso = HISTORIAL.NumeroIngreso Where Ingresos.DNI = " & "'" & DNI & "' And HISTORIAL.CodigoHistorial LIKE 'T###' ORDER BY Historial.FechaDesde")
    If rstReporteGeneralPacientes.BOF = False Then
        With rstReporteGeneralPacientes
            .MoveFirst
            While .EOF = False
                i = i + 1
                dgTratamientoYProfilaxis.RowHeight(i) = 300
                If .Fields("Jerarquia") = "Principal" Then
                    dgTratamientoYProfilaxis.TextMatrix(i, 0) = " + " & .Fields("Descripcion") & " (" & .Fields("Fecha") & ")"
                Else
                    dgTratamientoYProfilaxis.TextMatrix(i, 0) = " + " & .Fields("Jerarquia") & " - " & .Fields("Descripcion") & " (" & .Fields("Fecha") & ")"
                End If
                dgTratamientoYProfilaxis.TextMatrix(i, 1) = .Fields("FechaDesde")
                If Len(.Fields("FechaHasta")) <> "0" Then
                    dgTratamientoYProfilaxis.TextMatrix(i, 2) = .Fields("FechaHasta")
                Else
                    dgTratamientoYProfilaxis.TextMatrix(i, 2) = ""
                End If
                If Len(.Fields("Dato")) <> "0" Then
                    dgTratamientoYProfilaxis.TextMatrix(i, 3) = .Fields("Dato")
                Else
                    dgTratamientoYProfilaxis.TextMatrix(i, 3) = ""
                End If
                dgTratamientoYProfilaxis.Rows = dgTratamientoYProfilaxis.Rows + 1
                .MoveNext
            Wend
        End With
    End If
    dgTratamientoYProfilaxis.Rows = dgTratamientoYProfilaxis.Rows - 1
    Set rstReporteGeneralPacientes = Nothing
End Sub

Sub CargarProfilaxis()
    dgTratamientoYProfilaxis.Rows = dgTratamientoYProfilaxis.Rows + 1
    i = i + 1
    dgTratamientoYProfilaxis.RowHeight(i) = 300
    dgTratamientoYProfilaxis.TextMatrix(i, 0) = "PROFILAXIS - Fecha Consulta"
    dgTratamientoYProfilaxis.Rows = dgTratamientoYProfilaxis.Rows + 1
    Call SetRecordset(rstReporteGeneralPacientes, "Select  INGRESOS.Fecha, HISTORIAL.FechaDesde, HISTORIAL.FechaHasta, HISTORIAL.Dato, TIPOHISTORIAL.Descripcion, TIPOHISTORIAL.Jerarquia From INGRESOS INNER JOIN (TIPOHISTORIAL INNER JOIN HISTORIAL ON TIPOHISTORIAL.CodigoHistorial = HISTORIAL.CodigoHistorial) ON INGRESOS.NumeroIngreso = HISTORIAL.NumeroIngreso Where Ingresos.DNI = " & "'" & DNI & "' And HISTORIAL.CodigoHistorial LIKE 'P###' ORDER BY Historial.FechaDesde")
    If rstReporteGeneralPacientes.BOF = False Then
        With rstReporteGeneralPacientes
            .MoveFirst
            While .EOF = False
                i = i + 1
                dgTratamientoYProfilaxis.RowHeight(i) = 300
                If .Fields("Jerarquia") = "Principal" Then
                    dgTratamientoYProfilaxis.TextMatrix(i, 0) = " + " & .Fields("Descripcion") & " (" & .Fields("Fecha") & ")"
                Else
                    dgTratamientoYProfilaxis.TextMatrix(i, 0) = " + " & .Fields("Jerarquia") & " - " & .Fields("Descripcion") & " (" & .Fields("Fecha") & ")"
                End If
                dgTratamientoYProfilaxis.TextMatrix(i, 1) = .Fields("FechaDesde")
                If Len(.Fields("FechaHasta")) <> "0" Then
                    dgTratamientoYProfilaxis.TextMatrix(i, 2) = .Fields("FechaHasta")
                Else
                    dgTratamientoYProfilaxis.TextMatrix(i, 2) = ""
                End If
                If Len(.Fields("Dato")) <> "0" Then
                    dgTratamientoYProfilaxis.TextMatrix(i, 3) = .Fields("Dato")
                Else
                    dgTratamientoYProfilaxis.TextMatrix(i, 3) = ""
                End If
                dgTratamientoYProfilaxis.Rows = dgTratamientoYProfilaxis.Rows + 1
                .MoveNext
            Wend
        End With
    End If
    dgTratamientoYProfilaxis.Rows = dgTratamientoYProfilaxis.Rows - 1
    Set rstReporteGeneralPacientes = Nothing
End Sub
