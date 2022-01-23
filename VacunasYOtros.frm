VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form VacunasYOtros 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vacunas y Otros"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frLaboratorio 
      Caption         =   "Carga Datos Vacunas y Otros"
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
      Height          =   4575
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   4335
      Begin VB.TextBox txtEdicion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   1440
         TabIndex        =   9
         Top             =   1800
         Visible         =   0   'False
         Width           =   915
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgVacunasYOtros 
         Height          =   4095
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   7223
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   6240
      Width           =   4335
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Guardar Datos Vacunas y Otros"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   4095
      End
      Begin VB.CommandButton cmdEliminarUno 
         Caption         =   "Eliminar Valor Seleccionado"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   4095
      End
      Begin VB.CommandButton cmdEliminarTodos 
         Caption         =   "Eliminar Todos los Datos Vacunas y Otros"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Consulta"
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
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtApellidoyNombre 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   4095
      End
      Begin VB.TextBox txtFechaConsulta 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   1
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha Consulta"
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
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
   End
End
Attribute VB_Name = "VacunasYOtros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variable para la clase
Dim EditarGrilla As CEditarFlexGrid
Dim i As Integer

Private Sub cmdAgregar_Click()
    Dim X As Integer
    If Validar = True Then
        With dgVacunasYOtros
            If rstDatosCargaHistorial.BOF = False Then
                rstDatosCargaHistorial.MoveFirst
                For X = 1 To (.Rows - 1)
                    If Trim(.TextMatrix(X, 3)) <> "" Then
                        Dim ValorBuscado As String
                        ValorBuscado = Format(.TextMatrix(X, 0), "'&&&&&&&&&&&&&&&&&&&&&&&" _
                        & "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&" _
                        & "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&" _
                        & "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&'")
                        rstDatosCargaHistorial.FindFirst "CodigoHistorial = " & ValorBuscado
                        If rstDatosCargaHistorial.NoMatch = True Then
                            rstDatosCargaHistorial.AddNew
                            rstDatosCargaHistorial.Fields("NumeroIngreso") = rstDatosConsultaHistorial.Fields("NumeroIngreso")
                            rstDatosCargaHistorial.Fields("CodigoHistorial") = .TextMatrix(X, 0)
                        Else
                            rstDatosCargaHistorial.Edit
                        End If
                            rstDatosCargaHistorial.Fields("FechaDesde") = txtFechaConsulta.Text
                            rstDatosCargaHistorial.Fields("Dato") = .TextMatrix(X, 3)
                            rstDatosCargaHistorial.Update
                    End If
                Next X
                ValorBuscado = ""
            Else
                Call SetRecordset(rstCargaHistorial, "HISTORIAL")
                For X = 1 To (.Rows - 1)
                    If Trim(.TextMatrix(X, 3)) <> "" Then
                        rstCargaHistorial.AddNew
                        rstCargaHistorial.Fields("NumeroIngreso") = rstDatosConsultaHistorial.Fields("NumeroIngreso")
                        rstCargaHistorial.Fields("CodigoHistorial") = .TextMatrix(X, 0)
                        rstCargaHistorial.Fields("FechaDesde") = txtFechaConsulta.Text
                        rstCargaHistorial.Fields("Dato") = .TextMatrix(X, 3)
                        rstCargaHistorial.Update
                    End If
                Next X
                Set rstCargaHistorial = Nothing
            End If
        End With
        strBuscarPaciente = rstDatosConsultaHistorial.Fields("DNI")
        Unload VacunasYOtros
        ListadoConsultas.Show
    End If
    X = 0
End Sub

Private Sub cmdEliminarTodos_Click()
    If rstDatosCargaHistorial.BOF = False Then
        Dim Borrar As Integer
        Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE todos los datos de VACUNAS Y OTROS de la presente CONSULTA?", vbQuestion + vbYesNo, "BORRANDO HISTORIAL")
        If Borrar = 6 Then
            With rstDatosCargaHistorial
                .MoveFirst
                Do Until .EOF = True
                    .Delete
                    .MoveNext
                Loop
            End With
            ConfigurarVacunasYotros
            CargarVacunas
            CargarOtros
        End If
        Borrar = 0
    Else
        MsgBox "La consulta NO POSEE datos para eliminar", vbInformation + vbOKOnly, "IMPOSIBLE ELIMINAR"
    End If
End Sub

Private Sub cmdEliminarUno_Click()
    Dim X As Integer
    X = dgVacunasYOtros.Row
    If dgVacunasYOtros.TextMatrix(X, 0) <> "SECCION" Then
        Dim ValorBuscado As String
        ValorBuscado = Format(dgVacunasYOtros.TextMatrix(X, 0), "'&&&&&&&&&&&&&&&&&&&&&&&" _
        & "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&" _
        & "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&" _
        & "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&'")
        rstDatosCargaHistorial.FindFirst "CodigoHistorial = " & ValorBuscado
        If rstDatosCargaHistorial.NoMatch = True Then
            MsgBox "El VALOR que intenta borrar se encuentra vacio" & vbCrLf & "Verifique si la fila seleccionada es la correcta", vbInformation + vbOKOnly, "IMPOSIBLE ELIMINAR"
        Else
            Dim Borrar As Integer
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la VARIABLE: " & dgVacunasYOtros.TextMatrix(X, 1) & " ?", vbQuestion + vbYesNo, "BORRANDO HISTORIAL")
            If Borrar = 6 Then
                rstDatosCargaHistorial.Delete
                ConfigurarVacunasYotros
                CargarVacunas
                CargarOtros
            End If
            Borrar = 0
        End If
        ValorBuscado = ""
    End If
    X = 0
End Sub

Private Sub Form_Load()
    Call CenterMe(VacunasYOtros)
    With VacunasYOtros
        .Width = 4700
        .Height = 8500
    End With
    With rstDatosConsultaHistorial
        txtApellidoyNombre.Text = .Fields("Apellido") & ", " & .Fields("Nombre")
        txtFechaConsulta.Text = .Fields("Fecha")
    End With
    ConfigurarVacunasYotros
    CargarVacunas
    CargarOtros
    'Nueva instancia
    Set EditarGrilla = New CEditarFlexGrid
    
    'Inicia ( se le envia el Flex y el text )
    EditarGrilla.Iniciar dgVacunasYOtros, txtEdicion
End Sub

Sub ConfigurarVacunasYotros()
    With dgVacunasYOtros
        .Clear
        .Cols = 4
        .Rows = 2
        .TextMatrix(0, 0) = "Código Historial"
        .TextMatrix(0, 1) = "Descripción"
        .TextMatrix(0, 2) = "TipoDatos"
        .TextMatrix(0, 3) = "Datos"
        .ColWidth(0) = 1
        .ColWidth(1) = 2500
        .ColWidth(2) = 1
        .ColWidth(3) = 1250
        .FixedCols = 3
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
    End With
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
                dgVacunasYOtros.TextMatrix(i, 2) = .Fields("TipoDato")
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
                            dgVacunasYOtros.TextMatrix(i, 2) = .Fields("TipoDato")
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
    If rstDatosCargaHistorial.BOF = False Then
        With rstDatosCargaHistorial
            .MoveFirst
            While .EOF = False
            Dim X As Integer
                For X = 1 To (dgVacunasYOtros.Rows - 1)
                    If Left(.Fields("CodigoHistorial"), 1) = "V" Then
                        If dgVacunasYOtros.TextMatrix(X, 0) = .Fields("CodigoHistorial") Then
                            dgVacunasYOtros.TextMatrix(X, 3) = .Fields("Dato")
                            Exit For
                        End If
                    End If
                Next X
                X = 0
                .MoveNext
            Wend
        End With
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
                dgVacunasYOtros.TextMatrix(i, 2) = .Fields("TipoDato")
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
                            dgVacunasYOtros.TextMatrix(i, 2) = .Fields("TipoDato")
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
    If rstDatosCargaHistorial.BOF = False Then
        With rstDatosCargaHistorial
            .MoveFirst
            While .EOF = False
            Dim X As Integer
                For X = 1 To (dgVacunasYOtros.Rows - 1)
                    If Left(.Fields("CodigoHistorial"), 1) = "O" Then
                        If dgVacunasYOtros.TextMatrix(X, 0) = .Fields("CodigoHistorial") Then
                            dgVacunasYOtros.TextMatrix(X, 3) = .Fields("Dato")
                            Exit For
                        End If
                    End If
                Next X
                X = 0
                .MoveNext
            Wend
        End With
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    strBuscarPaciente = rstDatosConsultaHistorial.Fields("DNI")
    Set EditarGrilla = Nothing
    Set rstDatosConsultaHistorial = Nothing
    Set rstDatosCargaHistorial = Nothing
    ListadoConsultas.Show
End Sub

Function Validar() As Boolean
    Dim X As Integer
    With dgVacunasYOtros
        For X = 1 To (.Rows - 1)
            If .TextMatrix(X, 0) = "SECCION" Then
                .TextMatrix(X, 3) = ""
            Else
                Select Case .TextMatrix(X, 2)
                Case Is = "Ninguno"
                    If Trim(.TextMatrix(X, 3)) <> "" Then
                        MsgBox "La Variable " & .TextMatrix(X, 1) & " no puedo guardar datos por ser su propiedad igual a NINGUNO", vbOKOnly, "IMPOSIBLE CARGAR DATOS"
                        .TextMatrix(X, 3) = ""
                    End If
                Case Is = "Numero"
                    If Trim(.TextMatrix(X, 3)) <> "" Then
                        If EsNumeroNoVacio(.TextMatrix(X, 3), "30", .TextMatrix(X, 1)) = False Then
                            Validar = False
                            .SetFocus
                            .Row = X
                            .Col = 3
                            Exit Function
                        End If
                    End If
                Case Is = "Texto"
                    If Trim(.TextMatrix(X, 3)) <> "" Then
                        If EsTextoNoVacio(.TextMatrix(X, 3), "30", .TextMatrix(X, 1)) = False Then
                            Validar = False
                            .SetFocus
                            .Row = X
                            .Col = 3
                            Exit Function
                        End If
                    End If
                Case Is = "Fecha"
                    If Trim(.TextMatrix(X, 3)) <> "" Then
                        If EsFechaNoVacio(.TextMatrix(X, 3), "30", .TextMatrix(X, 1)) = False Then
                            Validar = False
                            .SetFocus
                            .Row = X
                            .Col = 3
                            Exit Function
                        End If
                    End If
                End Select
            End If
        Next X
    End With
    X = 0
    Validar = True
    
End Function
