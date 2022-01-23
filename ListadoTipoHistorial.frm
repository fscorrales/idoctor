VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoTipoHistorial 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipo Historial"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
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
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   4335
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar Nuevo Tipo Historial"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   4095
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar Tipo Historial Seleccionado"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   4095
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar Tipo Historial Seleccionado"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Tipo Historial"
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
      Width           =   4335
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgTipoHistorial 
         Height          =   4455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   7858
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
End
Attribute VB_Name = "ListadoTipoHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer

Private Sub cmdAgregar_Click()
    Unload ListadoTipoHistorial
    TipoHistorial.Show
End Sub

Private Sub cmdEditar_Click()
    Dim i As String
    i = dgTipoHistorial.Row
    If dgTipoHistorial.TextMatrix(i, 0) <> "SECCION" Then
        Call SetRecordset(rstEditarTipoHistorial, "Select * from TIPOHISTORIAL Where CodigoHistorial = " & "'" & dgTipoHistorial.TextMatrix(i, 0) & "'")
        Unload ListadoTipoHistorial
        With TipoHistorial
            .Show
            .txtCodigoHistorial.Text = Convertir(rstEditarTipoHistorial.Fields("CodigoHistorial"))
            .txtCodigoHistorial.Enabled = False
            .txtDescripcion.Text = rstEditarTipoHistorial.Fields("Descripcion")
            .txtTipoDato.Text = rstEditarTipoHistorial.Fields("TipoDato")
        End With
    End If
    i = ""
    bolEditandoTipoHistorial = True
End Sub

Private Sub cmdEliminar_Click()
    Dim i As String
    Dim DescripcionConvertida As String
    Dim Borrar As Integer
    i = dgTipoHistorial.Row
    DescripcionConvertida = Right(LTrim(dgTipoHistorial.TextMatrix(i, 1)), Len(LTrim(dgTipoHistorial.TextMatrix(i, 1))) - 2)
    If dgTipoHistorial.TextMatrix(i, 0) <> "SECCION" Then
        Call SetRecordset(rstComprobarEliminacionTipoHistorial, "Select * From TipoHistorial Where Jerarquia = " & "'" & DescripcionConvertida & "'")
        If rstComprobarEliminacionTipoHistorial.BOF = False Then
            MsgBox "La variable " & DescripcionConvertida & " es de tipo PRINCIPAL y no podrá ser elimanada en tanto tenga contenido.", vbOKOnly + vbCritical, "IMPOSIBLE PROCEDER A LA ELIMINACIÓN"
        Else
            Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la VARIABLE: " & DescripcionConvertida & "?" & vbCrLf & "Tenga en cuenta que todos los datos cargados en esta variable serán ELIMINADOS", vbQuestion + vbYesNo, "BORRANDO TIPO HISTORIAL")
            If Borrar = 6 Then
                Call Buscar(dgTipoHistorial.TextMatrix(i, 0), rstEliminarTipoHistorial, "TIPOHISTORIAL", "pkTipoHistorial")
                rstEliminarTipoHistorial.Delete
                Set rstEliminarTipoHistorial = Nothing
                ConfigurarTipoHistorial
                CargarLaboratorio
                CargarVacunas
                CargarTratamiento
                CargarProfilaxis
                CargarOtros
            End If
        End If
    Set rstComprobarEliminacionTipoHistorial = Nothing
    Borrar = 0
    i = ""
    End If
End Sub

Private Sub Form_Load()
    Call CenterMe(ListadoTipoHistorial)
    With ListadoTipoHistorial
        .Width = 4700
        .Height = 7700
    End With
    ConfigurarTipoHistorial
    CargarLaboratorio
    CargarVacunas
    CargarTratamiento
    CargarProfilaxis
    CargarOtros
End Sub

Sub ConfigurarTipoHistorial()
    With dgTipoHistorial
        .Clear
        .Cols = 3
        .Rows = 2
        .TextMatrix(0, 0) = "Código Historial"
        .TextMatrix(0, 1) = "Descripción"
        .TextMatrix(0, 2) = "Datos"
        .ColWidth(0) = 1
        .ColWidth(1) = 3000
        .ColWidth(2) = 1000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
        .ColAlignment(0) = 1
        .ColAlignment(0) = 1
    End With
End Sub

Sub CargarLaboratorio()
    i = 1
    dgTipoHistorial.Rows = 2
    dgTipoHistorial.RowHeight(i) = 300
    dgTipoHistorial.TextMatrix(i, 0) = "SECCION"
    dgTipoHistorial.TextMatrix(i, 1) = "LABORATORIO"
    dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
    Call SetRecordset(rstListadoHistorialPrincipal, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'L###' AND Jerarquia = 'Principal' ORDER BY CodigoHistorial")
    If rstListadoHistorialPrincipal.BOF = False Then
        With rstListadoHistorialPrincipal
            .MoveFirst
            While .EOF = False
                i = i + 1
                dgTipoHistorial.RowHeight(i) = 300
                dgTipoHistorial.TextMatrix(i, 0) = .Fields("CodigoHistorial")
                dgTipoHistorial.TextMatrix(i, 1) = " + " & .Fields("Descripcion")
                dgTipoHistorial.TextMatrix(i, 2) = .Fields("TipoDato")
                dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
                Call SetRecordset(rstListadoHistorialAccesorio, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'L###' AND Jerarquia = " & "'" & rstListadoHistorialPrincipal.Fields("Descripcion") & "'")
                If rstListadoHistorialAccesorio.BOF = False Then
                    With rstListadoHistorialAccesorio
                        .MoveFirst
                        While .EOF = False
                            i = i + 1
                            dgTipoHistorial.RowHeight(i) = 300
                            dgTipoHistorial.TextMatrix(i, 0) = .Fields("CodigoHistorial")
                            dgTipoHistorial.TextMatrix(i, 1) = "     - " & .Fields("Descripcion")
                            dgTipoHistorial.TextMatrix(i, 2) = .Fields("TipoDato")
                            dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
                            .MoveNext
                        Wend
                    End With
                End If
                .MoveNext
            Wend
        End With
    End If
    dgTipoHistorial.Rows = dgTipoHistorial.Rows - 1
    Set rstListadoHistorialPrincipal = Nothing
    Set rstListadoHistorialAccesorio = Nothing
End Sub

Sub CargarVacunas()
    dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
    i = i + 1
    dgTipoHistorial.RowHeight(i) = 300
    dgTipoHistorial.TextMatrix(i, 0) = "SECCION"
    dgTipoHistorial.TextMatrix(i, 1) = "VACUNAS"
    dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
    Call SetRecordset(rstListadoHistorialPrincipal, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'V###' AND Jerarquia = 'Principal' ORDER BY CodigoHistorial")
    If rstListadoHistorialPrincipal.BOF = False Then
        With rstListadoHistorialPrincipal
            .MoveFirst
            While .EOF = False
                i = i + 1
                dgTipoHistorial.RowHeight(i) = 300
                dgTipoHistorial.TextMatrix(i, 0) = .Fields("CodigoHistorial")
                dgTipoHistorial.TextMatrix(i, 1) = " + " & .Fields("Descripcion")
                dgTipoHistorial.TextMatrix(i, 2) = .Fields("TipoDato")
                dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
                Call SetRecordset(rstListadoHistorialAccesorio, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'V###' AND Jerarquia = " & "'" & rstListadoHistorialPrincipal.Fields("Descripcion") & "'")
                If rstListadoHistorialAccesorio.BOF = False Then
                    With rstListadoHistorialAccesorio
                        .MoveFirst
                        While .EOF = False
                            i = i + 1
                            dgTipoHistorial.RowHeight(i) = 300
                            dgTipoHistorial.TextMatrix(i, 0) = .Fields("CodigoHistorial")
                            dgTipoHistorial.TextMatrix(i, 1) = "     - " & .Fields("Descripcion")
                            dgTipoHistorial.TextMatrix(i, 2) = .Fields("TipoDato")
                            dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
                            .MoveNext
                        Wend
                    End With
                End If
                .MoveNext
            Wend
        End With
    End If
    dgTipoHistorial.Rows = dgTipoHistorial.Rows - 1
    Set rstListadoHistorialPrincipal = Nothing
    Set rstListadoHistorialAccesorio = Nothing
End Sub

Sub CargarTratamiento()
    dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
    i = i + 1
    dgTipoHistorial.RowHeight(i) = 300
    dgTipoHistorial.TextMatrix(i, 0) = "SECCION"
    dgTipoHistorial.TextMatrix(i, 1) = "TRATAMIENTO"
    dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
    Call SetRecordset(rstListadoHistorialPrincipal, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'T###' AND Jerarquia = 'Principal' ORDER BY CodigoHistorial")
    If rstListadoHistorialPrincipal.BOF = False Then
        With rstListadoHistorialPrincipal
            .MoveFirst
            While .EOF = False
                i = i + 1
                dgTipoHistorial.RowHeight(i) = 300
                dgTipoHistorial.TextMatrix(i, 0) = .Fields("CodigoHistorial")
                dgTipoHistorial.TextMatrix(i, 1) = " + " & .Fields("Descripcion")
                dgTipoHistorial.TextMatrix(i, 2) = .Fields("TipoDato")
                dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
                Call SetRecordset(rstListadoHistorialAccesorio, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'T###' AND Jerarquia = " & "'" & rstListadoHistorialPrincipal.Fields("Descripcion") & "'")
                If rstListadoHistorialAccesorio.BOF = False Then
                    With rstListadoHistorialAccesorio
                        .MoveFirst
                        While .EOF = False
                            i = i + 1
                            dgTipoHistorial.RowHeight(i) = 300
                            dgTipoHistorial.TextMatrix(i, 0) = .Fields("CodigoHistorial")
                            dgTipoHistorial.TextMatrix(i, 1) = "     - " & .Fields("Descripcion")
                            dgTipoHistorial.TextMatrix(i, 2) = .Fields("TipoDato")
                            dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
                            .MoveNext
                        Wend
                    End With
                End If
                .MoveNext
            Wend
        End With
    End If
    dgTipoHistorial.Rows = dgTipoHistorial.Rows - 1
    Set rstListadoHistorialPrincipal = Nothing
    Set rstListadoHistorialAccesorio = Nothing
End Sub

Sub CargarProfilaxis()
    dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
    i = i + 1
    dgTipoHistorial.RowHeight(i) = 300
    dgTipoHistorial.TextMatrix(i, 0) = "SECCION"
    dgTipoHistorial.TextMatrix(i, 1) = "PROFILAXIS"
    dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
    Call SetRecordset(rstListadoHistorialPrincipal, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'P###' AND Jerarquia = 'Principal' ORDER BY CodigoHistorial")
    If rstListadoHistorialPrincipal.BOF = False Then
        With rstListadoHistorialPrincipal
            .MoveFirst
            While .EOF = False
                i = i + 1
                dgTipoHistorial.RowHeight(i) = 300
                dgTipoHistorial.TextMatrix(i, 0) = .Fields("CodigoHistorial")
                dgTipoHistorial.TextMatrix(i, 1) = " + " & .Fields("Descripcion")
                dgTipoHistorial.TextMatrix(i, 2) = .Fields("TipoDato")
                dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
                Call SetRecordset(rstListadoHistorialAccesorio, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'P###' AND Jerarquia = " & "'" & rstListadoHistorialPrincipal.Fields("Descripcion") & "'")
                If rstListadoHistorialAccesorio.BOF = False Then
                    With rstListadoHistorialAccesorio
                        .MoveFirst
                        While .EOF = False
                            i = i + 1
                            dgTipoHistorial.RowHeight(i) = 300
                            dgTipoHistorial.TextMatrix(i, 0) = .Fields("CodigoHistorial")
                            dgTipoHistorial.TextMatrix(i, 1) = "     - " & .Fields("Descripcion")
                            dgTipoHistorial.TextMatrix(i, 2) = .Fields("TipoDato")
                            dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
                            .MoveNext
                        Wend
                    End With
                End If
                .MoveNext
            Wend
        End With
    End If
    dgTipoHistorial.Rows = dgTipoHistorial.Rows - 1
    Set rstListadoHistorialPrincipal = Nothing
    Set rstListadoHistorialAccesorio = Nothing
End Sub

Sub CargarOtros()
    dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
    i = i + 1
    dgTipoHistorial.RowHeight(i) = 300
    dgTipoHistorial.TextMatrix(i, 0) = "SECCION"
    dgTipoHistorial.TextMatrix(i, 1) = "OTROS"
    dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
    Call SetRecordset(rstListadoHistorialPrincipal, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'O###' AND Jerarquia = 'Principal' ORDER BY CodigoHistorial")
    If rstListadoHistorialPrincipal.BOF = False Then
        With rstListadoHistorialPrincipal
            .MoveFirst
            While .EOF = False
                i = i + 1
                dgTipoHistorial.RowHeight(i) = 300
                dgTipoHistorial.TextMatrix(i, 0) = .Fields("CodigoHistorial")
                dgTipoHistorial.TextMatrix(i, 1) = " + " & .Fields("Descripcion")
                dgTipoHistorial.TextMatrix(i, 2) = .Fields("TipoDato")
                dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
                Call SetRecordset(rstListadoHistorialAccesorio, "SELECT * FROM TipoHistorial Where CodigoHistorial LIKE 'O###' AND Jerarquia = " & "'" & rstListadoHistorialPrincipal.Fields("Descripcion") & "'")
                If rstListadoHistorialAccesorio.BOF = False Then
                    With rstListadoHistorialAccesorio
                        .MoveFirst
                        While .EOF = False
                            i = i + 1
                            dgTipoHistorial.RowHeight(i) = 300
                            dgTipoHistorial.TextMatrix(i, 0) = .Fields("CodigoHistorial")
                            dgTipoHistorial.TextMatrix(i, 1) = "     - " & .Fields("Descripcion")
                            dgTipoHistorial.TextMatrix(i, 2) = .Fields("TipoDato")
                            dgTipoHistorial.Rows = dgTipoHistorial.Rows + 1
                            .MoveNext
                        Wend
                    End With
                End If
                .MoveNext
            Wend
        End With
    End If
    dgTipoHistorial.Rows = dgTipoHistorial.Rows - 1
    Set rstListadoHistorialPrincipal = Nothing
    Set rstListadoHistorialAccesorio = Nothing
End Sub

Private Function Convertir(Codigo As String) As String
    Select Case Left(Codigo, 1)
    Case Is = "L"
        Convertir = "Laboratorio"
    Case Is = "V"
        Convertir = "Vacunas"
    Case Is = "T"
        Convertir = "Tratamiento"
    Case Is = "P"
        Convertir = "Profilaxis"
    Case Is = "O"
        Convertir = "Otros"
    End Select
End Function
