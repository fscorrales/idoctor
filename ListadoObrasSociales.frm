VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoObrasSociales 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Obras Sociales"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Listado Obras Sociales"
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
      TabIndex        =   4
      Top             =   120
      Width           =   4335
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgObrasSociales 
         Height          =   4455
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   7858
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
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   4335
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar Nueva Obra Social"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4095
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar Obra Social Seleccionada"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   4095
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar Obra Social Seleccionada"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   4095
      End
   End
End
Attribute VB_Name = "ListadoObrasSociales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    Unload ListadoObrasSociales
    ObrasSociales.Show
End Sub

Private Sub cmdEditar_Click()
    Dim i As String
    i = dgObrasSociales.Row
    Call SetRecordset(rstEditarObraSocial, "Select * from OBRASSOCIALES Where ObraSocial = " & "'" & dgObrasSociales.TextMatrix(i, 0) & "'")
    Unload ListadoObrasSociales
    With ObrasSociales
        .Show
        .txtObraSocial.Text = rstEditarObraSocial.Fields("ObraSocial")
    End With
    i = ""
    bolEditandoObraSocial = True

End Sub

Private Sub cmdEliminar_Click()
    Dim i As String
    Dim Borrar As Integer
    i = dgObrasSociales.Row
    Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la Obra Social: " & dgObrasSociales.TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que los datos de Afiliación a dicha Obra Social también serán ELIMINADOS", vbQuestion + vbYesNo, "BORRANDO OBRA SOCIAL")
    If Borrar = 6 Then
        Call Buscar(dgObrasSociales.TextMatrix(i, 0), rstEliminarObraSocial, "OBRASSOCIALES", "pkObraSocial")
        rstEliminarObraSocial.Delete
        Set rstEliminarObraSocial = Nothing
        ConfigurarObrasSociales
        CargarObrasSociales
    End If
    Borrar = 0
    i = ""
End Sub

Private Sub Form_Load()
    Call CenterMe(ListadoObrasSociales)
    With ListadoObrasSociales
        .Width = 4700
        .Height = 7700
    End With
    ConfigurarObrasSociales
    CargarObrasSociales
End Sub

Sub ConfigurarObrasSociales()
    With dgObrasSociales
        .Clear
        .Cols = 1
        .Rows = 2
        .TextMatrix(0, 0) = "Obras Sociales"
        .ColWidth(0) = 4000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
    End With
End Sub

Sub CargarObrasSociales()
    Dim i As Integer
    i = 0
    dgObrasSociales.Rows = 2
    Call SetRecordset(rstListadoObrasSociales, "Select * From OBRASSOCIALES Order by ObraSocial")
    If rstListadoObrasSociales.BOF = False Then
        With rstListadoObrasSociales
            .MoveFirst
            While .EOF = False
                i = i + 1
                dgObrasSociales.RowHeight(i) = 300
                dgObrasSociales.TextMatrix(i, 0) = .Fields("ObraSocial")
                .MoveNext
                dgObrasSociales.Rows = dgObrasSociales.Rows + 1
            Wend
        End With
        dgObrasSociales.Rows = dgObrasSociales.Rows - 1
    End If
    Set rstListadoObrasSociales = Nothing
End Sub


