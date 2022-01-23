VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoLocalidades 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Localidades"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   4605
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
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar Localidad Seleccionada"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   4095
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar Localidad Seleccionada"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   4095
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar Nueva Localidad"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Localidades"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgLocalidades 
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
Attribute VB_Name = "ListadoLocalidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregar_Click()
    Unload ListadoLocalidades
    Localidades.Show
End Sub

Private Sub cmdEditar_Click()
    Dim i As String
    i = dgLocalidades.Row
    Call SetRecordset(rstEditarLocalidad, "Select * from LOCALIDADES Where Localidad = " & "'" & dgLocalidades.TextMatrix(i, 0) & "'")
    Unload ListadoLocalidades
    With Localidades
        .Show
        .txtLocalidad.Text = rstEditarLocalidad.Fields("Localidad")
    End With
    i = ""
    bolEditandoLocalidad = True

End Sub

Private Sub cmdEliminar_Click()
    Dim i As String
    Dim Borrar As Integer
    i = dgLocalidades.Row
    Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE la Localidad: " & dgLocalidades.TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que todos los Pacientes de esa Localidad junto a sus Consultas también serán ELIMINADOS", vbQuestion + vbYesNo, "BORRANDO LOCALIDAD")
    If Borrar = 6 Then
        Call Buscar(dgLocalidades.TextMatrix(i, 0), rstEliminarLocalidad, "LOCALIDADES", "pkLocalidad")
        rstEliminarLocalidad.Delete
        Set rstEliminarLocalidad = Nothing
        ConfigurarLocalidades
        CargarLocalidades
    End If
    Borrar = 0
    i = ""
End Sub

Private Sub Form_Load()
    Call CenterMe(ListadoLocalidades)
    With ListadoLocalidades
        .Width = 4700
        .Height = 7700
    End With
    ConfigurarLocalidades
    CargarLocalidades
End Sub

Sub ConfigurarLocalidades()
    With dgLocalidades
        .Clear
        .Cols = 1
        .Rows = 2
        .TextMatrix(0, 0) = "Localidades"
        .ColWidth(0) = 4000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
    End With
End Sub

Sub CargarLocalidades()
    Dim i As Integer
    i = 0
    dgLocalidades.Rows = 2
    Call SetRecordset(rstListadoLocalidades, "Select * From LOCALIDADES Order by Localidad")
    If rstListadoLocalidades.BOF = False Then
        With rstListadoLocalidades
            .MoveFirst
            While .EOF = False
                i = i + 1
                dgLocalidades.RowHeight(i) = 300
                dgLocalidades.TextMatrix(i, 0) = .Fields("Localidad")
                .MoveNext
                dgLocalidades.Rows = dgLocalidades.Rows + 1
            Wend
        End With
        dgLocalidades.Rows = dgLocalidades.Rows - 1
    End If
    Set rstListadoLocalidades = Nothing
End Sub

