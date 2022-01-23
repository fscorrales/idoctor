VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ListadoDiagnosticos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Diagnósticos"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
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
         Caption         =   "Eliminar Diagnóstico Seleccionado"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   4095
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar Diagnóstico Seleccionado"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   4095
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar Nuevo Diagnóstico"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Listado de Diagnósticos"
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgDiagnosticos 
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
Attribute VB_Name = "ListadoDiagnosticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAgregar_Click()
    Unload ListadoDiagnosticos
    Diagnosticos.Show
End Sub

Private Sub cmdEditar_Click()
    Dim i As String
    i = dgDiagnosticos.Row
    Call SetRecordset(rstEditarDiagnostico, "Select * from DIAGNOSTICOS Where Diagnostico = " & "'" & dgDiagnosticos.TextMatrix(i, 0) & "'")
    Unload ListadoDiagnosticos
    With Diagnosticos
        .Show
        .txtDiagnostico.Text = rstEditarDiagnostico.Fields("Diagnostico")
    End With
    i = ""
    bolEditandoDiagnostico = True

End Sub

Private Sub cmdEliminar_Click()
    Dim i As String
    Dim Borrar As Integer
    i = dgDiagnosticos.Row
    Borrar = MsgBox("Desea Borrar DEFINITIVAMENTE el Diagnóstico: " & dgDiagnosticos.TextMatrix(i, 0) & "?" & vbCrLf & "Tenga en cuenta que las Consultas Asociadas al Diagnóstico también serán ELIMINADAS", vbQuestion + vbYesNo, "BORRANDO DIAGNÓSTICO")
    If Borrar = 6 Then
        Call Buscar(dgDiagnosticos.TextMatrix(i, 0), rstEliminarDiagnostico, "DIAGNOSTICOS", "pkDiagnostico")
        rstEliminarDiagnostico.Delete
        Set rstEliminarDiagnostico = Nothing
        ConfigurarDiagnosticos
        CargarDiagnosticos
    End If
    Borrar = 0
    i = ""
End Sub

Private Sub Form_Load()
    Call CenterMe(ListadoDiagnosticos)
    With ListadoDiagnosticos
        .Width = 4700
        .Height = 7700
    End With
    ConfigurarDiagnosticos
    CargarDiagnosticos
End Sub

Sub ConfigurarDiagnosticos()
    With dgDiagnosticos
        .Clear
        .Cols = 1
        .Rows = 2
        .TextMatrix(0, 0) = "Diagnósticos"
        .ColWidth(0) = 4000
        .FixedCols = 0
        .FocusRect = flexFocusHeavy
        .HighLight = flexHighlightWithFocus
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ColAlignment(0) = 1
    End With
End Sub

Sub CargarDiagnosticos()
    Dim i As Integer
    i = 0
    dgDiagnosticos.Rows = 2
    Call SetRecordset(rstListadoDiagnosticos, "Select * From DIAGNOSTICOS Order by Diagnostico")
    If rstListadoDiagnosticos.BOF = False Then
        With rstListadoDiagnosticos
            .MoveFirst
            While .EOF = False
                i = i + 1
                dgDiagnosticos.RowHeight(i) = 300
                dgDiagnosticos.TextMatrix(i, 0) = .Fields("Diagnostico")
                .MoveNext
                dgDiagnosticos.Rows = dgDiagnosticos.Rows + 1
            Wend
        End With
        dgDiagnosticos.Rows = dgDiagnosticos.Rows - 1
    End If
    Set rstListadoDiagnosticos = Nothing
End Sub



