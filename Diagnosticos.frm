VERSION 5.00
Begin VB.Form Diagnosticos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Diagnósticos"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGuardarDiagnostico 
      Caption         =   "Guardar Diagnóstico"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Diagnóstico"
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
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.TextBox txtDiagnostico 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4095
      End
   End
End
Attribute VB_Name = "Diagnosticos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGuardarDiagnostico_Click()
    If Validar = True Then
        If bolEditandoDiagnostico = False Then
            Call SetRecordset(rstCargaDiagnostico, "DIAGNOSTICOS")
            Call GuardarDiagnostico(rstCargaDiagnostico)
        Else
            Call GuardarDiagnostico(rstEditarDiagnostico)
        End If
        
        If bolCargaDiagnosticoDesdeConsultas = True Then
            strPasajeAConsultas = txtDiagnostico.Text
        End If
        
        Unload Diagnosticos
    End If
End Sub

Private Function Validar() As Boolean
    If EsTextoNoVacio(txtDiagnostico.Text, "50", "DIAGNOSTICO") = False Then
        txtDiagnostico.SetFocus
        Validar = False
        Exit Function
    End If
    If bolEditandoDiagnostico = False Then
        If ValorDuplicado(txtDiagnostico.Text, rstDuplicadoDiagnostico, "DIAGNOSTICOS", "pkDiagnostico") = True Then
            txtDiagnostico.SetFocus
            Validar = False
            Exit Function
        End If
    Else
        If rstEditarDiagnostico.Fields("Diagnostico") <> txtDiagnostico.Text Then
            If ValorDuplicado(txtDiagnostico.Text, rstDuplicadoDiagnostico, "DIAGNOSTICOS", "pkDiagnostico") = True Then
                txtDiagnostico.SetFocus
                Validar = False
                Exit Function
            End If
        End If
    End If
    
    Validar = True
End Function

Private Sub GuardarDiagnostico(Recordset As Recordset)
    With Recordset
        If bolEditandoDiagnostico = False Then
            .AddNew
        Else
            .Edit
        End If
        .Fields("Diagnostico") = Format(txtDiagnostico.Text, ">")
        .Update
    End With
    
    Set Recordset = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bolEditandoDiagnostico = True Then
        bolEditandoDiagnostico = False
        Set rstEditarDiagnostico = Nothing
        ListadoDiagnosticos.Show
    End If
    If bolCargaDiagnosticoDesdeConsultas = False Then
        ListadoDiagnosticos.Show
    Else
        With Consultas
            .Show
            .txtDiagnostico.Clear
            .CargaComboDiagnostico
            .txtDiagnostico.Text = strPasajeAConsultas
            strPasajeAConsultas = ""
            .txtDiagnostico.SetFocus
        End With
        bolCargaDiagnosticoDesdeConsultas = False
    End If
End Sub

Private Sub Form_Load()
    Call CenterMe(Diagnosticos)
    With Diagnosticos
        .Width = 4700
        .Height = 1900
    End With
End Sub
