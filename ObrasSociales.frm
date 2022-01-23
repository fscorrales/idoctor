VERSION 5.00
Begin VB.Form ObrasSociales 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Obras Sociales"
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGuardarObraSocial 
      Caption         =   "Guardar Obra Social"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Obra Social"
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
      Begin VB.TextBox txtObraSocial 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4095
      End
   End
End
Attribute VB_Name = "ObrasSociales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call CenterMe(ObrasSociales)
    With ObrasSociales
        .Width = 4700
        .Height = 1900
    End With
End Sub

Private Sub cmdGuardarObraSocial_Click()
    If Validar = True Then
        If bolEditandoObraSocial = False Then
            Call SetRecordset(rstCargaObraSocial, "OBRASSOCIALES")
            Call GuardarObraSocial(rstCargaObraSocial)
        Else
            Call GuardarObraSocial(rstEditarObraSocial)
        End If
        
        If bolCargaObraSocialPrimariaDesdePacientes = True Or bolCargaObraSocialSecundariaDesdePacientes = True Then
            strPasajeAPacientes = txtObraSocial.Text
        End If
        
        Unload ObrasSociales
    End If
    

End Sub

Private Function Validar() As Boolean
    If EsTextoNoVacio(txtObraSocial.Text, "50", "OBRASSOCIALES") = False Then
        txtObraSocial.SetFocus
        Validar = False
        Exit Function
    End If
    If bolEditandoObraSocial = False Then
        If ValorDuplicado(txtObraSocial.Text, rstDuplicadoObraSocial, "OBRASSOCIALES", "pkObraSocial") = True Then
            txtObraSocial.SetFocus
            Validar = False
            Exit Function
        End If
    Else
        If rstEditarObraSocial.Fields("ObraSocial") <> txtObraSocial.Text Then
            If ValorDuplicado(txtObraSocial.Text, rstDuplicadoObraSocial, "OBRASSOCIALES", "pkObraSocial") = True Then
                txtObraSocial.SetFocus
                Validar = False
                Exit Function
            End If
        End If
    End If
    
    Validar = True
End Function

Private Sub GuardarObraSocial(Recordset As Recordset)
    With Recordset
        If bolEditandoObraSocial = False Then
            .AddNew
        Else
            .Edit
        End If
        .Fields("ObraSocial") = Format(txtObraSocial.Text, ">")
        .Update
    End With
    
    Set Recordset = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bolEditandoObraSocial = True Then
        bolEditandoObraSocial = False
        Set rstEditarObraSocial = Nothing
        ListadoObrasSociales.Show
    End If
    
    If bolCargaObraSocialPrimariaDesdePacientes = False And bolCargaObraSocialSecundariaDesdePacientes = False Then
        ListadoObrasSociales.Show
    ElseIf bolCargaObraSocialPrimariaDesdePacientes = True Then
        With Pacientes
            .Show
            .txtObraSocialPrimaria.Clear
            .CargaComboObraSocialPrimaria
            .txtObraSocialPrimaria.Text = strPasajeAPacientes
            strPasajeAPacientes = ""
            .txtObraSocialPrimaria.SetFocus
        End With
        bolCargaObraSocialPrimariaDesdePacientes = False
    ElseIf bolCargaObraSocialSecundariaDesdePacientes = True Then
        With Pacientes
            .Show
            .txtObraSocialSecundaria.Clear
            .CargaComboObraSocialSecundaria
            .txtObraSocialSecundaria.Text = strPasajeAPacientes
            strPasajeAPacientes = ""
            .txtObraSocialSecundaria.SetFocus
        End With
        bolCargaObraSocialSecundariaDesdePacientes = False
    End If

End Sub


