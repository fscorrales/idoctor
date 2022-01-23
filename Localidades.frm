VERSION 5.00
Begin VB.Form Localidades 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Localidades"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGuardarLocalidad 
      Caption         =   "Guardar Localidad"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ingreso de Localidad"
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
      Begin VB.TextBox txtLocalidad 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4095
      End
   End
End
Attribute VB_Name = "Localidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call CenterMe(Localidades)
    With Localidades
        .Width = 4700
        .Height = 1900
    End With
End Sub

Private Sub cmdGuardarLocalidad_Click()
    If Validar = True Then
        If bolEditandoLocalidad = False Then
            Call SetRecordset(rstCargaLocalidad, "LOCALIDADES")
            Call GuardarLocalidad(rstCargaLocalidad)
        Else
            Call GuardarLocalidad(rstEditarLocalidad)
        End If
        
        If bolCargaLocalidadDesdePacientes = True Then
            strPasajeAPacientes = txtLocalidad.Text
        End If
        
        Unload Localidades
    End If
End Sub

Private Function Validar() As Boolean
    If EsTextoNoVacio(txtLocalidad.Text, "50", "LOCALIDADES") = False Then
        txtLocalidad.SetFocus
        Validar = False
        Exit Function
    End If
    If bolEditandoLocalidad = False Then
        If ValorDuplicado(txtLocalidad.Text, rstDuplicadoLocalidad, "LOCALIDADES", "pkLocalidad") = True Then
            txtLocalidad.SetFocus
            Validar = False
            Exit Function
        End If
    Else
        If rstEditarLocalidad.Fields("Localidad") <> txtLocalidad.Text Then
            If ValorDuplicado(txtLocalidad.Text, rstDuplicadoLocalidad, "LOCALIDADES", "pkLocalidad") = True Then
                txtLocalidad.SetFocus
                Validar = False
                Exit Function
            End If
        End If
    End If
    
    Validar = True
End Function

Private Sub GuardarLocalidad(Recordset As Recordset)
    With Recordset
        If bolEditandoLocalidad = False Then
            .AddNew
        Else
            .Edit
        End If
        .Fields("Localidad") = Format(txtLocalidad.Text, ">")
        .Update
    End With
    
    Set Recordset = Nothing
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bolEditandoLocalidad = True Then
        bolEditandoLocalidad = False
        Set rstEditarLocalidad = Nothing
        ListadoLocalidades.Show
    End If
    If bolCargaLocalidadDesdePacientes = False Then
        ListadoLocalidades.Show
    Else
        With Pacientes
            .Show
            .txtLocalidad.Clear
            .CargaComboLocalidad
            .txtLocalidad.Text = strPasajeAPacientes
            strPasajeAPacientes = ""
            .txtLocalidad.SetFocus
        End With
        bolCargaLocalidadDesdePacientes = False
    End If
End Sub

