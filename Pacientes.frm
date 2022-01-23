VERSION 5.00
Begin VB.Form Pacientes 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pacientes"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdGuardarPaciente 
      Caption         =   "Guardar Datos del Paciente"
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   5760
      Width           =   4095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Datos Obras Sociales"
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
      Height          =   2175
      Left            =   120
      TabIndex        =   23
      Top             =   3480
      Width           =   8655
      Begin VB.TextBox txtAfiliadoSecundario 
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   1800
         Width           =   2655
      End
      Begin VB.ComboBox txtObraSocialSecundaria 
         Height          =   315
         Left            =   1560
         TabIndex        =   11
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CommandButton cmdAgregarObraSocialSecundaria 
         Caption         =   "Agregar Obra Social"
         Height          =   375
         Left            =   4440
         TabIndex        =   28
         Top             =   1320
         Width           =   4095
      End
      Begin VB.CommandButton cmdAgregarObraSocialPrimaria 
         Caption         =   "Agregar Obra Social"
         Height          =   375
         Left            =   4440
         TabIndex        =   27
         Top             =   360
         Width           =   4095
      End
      Begin VB.ComboBox txtObraSocialPrimaria 
         Height          =   315
         Left            =   1560
         TabIndex        =   9
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtAfiliadoPrimario 
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label11 
         Caption         =   "Número Afiliado"
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
         TabIndex        =   30
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Secundaria"
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
         TabIndex        =   29
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Primaria"
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
         TabIndex        =   25
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Número Afiliado"
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
         TabIndex        =   24
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Domicilio"
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
      TabIndex        =   18
      Top             =   1560
      Width           =   8655
      Begin VB.CommandButton cmdAgregarLocalidad 
         Caption         =   "Agregar Localidad"
         Height          =   375
         Left            =   4440
         TabIndex        =   26
         Top             =   840
         Width           =   4095
      End
      Begin VB.ComboBox txtLocalidad 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox txtDireccion 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtTelefono1 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtTelefono2 
         Height          =   285
         Left            =   5880
         TabIndex        =   8
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Dirección"
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
         TabIndex        =   22
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Localidad"
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
         TabIndex        =   21
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Teléfono 1"
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
         TabIndex        =   20
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Teléfono 2"
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
         Left            =   4440
         TabIndex        =   19
         Top             =   1320
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Personales"
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
      Width           =   8655
      Begin VB.TextBox txtNacimiento 
         Height          =   285
         Left            =   5880
         TabIndex        =   4
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox txtDNI 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   5880
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtApellido 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Dia Nacimiento"
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
         Left            =   4440
         TabIndex        =   17
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "D.N.I."
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
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
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
         Left            =   4440
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Apellido"
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
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Pacientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregarLocalidad_Click()
    Pacientes.Hide
    Localidades.Show
    bolCargaLocalidadDesdePacientes = True
End Sub

Private Sub cmdAgregarObraSocialPrimaria_Click()
    Pacientes.Hide
    ObrasSociales.Show
    bolCargaObraSocialPrimariaDesdePacientes = True
End Sub

Private Sub cmdAgregarObraSocialSecundaria_Click()
    Pacientes.Hide
    ObrasSociales.Show
    bolCargaObraSocialSecundariaDesdePacientes = True
End Sub

Private Sub cmdGuardarPaciente_Click()
    If Validar = True Then
    
        If bolEditandoPaciente = False Then
            Call SetRecordset(rstCargaPaciente, "PACIENTES")
            Call GuardarPaciente(rstCargaPaciente)
        Else
            Call GuardarPaciente(rstEditarPaciente)
        End If
        
        If bolEditandoPaciente = False Then
            Call SetRecordset(rstCargaAfiliaciones, "AFILIACIONES")
            Call GuardarAfiliaciones(rstCargaAfiliaciones)
        Else
            Call GuardarAfiliaciones(rstEditarAfiliacion)
        End If
          
        Unload Pacientes
    End If
End Sub

Private Sub Form_Load()
    Call CenterMe(Pacientes)
    With Pacientes
        .Width = 9000
        .Height = 6500
    End With
    CargaCombos
End Sub

Sub CargaCombos()
    CargaComboLocalidad
    CargaComboObraSocialPrimaria
    CargaComboObraSocialSecundaria
End Sub

Function Validar() As Boolean
    If EsTextoNoVacio(txtApellido.Text, "50", "APELLIDO") = False Then
        txtApellido.SetFocus
        Validar = False
        Exit Function
    End If
    If EsTextoNoVacio(txtNombre.Text, "50", "NOMBRE") = False Then
        txtNombre.SetFocus
        Validar = False
        Exit Function
    End If
    If EsNumeroNoVacio(txtDNI.Text, "15", "DNI") = False Then
        txtDNI.SetFocus
        Validar = False
        Exit Function
    End If
    If EsFechaNoVacio(txtNacimiento.Text, "20", "FECHA DE NACIMIENTO") = False Then
        txtNacimiento.SetFocus
        Validar = False
        Exit Function
    End If
    If Trim(txtDireccion.Text) <> "" Then
        If EsTextoNoVacio(txtDireccion.Text, "50", "DOMICILIO") = False Then
            txtDireccion.SetFocus
            Validar = False
            Exit Function
        End If
    End If
    If EsTextoNoVacio(txtLocalidad.Text, "50", "LOCALIDAD") = False Then
        txtLocalidad.SetFocus
        Validar = False
        Exit Function
    End If
    If ExisteEnTablaPrincipal(txtLocalidad.Text, rstExisteLocalidad, "LOCALIDADES", "pkLocalidad") = False Then
        txtLocalidad.SetFocus
        Validar = False
        Exit Function
    End If
    If Trim(txtTelefono1.Text) <> "" Then
        If EsNumeroNoVacio(txtTelefono1.Text, "20", "TELÉFONO NRO 1") = False Then
            txtTelefono1.SetFocus
            Validar = False
            Exit Function
        End If
    End If
    If Trim(txtTelefono2.Text) <> "" Then
        If EsNumeroNoVacio(txtTelefono2.Text, "20", "TELÉFONO NRO 2") = False Then
            txtTelefono2.SetFocus
            Validar = False
            Exit Function
        End If
    End If
    If EsTextoNoVacio(txtObraSocialPrimaria.Text, "50", "OBRA SOCIAL PRIMARIA") = False Then
        txtObraSocialPrimaria.SetFocus
        Validar = False
        Exit Function
    End If
    If ExisteEnTablaPrincipal(txtObraSocialPrimaria.Text, rstExisteObraSocial, "OBRASSOCIALES", "pkObraSocial") = False Then
        txtObraSocialPrimaria.SetFocus
        Validar = False
        Exit Function
    End If
    If Trim(txtObraSocialSecundaria.Text) <> "" Then
        If EsTextoNoVacio(txtObraSocialSecundaria.Text, "50", "OBRA SOCIAL SECUNDARIA") = False Then
            txtObraSocialSecundaria.SetFocus
            Validar = False
            Exit Function
        End If
        If ExisteEnTablaPrincipal(txtObraSocialSecundaria.Text, rstExisteObraSocial, "OBRASSOCIALES", "pkObraSocial") = False Then
            txtObraSocialSecundaria.SetFocus
            Validar = False
            Exit Function
        End If
        If txtObraSocialSecundaria.Text = txtObraSocialPrimaria.Text Then
            MsgBox "Las Obras Sociales PRIMARIAS y SECUNDARIAS no pueden ser iguales" & vbCrLf & "INTENTE DE NUEVO", vbCritical + vbOKOnly, "OBRAS SOCIALES INCOMPATIBLES"
            txtObraSocialSecundaria.SetFocus
            Validar = False
            Exit Function
        End If
    End If
    
    If bolEditandoPaciente = False Then
        If ValorDuplicado(txtDNI.Text, rstDuplicadoDNI, "PACIENTES", "pkPaciente") = True Then
            txtDNI.SetFocus
            Validar = False
            Exit Function
        End If
    Else
        If rstEditarPaciente.Fields("DNI") <> txtDNI.Text Then
            If ValorDuplicado(txtDNI.Text, rstDuplicadoDNI, "PACIENTES", "pkPaciente") = True Then
                txtDNI.SetFocus
                Validar = False
                Exit Function
            End If
        End If
    End If
    
    Validar = True
End Function

Private Sub Form_Unload(Cancel As Integer)
    If bolEditandoPaciente = True Then
        bolEditandoPaciente = False
        Set rstEditarPaciente = Nothing
        Set rstEditarAfiliacion = Nothing
        If bolEditandoPacienteDesdeConsulta = False Then
            ListadoPacientes.Show
        Else
            bolEditandoPacienteDesdeConsulta = False
            rstDatosObrasSocialesConsulta.MoveFirst
            Consultas.Show
        End If
        Exit Sub
    End If
    If AgregarPacientesDesdeListadoConsultas = True Then
        AgregarPacientesDesdeListadoConsultas = False
        ListadoConsultas.Show
        Exit Sub
    End If
    ListadoPacientes.Show
End Sub

Private Sub GuardarPaciente(Recordset As Recordset)
    With Recordset
        If bolEditandoPaciente = False Then
            .AddNew
        Else
            .Edit
        End If
        .Fields("DNI") = txtDNI.Text
        .Fields("Apellido") = Format(txtApellido.Text, ">")
        .Fields("Nombre") = Format(txtNombre.Text, ">")
        If Trim(txtNacimiento.Text) <> "" Then
            .Fields("FechaNacimiento") = txtNacimiento.Text
        Else
            .Fields("FechaNacimiento") = Date
        End If
        If Trim(txtDireccion.Text) <> "" Then
            .Fields("Domicilio") = txtDireccion.Text
        Else
            .Fields("Domicilio") = ""
        End If
        .Fields("Localidad") = Format(txtLocalidad.Text, ">")
        If Trim(txtTelefono1.Text) <> "" Then
            .Fields("Telefono1") = txtTelefono1.Text
        Else
            .Fields("Telefono1") = ""
        End If
        If Trim(txtTelefono2.Text) <> "" Then
            .Fields("Telefono2") = txtTelefono2.Text
        Else
            .Fields("Telefono2") = ""
        End If
        .Update
    End With
    
    Set Recordset = Nothing
    
End Sub

Private Sub GuardarAfiliaciones(Recordset As Recordset)
    With Recordset
        If bolEditandoPaciente = False Then
            .AddNew
        Else
            .Edit
        End If
        .Fields("DNI") = txtDNI.Text
        .Fields("ObraSocial") = Format(txtObraSocialPrimaria.Text, ">")
        .Fields("TipoAfiliacion") = "PRIMARIA"
        If Trim(txtAfiliadoPrimario.Text) <> "" Then
            .Fields("NumeroAfiliado") = txtAfiliadoPrimario.Text
        Else
            .Fields("NumeroAfiliado") = ""
        End If
        .Update
    End With
    If Trim(txtObraSocialSecundaria.Text) <> "" Then
        With Recordset
            If .RecordCount = 1 Then
                .AddNew
            Else
                .MoveNext
                .Edit
            End If
            .Fields("DNI") = txtDNI.Text
            .Fields("ObraSocial") = Format(txtObraSocialSecundaria.Text, ">")
            .Fields("TipoAfiliacion") = "SECUNDARIA"
            If Trim(txtAfiliadoSecundario.Text) <> "" Then
                .Fields("NumeroAfiliado") = txtAfiliadoSecundario.Text
            Else
                .Fields("NumeroAfiliado") = ""
            End If
            .Update
        End With
    End If
    
    Set Recordset = Nothing

End Sub

Sub DescargaCombos()
    txtLocalidad.Clear
    txtObraSocialPrimaria.Clear
    txtObraSocialSecundaria.Clear
End Sub

Sub CargaComboLocalidad()
    Call SetRecordset(rstComboLocalidad, "Select * From LOCALIDADES Order by Localidad")
    
    If rstComboLocalidad.BOF = False Then
        With rstComboLocalidad
        .MoveFirst
            While .EOF = False
                txtLocalidad.AddItem .Fields("Localidad")
                .MoveNext
            Wend
        End With
    End If
    
    Set rstComboLocalidad = Nothing
End Sub

Sub CargaComboObraSocialPrimaria()
    Call SetRecordset(rstComboObraSocial, "Select * From OBRASSOCIALES Order by ObraSocial")

    If rstComboObraSocial.BOF = False Then
        With rstComboObraSocial
        .MoveFirst
            While .EOF = False
                txtObraSocialPrimaria.AddItem .Fields("ObraSocial")
                .MoveNext
            Wend
        End With
    End If
    
    Set rstComboObraSocial = Nothing
End Sub

Sub CargaComboObraSocialSecundaria()
    Call SetRecordset(rstComboObraSocial, "Select * From OBRASSOCIALES Order by ObraSocial")

    If rstComboObraSocial.BOF = False Then
        With rstComboObraSocial
        .MoveFirst
            While .EOF = False
                txtObraSocialSecundaria.AddItem .Fields("ObraSocial")
                .MoveNext
            Wend
        End With
    End If
    
    Set rstComboObraSocial = Nothing
End Sub
