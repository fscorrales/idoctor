VERSION 5.00
Begin VB.Form Consultas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingreso de Consulta"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8910
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
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
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   8655
      Begin VB.CommandButton cmdEditarPaciente 
         Caption         =   "Editar Datos Paciente"
         Height          =   375
         Left            =   2280
         TabIndex        =   22
         Top             =   1680
         Width           =   4095
      End
      Begin VB.TextBox txtNacimiento 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtObrasSociales 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtDomicilio 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   840
         Width           =   3735
      End
      Begin VB.TextBox txtApellidoyNombre 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox txtDNI 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtTelefonos 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6840
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Edad"
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
         Left            =   5760
         TabIndex        =   20
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Obras Sociales"
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
         TabIndex        =   19
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfonos"
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
         Left            =   5760
         TabIndex        =   17
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Apellido y Nombre"
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
         TabIndex        =   15
         Top             =   360
         Width           =   1695
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
         Left            =   5760
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Domicilio"
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
         TabIndex        =   13
         Top             =   840
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdGuardarConsulta 
      Caption         =   "Guardar Consulta"
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   4800
      Width           =   4095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de la Consulta"
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
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   8655
      Begin VB.TextBox txtObservaciones 
         Height          =   885
         Left            =   1560
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1320
         Width           =   6975
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
      Begin VB.ComboBox txtDiagnostico 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton cmdAgregarDiagnostico 
         Caption         =   "Agregar Diagnóstico"
         Height          =   375
         Left            =   4440
         TabIndex        =   4
         Top             =   840
         Width           =   4095
      End
      Begin VB.Label Label6 
         Caption         =   "Observaciones"
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
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "Diagnóstico"
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
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha"
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
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Consultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgregarDiagnostico_Click()
    Consultas.Hide
    Diagnosticos.Show
    bolCargaDiagnosticoDesdeConsultas = True
End Sub

Private Sub cmdEditarPaciente_Click()
    Set rstEditarPaciente = rstDatosPacienteConsulta
    Set rstEditarAfiliacion = rstDatosObrasSocialesConsulta
    bolEditandoPaciente = True
    bolEditandoPacienteDesdeConsulta = True
    Unload Consultas
    With Pacientes
        .Show
        .txtApellido.Text = rstEditarPaciente.Fields("Apellido")
        .txtNombre.Text = rstEditarPaciente.Fields("Nombre")
        .txtDNI.Text = rstEditarPaciente.Fields("DNI")
        .txtNacimiento.Text = rstEditarPaciente.Fields("FechaNacimiento")
        If Len(rstEditarPaciente.Fields("Domicilio")) <> "0" Then
            .txtDireccion.Text = rstEditarPaciente.Fields("Domicilio")
        Else
            .txtDireccion.Text = ""
        End If
        If Len(rstEditarPaciente.Fields("Domicilio")) <> "0" Then
            .txtDireccion.Text = rstEditarPaciente.Fields("Domicilio")
        Else
            .txtDireccion.Text = ""
        End If
        .txtLocalidad.Text = rstEditarPaciente.Fields("Localidad")
        If Len(rstEditarPaciente.Fields("Telefono1")) <> "0" Then
            .txtTelefono1.Text = rstEditarPaciente.Fields("Telefono1")
        Else
            .txtTelefono1.Text = ""
        End If
        If Len(rstEditarPaciente.Fields("Telefono2")) <> "0" Then
            .txtTelefono2.Text = rstEditarPaciente.Fields("Telefono2")
        Else
            .txtTelefono2.Text = ""
        End If
        .txtObraSocialPrimaria.Text = rstEditarAfiliacion.Fields("ObraSocial")
        If Len(rstEditarAfiliacion.Fields("NumeroAfiliado")) <> "0" Then
            .txtAfiliadoPrimario.Text = rstEditarAfiliacion.Fields("NumeroAfiliado")
        Else
            .txtAfiliadoPrimario.Text = ""
        End If
        If rstEditarAfiliacion.RecordCount > "1" Then
            rstEditarAfiliacion.MoveNext
            .txtObraSocialSecundaria.Text = rstEditarAfiliacion.Fields("ObraSocial")
            If Len(rstEditarAfiliacion.Fields("NumeroAfiliado")) <> "0" Then
                .txtAfiliadoSecundario.Text = rstEditarAfiliacion.Fields("NumeroAfiliado")
            Else
                .txtAfiliadoSecundario.Text = ""
            End If
            rstEditarAfiliacion.MovePrevious
        End If
    End With
    i = ""
End Sub

Private Sub cmdGuardarConsulta_Click()
    If Validar = True Then
        If bolEditandoConsulta = False Then
            Call SetRecordset(rstCargaConsulta, "Select * From INGRESOS")
            Call GuardarConsulta(rstCargaConsulta)
        Else
            Call GuardarConsulta(rstEditarConsulta)
        End If
        strBuscarPaciente = rstDatosPacienteConsulta.Fields("DNI")
        Unload Consultas
        ListadoConsultas.Show
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bolEditandoPacienteDesdeConsulta = False Then
        Set rstDatosPacienteConsulta = Nothing
        Set rstDatosObrasSocialesConsulta = Nothing
        If bolEditandoConsulta = True Then
            bolEditandoConsulta = False
            Set rstEditarPaciente = Nothing
            ListadoConsultas.Show
            Exit Sub
        End If
        ListadoConsultas.Show
    End If
End Sub

Private Sub Form_Load()
    Call CenterMe(Consultas)
    With Consultas
        .Width = 9000
        .Height = 5600
    End With
    CargarDatosPersonales
    CargaComboDiagnostico
    If bolEditandoConsulta = True Then
        txtFecha.Text = rstEditarConsulta.Fields("Fecha")
        txtDiagnostico.Text = rstEditarConsulta.Fields("Diagnostico")
        If Len(rstEditarConsulta.Fields("Observaciones")) <> "0" Then
            txtObservaciones.Text = rstEditarConsulta.Fields("Observaciones")
        Else
            txtObservaciones.Text = ""
        End If
    End If
End Sub

Public Sub CargarDatosPersonales()
    With rstDatosPacienteConsulta
        txtApellidoyNombre.Text = .Fields("Apellido") & ", " & .Fields("Nombre")
        txtDNI.Text = .Fields("DNI")
        If Len(.Fields("Domicilio")) <> 0 And Len(.Fields("Localidad")) <> 0 Then
            txtDomicilio.Text = .Fields("Domicilio") & " - " & .Fields("Localidad")
        ElseIf Len(.Fields("Localidad")) <> 0 Then
            txtDomicilio.Text = .Fields("Localidad")
        ElseIf Len(.Fields("Domicilio")) <> 0 Then
            txtDomicilio.Text = .Fields("Domicilio")
        Else
            txtDomicilio.Text = ""
        End If
        If Len(.Fields("Telefono1")) <> 0 And Len(.Fields("Telefono2")) <> 0 Then
            txtTelefonos.Text = .Fields("Telefono1") & " y " & .Fields("Telefono2")
        ElseIf Len(.Fields("Telefono1")) <> 0 Then
            txtTelefonos.Text = .Fields("Telefono1")
        ElseIf Len(.Fields("Telefono2")) <> 0 Then
            txtTelefonos.Text = .Fields("Telefono2")
        Else
            txtTelefonos.Text = ""
        End If
        If Len(.Fields("FechaNacimiento")) <> 0 Then
            txtNacimiento.Text = CalcularEdad(.Fields("FechaNacimiento")) & " años" & " (" & .Fields("FechaNacimiento") & ")"
        Else
            txtNacimiento.Text = ""
        End If
    End With
    With rstDatosObrasSocialesConsulta
        txtObrasSociales.Text = .Fields("ObraSocial") & " (" & .Fields("NumeroAfiliado") & ")"
        If rstDatosObrasSocialesConsulta.RecordCount > 1 Then
            .MoveNext
            txtObrasSociales.Text = txtObrasSociales.Text & " - " & .Fields("ObraSocial") & " (" & .Fields("NumeroAfiliado") & ")"
            .MovePrevious
        End If
    End With
End Sub

Sub CargaComboDiagnostico()
    Call SetRecordset(rstComboDiagnostico, "Select * From DIAGNOSTICOS Order by Diagnostico")
    
    If rstComboDiagnostico.BOF = False Then
        With rstComboDiagnostico
        .MoveFirst
            While .EOF = False
                txtDiagnostico.AddItem .Fields("Diagnostico")
                .MoveNext
            Wend
        End With
    End If
    
    Set rstComboDiagnostico = Nothing
End Sub

Private Sub GuardarConsulta(Recordset As Recordset)
    With Recordset
        If bolEditandoConsulta = False Then
            .AddNew
            .Fields("NumeroIngreso") = AsignarNumero("SELECT NumeroIngreso From Ingresos ORDER BY NumeroIngreso")
        Else
            .Edit
        End If
            .Fields("DNI") = txtDNI.Text
            .Fields("Fecha") = txtFecha.Text
            .Fields("Diagnostico") = Format(txtDiagnostico.Text, ">")
            If Trim(txtObservaciones.Text) <> "" Then
                .Fields("Observaciones") = txtObservaciones.Text
            Else
                .Fields("Observaciones") = ""
            End If
        .Update
    End With
    
    Set Recordset = Nothing
    
End Sub

Function Validar() As Boolean
    If EsFechaNoVacio(txtFecha.Text, "20", "Fecha") = False Then
        txtFecha.SetFocus
        Validar = False
        Exit Function
    End If
    If EsTextoNoVacio(txtDiagnostico.Text, "50", "Diagnostico") = False Then
        txtDiagnostico.SetFocus
        Validar = False
        Exit Function
    End If
    If ExisteEnTablaPrincipal(txtDiagnostico.Text, rstExisteDiagnostico, "DIAGNOSTICOS", "pkDiagnostico") = False Then
        txtDiagnostico.SetFocus
        Validar = False
        Exit Function
    End If
    
    Validar = True
    
End Function

Private Function AsignarNumero(SQL As String) As String
    Call SetRecordset(rstAsignarNumero, SQL)
    If rstAsignarNumero.BOF = False Then
        With rstAsignarNumero
            .MoveLast
            AsignarNumero = .Fields("NumeroIngreso") + 1
        End With
    Else
        AsignarNumero = "1"
    End If
    Set rstAsignarNumero = Nothing
End Function
