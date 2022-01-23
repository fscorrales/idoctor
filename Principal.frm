VERSION 5.00
Begin VB.MDIForm Principal 
   BackColor       =   &H8000000C&
   Caption         =   "IDoctor"
   ClientHeight    =   3195
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   4680
   Icon            =   "Principal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu MnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu MnuIngresarConsulta 
         Caption         =   "Ingresar Consulta"
      End
      Begin VB.Menu Line02 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPacientes 
         Caption         =   "Pacientes"
         Begin VB.Menu MnuNuevoPaciente 
            Caption         =   "Nuevo Paciente"
         End
         Begin VB.Menu MnuListadoPacientes 
            Caption         =   "Listado Pacientes"
         End
      End
      Begin VB.Menu MnuDiagnosticos 
         Caption         =   "Diagnosticos"
         Begin VB.Menu MnuNuevoDiagnositco 
            Caption         =   "Nuevo Diagnostico"
         End
         Begin VB.Menu MnuListadoDiagnosticos 
            Caption         =   "Listado Diagnosticos"
         End
      End
      Begin VB.Menu MnuObrasSociales 
         Caption         =   "Obras Sociales"
         Begin VB.Menu MnuNuevaObraSocial 
            Caption         =   "Nueva Obra Social"
         End
         Begin VB.Menu MnuListadoObrasSociales 
            Caption         =   "Listado Obras Sociales"
         End
      End
      Begin VB.Menu MnuLocalidades 
         Caption         =   "Localidades"
         Begin VB.Menu MnuNuevaLocalidad 
            Caption         =   "Nueva Localidad"
         End
         Begin VB.Menu MnuListadoLocalidades 
            Caption         =   "Listado Localidades"
         End
      End
      Begin VB.Menu MnuHistorial 
         Caption         =   "Historial"
         Begin VB.Menu MnuNuevoTipoHistorial 
            Caption         =   "Nuevo Tipo Historial"
         End
         Begin VB.Menu MnuListadoTipoHistorial 
            Caption         =   "Listado Tipo Historial"
         End
      End
      Begin VB.Menu Line01 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu MnuReportes 
      Caption         =   "&Reportes"
      Begin VB.Menu MnuReporteGeneralPaciente 
         Caption         =   "Reporte General Por Paciente"
      End
   End
   Begin VB.Menu MnuAdministrador 
      Caption         =   "&Administrador"
      Begin VB.Menu MnuCargaYRecuperoBD 
         Caption         =   "Carga y Recupero BD"
      End
   End
End
Attribute VB_Name = "Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    DesconectarRecordset
    BlanquearVariables
    Call Conexion(dbBase, App.Path & "\" & BaseDeDatos())
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    DesconectarRecordset
End Sub

Private Sub MnuCargaYRecuperoBD_Click()
    dbBase.Close
    CargaBD.Show
End Sub

Private Sub MnuIngresarConsulta_Click()
    ListadoConsultas.Show
End Sub

Private Sub MnuListadoDiagnosticos_Click()
    ListadoDiagnosticos.Show
End Sub

Private Sub MnuListadoLocalidades_Click()
    ListadoLocalidades.Show
End Sub

Private Sub MnuListadoObrasSociales_Click()
    ListadoObrasSociales.Show
End Sub

Private Sub MnuListadoPacientes_Click()
    ListadoPacientes.Show
End Sub

Private Sub MnuListadoTipoHistorial_Click()
    ListadoTipoHistorial.Show
End Sub

Private Sub MnuNuevaLocalidad_Click()
    Localidades.Show
End Sub

Private Sub MnuNuevaObraSocial_Click()
    ObrasSociales.Show
End Sub

Private Sub MnuNuevoDiagnositco_Click()
    Diagnosticos.Show
End Sub

Private Sub MnuNuevoPaciente_Click()
    Pacientes.Show
End Sub

Private Sub MnuNuevoTipoHistorial_Click()
    TipoHistorial.Show
End Sub

Private Sub MnuReporteGeneralPaciente_Click()
    ReporteGeneralPaciente.Show
End Sub

Private Sub MnuSalir_Click()
    dbBase.Close
    End
End Sub
