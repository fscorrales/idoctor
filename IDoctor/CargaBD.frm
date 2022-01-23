VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.OCX"
Begin VB.Form CargaBD 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5130
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog dlgRescatar 
      Left            =   1560
      Top             =   4560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtInforme 
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton cmdRescatar 
      Caption         =   "&Rescatar"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCrear 
      Caption         =   "&Crear"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   4560
      Width           =   1215
   End
End
Attribute VB_Name = "CargaBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub cmdCrear_Click()
    Dim Respuesta As Integer
    Direccion = (App.Path & "\" & BaseDeDatos())
    cmdCrear.Enabled = False
    
    Call Info("1: Inicio de Creacion de la BD con DAO")
    If FileSystem.Dir(Direccion) <> "" Then ' Si la BD ya existe
        Respuesta = MsgBox("Una base de datos ya Existe, esta seguro que desea REEMPLAZARLA?", 4, "Alerta")
        If Respuesta = 6 Then 'Al aceptar
            FileSystem.Kill (Direccion)
            Call Info(" - Archivo Eliminado - ")
        Else 'Al apretar Cancelar
            MsgBox "Se canceló la operación", 0 + 48
            End
        End If
    Else 'La BD no existe
        Call Info(" - El Archivo NO EXISTE - ")
    End If
    
    Call Info("2: Preparando Base de Datos")
    Dim dbBase As Database
    
    Call Info("3: Preparando Espacio de Trabajo")
    Dim wsEspacio As Workspace
    
    Call Info("4: Preparando Tablas")
    Dim tbdPACIENTES As TableDef
    Dim tbdLOCALIDADES As TableDef
    Dim tbdOBRASSOCIALES As TableDef
    Dim tbdAFILIACIONES As TableDef
    Dim tbdINGRESOS As TableDef
    Dim tbdDIAGNOSTICOS As TableDef
    
    Call Info("5: Preparando Índices")
    Dim idxPacientes As Index
    Dim idxLocalidades As Index
    Dim idxObrasSociales As Index
    Dim idxAfiliaciones As Index
    Dim idxIngresos As Index
    Dim idxDiagnosticos As Index
    
    Call Info("6: Preparando Relaciones")
    Dim relPacientePorLocalidad As Relation
    Dim relAfiliacionPorPaciente As Relation
    Dim relAfiliacionPorObraSocial As Relation
    Dim relIngresoPorPaciente As Relation
    Dim relIngresoPorDiagnostico As Relation
    
    Call Info("7: Activando el Espacio de Trabajo")
    Set wsEspacio = DBEngine.Workspaces(0)
    
    Call Info("8: Generando la Base de Datos")
    Set dbBase = wsEspacio.CreateDatabase(Direccion, dbLangGeneral, dbVersion30)

    Call Info("9: Generando las Tablas")
    Set tbdPACIENTES = dbBase.CreateTableDef("PACIENTES")
    Set tbdLOCALIDADES = dbBase.CreateTableDef("LOCALIDADES")
    Set tbdINGRESOS = dbBase.CreateTableDef("INGRESOS")
    Set tbdOBRASSOCIALES = dbBase.CreateTableDef("OBRASSOCIALES")
    Set tbdAFILIACIONES = dbBase.CreateTableDef("AFILIACIONES")
    Set tbdDIAGNOSTICOS = dbBase.CreateTableDef("DIAGNOSTICOS")

    Call Info("10: Generando Campos e Integrando Tablas")
    
    Call Info("->PACIENTES")
    With tbdPACIENTES
        .Fields.Append .CreateField("DNI", dbText, 15)
        With .Fields("DNI")
            .AllowZeroLength = False
            .Required = True
        End With
        .Fields.Append .CreateField("Apellido", dbText, 50)
        With .Fields("Apellido")
            .AllowZeroLength = False
            .Required = True
        End With
        .Fields.Append .CreateField("Nombre", dbText, 50)
        With .Fields("Nombre")
            .AllowZeroLength = True
            .Required = False
        End With
        .Fields.Append .CreateField("FechaNacimiento", dbDate)
        With .Fields("FechaNacimiento")
            .AllowZeroLength = True
            .Required = False
        End With
        .Fields.Append .CreateField("Domicilio", dbText, 50)
        With .Fields("Domicilio")
            .AllowZeroLength = True
            .Required = False
        End With
        .Fields.Append .CreateField("Localidad", dbText, 50)
        With .Fields("Localidad")
            .AllowZeroLength = False
            .Required = True
        End With
        .Fields.Append .CreateField("Telefono1", dbText, 20)
        With .Fields("Telefono1")
            .AllowZeroLength = True
            .Required = False
        End With
        .Fields.Append .CreateField("Telefono2", dbText, 20)
        With .Fields("Telefono2")
            .AllowZeroLength = True
            .Required = False
        End With
    End With
    
    Call Info("->LOCALIDADES")
    With tbdLOCALIDADES
        .Fields.Append .CreateField("Localidad", dbText, 50)
        With .Fields("Localidad")
            .AllowZeroLength = False
            .Required = True
        End With
    End With
    
    Call Info("->INGRESOS")
    With tbdINGRESOS
        .Fields.Append .CreateField("DNI", dbText, 15)
        With .Fields("DNI")
            .AllowZeroLength = False
            .Required = True
        End With
        .Fields.Append .CreateField("Fecha", dbDate)
        With .Fields("Fecha")
            .AllowZeroLength = False
            .Required = True
        End With
        .Fields.Append .CreateField("Diagnostico", dbText, 50)
        With .Fields("Diagnostico")
            .AllowZeroLength = False
            .Required = True
        End With
        .Fields.Append .CreateField("Observaciones", dbMemo)
        With .Fields("Observaciones")
            .AllowZeroLength = True
            .Required = False
        End With
    End With

    Call Info("->OBRAS SOCIALES")
    With tbdOBRASSOCIALES
        .Fields.Append .CreateField("ObraSocial", dbText, 50)
         With .Fields("ObraSocial")
            .AllowZeroLength = False
            .Required = True
        End With
    End With
    
    Call Info("->AFILIACIONES")
    With tbdAFILIACIONES
        .Fields.Append .CreateField("DNI", dbText, 15)
        With .Fields("DNI")
            .AllowZeroLength = False
            .Required = True
        End With
        .Fields.Append .CreateField("ObraSocial", dbText, 50)
        With .Fields("ObraSocial")
            .AllowZeroLength = False
            .Required = True
        End With
        .Fields.Append .CreateField("TipoAfiliacion", dbText, 10)
        With .Fields("TipoAfiliacion")
            .AllowZeroLength = False
            .Required = True
        End With
        .Fields.Append .CreateField("NumeroAfiliado", dbText, 30)
        With .Fields("NumeroAfiliado")
            .AllowZeroLength = True
            .Required = False
        End With
    End With

    Call Info("->DIAGNOSTICOS")
    With tbdDIAGNOSTICOS
        .Fields.Append .CreateField("Diagnostico", dbText, 50)
         With .Fields("Diagnostico")
            .AllowZeroLength = False
            .Required = True
        End With
    End With
    
    Call Info("11: Integrando la Base de Datos")
    With dbBase
        .TableDefs.Append tbdPACIENTES
        .TableDefs.Append tbdLOCALIDADES
        .TableDefs.Append tbdINGRESOS
        .TableDefs.Append tbdOBRASSOCIALES
        .TableDefs.Append tbdAFILIACIONES
        .TableDefs.Append tbdDIAGNOSTICOS
    End With

    Call Info("12: Creando Índices")

    Call Info("->PACIENTES")
    Set idxPacientes = tbdPACIENTES.CreateIndex("pkPaciente")
    With idxPacientes
        .Fields.Append .CreateField("DNI")
        .Primary = True
        .Unique = True
    End With
    tbdPACIENTES.Indexes.Append idxPacientes
    
    Call Info("->LOCALIDADES")
    Set idxLocalidades = tbdLOCALIDADES.CreateIndex("pkLocalidad")
    With idxLocalidades
        .Fields.Append .CreateField("Localidad")
        .Primary = True
        .Unique = True
    End With
    tbdLOCALIDADES.Indexes.Append idxLocalidades
    
    Call Info("->INGRESOS")
    Set idxIngresos = tbdINGRESOS.CreateIndex("pkIngreso")
    With idxIngresos
        .Fields.Append .CreateField("DNI")
        .Fields.Append .CreateField("Fecha")
        .Primary = True
        .Unique = True
    End With
    tbdINGRESOS.Indexes.Append idxIngresos
    
    Call Info("->OBRAS SOCIALES")
    Set idxObrasSociales = tbdOBRASSOCIALES.CreateIndex("pkObraSocial")
    With idxObrasSociales
        .Fields.Append .CreateField("ObraSocial")
        .Primary = True
        .Unique = True
    End With
    tbdOBRASSOCIALES.Indexes.Append idxObrasSociales
    
    Call Info("->AFILIACIONES")
    Set idxAfiliaciones = tbdAFILIACIONES.CreateIndex("pkAfiliaciones")
    With idxAfiliaciones
        .Fields.Append .CreateField("DNI")
        .Fields.Append .CreateField("ObraSocial")
        .Fields.Append .CreateField("TipoAfiliacion")
        .Primary = True
        .Unique = True
    End With
    tbdAFILIACIONES.Indexes.Append idxAfiliaciones
    
    Call Info("->DIAGNOSTICOS")
    Set idxDiagnostico = tbdDIAGNOSTICOS.CreateIndex("pkDiagnostico")
    With idxDiagnostico
        .Fields.Append .CreateField("Diagnostico")
        .Primary = True
        .Unique = True
    End With
    tbdDIAGNOSTICOS.Indexes.Append idxDiagnostico
        
    Call Info("13: Creando Relaciones")
    
    Call Info("->LOCALIDADES")
    Set relPacientePorLocalidad = dbBase.CreateRelation("PacientePorLocalidad", tbdLOCALIDADES.Name, _
    tbdPACIENTES.Name, dbRelationUpdateCascade + dbRelationDeleteCascade)
    With relPacientePorLocalidad
        .Fields.Append .CreateField("Localidad")
        .Fields!Localidad.ForeignName = "Localidad"
    End With
    dbBase.Relations.Append relPacientePorLocalidad
    
    Call Info("->OBRAS SOCIALES")
    Set relAfiliacionPorObraSocial = dbBase.CreateRelation("AfiliacionPorObraSocial", tbdOBRASSOCIALES.Name, _
    tbdAFILIACIONES.Name, dbRelationUpdateCascade + dbRelationDeleteCascade)
    With relAfiliacionPorObraSocial
        .Fields.Append .CreateField("ObraSocial")
        .Fields!ObraSocial.ForeignName = "ObraSocial"
    End With
    dbBase.Relations.Append relAfiliacionPorObraSocial
    
    Call Info("->AFILIACIONES")
    Set relAfiliacionPorPaciente = dbBase.CreateRelation("AfiliacionPorPaciente", tbdPACIENTES.Name, _
    tbdAFILIACIONES.Name, dbRelationUpdateCascade + dbRelationDeleteCascade)
    With relAfiliacionPorPaciente
        .Fields.Append .CreateField("DNI")
        .Fields!DNI.ForeignName = "DNI"
    End With
    dbBase.Relations.Append relAfiliacionPorPaciente

    Call Info("->INGRESOS")
    Set relIngresoPorPaciente = dbBase.CreateRelation("IngresoPorPaciente", tbdPACIENTES.Name, _
    tbdINGRESOS.Name, dbRelationUpdateCascade + dbRelationDeleteCascade)
    With relIngresoPorPaciente
        .Fields.Append .CreateField("DNI")
        .Fields!DNI.ForeignName = "DNI"
    End With
    dbBase.Relations.Append relIngresoPorPaciente
        
    Call Info("->DIAGNOSTICO")
    Set relIngresoPorDiagnostico = dbBase.CreateRelation("IngresoPorDiagnostico", tbdDIAGNOSTICOS.Name, _
    tbdINGRESOS.Name, dbRelationUpdateCascade + dbRelationDeleteCascade)
    With relIngresoPorDiagnostico
        .Fields.Append .CreateField("Diagnostico")
        .Fields!Diagnostico.ForeignName = "Diagnostico"
    End With
    dbBase.Relations.Append relIngresoPorDiagnostico
    
    dbBase.Close
    cmdRescatar.Enabled = True
End Sub

Private Function Info(StrAgregar As String)
    txtInforme.Text = txtInforme.Text & vbCrLf & StrAgregar
End Function

Private Sub cmdRescatar_Click()

    Dim BaseNueva As Database
    Dim BaseVieja As ADODB.Connection
    Dim rstNuevo As Recordset
    Dim rstViejo As ADODB.Recordset
    Dim Respuesta As String
    
    cmdRescatar.Enabled = False
    
    Call Info("Iniciando Rescate de datos")
    dlgRescatar.Filter = "Todos los Access(*.mdb)|*.mdb|"
    dlgRescatar.ShowOpen
    
    Call Info("1: Conectando Bases de Datos")
    Set BaseVieja = New ADODB.Connection
    With BaseVieja
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        .ConnectionString = "Data Source=" & dlgRescatar.FileName
        .Open
    End With
    
    Set BaseNueva = OpenDatabase(App.Path & "\" & BaseDeDatos())
    
    Call Info("2: Seteando y Cargando Tablas")
    
    Call Info("-> LOCALIDADES")
    Set rstViejo = New ADODB.Recordset
    rstViejo.Open "Select * FROM TblLocalidad", BaseVieja, adOpenDynamic, adLockOptimistic
    Set rstNuevo = BaseNueva.OpenRecordset("LOCALIDADES")
    rstViejo.MoveFirst
    Do Until rstViejo.EOF
        rstNuevo.AddNew
        With rstNuevo
            .Fields("Localidad") = rstViejo.Fields("Localidad")
        End With
        rstNuevo.Update
        rstViejo.MoveNext
    Loop
    
    Call Info("-> OBRAS SOCIALES")
    Set rstViejo = New ADODB.Recordset
    rstViejo.Open "Select * FROM TblObraSocial", BaseVieja, adOpenDynamic, adLockOptimistic
    Set rstNuevo = BaseNueva.OpenRecordset("OBRASSOCIALES")
    rstViejo.MoveFirst
    Do Until rstViejo.EOF
        rstNuevo.AddNew
        With rstNuevo
            .Fields("ObraSocial") = rstViejo.Fields("ObraSocial")
        End With
        rstNuevo.Update
        rstViejo.MoveNext
    Loop
    
    Call Info("-> PACIENTES")
    Set rstViejo = New ADODB.Recordset
    rstViejo.Open "Select TblPacientes.*, TblLocalidad.Localidad FROM TblLocalidad INNER JOIN TblPacientes ON TblLocalidad.IdLocalidad = TblPacientes.IdLocalidad", BaseVieja, adOpenDynamic, adLockOptimistic
    Set rstNuevo = BaseNueva.OpenRecordset("PACIENTES")
    rstViejo.MoveFirst
    Do Until rstViejo.EOF
        rstNuevo.AddNew
        With rstNuevo
            .Fields("DNI") = rstViejo.Fields("Documento")
            .Fields("Apellido") = rstViejo.Fields("Apellidos")
            .Fields("Nombre") = rstViejo.Fields("Nombres")
            .Fields("FechaNacimiento") = rstViejo.Fields("Edad")
            .Fields("Domicilio") = rstViejo.Fields("Domicilio")
            .Fields("Localidad") = rstViejo.Fields("Localidad")
            If Len(rstViejo.Fields("Telefono1")) <> "0" Then
                .Fields("Telefono1") = rstViejo.Fields("Telefono1")
            Else
                .Fields("Telefono1") = ""
            End If
            If Len(rstViejo.Fields("Telefono2")) <> "0" Then
                .Fields("Telefono2") = rstViejo.Fields("Telefono2")
            Else
                .Fields("Telefono2") = ""
            End If
        End With
        rstNuevo.Update
        rstViejo.MoveNext
    Loop

    Call Info("-> AFILIACIONES")
    Set rstViejo = New ADODB.Recordset
    rstViejo.Open "Select TblPacientes.*, TblObraSocial.ObraSocial FROM TblObraSocial INNER JOIN TblPacientes ON TblObraSocial.IdObraSocial = TblPacientes.IdObraSocial", BaseVieja, adOpenDynamic, adLockOptimistic
    Set rstNuevo = BaseNueva.OpenRecordset("AFILIACIONES")
    rstViejo.MoveFirst
    Do Until rstViejo.EOF
        rstNuevo.AddNew
        With rstNuevo
             .Fields("DNI") = rstViejo.Fields("Documento")
            If Len(rstViejo.Fields("ObraSocial")) <> "0" Then
                .Fields("ObraSocial") = rstViejo.Fields("ObraSocial")
            Else
                .Fields("ObraSocial") = "I.O.S.C.O.R."
            End If
            .Fields("TipoAfiliacion") = "PRIMARIA"
            If Len(rstViejo.Fields("NumeroSocio")) <> "0" Then
                .Fields("NumeroAfiliado") = rstViejo.Fields("NumeroSocio")
            Else
                .Fields("NumeroAfiliado") = ""
            End If
        End With
        rstNuevo.Update
        rstViejo.MoveNext
    Loop
    
    Call Info("-> DIGNOSTICOS")
    Set rstViejo = New ADODB.Recordset
    rstViejo.Open "Select * FROM TblDiagnostico", BaseVieja, adOpenDynamic, adLockOptimistic
    Set rstNuevo = BaseNueva.OpenRecordset("DIAGNOSTICOS")
    rstViejo.MoveFirst
    Do Until rstViejo.EOF
        rstNuevo.AddNew
        With rstNuevo
            .Fields("Diagnostico") = rstViejo.Fields("Diagnostico")
        End With
        rstNuevo.Update
        rstViejo.MoveNext
    Loop
                
    Call Info("-> INGRESOS")
    Set rstViejo = New ADODB.Recordset
    rstViejo.Open "Select TblIngresos.*, TblPacientes.Documento, TblDiagnostico.Diagnostico FROM TblPacientes INNER JOIN (TblDiagnostico INNER JOIN TblIngresos ON TblDiagnostico.IdDiagnostico = TblIngresos.IdDiagnostico) ON TblPacientes.IdPaciente = TblIngresos.IdPaciente", BaseVieja, adOpenDynamic, adLockOptimistic
    Set rstNuevo = BaseNueva.OpenRecordset("INGRESOS")
    rstViejo.MoveFirst
    Do Until rstViejo.EOF
        rstNuevo.AddNew
        With rstNuevo
            .Fields("DNI") = rstViejo.Fields("Documento")
            .Fields("Fecha") = rstViejo.Fields("Fecha")
            .Fields("Diagnostico") = rstViejo.Fields("Diagnostico")
            .Fields("Observaciones") = rstViejo.Fields("Observaciones")
        End With
        rstNuevo.Update
        rstViejo.MoveNext
    Loop
    
    
    BaseNueva.Close
    BaseVieja.Close
    Set rstNuevo = Nothing
    Set rstViejo = Nothing
    Call Conexion(dbBase, App.Path & "\" & BaseDeDatos())
End Sub

Private Sub Form_Load()
    cmdRescatar.Enabled = False
    With CargaBD
        .Height = 5600
        .Width = 3700
    End With
    Call CenterMe(CargaBD)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call Conexion(dbBase, App.Path & "\" & BaseDeDatos())
End Sub
