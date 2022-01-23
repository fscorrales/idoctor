VERSION 5.00
Begin VB.Form TipoHistorial 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Tipo Historial"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox txtCodigoHistorial 
      Height          =   315
      ItemData        =   "TipoHistorial.frx":0000
      Left            =   1920
      List            =   "TipoHistorial.frx":0002
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Historial"
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
      TabIndex        =   3
      Top             =   120
      Width           =   5775
      Begin VB.TextBox txtDescripcion 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Top             =   1320
         Width           =   3735
      End
      Begin VB.ComboBox txtTipoDato 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   1800
         Width           =   3735
      End
      Begin VB.ComboBox txtJerarquia 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Datos a almacenar"
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
         TabIndex        =   9
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Descripción"
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
         TabIndex        =   8
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo"
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
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Jerarquía"
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
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdGuardarDatosHistorial 
      Caption         =   "Guardar Datos Historial"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   2640
      Width           =   4095
   End
End
Attribute VB_Name = "TipoHistorial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGuardarDatosHistorial_Click()
    If Validar = True Then
        If bolEditandoTipoHistorial = False Then
            Call SetRecordset(rstCargaTipoHistorial, "TIPOHISTORIAL")
            Call GuardarTipoHistorial(rstCargaTipoHistorial)
        Else
            Call GuardarTipoHistorial(rstEditarTipoHistorial)
        End If
        
        Unload TipoHistorial
        ListadoTipoHistorial.Show
    End If
End Sub

Private Sub GuardarTipoHistorial(Recordset As Recordset)
    With Recordset
        If bolEditandoTipoHistorial = False Then
            .AddNew
        Else
            .Edit
        End If
        .Fields("CodigoHistorial") = AsignarNumero("SELECT CodigoHistorial FROM TipoHistorial Where CodigoHistorial LIKE " & "'" & ConvertirCodigo(txtCodigoHistorial.Text) & "###' ORDER BY CodigoHistorial")
        .Fields("Descripcion") = txtDescripcion.Text
        .Fields("TipoDato") = txtTipoDato.Text
        .Fields("Jerarquia") = txtJerarquia.Text
        .Update
    End With
    
    Set Recordset = Nothing
    
End Sub

Private Sub Form_Load()
    Call CenterMe(TipoHistorial)
    With TipoHistorial
        .Width = 6100
        .Height = 3500
    End With
    txtCodigoHistorial.AddItem "Laboratorio"
    txtCodigoHistorial.AddItem "Vacunas"
    txtCodigoHistorial.AddItem "Tratamiento"
    txtCodigoHistorial.AddItem "Profilaxis"
    txtCodigoHistorial.AddItem "Otros"
    txtTipoDato.AddItem "Ninguno"
    txtTipoDato.AddItem "Numero"
    txtTipoDato.AddItem "Texto"
    txtTipoDato.AddItem "Fecha"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bolEditandoTipoHistorial = True Then
        bolEditandoTipoHistorial = False
        ListadoTipoHistorial.Show
    End If
End Sub

Private Sub txtCodigoHistorial_LostFocus()
    txtJerarquia.Clear
    txtJerarquia.AddItem "Principal"
    Call SetRecordset(rstComboJerarquiaHistorial, "SELECT Descripcion FROM TipoHistorial Where CodigoHistorial LIKE " & "'" & ConvertirCodigo(txtCodigoHistorial.Text) & "###' and Jerarquia = 'Principal'")
    If rstComboJerarquiaHistorial.BOF = False Then
        With rstComboJerarquiaHistorial
            .MoveFirst
            While .EOF = False
                txtJerarquia.AddItem .Fields("Descripcion")
                .MoveNext
            Wend
        End With
    End If
    Set rstComboJerarquiaHistorial = Nothing
    If bolEditandoTipoHistorial = True Then
        txtJerarquia.Text = rstEditarTipoHistorial.Fields("Jerarquia")
    End If
End Sub

Private Function ConvertirCodigo(Descripcion As String) As String
    Select Case Descripcion
        Case Is = "Laboratorio"
            ConvertirCodigo = "L"
        Case Is = "Vacunas"
            ConvertirCodigo = "V"
        Case Is = "Tratamiento"
            ConvertirCodigo = "T"
        Case Is = "Profilaxis"
            ConvertirCodigo = "P"
        Case Is = "Otros"
            ConvertirCodigo = "O"
    End Select
End Function

Private Function Validar() As Boolean
    If EsIgualTextoEspecificado(txtCodigoHistorial.Text, "CÓDIGO HISTORIAL", "Laboratorio", "Vacunas", "Tratamiento", "Profilaxis", "Otros") = False Then
        txtCodigoHistorial.SetFocus
        Validar = False
        Exit Function
    End If
    If txtJerarquia.Text <> "Principal" Then
        Call Encontrar(txtJerarquia, rstComboJerarquiaHistorial, "SELECT DISTINCT Descripcion FROM TipoHistorial WHERE CodigoHistorial LIKE " & "'" & ConvertirCodigo(txtCodigoHistorial.Text) & "###' and Jerarquia = 'Principal'", "Descripcion")
        If rstComboJerarquiaHistorial.NoMatch = True Then
            Set rstComboJerarquiaHistorial = Nothing
            MsgBox "El Dato Ingresado es incorrecto por no encontrarse en la LISTA ESPECIFICADA" & vbCrLf & "INTENTE DE NUEVO", vbCritical + vbOKOnly, "JERARQUÍA INCORRECTA"
            txtJerarquia.SetFocus
            Validar = False
            Exit Function
        End If
        Set rstComboJerarquiaHistorial = Nothing
    End If
    If EsTextoNoVacio(txtDescripcion.Text, "30", "DESCRIPCIÓN") = False Then
        txtDescripcion.SetFocus
        Validar = False
        Exit Function
    End If
    If EsIgualTextoEspecificado(txtTipoDato.Text, "DATOS A ALMACENAR", "Ninguno", "Numero", "Texto", "Fecha") = False Then
        txtTipoDato.SetFocus
        Validar = False
        Exit Function
    End If
    If txtDescripcion.Text = txtJerarquia.Text Then
        MsgBox "La DESCRIPCIÓN de la variable no puede ser igual a su JERARQUÍA", vbOKOnly + vbCritical, "JERAQUÍA MAL ASIGNADA"
        txtJerarquia.Text = ""
        txtDescripcion.SetFocus
        Validar = False
        Exit Function
    End If
    
    Validar = True
End Function

Private Function AsignarNumero(SQL As String) As String
    Call SetRecordset(rstAsignarNumero, SQL)
    Dim n As Single
    If rstAsignarNumero.BOF = False Then
        With rstAsignarNumero
            .MoveLast
            n = Format(Right(.Fields("CodigoHistorial"), 3), GeneralNumber) + 1
            AsignarNumero = ConvertirCodigo(txtCodigoHistorial.Text) & Format(n, "000")
        End With
    Else
        AsignarNumero = ConvertirCodigo(txtCodigoHistorial.Text) & "001"
    End If
    Set rstAsignarNumero = Nothing
End Function
