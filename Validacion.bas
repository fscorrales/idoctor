Attribute VB_Name = "Validacion"
Public Function EsTextoNoVacio(TextoAnalizado As String, TamañoCampo As Byte, NombredelCampo As String) As Boolean
    If Trim(TextoAnalizado) = "" Or IsNumeric(TextoAnalizado) = True Or Len(TextoAnalizado) > TamañoCampo Then
        MsgBox "El Dato Ingresado no cumple con el requisito de ser TEXTO o sobrepasa la EXTENSIÓN MÁXIMA" & vbCrLf & "INTENTE DE NUEVO", vbCritical + vbOKOnly, NombredelCampo & " INCORRECTO"
        EsTextoNoVacio = False
        Exit Function
    End If
    EsTextoNoVacio = True
End Function

Public Function EsNumeroNoVacio(TextoAnalizado As String, TamañoCampo As Byte, NombredelCampo As String) As Boolean
    If Trim(TextoAnalizado) = "" Or IsNumeric(TextoAnalizado) = False Or Len(TextoAnalizado) > TamañoCampo Then
        MsgBox "El Dato Ingresado no cumple con el requisito de ser NÚMERO o sobrepasa la EXTENSIÒN MÁXIMA" & vbCrLf & "INTENTE DE NUEVO", vbCritical + vbOKOnly, NombredelCampo & " INCORRECTO"
        EsNumeroNoVacio = False
        Exit Function
    End If
    EsNumeroNoVacio = True
End Function

Public Function EsFechaNoVacio(TextoAnalizado As String, TamañoCampo As Byte, NombredelCampo As String) As Boolean
    If Trim(TextoAnalizado) = "" Or IsDate(TextoAnalizado) = False Or Len(TextoAnalizado) > TamañoCampo Then
        MsgBox "El Dato Ingresado no cumple con el requisito de ser FECHA o sobrepasa la EXTENSIÒN MÁXIMA" & vbCrLf & "INTENTE DE NUEVO", vbCritical + vbOKOnly, NombredelCampo & " INCORRECTO"
        EsFechaNoVacio = False
        Exit Function
    End If
    EsFechaNoVacio = True
End Function

Public Function EsIgualTextoEspecificado(TextoAnalizado As String, NombredelCampo As String, ParamArray Valores() As Variant) As Boolean
    Dim Valor As Variant
    For Each Valor In Valores
        If Valor = TextoAnalizado Then
            EsIgualTextoEspecificado = True
            Exit Function
        End If
    Next Valor
    MsgBox "El Dato Ingresado es incorrecto por no encontrarse en la LISTA ESPECIFICADA" & vbCrLf & "INTENTE DE NUEVO", vbCritical + vbOKOnly, NombredelCampo & " INCORRECTO"
    EsIgualTextoEspecificado = False
End Function

Public Function ExisteEnTablaPrincipal(ValorBuscado As String, Recordset As Recordset, SQL As String, Indice As String) As Boolean
    Call SetRecordset(Recordset, SQL)
    With Recordset
        .Index = Indice
        .Seek "=", ValorBuscado
    End With
    If Recordset.NoMatch = True Then
        MsgBox "El Valor Ingresado NO EXISTE en la TABLA PRINCIPAL" & vbCrLf & "INTENTE DE NUEVO", vbCritical + vbOKOnly, "VIOLACIÓN A LA INTEGRIDAD REFERENCIAL"
        ExisteEnTablaPrincipal = False
        Set Recordset = Nothing
        Exit Function
    End If
    ExisteEnTablaPrincipal = True
    Set Recordset = Nothing
End Function

Public Function ValorDuplicado(ValorBuscado As String, Recordset As Recordset, SQL As String, Indice As String) As Boolean
    Call SetRecordset(Recordset, SQL)
    With Recordset
        .Index = Indice
        .Seek "=", ValorBuscado
    End With
    If Recordset.NoMatch = False Then
        MsgBox "El Valor Ingresado ya EXISTE y la duplicación del mismo es IMPROCEDENTE" & vbCrLf & "INTENTE DE NUEVO", vbCritical + vbOKOnly, "VALOR DUPLICADO"
        ValorDuplicado = True
        Set Recordset = Nothing
        Exit Function
    End If
    ValorDuplicado = False
    Set Recordset = Nothing
End Function
