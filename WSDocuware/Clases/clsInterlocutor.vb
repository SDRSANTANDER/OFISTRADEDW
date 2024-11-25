Imports SAPbobsCOM
Imports System.Reflection
Imports System.Globalization
Imports WSDocuware.Utilidades

Public Class clsInterlocutor

#Region "Públicas"

    Public Function NuevoInterlocutor(ByVal objInterlocutor As EntInterlocutorSAP, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oInterlocutor As BusinessPartners = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUEVO INTERLOCUTOR"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de interlocutor para " & objInterlocutor.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objInterlocutor.UserSAP, objInterlocutor.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If objInterlocutor.Serie <= 0 AndAlso String.IsNullOrEmpty(objInterlocutor.Codigo) Then Throw New Exception("Serie o código no suministrado")
            If String.IsNullOrEmpty(objInterlocutor.Nombre) Then Throw New Exception("Nombre no suministrado")
            If String.IsNullOrEmpty(objInterlocutor.Tipo) Then Throw New Exception("Tipo no suministrado")
            If String.IsNullOrEmpty(objInterlocutor.NIFTercero) Then Throw New Exception("NIF no suministrado")

            'Objeto interlocutor
            oInterlocutor = oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners)

            'Comprueba si existe el interlocutor
            If objInterlocutor.Serie <= 0 AndAlso Not String.IsNullOrEmpty(objInterlocutor.Codigo) Then

                If oInterlocutor.GetByKey(objInterlocutor.Codigo) Then Throw New Exception("Ya existe el interlocutor con CardCode: " & objInterlocutor.Codigo)

            End If

            'El código depende de la serie

            'Series
            If objInterlocutor.Serie > 0 Then oInterlocutor.Series = objInterlocutor.Serie

            'CardCode
            If Not String.IsNullOrEmpty(objInterlocutor.Codigo) Then oInterlocutor.CardCode = objInterlocutor.Codigo

            'CardName
            If Not String.IsNullOrEmpty(objInterlocutor.Nombre) Then oInterlocutor.CardName = objInterlocutor.Nombre

            'ForeignName
            If Not String.IsNullOrEmpty(objInterlocutor.NombreExtranjero) Then oInterlocutor.ForeignName = objInterlocutor.NombreExtranjero

            'CardType
            If IsNumeric(objInterlocutor.Tipo) Then oInterlocutor.CardType = CInt(objInterlocutor.Tipo)

            'GroupCode
            If IsNumeric(objInterlocutor.Grupo) Then oInterlocutor.GroupCode = CInt(objInterlocutor.Grupo)

            'Currency
            If Not String.IsNullOrEmpty(objInterlocutor.Moneda) Then oInterlocutor.Currency = objInterlocutor.Moneda

            'NIF
            If Not String.IsNullOrEmpty(objInterlocutor.NIFTercero) Then oInterlocutor.FederalTaxID = objInterlocutor.NIFTercero

            'Phone1
            If Not String.IsNullOrEmpty(objInterlocutor.Telefono) Then oInterlocutor.Phone1 = objInterlocutor.Telefono

            'E_Mail
            If Not String.IsNullOrEmpty(objInterlocutor.Mail) Then oInterlocutor.EmailAddress = objInterlocutor.Mail

            'frozenFor
            oInterlocutor.Frozen = IIf(objInterlocutor.Activo = SN.Si, BoYesNoEnum.tYES, BoYesNoEnum.tNO)

            'Notes
            Dim sNotes As String = objInterlocutor.Comentarios
            If Not String.IsNullOrEmpty(sNotes) AndAlso sNotes.Length > 100 Then sNotes = sNotes.Substring(0, 100)
            If Not String.IsNullOrEmpty(sNotes) Then oInterlocutor.Notes = objInterlocutor.Comentarios

            '---------------------------------
            ' Campos de usuario
            '---------------------------------

            'Campos de usuario
            If Not objInterlocutor.CamposUsuario Is Nothing AndAlso objInterlocutor.CamposUsuario.Count > 0 Then

                For Each oCampoUsuario In objInterlocutor.CamposUsuario

                    Dim oUserField As Field = oInterlocutor.UserFields.Fields.Item(oCampoUsuario.Campo)

                    Select Case oUserField.Type

                        Case BoFieldTypes.db_Numeric
                            If oUserField.SubType = BoFldSubTypes.st_Time Then
                                If DateTime.TryParseExact(oCampoUsuario.Valor, "HHmm", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                    oInterlocutor.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "HHmm", CultureInfo.CurrentCulture)
                            Else
                                If IsNumeric(oCampoUsuario.Valor) Then _
                                    oInterlocutor.Lines.UserFields.Item(oCampoUsuario.Campo).Value = CInt(oCampoUsuario.Valor)
                            End If

                        Case BoFieldTypes.db_Float
                            If IsNumeric(oCampoUsuario.Valor) Then _
                                oInterlocutor.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CDbl(oCampoUsuario.Valor.Replace(".", ","))

                        Case BoFieldTypes.db_Date
                            If DateTime.TryParseExact(oCampoUsuario.Valor, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                oInterlocutor.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                        Case Else
                            If Not String.IsNullOrEmpty(oCampoUsuario.Valor) Then _
                                oInterlocutor.UserFields.Fields.Item(oCampoUsuario.Campo).Value = oCampoUsuario.Valor

                    End Select

                Next

            End If

            '---------------------------------
            ' Creación
            '---------------------------------

            'Añadimos interlocutor
            If oInterlocutor.Add() = 0 Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Interlocutor creado con éxito"
                retVal.MENSAJEAUX = getCardCodeDeCardName(objInterlocutor.Nombre, Sociedad)

            Else

                retVal.CODIGO = Respuesta.Ko
                retVal.MENSAJE = oCompany.GetLastErrorCode & "#" & oCompany.GetLastErrorDescription
                retVal.MENSAJEAUX = ""

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        Finally
            oComun.LiberarObjCOM(oInterlocutor)
        End Try

        Return retVal

    End Function

    Public Function ActualizarInterlocutor(ByVal objInterlocutor As EntInterlocutorSAP, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oInterlocutor As BusinessPartners = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "ACTUALIZAR INTERLOCUTOR"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de interlocutor: " & objInterlocutor.Codigo)

            oCompany = ConexionSAP.getCompany(objInterlocutor.UserSAP, objInterlocutor.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objInterlocutor.Codigo) Then Throw New Exception("Código no suministrado")

            'Objeto interlocutor
            oInterlocutor = oCompany.GetBusinessObject(BoObjectTypes.oBusinessPartners)

            'Comprueba si existe el interlocutor
            If Not oInterlocutor.GetByKey(objInterlocutor.Codigo) Then

                Throw New Exception("No se encuentra el interlocutor con CardCode: " & objInterlocutor.Codigo)

            Else

                'CardName
                If Not String.IsNullOrEmpty(objInterlocutor.Nombre) Then oInterlocutor.CardName = objInterlocutor.Nombre

                'ForeignName
                If Not String.IsNullOrEmpty(objInterlocutor.NombreExtranjero) Then oInterlocutor.ForeignName = objInterlocutor.NombreExtranjero

                'CardType
                If IsNumeric(objInterlocutor.Tipo) Then oInterlocutor.CardType = CInt(objInterlocutor.Tipo)

                'GroupCode
                If IsNumeric(objInterlocutor.Grupo) Then oInterlocutor.GroupCode = CInt(objInterlocutor.Grupo)

                'Currency
                If Not String.IsNullOrEmpty(objInterlocutor.Moneda) Then oInterlocutor.Currency = objInterlocutor.Moneda

                'NIF
                If Not String.IsNullOrEmpty(objInterlocutor.NIFTercero) Then oInterlocutor.FederalTaxID = objInterlocutor.NIFTercero

                'Phone1
                If Not String.IsNullOrEmpty(objInterlocutor.Telefono) Then oInterlocutor.Phone1 = objInterlocutor.Telefono

                'E_Mail
                If Not String.IsNullOrEmpty(objInterlocutor.Mail) Then oInterlocutor.EmailAddress = objInterlocutor.Mail

                'frozenFor
                oInterlocutor.Frozen = IIf(objInterlocutor.Activo = SN.Si, BoYesNoEnum.tYES, BoYesNoEnum.tNO)

                'Notes
                Dim sNotes As String = objInterlocutor.Comentarios
                If Not String.IsNullOrEmpty(sNotes) AndAlso sNotes.Length > 100 Then sNotes = sNotes.Substring(0, 100)
                If Not String.IsNullOrEmpty(sNotes) Then oInterlocutor.Notes = objInterlocutor.Comentarios

                'Campos de usuario
                If Not objInterlocutor.CamposUsuario Is Nothing AndAlso objInterlocutor.CamposUsuario.Count > 0 Then

                    For Each oCampoUsuario In objInterlocutor.CamposUsuario

                        Dim oUserField As Field = oInterlocutor.UserFields.Fields.Item(oCampoUsuario.Campo)

                        Select Case oUserField.Type

                            Case BoFieldTypes.db_Numeric
                                If oUserField.SubType = BoFldSubTypes.st_Time Then
                                    If DateTime.TryParseExact(oCampoUsuario.Valor, "HHmm", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                    oInterlocutor.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "HHmm", CultureInfo.CurrentCulture)
                                Else
                                    If IsNumeric(oCampoUsuario.Valor) Then _
                                    oInterlocutor.Lines.UserFields.Item(oCampoUsuario.Campo).Value = CInt(oCampoUsuario.Valor)
                                End If

                            Case BoFieldTypes.db_Float
                                If IsNumeric(oCampoUsuario.Valor) Then _
                                oInterlocutor.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CDbl(oCampoUsuario.Valor.Replace(".", ","))

                            Case BoFieldTypes.db_Date
                                If DateTime.TryParseExact(oCampoUsuario.Valor, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                oInterlocutor.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                            Case Else
                                If Not String.IsNullOrEmpty(oCampoUsuario.Valor) Then _
                                oInterlocutor.UserFields.Fields.Item(oCampoUsuario.Campo).Value = oCampoUsuario.Valor

                        End Select

                    Next

                End If

                '---------------------------------
                ' Actualización
                '---------------------------------

                'Actualizamos interlocutor
                If oInterlocutor.Update() = 0 Then

                    retVal.CODIGO = Respuesta.Ok
                    retVal.MENSAJE = "Interlocutor actualizado con éxito"
                    retVal.MENSAJEAUX = objInterlocutor.Codigo

                Else

                    retVal.CODIGO = Respuesta.Ko
                    retVal.MENSAJE = oCompany.GetLastErrorCode & "#" & oCompany.GetLastErrorDescription
                    retVal.MENSAJEAUX = ""

                End If

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        Finally
            oComun.LiberarObjCOM(oInterlocutor)
        End Try

        Return retVal

    End Function

#End Region

#Region "Privadas"

    Private Function getCardCodeDeCardName(ByVal CardName As String,
                                           ByVal Sociedad As eSociedad) As String

        'Devuelve el CardCode del interlocutor

        Dim retVal As String = ""

        Try

            'Buscamos por CardName
            Dim SQL As String = ""

            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " T0." & putQuotes("CardCode") & " As CardCode " & vbCrLf
            SQL &= " FROM " & getDataBaseRef("OCRD", Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("CardName") & "  = N'" & CardName & "'" & vbCrLf

            SQL &= " ORDER BY " & vbCrLf
            SQL &= " T0." & putQuotes("CreateDate") & " DESC, " & vbCrLf
            SQL &= " T0." & putQuotes("CreateTS") & " DESC " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

#End Region

End Class
