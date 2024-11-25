Imports SAPbobsCOM
Imports System.Reflection
Imports System.Globalization
Imports WSDocuware.Utilidades

Public Class clsOportunidad

#Region "Públicas"

    Public Function NuevaOportunidad(ByVal objOportunidad As EntOportunidadCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oOpportunity As SalesOpportunities = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUEVA OPORTUNIDAD"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de oportunidad para " & objOportunidad.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objOportunidad.UserSAP, objOportunidad.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objOportunidad.NIFTercero) Then Throw New Exception("NIF no suministrado")
            If String.IsNullOrEmpty(objOportunidad.Nombre) Then Throw New Exception("Nombre no suministrado")

            If Not DateTime.TryParseExact(objOportunidad.FechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha inicio no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objOportunidad.FechaCierrePrevista.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha cierre prevista no suministrada o incorrecta")

            If objOportunidad.ImportePotencial = 0 Then Throw New Exception("Importe potencial no suministrado")

            If String.IsNullOrEmpty(objOportunidad.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objOportunidad.DOCURLDW) Then Throw New Exception("URL Docuware no suministrado")

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objOportunidad.ObjTypeDestino)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objOportunidad.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objOportunidad.NIFTercero, objOportunidad.RazonSocial, objOportunidad.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objOportunidad.NIFTercero & ", Razón social: " & objOportunidad.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Comprueba que no existe el Oportunidad
            Dim NumOportunidad As String = getNumberDeOportunidad(TablaDestino, CardCode, "", -1, objOportunidad.Nombre, Sociedad)
            If Not String.IsNullOrEmpty(NumOportunidad) Then Throw New Exception("Ya existe oportunidad " & NumOportunidad & " para NIF: " & objOportunidad.NIFTercero & ", Razón social: " & objOportunidad.RazonSocial & ", Nombre: " & objOportunidad.Nombre)

            'Objeto Oportunidad
            oOpportunity = oCompany.GetBusinessObject(objOportunidad.ObjTypeDestino)

            'Cabecera
            oOpportunity.OpportunityType = IIf(objOportunidad.Ambito = Ambito.Compras, OpportunityTypeEnum.boOpPurchasing, OpportunityTypeEnum.boOpSales)

            oOpportunity.CardCode = CardCode
            If objOportunidad.PersonaContacto > 0 Then oOpportunity.ContactPerson = objOportunidad.PersonaContacto

            If objOportunidad.EmpDptoCompraVenta > 0 Then oOpportunity.SalesPerson = objOportunidad.EmpDptoCompraVenta

            If Not String.IsNullOrEmpty(objOportunidad.Nombre) Then oOpportunity.OpportunityName = objOportunidad.Nombre

            oOpportunity.StartDate = Date.ParseExact(objOportunidad.FechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Potencial
            oOpportunity.PredictedClosingDate = Date.ParseExact(objOportunidad.FechaCierrePrevista.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            'oOpportunity.MaxLocalTotal = objOportunidad.ImportePotencial

            'Resumen
            If objOportunidad.Status > -1 Then oOpportunity.Status = objOportunidad.Status

            'Campos Docuware
            oOpportunity.UserFields.Item("U_SEIIDDW").Value = objOportunidad.DOCIDDW
            oOpportunity.UserFields.Item("U_SEIURLDW").Value = objOportunidad.DOCURLDW

            'Líneas
            For i = 0 To oOpportunity.Lines.Count - 1
                oOpportunity.Lines.SetCurrentLine(i)
                oOpportunity.Lines.StartDate = oOpportunity.StartDate
                oOpportunity.Lines.MaxLocalTotal = objOportunidad.ImportePotencial
            Next

            'Añadimos oportunidad
            If oOpportunity.Add() = 0 Then

                'Obtiene el número del documento añadido
                Dim OpprID As String = oCompany.GetNewObjectKey

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Oportunidad creada con éxito"
                retVal.MENSAJEAUX = OpprID

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
            oComun.LiberarObjCOM(oOpportunity)
        End Try

        Return retVal

    End Function

    Public Function ActualizarOportunidad(ByVal objOportunidad As EntOportunidadCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oOpportunity As SalesOpportunities = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "ACTUALIZAR OPORTUNIDAD"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de oportunidad: " & objOportunidad.Numero)

            oCompany = ConexionSAP.getCompany(objOportunidad.UserSAP, objOportunidad.PassSAP, Sociedad)

            If oCompany Is Nothing OrElse Not oCompany.Connected Then
                Throw New Exception("No se puede conectar a SAP")
            End If

            'Obligatorios
            If String.IsNullOrEmpty(objOportunidad.Numero) Then Throw New Exception("Número no suministrado")

            If objOportunidad.FechaInicio > 0 AndAlso Not DateTime.TryParseExact(objOportunidad.FechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha inicio no suministrada o incorrecta")
            If objOportunidad.FechaCierrePrevista > 0 AndAlso Not DateTime.TryParseExact(objOportunidad.FechaCierrePrevista.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha cierre prevista no suministrada o incorrecta")

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objOportunidad.ObjTypeDestino)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objOportunidad.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objOportunidad.NIFTercero, objOportunidad.RazonSocial, objOportunidad.Ambito, Sociedad)
            If Not String.IsNullOrEmpty(CardCode) Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Objeto Oportunidad
            oOpportunity = oCompany.GetBusinessObject(objOportunidad.ObjTypeDestino)
            If Not oOpportunity.GetByKey(CInt(objOportunidad.Numero)) Then Throw New Exception("No puedo recuperar la oportunidad origen con número: " & objOportunidad.Numero)

            'Cabecera
            If Not String.IsNullOrEmpty(CardCode) Then oOpportunity.CardCode = CardCode

            If objOportunidad.PersonaContacto > 0 Then oOpportunity.ContactPerson = objOportunidad.PersonaContacto

            If objOportunidad.EmpDptoCompraVenta > 0 Then oOpportunity.SalesPerson = objOportunidad.EmpDptoCompraVenta

            If Not String.IsNullOrEmpty(objOportunidad.Nombre) Then oOpportunity.OpportunityName = objOportunidad.Nombre

            If objOportunidad.FechaInicio > 0 Then oOpportunity.StartDate = Date.ParseExact(objOportunidad.FechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Potencial
            If objOportunidad.FechaCierrePrevista > 0 Then oOpportunity.PredictedClosingDate = Date.ParseExact(objOportunidad.FechaCierrePrevista.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Resumen
            If objOportunidad.Status > -1 Then oOpportunity.Status = objOportunidad.Status

            'Campos Docuware
            If Not String.IsNullOrEmpty(objOportunidad.DOCIDDW) Then oOpportunity.UserFields.Item("U_SEIIDDW").Value = objOportunidad.DOCIDDW
            If Not String.IsNullOrEmpty(objOportunidad.DOCURLDW) Then oOpportunity.UserFields.Item("U_SEIURLDW").Value = objOportunidad.DOCURLDW

            'Líneas
            For i = 0 To oOpportunity.Lines.Count - 1
                oOpportunity.Lines.SetCurrentLine(i)
                oOpportunity.Lines.StartDate = oOpportunity.StartDate
                If objOportunidad.ImportePotencial > 0 Then oOpportunity.Lines.MaxLocalTotal = objOportunidad.ImportePotencial
            Next

            'Añadimos oportunidad
            If oOpportunity.Update() = 0 Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Oportunidad actualizada con éxito"
                retVal.MENSAJEAUX = objOportunidad.Numero

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
            oComun.LiberarObjCOM(oOpportunity)
        End Try

        Return retVal

    End Function

    Public Function getComprobarOportunidad(ByVal objOportunidad As EntOportunidadCab, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un Oportunidad

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "OPORTUNIDAD"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ")  Comprobar oportunidad para ObjType: " & objOportunidad.ObjTypeDestino & ", Ambito: " & objOportunidad.Ambito & ", NIFTercero: " & objOportunidad.NIFTercero & ", Razón social: " & objOportunidad.RazonSocial & ", DOCIDDW: " & objOportunidad.DOCIDDW & ", Nombre: " & objOportunidad.Nombre)

            'Obligatorios
            If String.IsNullOrEmpty(objOportunidad.NIFTercero) AndAlso String.IsNullOrEmpty(objOportunidad.RazonSocial) Then Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objOportunidad.DOCIDDW) AndAlso String.IsNullOrEmpty(objOportunidad.Nombre) Then Throw New Exception("ID Docuware o descripción no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objOportunidad.NIFTercero, objOportunidad.RazonSocial, objOportunidad.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objOportunidad.NIFTercero & ", Razón social: " & objOportunidad.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objOportunidad.ObjTypeDestino)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objOportunidad.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ")  Encontrada tabla: " & Tabla)

            'Devuelve el Number del Oportunidad
            Dim DocNum As String = getNumberDeOportunidad(Tabla, CardCode, objOportunidad.DOCIDDW, objOportunidad.Numero, objOportunidad.Nombre, Sociedad)

            If Not String.IsNullOrEmpty(DocNum) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Oportunidad encontrada"
                retVal.MENSAJEAUX = DocNum

            Else

                retVal.CODIGO = Respuesta.Ko
                retVal.MENSAJE = "No se encuentra la oportunidad"
                retVal.MENSAJEAUX = ""

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

#End Region

#Region "Privadas"

    Private Function getNumberOportunidadPorDWID(ByVal Tabla As String,
                                                 ByVal IDDW As String,
                                                 ByVal Sociedad As eSociedad) As String

        'Devuelve el Number del Oportunidad

        Dim retVal As String = ""

        Try

            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("OpprId") & vbCrLf
            SQL &= " FROM " & putQuotes(Tabla) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And COALESCE(T0." & putQuotes("U_SEIIDDW") & ",'')  = '" & IDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Sub setDatosDWOportunidad(ByVal Tabla As String,
                                      ByVal CardCode As String,
                                      ByVal Number As String,
                                      ByVal IDDW As String,
                                      ByVal URLDW As String,
                                      ByVal Sociedad As eSociedad)

        'Actualiza los campos de DW en el documento

        Try

            'Actualizamos por CardCode, DocDate y DocNum
            Dim SQL As String = ""
            SQL = "  UPDATE " & putQuotes(Tabla) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEIIDDW") & " ='" & IDDW & "', " & vbCrLf
            SQL &= " " & putQuotes("U_SEIURLDW") & " ='" & URLDW & "' " & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And COALESCE(" & putQuotes("CardCode") & ",'') = '" & CardCode & "'" & vbCrLf
            SQL &= " And COALESCE(" & putQuotes("OpprId") & ",-1) = '" & Number & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("Actualización nula al ejecutar sentencia " & SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

    End Sub

    Private Function getNumberDeOportunidad(ByVal Tabla As String,
                                            ByVal CardCode As String,
                                            ByVal DOCIDDW As String,
                                            ByVal Numero As String,
                                            ByVal Nombre As String,
                                            ByVal Sociedad As eSociedad) As String

        'Devuelve el Number del Oportunidad

        Dim retVal As String = ""

        Try

            'Buscamos por Docnum
            Dim SQL As String = ""
            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " T0." & putQuotes("OpprId") & vbCrLf
            SQL &= " FROM " & putQuotes(Tabla) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("CardCode") & " = '" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then
                SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = '" & DOCIDDW & "'" & vbCrLf
            End If

            If Not String.IsNullOrEmpty(Numero) Then
                SQL &= " And T0." & putQuotes("OpprId") & " = '" & Numero & "'" & vbCrLf
            End If

            If Not String.IsNullOrEmpty(Nombre) Then
                SQL &= " And T0." & putQuotes("Name") & " = '" & Nombre & "'" & vbCrLf
            End If

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getNumberDeOpprID(ByVal Tabla As String,
                                       ByVal OpprID As String,
                                       ByVal Sociedad As eSociedad) As String

        'Devuelve el Number del Oportunidad

        Dim retVal As String = ""

        Try

            'Buscamos por Docnum
            Dim SQL As String = ""
            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " T0." & putQuotes("OpprId") & vbCrLf
            SQL &= " FROM " & putQuotes(Tabla) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("OpprId") & " = '" & OpprID & "'" & vbCrLf

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
