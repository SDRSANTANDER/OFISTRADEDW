Imports SAPbobsCOM
Imports System.Globalization
Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class clsLote

#Region "Públicas"

    Public Function ActualizarLote(ByVal objLote As EntLoteSAP, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oCompany As Company
        Dim oCompanyService As CompanyService = Nothing
        Dim oBatchNumbersService As BatchNumberDetailsService = Nothing
        Dim oBatchNumberDetailParams As BatchNumberDetailParams = Nothing
        Dim oBatchNumberDetail As BatchNumberDetail = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "ACTUALIZAR LOTE"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de lote: " & objLote.NumLote)

            oCompany = ConexionSAP.getCompany(objLote.UserSAP, objLote.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            oCompanyService = oCompany.GetCompanyService()

            'Obligatorios
            If String.IsNullOrEmpty(objLote.Articulo) AndAlso (String.IsNullOrEmpty(objLote.RefOrigen) OrElse String.IsNullOrEmpty(objLote.TipoRefOrigen) OrElse Not objLote.ObjTypeOrigen > 0) Then _
                Throw New Exception("Artículo o referencia origen no suministrada")

            If String.IsNullOrEmpty(objLote.NumLote) Then Throw New Exception("Número de lote no suministrado")

            If objLote.CamposUsuario Is Nothing OrElse objLote.CamposUsuario.Count = 0 Then Throw New Exception("Campos de usuario no suministrados")


            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objLote.NIFTercero, objLote.RazonSocial, objLote.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then
                clsLog.Log.Info("(" & sLogInfo & ") No encuentro IC con NIF: " & objLote.NIFTercero & ", Razón social: " & objLote.RazonSocial)
            Else
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Obtiene el artículo a partir de los datos del documento
            If String.IsNullOrEmpty(objLote.Articulo) Then

                'Buscamos tabla por ObjectType
                Dim Tabla As String = getTablaDeObjType(objLote.ObjTypeOrigen)
                If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objLote.ObjTypeOrigen)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

                'Obtenemos el artículo
                objLote.Articulo = getArticuloDeRefOrigen(Tabla, objLote.RefOrigen, objLote.TipoRefOrigen, objLote.LineNum, CardCode, Sociedad)
                If String.IsNullOrEmpty(objLote.Articulo) Then Throw New Exception("No encuentro artículo en documento con RefOrigen: '" & objLote.RefOrigen & "' y LineNum: '" & objLote.LineNum & "'")
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado artículo: " & objLote.Articulo)

            Else

                'Obtiene el artículo a partir del JSON

                'Artículo (si no se encuentra el código se busca por la referencia de proveedor)
                If Not String.IsNullOrEmpty(CardCode) Then
                    Dim ItemCode As String = oComun.getItemCode(CardCode, objLote.Articulo, objLote.RefExt, Sociedad)
                    If String.IsNullOrEmpty(ItemCode) Then Throw New Exception("No encuentro artículo con código: " & objLote.Articulo & " o referencia externa: " & objLote.RefExt & " para IC: " & CardCode)
                    clsLog.Log.Info("(" & sLogInfo & ") Encontrado artículo: " & ItemCode)

                    objLote.Articulo = ItemCode
                End If

            End If

            'Comprueba que existe el lote con el artículo
            Dim AbsEntry As String = getAbsEntryDeLote(TablaSAP.Lotes, objLote.Articulo, objLote.NumLote, Sociedad)
            If String.IsNullOrEmpty(AbsEntry) Then Throw New Exception("No encuentro lote con artículo: '" & objLote.Articulo & "' y NumLote: '" & objLote.NumLote & "'")
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado AbsEntry: " & AbsEntry)

            'Objeto lote
            oBatchNumbersService = oCompanyService.GetBusinessService(ServiceTypes.BatchNumberDetailsService)
            oBatchNumberDetailParams = oBatchNumbersService.GetDataInterface(BatchNumberDetailsServiceDataInterfaces.bndsBatchNumberDetailParams)

            oBatchNumberDetailParams.DocEntry = CInt(AbsEntry)

            oBatchNumberDetail = oBatchNumbersService.Get(oBatchNumberDetailParams)

            'Campos de usuario
            If Not objLote.CamposUsuario Is Nothing AndAlso objLote.CamposUsuario.Count > 0 Then

                For Each oCampoUsuario In objLote.CamposUsuario

                    Dim oUserField As Field = oBatchNumberDetail.UserFields.Item(oCampoUsuario.Campo)

                    Select Case oUserField.Type

                        Case BoFieldTypes.db_Numeric
                            If oUserField.SubType = BoFldSubTypes.st_Time Then
                                If DateTime.TryParseExact(oCampoUsuario.Valor, "HHmm", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                    oBatchNumberDetail.UserFields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "HHmm", CultureInfo.CurrentCulture)
                            Else
                                If IsNumeric(oCampoUsuario.Valor) Then _
                                    oBatchNumberDetail.Lines.UserFields.Item(oCampoUsuario.Campo).Value = CInt(oCampoUsuario.Valor)
                            End If

                        Case BoFieldTypes.db_Float
                            If IsNumeric(oCampoUsuario.Valor) Then _
                                oBatchNumberDetail.UserFields.Item(oCampoUsuario.Campo).Value = CDbl(oCampoUsuario.Valor.Replace(".", ","))

                        Case BoFieldTypes.db_Date
                            If DateTime.TryParseExact(oCampoUsuario.Valor, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                oBatchNumberDetail.UserFields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                        Case Else
                            If Not String.IsNullOrEmpty(oCampoUsuario.Valor) Then _
                                oBatchNumberDetail.UserFields.Item(oCampoUsuario.Campo).Value = oCampoUsuario.Valor

                    End Select

                Next

            End If

            'Actualiza el lote
            oBatchNumbersService.Update(oBatchNumberDetail)

            retVal.CODIGO = Respuesta.Ok
            retVal.MENSAJE = "Lote actualizado con éxito"
            retVal.MENSAJEAUX = objLote.NumLote

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        Finally
            oComun.LiberarObjCOM(oBatchNumberDetail)
            oComun.LiberarObjCOM(oCompanyService)
            oComun.LiberarObjCOM(oBatchNumbersService)
            oComun.LiberarObjCOM(oBatchNumberDetailParams)
        End Try

        Return retVal

    End Function

#End Region

#Region "Privadas"

    Private Function getAbsEntryDeLote(ByVal Tabla As String,
                                       ByVal Articulo As String,
                                       ByVal NumLote As String,
                                       ByVal Sociedad As eSociedad) As String

        'Devuelve el AbsEntry del lote

        Dim retVal As String = ""

        Try

            'Buscamos por artículo y número de lote
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("AbsEntry") & " As AbsEntry " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("ItemCode") & "  = N'" & Articulo & "'" & vbCrLf
            SQL &= " And T0." & putQuotes("DistNumber") & "  = N'" & NumLote & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso IsNumeric(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getArticuloDeRefOrigen(ByVal Tabla As String,
                                            ByVal RefOrigen As String,
                                            ByVal TipoRefOrigen As String,
                                            ByVal LineNum As Integer,
                                            ByVal CardCode As String,
                                            ByVal Sociedad As eSociedad) As String

        Dim retVal As String = ""

        Try

            Dim SQL As String = ""
            SQL = "  SELECT  " & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("ItemCode") & ",'') As ItemCode " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " INNER JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " ON T0." & putQuotes("DocEntry") & " = T1." & putQuotes("DocEntry") & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            Select Case TipoRefOrigen
                Case Utilidades.RefOrigen.DocNum
                    SQL &= " And T0." & putQuotes("DocNum") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Utilidades.RefOrigen.DocEntry
                    SQL &= " And T0." & putQuotes("DocEntry") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Else
                    SQL &= " And T0." & putQuotes("NumAtCard") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
            End Select

            SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            SQL &= " And T1." & putQuotes("LineNum") & "  = N'" & LineNum & "'" & vbCrLf
            'SQL &= " And T0." & putQuotes("DocStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf
            'SQL &= " And T1." & putQuotes("LineStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf
            SQL &= " And T0." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            Dim DT As DataTable = oCon.ObtenerDT(SQL)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

#End Region

End Class
