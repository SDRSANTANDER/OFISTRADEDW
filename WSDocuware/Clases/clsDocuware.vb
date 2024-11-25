Imports SAPbobsCOM
Imports System.Reflection
Imports System.Globalization
Imports WSDocuware.Utilidades

Public Class clsDocuware

#Region "Públicas"

    Public Function setDocuware(ByVal objDocuware As EntDocuware, ByVal Sociedad As eSociedad) As EntResultado

        'Actualiza campos docuware
        Dim retVal As New EntResultado

        Try

            Select Case objDocuware.ObjType

                Case ObjType.LlamadaServicio
                    retVal = ActualizarLlamadaServicio(objDocuware, Sociedad)

                Case ObjType.Actividad
                    retVal = ActualizarActividad(objDocuware, Sociedad)

                Case ObjType.AcuerdoGlobal
                    retVal = ActualizarAcuerdoGlobal(objDocuware, Sociedad)

                Case ObjType.Asiento
                    retVal = ActualizarAsiento(objDocuware, Sociedad)

                Case ObjType.SolicitudCompra
                    retVal = ActualizarSolicitud(objDocuware, Sociedad)

                Case ObjType.Oportunidad
                    retVal = ActualizarOportunidad(objDocuware, Sociedad)

                Case ObjType.OrdenProduccion
                    retVal = ActualizarOrdenProduccion(objDocuware, Sociedad)

                Case ObjType.PrecioEntrega
                    retVal = ActualizarPrecioEntrega(objDocuware, Sociedad)

                Case ObjType.EntradaMercancias, ObjType.SalidaMercancias, ObjType.Traslado
                    retVal = ActualizarInventario(objDocuware, Sociedad)

                Case ObjType.Cobro, ObjType.Pago
                    retVal = ActualizarCobroPago(objDocuware, Sociedad)

                Case Else
                    retVal = ActualizarDocumento(objDocuware, Sociedad)

            End Select

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

#Region "Solicitud"

    Private Function ActualizarSolicitud(ByVal objDocuware As EntDocuware, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "DOCUWARE SOLICITUD"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de solicitud: IDDW " & objDocuware.DOCIDDW)

            'Obligatorios
            If String.IsNullOrEmpty(objDocuware.RefOrigen) OrElse String.IsNullOrEmpty(objDocuware.TipoRefOrigen) Then _
                Throw New Exception("Referencia y/o tipo de referencia no suministrados")

            If String.IsNullOrEmpty(objDocuware.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocuware.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")

            'Buscamos IC por NIF o razón social
            Dim CardCode As String = ""
            If Not String.IsNullOrEmpty(objDocuware.NIFTercero) AndAlso Not String.IsNullOrEmpty(objDocuware.RazonSocial) Then
                CardCode = oComun.getCardCode(objDocuware.NIFTercero, objDocuware.RazonSocial, objDocuware.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocuware.NIFTercero & ", Razón social: " & objDocuware.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocuware.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocuware.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objDocuware.RefOrigen.Split("#")

            'Recorre cada uno de las solicitudes origen
            For Each RefOrigen In RefsOrigen

                'Comprueba que existe el solicitud origen y que se puede acceder a él
                'Pueden ser varios solicitudes origen con la misma referencia
                Dim DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)) = getDocEntryNumDeRefOrigenSoliciutd(Tabla, RefOrigen, objDocuware.TipoRefOrigen, CardCode, Sociedad)
                If DocEntryNums Is Nothing OrElse DocEntryNums.Count = 0 Then Throw New Exception("No existe solicitud con referencia: " & RefOrigen & " o su estado es cancelado")

                'Se actualizan los campos de usuario de DW
                setDatosDWSolicitud(Tabla, CardCode, DocEntryNums, objDocuware.DOCIDDW, objDocuware.DOCURLDW, objDocuware.DOCESTADODW, objDocuware.DOCMOTIVODW, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Solicitud actualizada con éxito"
                retVal.MENSAJEAUX = String.Join("#", DocEntryNums.Select(Function(p) p.Value))

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Private Function getDocEntryNumDeRefOrigenSoliciutd(ByVal Tabla As String,
                                                        ByVal RefOrigen As String,
                                                        ByVal TipoRefOrigen As String,
                                                        ByVal CardCode As String,
                                                        ByVal Sociedad As eSociedad) As List(Of KeyValuePair(Of Integer, Integer))

        'Devuelve el DocEntry del documento

        Dim retVal As New List(Of KeyValuePair(Of Integer, Integer))

        Try

            'Buscamos por DocEntryNum, DocEntry o NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("DocEntry") & " As DocEntry " & vbCrLf
            SQL &= ",T0." & putQuotes("DocNum") & " As DocNum " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " On T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            Select Case TipoRefOrigen
                Case Utilidades.RefOrigen.DocEntry
                    SQL &= " And T0." & putQuotes("DocEntry") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Utilidades.RefOrigen.DocNum
                    SQL &= " And T0." & putQuotes("DocNum") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Else
                    SQL &= " And T0." & putQuotes("NumAtCard") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
            End Select

            If Not String.IsNullOrEmpty(CardCode) Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            SQL &= " And T0." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(New KeyValuePair(Of Integer, Integer)(DR.Item("DocEntry"), DR.Item("DocNum")))
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Sub setDatosDWSolicitud(ByVal Tabla As String,
                                    ByVal CardCode As String,
                                    ByVal DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)),
                                    ByVal IDDW As String,
                                    ByVal URLDW As String,
                                    ByVal ESTADODW As Integer,
                                    ByVal MOTIVODW As String,
                                    ByVal Sociedad As eSociedad)

        'Actualiza los campos de DW en el documento

        Try

            'Actualizamos por CardCode, DocDate y DocEntryNum
            Dim SQL As String = ""

            SQL = "  UPDATE T0 SET" & vbCrLf
            SQL &= " T0." & putQuotes("U_SEIIDDW") & " =N'" & IDDW & "' " & vbCrLf
            SQL &= " ,T0." & putQuotes("U_SEIURLDW") & " =N'" & URLDW & "' " & vbCrLf

            If ESTADODW > -1 Then SQL &= " ,T0." & putQuotes("U_SEIESTDW") & " =N'" & ESTADODW & "' " & vbCrLf
            If Not String.IsNullOrEmpty(MOTIVODW) Then SQL &= " ,T0." & putQuotes("U_SEIMOTDW") & " =N'" & ESTADODW & "' " & vbCrLf

            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " On T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            SQL &= " And COALESCE(T1." & putQuotes("LineVendor") & ",'') = N'" & CardCode & "'" & vbCrLf

            SQL &= " And COALESCE(T0." & putQuotes("DocEntry") & ",-1)  IN ('-1' " & vbCrLf
            For Each DocEntryNum In DocEntryNums
                SQL &= ", N'" & DocEntryNum.Key & "'" & vbCrLf
            Next
            SQL &= ") " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("No se ha podido actualizar la solicitud: " & putQuotes(Tabla) & " para CardCode: " & CardCode & " y DocNum: " & String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList))

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

    End Sub

#End Region

#Region "Documento"

    Private Function ActualizarDocumento(ByVal objDocuware As EntDocuware, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "DOCUWARE DOCUMENTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de documento: IDDW " & objDocuware.DOCIDDW)

            'Obligatorios
            If String.IsNullOrEmpty(objDocuware.RefOrigen) OrElse String.IsNullOrEmpty(objDocuware.TipoRefOrigen) Then _
                Throw New Exception("Referencia y/o tipo de referencia no suministrados")

            If String.IsNullOrEmpty(objDocuware.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocuware.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")

            'Buscamos IC por NIF o razón social
            Dim CardCode As String = ""
            If Not String.IsNullOrEmpty(objDocuware.NIFTercero) AndAlso Not String.IsNullOrEmpty(objDocuware.RazonSocial) Then
                CardCode = oComun.getCardCode(objDocuware.NIFTercero, objDocuware.RazonSocial, objDocuware.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocuware.NIFTercero & ", Razón social: " & objDocuware.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocuware.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocuware.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objDocuware.RefOrigen.Split("#")

            'Recorre cada uno de los documento origen
            For Each RefOrigen In RefsOrigen

                'Comprueba que existe el documento origen y que se puede acceder a él
                'Pueden ser varios documentos origen con la misma referencia
                Dim DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)) = getDocEntryNumDeRefOrigenDocumento(Tabla, RefOrigen, objDocuware.TipoRefOrigen, CardCode, Sociedad)
                If DocEntryNums Is Nothing OrElse DocEntryNums.Count = 0 Then Throw New Exception("No existe documento con referencia: " & RefOrigen & " o su estado es cancelado")

                'Se actualizan los campos de usuario de DW
                setDatosDWDocumento(Tabla, CardCode, DocEntryNums, objDocuware.DOCIDDW, objDocuware.DOCURLDW, objDocuware.DOCTRATADODW, objDocuware.DOCESTADODW, objDocuware.DOCMOTIVODW, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento actualizado con éxito"
                retVal.MENSAJEAUX = String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Private Function getDocEntryNumDeRefOrigenDocumento(ByVal Tabla As String,
                                                        ByVal RefOrigen As String,
                                                        ByVal TipoRefOrigen As String,
                                                        ByVal CardCode As String,
                                                        ByVal Sociedad As eSociedad) As List(Of KeyValuePair(Of Integer, Integer))

        'Devuelve el DocEntry del documento

        Dim retVal As New List(Of KeyValuePair(Of Integer, Integer))

        Try

            'Buscamos por DocEntryNum, DocEntry o NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= "  T0." & putQuotes("DocEntry") & " As DocEntry " & vbCrLf
            SQL &= " ,T0." & putQuotes("DocNum") & " As DocNum " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            Select Case TipoRefOrigen
                Case Utilidades.RefOrigen.DocEntry
                    SQL &= " And T0." & putQuotes("DocEntry") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Utilidades.RefOrigen.DocNum
                    SQL &= " And T0." & putQuotes("DocNum") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Else
                    SQL &= " And T0." & putQuotes("NumAtCard") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
            End Select

            If Not String.IsNullOrEmpty(CardCode) Then SQL &= " And T0." & putQuotes("CardCode") & " = N'" & CardCode & "'" & vbCrLf

            SQL &= " And T0." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(New KeyValuePair(Of Integer, Integer)(DR.Item("DocEntry"), DR.Item("DocNum")))
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Sub setDatosDWDocumento(ByVal Tabla As String,
                                    ByVal CardCode As String,
                                    ByVal DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)),
                                    ByVal IDDW As String,
                                    ByVal URLDW As String,
                                    ByVal TRATADODW As String,
                                    ByVal ESTADODW As Integer,
                                    ByVal MOTIVODW As String,
                                    ByVal Sociedad As eSociedad)

        'Actualiza los campos de DW en el documento

        Try

            'Actualizamos por CardCode, DocDate y DocEntryNum
            Dim SQL As String = ""
            SQL = "  UPDATE " & getDataBaseRef(Tabla, Sociedad) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEIIDDW") & " =N'" & IDDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIURLDW") & " =N'" & URLDW & "' " & vbCrLf

            If Not String.IsNullOrEmpty(TRATADODW) Then SQL &= " ," & putQuotes("U_SEIMOT") & " =N'" & TRATADODW & "' " & vbCrLf
            If ESTADODW > -1 Then SQL &= " ," & putQuotes("U_SEIESTDW") & " =N'" & ESTADODW & "' " & vbCrLf
            If Not String.IsNullOrEmpty(MOTIVODW) Then SQL &= " ," & putQuotes("U_SEIMOTDW") & " =N'" & ESTADODW & "' " & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) Then SQL &= " And COALESCE(" & putQuotes("CardCode") & ",'') = N'" & CardCode & "'" & vbCrLf

            SQL &= " And COALESCE(" & putQuotes("DocEntry") & ",-1)  IN ('-1' " & vbCrLf
            For Each DocEntryNum In DocEntryNums
                SQL &= ", N'" & DocEntryNum.Key & "'" & vbCrLf
            Next
            SQL &= ") " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("No se ha podido actualizar el documento: " & putQuotes(Tabla) & " para CardCode: " & CardCode & " y DocNum: " & String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList))

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

    End Sub

#End Region

#Region "Oportunidad"

    Private Function ActualizarOportunidad(ByVal objDocuware As EntDocuware, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "DOCUWARE OPORTUNIDAD"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de oportunidad: IDDW " & objDocuware.DOCIDDW)

            'Obligatorios
            If String.IsNullOrEmpty(objDocuware.RefOrigen) OrElse String.IsNullOrEmpty(objDocuware.TipoRefOrigen) Then _
                Throw New Exception("Referencia y/o tipo de referencia no suministrados")

            If String.IsNullOrEmpty(objDocuware.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocuware.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")

            'Buscamos IC por NIF o razón social
            Dim CardCode As String = ""
            If Not String.IsNullOrEmpty(objDocuware.NIFTercero) AndAlso Not String.IsNullOrEmpty(objDocuware.RazonSocial) Then
                CardCode = oComun.getCardCode(objDocuware.NIFTercero, objDocuware.RazonSocial, objDocuware.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocuware.NIFTercero & ", Razón social: " & objDocuware.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocuware.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocuware.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objDocuware.RefOrigen.Split("#")

            'Recorre cada uno de los oportunidad origen
            For Each RefOrigen In RefsOrigen

                'Comprueba que existe el oportunidad origen y que se puede acceder a él
                'Pueden ser varios oportunidads origen con la misma referencia
                Dim DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)) = getDocEntryNumDeRefOrigenOportunidad(Tabla, RefOrigen, objDocuware.TipoRefOrigen, CardCode, Sociedad)
                If DocEntryNums Is Nothing OrElse DocEntryNums.Count = 0 Then Throw New Exception("No existe oportunidad con referencia: " & RefOrigen & " o su estado es cancelado")

                'Se actualizan los campos de usuario de DW
                setDatosDWOportunidad(Tabla, CardCode, DocEntryNums, objDocuware.DOCIDDW, objDocuware.DOCURLDW, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Oportunidad actualizada con éxito"
                retVal.MENSAJEAUX = String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Private Function getDocEntryNumDeRefOrigenOportunidad(ByVal Tabla As String,
                                                        ByVal RefOrigen As String,
                                                        ByVal TipoRefOrigen As String,
                                                        ByVal CardCode As String,
                                                        ByVal Sociedad As eSociedad) As List(Of KeyValuePair(Of Integer, Integer))

        'Devuelve el DocEntry del oportunidad

        Dim retVal As New List(Of KeyValuePair(Of Integer, Integer))

        Try

            'Buscamos por DocEntryNum, DocEntry o NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= "  T0." & putQuotes("OpprId") & " As DocEntry " & vbCrLf
            SQL &= " ,T0." & putQuotes("OpprId") & " As DocNum " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            Select Case TipoRefOrigen
                Case Utilidades.RefOrigen.DocEntry
                    SQL &= " And T0." & putQuotes("OpprId") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Utilidades.RefOrigen.DocNum
                    SQL &= " And T0." & putQuotes("OpprId") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Else
                    SQL &= " And T0." & putQuotes("Name") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
            End Select

            If Not String.IsNullOrEmpty(CardCode) Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            'SQL &= " And T0." & putQuotes("DocStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf
            'SQL &= " And T0." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(New KeyValuePair(Of Integer, Integer)(DR.Item("DocEntry"), DR.Item("DocNum")))
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Sub setDatosDWOportunidad(ByVal Tabla As String,
                                      ByVal CardCode As String,
                                      ByVal DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)),
                                      ByVal IDDW As String,
                                      ByVal URLDW As String,
                                      ByVal Sociedad As eSociedad)

        'Actualiza los campos de DW en el oportunidad

        Try

            'Actualizamos por CardCode, DocDate y DocEntryNum
            Dim SQL As String = ""
            SQL = "  UPDATE " & getDataBaseRef(Tabla, Sociedad) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEIIDDW") & " =N'" & IDDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIURLDW") & " =N'" & URLDW & "' " & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) Then SQL &= " And COALESCE(" & putQuotes("CardCode") & ",'') = N'" & CardCode & "'" & vbCrLf

            SQL &= " And COALESCE(" & putQuotes("OpprId") & ",-1)  IN ('-1' " & vbCrLf
            For Each DocEntryNum In DocEntryNums
                SQL &= ", N'" & DocEntryNum.Key & "'" & vbCrLf
            Next
            SQL &= ") " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("No se ha podido actualizar la oportunidad: " & putQuotes(Tabla) & " para CardCode: " & CardCode & " y DocNum: " & String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList))

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

    End Sub

#End Region

#Region "Orden produccion"

    Private Function ActualizarOrdenProduccion(ByVal objDocuware As EntDocuware, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "DOCUWARE ORDEN PRODUCCION"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de orden produccion: IDDW " & objDocuware.DOCIDDW)

            'Obligatorios
            If String.IsNullOrEmpty(objDocuware.RefOrigen) OrElse String.IsNullOrEmpty(objDocuware.TipoRefOrigen) Then _
                Throw New Exception("Referencia y/o tipo de referencia no suministrados")

            If String.IsNullOrEmpty(objDocuware.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocuware.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocuware.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocuware.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objDocuware.RefOrigen.Split("#")

            'Recorre cada uno de las orden producciones origen
            For Each RefOrigen In RefsOrigen

                'Comprueba que existe la orden produccion origen y que se puede acceder a él
                'Pueden ser varios orden producciones origen con la misma referencia
                Dim DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)) = getDocEntryNumDeRefOrigenOrdenProduccion(Tabla, RefOrigen, objDocuware.TipoRefOrigen, Sociedad)
                If DocEntryNums Is Nothing OrElse DocEntryNums.Count = 0 Then Throw New Exception("No existe orden produccion con referencia: " & RefOrigen & " o su estado es cancelado")

                'Se actualizan los campos de usuario de DW
                setDatosDWOrdenProduccion(Tabla, DocEntryNums, objDocuware.DOCIDDW, objDocuware.DOCURLDW, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Orden produccion actualizada con éxito"
                retVal.MENSAJEAUX = String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Private Function getDocEntryNumDeRefOrigenOrdenProduccion(ByVal Tabla As String,
                                                              ByVal RefOrigen As String,
                                                              ByVal TipoRefOrigen As String,
                                                              ByVal Sociedad As eSociedad) As List(Of KeyValuePair(Of Integer, Integer))

        'Devuelve el DocEntry del documento

        Dim retVal As New List(Of KeyValuePair(Of Integer, Integer))

        Try

            'Buscamos por DocEntry, DocEntry o NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("DocEntry") & " As DocEntry " & vbCrLf
            SQL &= ",T0." & putQuotes("DocNum") & " As DocNum " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            Select Case TipoRefOrigen
                Case Utilidades.RefOrigen.DocEntry
                    SQL &= " And T0." & putQuotes("DocEntry") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Utilidades.RefOrigen.DocNum
                    SQL &= " And T0." & putQuotes("DocNum") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Else
                    SQL &= " And T0." & putQuotes("DocNum") & "  = N'-1'" & vbCrLf
            End Select

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(New KeyValuePair(Of Integer, Integer)(DR.Item("DocEntry"), DR.Item("DocNum")))
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Sub setDatosDWOrdenProduccion(ByVal Tabla As String,
                                          ByVal DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)),
                                          ByVal IDDW As String,
                                          ByVal URLDW As String,
                                          ByVal Sociedad As eSociedad)

        'Actualiza los campos de DW en el documento

        Try

            'Actualizamos por CardCode, DocDate y DocEntryNum
            Dim SQL As String = ""
            SQL = "  UPDATE " & getDataBaseRef(Tabla, Sociedad) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEIIDDW") & " =N'" & IDDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIURLDW") & " =N'" & URLDW & "' " & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            SQL &= " And COALESCE(" & putQuotes("DocEntry") & ",-1)  IN ('-1' " & vbCrLf
            For Each DocEntryNum In DocEntryNums
                SQL &= ", N'" & DocEntryNum.Key & "'" & vbCrLf
            Next
            SQL &= ") " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("No se ha podido actualizar la orden produccion: " & putQuotes(Tabla) & " para DocNum: " & String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList))

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

    End Sub

#End Region

#Region "Llamada servicio"

    Private Function ActualizarLlamadaServicio(ByVal objDocuware As EntDocuware, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "DOCUWARE LLAMADA SERVICIO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de llamada servicio: IDDW " & objDocuware.DOCIDDW)

            'Obligatorios
            If String.IsNullOrEmpty(objDocuware.RefOrigen) OrElse String.IsNullOrEmpty(objDocuware.TipoRefOrigen) Then _
                Throw New Exception("Referencia y/o tipo de referencia no suministrados")

            If String.IsNullOrEmpty(objDocuware.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocuware.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")

            'Buscamos IC por NIF o razón social
            Dim CardCode As String = ""
            If Not String.IsNullOrEmpty(objDocuware.NIFTercero) AndAlso Not String.IsNullOrEmpty(objDocuware.RazonSocial) Then
                CardCode = oComun.getCardCode(objDocuware.NIFTercero, objDocuware.RazonSocial, objDocuware.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocuware.NIFTercero & ", Razón social: " & objDocuware.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocuware.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocuware.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objDocuware.RefOrigen.Split("#")

            'Recorre cada uno de los llamada servicio origen
            For Each RefOrigen In RefsOrigen

                'Comprueba que existe la llamada servicio origen y que se puede acceder a él
                'Pueden ser varios llamada servicios origen con la misma referencia
                Dim DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)) = getDocEntryNumDeRefOrigenLlamadaServicio(Tabla, RefOrigen, objDocuware.TipoRefOrigen, CardCode, Sociedad)
                If DocEntryNums Is Nothing OrElse DocEntryNums.Count = 0 Then Throw New Exception("No existe llamada servicio con referencia: " & RefOrigen & " o su estado es cancelado")

                'Se actualizan los campos de usuario de DW
                setDatosDWLlamadaservicio(Tabla, CardCode, DocEntryNums, objDocuware.DOCIDDW, objDocuware.DOCURLDW, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Llamada servicio actualizado con éxito"
                retVal.MENSAJEAUX = String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Private Function getDocEntryNumDeRefOrigenLlamadaServicio(ByVal Tabla As String,
                                                              ByVal RefOrigen As String,
                                                              ByVal TipoRefOrigen As String,
                                                              ByVal CardCode As String,
                                                              ByVal Sociedad As eSociedad) As List(Of KeyValuePair(Of Integer, Integer))

        'Devuelve el DocEntry dla llamada servicio

        Dim retVal As New List(Of KeyValuePair(Of Integer, Integer))

        Try

            'Buscamos por DocEntryNum, DocEntry o NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= "  T0." & putQuotes("callID") & " As CallID " & vbCrLf
            SQL &= " ,T0." & putQuotes("DocNum") & " As DocNum " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            Select Case TipoRefOrigen
                Case Utilidades.RefOrigen.DocEntry
                    SQL &= " And T0." & putQuotes("callID") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Utilidades.RefOrigen.DocNum
                    SQL &= " And T0." & putQuotes("DocNum") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Else
                    SQL &= " And T0." & putQuotes("subject") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
            End Select

            If Not String.IsNullOrEmpty(CardCode) Then SQL &= " And T0." & putQuotes("customer") & "  = N'" & CardCode & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(New KeyValuePair(Of Integer, Integer)(DR.Item("CallID"), DR.Item("DocNum")))
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Sub setDatosDWLlamadaservicio(ByVal Tabla As String,
                                          ByVal CardCode As String,
                                          ByVal DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)),
                                          ByVal IDDW As String,
                                          ByVal URLDW As String,
                                          ByVal Sociedad As eSociedad)

        'Actualiza los campos de DW en la llamada servicio

        Try

            'Actualizamos por CardCode, DocDate y DocEntry
            Dim SQL As String = ""
            SQL = "  UPDATE " & getDataBaseRef(Tabla, Sociedad) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEIIDDW") & " =N'" & IDDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIURLDW") & " =N'" & URLDW & "' " & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) Then SQL &= " And COALESCE(" & putQuotes("customer") & ",'') = N'" & CardCode & "'" & vbCrLf

            SQL &= " And COALESCE(" & putQuotes("callID") & ",-1)  IN ('-1' " & vbCrLf
            For Each DocEntryNum In DocEntryNums
                SQL &= ", N'" & DocEntryNum.Key & "'" & vbCrLf
            Next
            SQL &= ") " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("No se ha podido actualizar la llamada servicio: " & putQuotes(Tabla) & " para CardCode: " & CardCode & " y DocNum: " & String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList))

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

    End Sub

#End Region

#Region "Actividad"

    Private Function ActualizarActividad(ByVal objDocuware As EntDocuware, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "DOCUWARE ACTIVIDAD"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de actividad: IDDW " & objDocuware.DOCIDDW)

            'Obligatorios
            If String.IsNullOrEmpty(objDocuware.RefOrigen) OrElse String.IsNullOrEmpty(objDocuware.TipoRefOrigen) Then _
                Throw New Exception("Referencia y/o tipo de referencia no suministrados")

            If String.IsNullOrEmpty(objDocuware.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocuware.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")

            'Buscamos IC por NIF o razón social
            Dim CardCode As String = ""
            If Not String.IsNullOrEmpty(objDocuware.NIFTercero) AndAlso Not String.IsNullOrEmpty(objDocuware.RazonSocial) Then
                CardCode = oComun.getCardCode(objDocuware.NIFTercero, objDocuware.RazonSocial, objDocuware.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocuware.NIFTercero & ", Razón social: " & objDocuware.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocuware.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocuware.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objDocuware.RefOrigen.Split("#")

            'Recorre cada uno de los actividad origen
            For Each RefOrigen In RefsOrigen

                'Comprueba que existe la actividad origen y que se puede acceder a él
                'Pueden ser varios actividads origen con la misma referencia
                Dim DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)) = getDocEntryNumDeRefOrigenActividad(Tabla, RefOrigen, objDocuware.TipoRefOrigen, CardCode, Sociedad)
                If DocEntryNums Is Nothing OrElse DocEntryNums.Count = 0 Then Throw New Exception("No existe actividad con referencia: " & RefOrigen & " o su estado es cancelado")

                'Se actualizan los campos de usuario de DW
                setDatosDWActividad(Tabla, CardCode, DocEntryNums, objDocuware.DOCIDDW, objDocuware.DOCURLDW, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Actividad actualizada con éxito"
                retVal.MENSAJEAUX = String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Private Function getDocEntryNumDeRefOrigenActividad(ByVal Tabla As String,
                                                        ByVal RefOrigen As String,
                                                        ByVal TipoRefOrigen As String,
                                                        ByVal CardCode As String,
                                                        ByVal Sociedad As eSociedad) As List(Of KeyValuePair(Of Integer, Integer))

        'Devuelve el DocEntry dla actividad

        Dim retVal As New List(Of KeyValuePair(Of Integer, Integer))

        Try

            'Buscamos por DocEntryNum, DocEntry o NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= "  T0." & putQuotes("ClgCode") & " As ClgCode " & vbCrLf
            SQL &= " ,T0." & putQuotes("ClgCode") & " As DocNum " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            Select Case TipoRefOrigen
                Case Utilidades.RefOrigen.DocEntry
                    SQL &= " And T0." & putQuotes("ClgCode") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Utilidades.RefOrigen.DocNum
                    SQL &= " And T0." & putQuotes("ClgCode") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Else
                    SQL &= " And T0." & putQuotes("ClgCode") & "  = N'-1'" & vbCrLf
            End Select

            If Not String.IsNullOrEmpty(CardCode) Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(New KeyValuePair(Of Integer, Integer)(DR.Item("ClgCode"), DR.Item("DocNum")))
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Sub setDatosDWActividad(ByVal Tabla As String,
                                          ByVal CardCode As String,
                                          ByVal DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)),
                                          ByVal IDDW As String,
                                          ByVal URLDW As String,
                                          ByVal Sociedad As eSociedad)

        'Actualiza los campos de DW en la actividad

        Try

            'Actualizamos por CardCode, DocDate y DocEntry
            Dim SQL As String = ""
            SQL = "  UPDATE " & getDataBaseRef(Tabla, Sociedad) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEIIDDW") & " =N'" & IDDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIURLDW") & " =N'" & URLDW & "' " & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) Then SQL &= " And COALESCE(" & putQuotes("CardCode") & ",'') = N'" & CardCode & "'" & vbCrLf

            SQL &= " And COALESCE(" & putQuotes("ClgCode") & ",-1)  IN ('-1' " & vbCrLf
            For Each DocEntryNum In DocEntryNums
                SQL &= ", N'" & DocEntryNum.Key & "'" & vbCrLf
            Next
            SQL &= ") " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("No se ha podido actualizar la actividad: " & putQuotes(Tabla) & " para CardCode: " & CardCode & " y DocNum: " & String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList))

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

    End Sub

#End Region

#Region "Acuerdo global"

    Private Function ActualizarAcuerdoGlobal(ByVal objDocuware As EntDocuware, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "DOCUWARE ACUERDO GLOBAL"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de acuerdo global: IDDW " & objDocuware.DOCIDDW)

            'Obligatorios
            If String.IsNullOrEmpty(objDocuware.RefOrigen) OrElse String.IsNullOrEmpty(objDocuware.TipoRefOrigen) Then _
                Throw New Exception("Referencia y/o tipo de referencia no suministrados")

            If String.IsNullOrEmpty(objDocuware.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocuware.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")

            'Buscamos IC por NIF o razón social
            Dim CardCode As String = ""
            If Not String.IsNullOrEmpty(objDocuware.NIFTercero) AndAlso Not String.IsNullOrEmpty(objDocuware.RazonSocial) Then
                CardCode = oComun.getCardCode(objDocuware.NIFTercero, objDocuware.RazonSocial, objDocuware.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocuware.NIFTercero & ", Razón social: " & objDocuware.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocuware.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocuware.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objDocuware.RefOrigen.Split("#")

            'Recorre cada uno de los acuerdo global origen
            For Each RefOrigen In RefsOrigen

                'Comprueba que existe la acuerdo global origen y que se puede acceder a él
                'Pueden ser varios acuerdo globals origen con la misma referencia
                Dim DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)) = getDocEntryNumDeRefOrigenAcuerdoGlobal(Tabla, RefOrigen, objDocuware.TipoRefOrigen, CardCode, Sociedad)
                If DocEntryNums Is Nothing OrElse DocEntryNums.Count = 0 Then Throw New Exception("No existe acuerdo global con referencia: " & RefOrigen & " o su estado es cancelado")

                'Se actualizan los campos de usuario de DW
                setDatosDWAcuerdoGlobal(Tabla, CardCode, DocEntryNums, objDocuware.DOCIDDW, objDocuware.DOCURLDW, objDocuware.DOCESTADODW, objDocuware.DOCMOTIVODW, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Acuerdo global actualizado con éxito"
                retVal.MENSAJEAUX = String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Private Function getDocEntryNumDeRefOrigenAcuerdoGlobal(ByVal Tabla As String,
                                                            ByVal RefOrigen As String,
                                                            ByVal TipoRefOrigen As String,
                                                            ByVal CardCode As String,
                                                            ByVal Sociedad As eSociedad) As List(Of KeyValuePair(Of Integer, Integer))

        'Devuelve el DocEntry dla acuerdo global

        Dim retVal As New List(Of KeyValuePair(Of Integer, Integer))

        Try

            'Buscamos por DocEntryNum, DocEntry o NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= "  T0." & putQuotes("AbsID") & " As AbsID " & vbCrLf
            SQL &= " ,T0." & putQuotes("AbsID") & " As DocNum " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            Select Case TipoRefOrigen
                Case Utilidades.RefOrigen.DocEntry
                    SQL &= " And T0." & putQuotes("AbsID") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Utilidades.RefOrigen.DocNum
                    SQL &= " And T0." & putQuotes("AbsID") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Else
                    SQL &= " And T0." & putQuotes("Descript") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
            End Select

            If Not String.IsNullOrEmpty(CardCode) Then SQL &= " And T0." & putQuotes("BPCode") & "  = N'" & CardCode & "'" & vbCrLf

            SQL &= " And T0." & putQuotes("Cancelled") & " = N'" & SN.No & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(New KeyValuePair(Of Integer, Integer)(DR.Item("AbsID"), DR.Item("DocNum")))
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Sub setDatosDWAcuerdoGlobal(ByVal Tabla As String,
                                        ByVal CardCode As String,
                                        ByVal DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)),
                                        ByVal IDDW As String,
                                        ByVal URLDW As String,
                                        ByVal ESTADODW As Integer,
                                        ByVal MOTIVODW As String,
                                        ByVal Sociedad As eSociedad)

        'Actualiza los campos de DW en la acuerdo global

        Try

            'Actualizamos por CardCode, DocDate y DocEntry
            Dim SQL As String = ""
            SQL = "  UPDATE " & getDataBaseRef(Tabla, Sociedad) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEIIDDW") & " =N'" & IDDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIURLDW") & " =N'" & URLDW & "' " & vbCrLf

            If ESTADODW > -1 Then SQL &= " ," & putQuotes("U_SEIESTDW") & " =N'" & ESTADODW & "' " & vbCrLf
            If Not String.IsNullOrEmpty(MOTIVODW) Then SQL &= " ," & putQuotes("U_SEIMOTDW") & " =N'" & ESTADODW & "' " & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) Then SQL &= " And COALESCE(" & putQuotes("BPCode") & ",'') = N'" & CardCode & "'" & vbCrLf

            SQL &= " And COALESCE(" & putQuotes("AbsID") & ",-1)  IN ('-1' " & vbCrLf
            For Each DocEntryNum In DocEntryNums
                SQL &= ", N'" & DocEntryNum.Key & "'" & vbCrLf
            Next
            SQL &= ") " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("No se ha podido actualizar el acuerdo global: " & putQuotes(Tabla) & " para CardCode: " & CardCode & " y DocNum: " & String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList))

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

    End Sub

#End Region

#Region "Asiento"

    Private Function ActualizarAsiento(ByVal objDocuware As EntDocuware, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "DOCUWARE ASIENTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de asiento: IDDW " & objDocuware.DOCIDDW)

            'Obligatorios
            If String.IsNullOrEmpty(objDocuware.RefOrigen) OrElse String.IsNullOrEmpty(objDocuware.TipoRefOrigen) Then _
                Throw New Exception("Referencia y/o tipo de referencia no suministrados")

            If String.IsNullOrEmpty(objDocuware.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocuware.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocuware.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocuware.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objDocuware.RefOrigen.Split("#")

            'Recorre cada uno de los Asiento origen
            For Each RefOrigen In RefsOrigen

                'Comprueba que existe la Asiento origen y que se puede acceder a él
                'Pueden ser varios Asientos origen con la misma referencia
                Dim DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)) = getDocEntryNumDeRefOrigenAsiento(Tabla, RefOrigen, objDocuware.TipoRefOrigen, Sociedad)
                If DocEntryNums Is Nothing OrElse DocEntryNums.Count = 0 Then Throw New Exception("No existe asiento con referencia: " & RefOrigen & " o su estado es cancelado")

                'Se actualizan los campos de usuario de DW
                setDatosDWAsiento(Tabla, DocEntryNums, objDocuware.DOCIDDW, objDocuware.DOCURLDW, objDocuware.DOCTRATADODW, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Asiento actualizado con éxito"
                retVal.MENSAJEAUX = String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Private Function getDocEntryNumDeRefOrigenAsiento(ByVal Tabla As String,
                                                              ByVal RefOrigen As String,
                                                              ByVal TipoRefOrigen As String,
                                                              ByVal Sociedad As eSociedad) As List(Of KeyValuePair(Of Integer, Integer))

        'Devuelve el DocEntry dla Asiento

        Dim retVal As New List(Of KeyValuePair(Of Integer, Integer))

        Try

            'Buscamos por DocEntryNum, DocEntry o NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= "  T0." & putQuotes("TransId") & " As TransId " & vbCrLf
            SQL &= " ,T0." & putQuotes("Number") & " As Number " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            Select Case TipoRefOrigen
                Case Utilidades.RefOrigen.DocEntry, Utilidades.RefOrigen.TransId
                    SQL &= " And T0." & putQuotes("TransId") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Utilidades.RefOrigen.DocNum, Utilidades.RefOrigen.TransNum
                    SQL &= " And T0." & putQuotes("Number") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Utilidades.RefOrigen.NumAtCard, Utilidades.RefOrigen.TransRef1
                    SQL &= " And T0." & putQuotes("Ref1") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
                Case Utilidades.RefOrigen.TransRef2
                    SQL &= " And T0." & putQuotes("Ref2") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
                Case Else
                    SQL &= " And T0." & putQuotes("Ref3") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
            End Select

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(New KeyValuePair(Of Integer, Integer)(DR.Item("TransId"), DR.Item("Number")))
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Sub setDatosDWAsiento(ByVal Tabla As String,
                                  ByVal DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)),
                                  ByVal IDDW As String,
                                  ByVal URLDW As String,
                                  ByVal TRATADODW As String,
                                  ByVal Sociedad As eSociedad)

        'Actualiza los campos de DW en la Asiento

        Try

            'Actualizamos por CardCode, DocDate y DocEntry
            Dim SQL As String = ""
            SQL = "  UPDATE " & getDataBaseRef(Tabla, Sociedad) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEIIDDW") & " =N'" & IDDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIURLDW") & " =N'" & URLDW & "' " & vbCrLf

            If Not String.IsNullOrEmpty(TRATADODW) Then SQL &= " ," & putQuotes("U_SEIMOT") & " =N'" & TRATADODW & "' " & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            SQL &= " And COALESCE(" & putQuotes("TransId") & ",-1)  IN ('-1' " & vbCrLf
            For Each DocEntryNum In DocEntryNums
                SQL &= ", N'" & DocEntryNum.Key & "'" & vbCrLf
            Next
            SQL &= ") " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("No se ha podido actualizar el asiento: " & putQuotes(Tabla) & " para DocNum: " & String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList))

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

    End Sub

#End Region

#Region "Precio entrega"

    Private Function ActualizarPrecioEntrega(ByVal objDocuware As EntDocuware, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "DOCUWARE PRECIO ENTREGA"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de precio entrega: IDDW " & objDocuware.DOCIDDW)

            'Obligatorios
            If String.IsNullOrEmpty(objDocuware.RefOrigen) OrElse String.IsNullOrEmpty(objDocuware.TipoRefOrigen) Then _
                Throw New Exception("Referencia y/o tipo de referencia no suministrados")

            If String.IsNullOrEmpty(objDocuware.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocuware.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")

            'Buscamos IC por NIF o razón social
            Dim CardCode As String = ""
            If Not String.IsNullOrEmpty(objDocuware.NIFTercero) AndAlso Not String.IsNullOrEmpty(objDocuware.RazonSocial) Then
                CardCode = oComun.getCardCode(objDocuware.NIFTercero, objDocuware.RazonSocial, objDocuware.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocuware.NIFTercero & ", Razón social: " & objDocuware.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocuware.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocuware.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objDocuware.RefOrigen.Split("#")

            'Recorre cada uno de los precio entrega origen
            For Each RefOrigen In RefsOrigen

                'Comprueba que existe el precio entrega origen y que se puede acceder a él
                'Pueden ser varios precio entregas origen con la misma referencia
                Dim DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)) = getDocEntryNumDeRefOrigenPrecioEntrega(Tabla, RefOrigen, objDocuware.TipoRefOrigen, CardCode, Sociedad)
                If DocEntryNums Is Nothing OrElse DocEntryNums.Count = 0 Then Throw New Exception("No existe precio entrega con referencia: " & RefOrigen & " o su estado es cancelado")

                'Se actualizan los campos de usuario de DW
                setDatosDWPrecioEntrega(Tabla, CardCode, DocEntryNums, objDocuware.DOCIDDW, objDocuware.DOCURLDW, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Precio entrega actualizado con éxito"
                retVal.MENSAJEAUX = String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Private Function getDocEntryNumDeRefOrigenPrecioEntrega(ByVal Tabla As String,
                                                            ByVal RefOrigen As String,
                                                            ByVal TipoRefOrigen As String,
                                                            ByVal CardCode As String,
                                                            ByVal Sociedad As eSociedad) As List(Of KeyValuePair(Of Integer, Integer))

        'Devuelve el DocEntry del precio entrega

        Dim retVal As New List(Of KeyValuePair(Of Integer, Integer))

        Try

            'Buscamos por DocEntryNum, DocEntry o NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= "  T0." & putQuotes("DocEntry") & " As DocEntry " & vbCrLf
            SQL &= " ,T0." & putQuotes("DocNum") & " As DocNum " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            Select Case TipoRefOrigen
                Case Utilidades.RefOrigen.DocEntry
                    SQL &= " And T0." & putQuotes("DocEntry") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Utilidades.RefOrigen.DocNum
                    SQL &= " And T0." & putQuotes("DocNum") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Else
                    SQL &= " And T0." & putQuotes("DocNum") & "  = N'-1'" & vbCrLf
            End Select

            If Not String.IsNullOrEmpty(CardCode) Then SQL &= " And T0." & putQuotes("CardCode") & " = N'" & CardCode & "'" & vbCrLf

            SQL &= " And T0." & putQuotes("Canceled") & " = N'" & SN.No & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(New KeyValuePair(Of Integer, Integer)(DR.Item("DocEntry"), DR.Item("DocNum")))
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Sub setDatosDWPrecioEntrega(ByVal Tabla As String,
                                        ByVal CardCode As String,
                                        ByVal DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)),
                                        ByVal IDDW As String,
                                        ByVal URLDW As String,
                                        ByVal Sociedad As eSociedad)

        'Actualiza los campos de DW en el precio entrega

        Try

            'Actualizamos por CardCode, DocDate y DocEntryNum
            Dim SQL As String = ""
            SQL = "  UPDATE " & getDataBaseRef(Tabla, Sociedad) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEIIDDW") & " =N'" & IDDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIURLDW") & " =N'" & URLDW & "' " & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) Then SQL &= " And COALESCE(" & putQuotes("CardCode") & ",'') = N'" & CardCode & "'" & vbCrLf

            SQL &= " And COALESCE(" & putQuotes("DocEntry") & ",-1)  IN ('-1' " & vbCrLf
            For Each DocEntryNum In DocEntryNums
                SQL &= ", N'" & DocEntryNum.Key & "'" & vbCrLf
            Next
            SQL &= ") " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("No se ha podido actualizar el precio entrega: " & putQuotes(Tabla) & " para CardCode: " & CardCode & " y DocNum: " & String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList))

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

    End Sub

#End Region

#Region "Inventario"

    Private Function ActualizarInventario(ByVal objDocuware As EntDocuware, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "DOCUWARE INVENTARIO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de inventario: IDDW " & objDocuware.DOCIDDW)

            'Obligatorios
            If String.IsNullOrEmpty(objDocuware.RefOrigen) OrElse String.IsNullOrEmpty(objDocuware.TipoRefOrigen) Then _
                Throw New Exception("Referencia y/o tipo de referencia no suministrados")

            If String.IsNullOrEmpty(objDocuware.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocuware.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")

            'Buscamos IC por NIF o razón social
            Dim CardCode As String = ""
            If Not String.IsNullOrEmpty(objDocuware.NIFTercero) AndAlso Not String.IsNullOrEmpty(objDocuware.RazonSocial) Then
                CardCode = oComun.getCardCode(objDocuware.NIFTercero, objDocuware.RazonSocial, objDocuware.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocuware.NIFTercero & ", Razón social: " & objDocuware.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocuware.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocuware.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objDocuware.RefOrigen.Split("#")

            'Recorre cada uno de los Inventario origen
            For Each RefOrigen In RefsOrigen

                'Comprueba que existe el Inventario origen y que se puede acceder a él
                'Pueden ser varios Inventarios origen con la misma referencia
                Dim DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)) = getDocEntryNumDeRefOrigenInventario(Tabla, RefOrigen, objDocuware.TipoRefOrigen, CardCode, Sociedad)
                If DocEntryNums Is Nothing OrElse DocEntryNums.Count = 0 Then Throw New Exception("No existe inventario con referencia: " & RefOrigen & " o su estado es cancelado")

                'Se actualizan los campos de usuario de DW
                setDatosDWInventario(Tabla, CardCode, DocEntryNums, objDocuware.DOCIDDW, objDocuware.DOCURLDW, objDocuware.DOCTRATADODW, objDocuware.DOCESTADODW, objDocuware.DOCMOTIVODW, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Inventario actualizado con éxito"
                retVal.MENSAJEAUX = String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Private Function getDocEntryNumDeRefOrigenInventario(ByVal Tabla As String,
                                                         ByVal RefOrigen As String,
                                                         ByVal TipoRefOrigen As String,
                                                         ByVal CardCode As String,
                                                         ByVal Sociedad As eSociedad) As List(Of KeyValuePair(Of Integer, Integer))

        'Devuelve el DocEntry del Inventario

        Dim retVal As New List(Of KeyValuePair(Of Integer, Integer))

        Try

            'Buscamos por DocEntryNum, DocEntry o NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= "  T0." & putQuotes("DocEntry") & " As DocEntry " & vbCrLf
            SQL &= " ,T0." & putQuotes("DocNum") & " As DocNum " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            Select Case TipoRefOrigen
                Case Utilidades.RefOrigen.DocEntry
                    SQL &= " And T0." & putQuotes("DocEntry") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Utilidades.RefOrigen.DocNum
                    SQL &= " And T0." & putQuotes("DocNum") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Else
                    SQL &= " And T0." & putQuotes("NumAtCard") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
            End Select

            If Not String.IsNullOrEmpty(CardCode) Then SQL &= " And T0." & putQuotes("CardCode") & " = N'" & CardCode & "'" & vbCrLf

            SQL &= " And T0." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(New KeyValuePair(Of Integer, Integer)(DR.Item("DocEntry"), DR.Item("DocNum")))
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Sub setDatosDWInventario(ByVal Tabla As String,
                                     ByVal CardCode As String,
                                     ByVal DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)),
                                     ByVal IDDW As String,
                                     ByVal URLDW As String,
                                     ByVal TRATADODW As String,
                                     ByVal ESTADODW As Integer,
                                     ByVal MOTIVODW As String,
                                     ByVal Sociedad As eSociedad)

        'Actualiza los campos de DW en el Inventario

        Try

            'Actualizamos por CardCode, DocDate y DocEntryNum
            Dim SQL As String = ""
            SQL = "  UPDATE " & getDataBaseRef(Tabla, Sociedad) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEIIDDW") & " =N'" & IDDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIURLDW") & " =N'" & URLDW & "' " & vbCrLf

            If Not String.IsNullOrEmpty(TRATADODW) Then SQL &= " ," & putQuotes("U_SEIMOT") & " =N'" & TRATADODW & "' " & vbCrLf
            If ESTADODW > -1 Then SQL &= " ," & putQuotes("U_SEIESTDW") & " =N'" & ESTADODW & "' " & vbCrLf
            If Not String.IsNullOrEmpty(MOTIVODW) Then SQL &= " ," & putQuotes("U_SEIMOTDW") & " =N'" & ESTADODW & "' " & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) Then SQL &= " And COALESCE(" & putQuotes("CardCode") & ",'') = N'" & CardCode & "'" & vbCrLf

            SQL &= " And COALESCE(" & putQuotes("DocEntry") & ",-1)  IN ('-1' " & vbCrLf
            For Each DocEntryNum In DocEntryNums
                SQL &= ", N'" & DocEntryNum.Key & "'" & vbCrLf
            Next
            SQL &= ") " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("No se ha podido actualizar el inventario: " & putQuotes(Tabla) & " para CardCode: " & CardCode & " y DocNum: " & String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList))

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

    End Sub

#End Region


#Region "Cobro/pago"

    Private Function ActualizarCobroPago(ByVal objDocuware As EntDocuware, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "DOCUWARE COBRO/PAGO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de cobro/pago: IDDW " & objDocuware.DOCIDDW)

            'Obligatorios
            If String.IsNullOrEmpty(objDocuware.RefOrigen) OrElse String.IsNullOrEmpty(objDocuware.TipoRefOrigen) Then _
                Throw New Exception("Referencia y/o tipo de referencia no suministrados")

            If String.IsNullOrEmpty(objDocuware.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocuware.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")

            'Buscamos IC por NIF o razón social
            Dim CardCode As String = ""
            If Not String.IsNullOrEmpty(objDocuware.NIFTercero) AndAlso Not String.IsNullOrEmpty(objDocuware.RazonSocial) Then
                CardCode = oComun.getCardCode(objDocuware.NIFTercero, objDocuware.RazonSocial, objDocuware.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocuware.NIFTercero & ", Razón social: " & objDocuware.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocuware.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocuware.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objDocuware.RefOrigen.Split("#")

            'Recorre cada uno de los CobroPago origen
            For Each RefOrigen In RefsOrigen

                'Comprueba que existe el cobro/pago origen y que se puede acceder a él
                'Pueden ser varios cobro/pago origen con la misma referencia
                Dim DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)) = getDocEntryNumDeRefOrigenCobroPago(Tabla, RefOrigen, objDocuware.TipoRefOrigen, CardCode, Sociedad)
                If DocEntryNums Is Nothing OrElse DocEntryNums.Count = 0 Then Throw New Exception("No existe cobro/pago con referencia: " & RefOrigen & " o su estado es cancelado")

                'Se actualizan los campos de usuario de DW
                setDatosDWCobroPago(Tabla, CardCode, DocEntryNums, objDocuware.DOCIDDW, objDocuware.DOCURLDW, objDocuware.DOCTRATADODW, objDocuware.DOCESTADODW, objDocuware.DOCMOTIVODW, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Cobro/pago actualizado con éxito"
                retVal.MENSAJEAUX = String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Private Function getDocEntryNumDeRefOrigenCobroPago(ByVal Tabla As String,
                                                        ByVal RefOrigen As String,
                                                        ByVal TipoRefOrigen As String,
                                                        ByVal CardCode As String,
                                                        ByVal Sociedad As eSociedad) As List(Of KeyValuePair(Of Integer, Integer))

        'Devuelve el DocEntry del cobro/pago

        Dim retVal As New List(Of KeyValuePair(Of Integer, Integer))

        Try

            'Buscamos por DocEntryNum, DocEntry o NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= "  T0." & putQuotes("DocEntry") & " As DocEntry " & vbCrLf
            SQL &= " ,T0." & putQuotes("DocNum") & " As DocNum " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            Select Case TipoRefOrigen
                Case Utilidades.RefOrigen.DocEntry
                    SQL &= " And T0." & putQuotes("DocEntry") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Utilidades.RefOrigen.DocNum
                    SQL &= " And T0." & putQuotes("DocNum") & "  = N'" & RefOrigen & "'" & vbCrLf
                Case Else
                    SQL &= " And T0." & putQuotes("CounterRef") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
            End Select

            If Not String.IsNullOrEmpty(CardCode) Then SQL &= " And T0." & putQuotes("CardCode") & " = N'" & CardCode & "'" & vbCrLf

            SQL &= " And T0." & putQuotes("Canceled") & " = N'" & SN.No & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(New KeyValuePair(Of Integer, Integer)(DR.Item("DocEntry"), DR.Item("DocNum")))
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Sub setDatosDWCobroPago(ByVal Tabla As String,
                                    ByVal CardCode As String,
                                    ByVal DocEntryNums As List(Of KeyValuePair(Of Integer, Integer)),
                                    ByVal IDDW As String,
                                    ByVal URLDW As String,
                                    ByVal TRATADODW As String,
                                    ByVal ESTADODW As Integer,
                                    ByVal MOTIVODW As String,
                                    ByVal Sociedad As eSociedad)

        'Actualiza los campos de DW en el CobroPago

        Try

            'Actualizamos por CardCode, DocDate y DocEntryNum
            Dim SQL As String = ""
            SQL = "  UPDATE " & getDataBaseRef(Tabla, Sociedad) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEIIDDW") & " =N'" & IDDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIURLDW") & " =N'" & URLDW & "' " & vbCrLf

            If Not String.IsNullOrEmpty(TRATADODW) Then SQL &= " ," & putQuotes("U_SEIMOT") & " =N'" & TRATADODW & "' " & vbCrLf
            If ESTADODW > -1 Then SQL &= " ," & putQuotes("U_SEIESTDW") & " =N'" & ESTADODW & "' " & vbCrLf
            If Not String.IsNullOrEmpty(MOTIVODW) Then SQL &= " ," & putQuotes("U_SEIMOTDW") & " =N'" & ESTADODW & "' " & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) Then SQL &= " And COALESCE(" & putQuotes("CardCode") & ",'') = N'" & CardCode & "'" & vbCrLf

            SQL &= " And COALESCE(" & putQuotes("DocEntry") & ",-1)  IN ('-1' " & vbCrLf
            For Each DocEntryNum In DocEntryNums
                SQL &= ", N'" & DocEntryNum.Key & "'" & vbCrLf
            Next
            SQL &= ") " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("No se ha podido actualizar el cobro/pago: " & putQuotes(Tabla) & " para CardCode: " & CardCode & " y DocNum: " & String.Join("#", DocEntryNums.Select(Function(p) p.Value).ToList))

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

    End Sub

#End Region

#End Region

End Class
