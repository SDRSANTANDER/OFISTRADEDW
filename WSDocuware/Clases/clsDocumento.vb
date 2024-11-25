Imports SAPbobsCOM
Imports System.Reflection
Imports System.Globalization
Imports WSDocuware.Utilidades

Public Class clsDocumento

#Region "Públicas"

    Public Function CrearDocumento(ByVal objDocumento As EntDocumentoCab, ByVal Sociedad As eSociedad) As EntResultado

        'Crea un documento como en borrador o en firme en base a otros o directamente 
        Dim retVal As New EntResultado

        Try

            'Buscamos si es solicitud por ObjectType
            Dim bEsSolicitud As Boolean = getEsDocumentoSolicitud(objDocumento.ObjTypeDestino)

            If Not bEsSolicitud Then

                If Not String.IsNullOrEmpty(objDocumento.RefOrigen) AndAlso objDocumento.ObjTypeOrigen > 0 Then

                    If objDocumento.ObjTypeDestino = ObjType.Cobro OrElse objDocumento.ObjTypeDestino = ObjType.Pago Then

                        'Crea cobro/pago
                        retVal = CopiarADocumentoCobroPago(objDocumento, Sociedad)

                    ElseIf objDocumento.ObjTypeDestino = ObjType.PrecioEntrega Then

                        'Crea cobro/pago
                        retVal = CopiarADocumentoPrecioEntrega(objDocumento, Sociedad)

                    ElseIf objDocumento.ObjTypeDestino = ObjType.ReciboProduccion Then

                        'Crear el documento de recibo de producción
                        retVal = NuevoDocumentoReciboProduccion(objDocumento, Sociedad)

                    Else

                        'Copia el documento
                        retVal = CopiarDeDocumento(objDocumento, Sociedad)

                    End If

                Else

                    If objDocumento.ObjTypeDestino = ObjType.Cobro OrElse objDocumento.ObjTypeDestino = ObjType.Pago Then

                        'Crea cobro/pago
                        If objDocumento.DocType = DocType.Cuenta _
                            OrElse (String.IsNullOrEmpty(objDocumento.RazonSocial) AndAlso String.IsNullOrEmpty(objDocumento.NIFTercero)) Then
                            retVal = NuevoDocumentoCobroPagoACuenta(objDocumento, Sociedad)
                        Else
                            retVal = NuevoDocumentoCobroPago(objDocumento, Sociedad)
                        End If

                    Else

                        'Crea el documento
                        retVal = NuevoDocumento(objDocumento, Sociedad)

                    End If

                End If

            Else

                'Crea el documento solicitud
                retVal = NuevoDocumentoSolicitud(objDocumento, Sociedad)

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Public Function ActualizarDocumentoDW(ByVal objDocumento As EntDocumentoCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "ACTUALIZAR DOCUMENTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de documento: " & objDocumento.NumAtCard & " para IDDW " & objDocumento.DOCIDDW)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumento.NIFTercero) Then Throw New Exception("NIF no suministrado")

            If Not DateTime.TryParseExact(objDocumento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha documento no suministrada")

            If String.IsNullOrEmpty(objDocumento.RefOrigen) OrElse String.IsNullOrEmpty(objDocumento.TipoRefOrigen) Then _
                Throw New Exception("Referencia y/o tipo de referencia no suministrados")

            If String.IsNullOrEmpty(objDocumento.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocumento.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objDocumento.NIFTercero, objDocumento.RazonSocial, objDocumento.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumento.NIFTercero & ", Razón social: " & objDocumento.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumento.ObjTypeDestino)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumento.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objDocumento.RefOrigen.Split("#")

            'Recorre cada uno de los documento origen
            For Each RefOrigen In RefsOrigen

                'Comprueba que existe el documento origen y que se puede acceder a él
                'Pueden ser varios documentos origen con la misma referencia
                Dim DocNums As List(Of String) = getDocNumDeRefOrigen(Tabla, RefOrigen, objDocumento.TipoRefOrigen, CardCode, Sociedad)
                If DocNums Is Nothing OrElse DocNums.Count = 0 Then Throw New Exception("No existe documento con referencia: " & RefOrigen & " o su estado es cancelado")

                'Se actualizan los campos de usuario de DW
                setDatosDWDocumento(Tabla, CardCode, objDocumento.DocDate, DocNums, objDocumento.DOCIDDW, objDocumento.DOCURLDW, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento actualizado con éxito"
                retVal.MENSAJEAUX = String.Join("#", DocNums)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Public Function ActualizarDocumento(ByVal objDocumento As EntDocumentoCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oDocDestino As Documents = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "ACTUALIZAR DOCUMENTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de documento para " & objDocumento.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objDocumento.UserSAP, objDocumento.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objDocumento.RefDestino) Then Throw New Exception("Referencia destino no suministrada")

            If objDocumento.Lineas Is Nothing Then Throw New Exception("Líneas no suministradas")

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objDocumento.ObjTypeDestino)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objDocumento.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objDocumento.NIFTercero, objDocumento.RazonSocial, objDocumento.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumento.NIFTercero & ", Razón social: " & objDocumento.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'RefsDestino contiene los NumAtCard
            Dim RefsDestino As String() = objDocumento.RefDestino.Split("#")

            'Comprueba que no lleguen líneas sin artículo/descripción 
            Dim bSinIdentificar As Boolean = getLinSinIdentificar(objDocumento)
            If bSinIdentificar Then Throw New Exception("Rellene el artículo/concepto a traspasar en todas las líneas")

            'Comprueba que el código de artículo sea ItemCode y no la referencia de proveedor
            For Each objLinea In objDocumento.Lineas

                'Índice
                Dim indLinea As Integer = objDocumento.Lineas.IndexOf(objLinea) + 1

                If Not String.IsNullOrEmpty(objLinea.Articulo) OrElse Not String.IsNullOrEmpty(objLinea.RefExt) Then
                    'Artículo (si no se encuentra el código se busca por la referencia de proveedor)
                    Dim ItemCode As String = oComun.getItemCode(CardCode, objLinea.Articulo, objLinea.RefExt, Sociedad)
                    If Not String.IsNullOrEmpty(objLinea.Articulo) AndAlso objLinea.Articulo = ItemCode Then Continue For

                    If String.IsNullOrEmpty(ItemCode) Then Throw New Exception("No encuentro artículo de la línea " & indLinea.ToString & " con código: " & objLinea.Articulo & " o referencia externa: " & objLinea.RefExt & " para IC: " & CardCode)
                    clsLog.Log.Info("(" & sLogInfo & ") Encontrado artículo: " & ItemCode)

                    objLinea.Articulo = ItemCode
                End If

            Next

            'Almacena los DocEntry
            Dim DocNums As New List(Of String)

            'Recorre cada uno de los documento origen
            For Each RefDestino In RefsDestino

                'Comprueba que existe el documento destino y que se puede acceder a él
                'Pueden ser varios documentos destino con la misma referencia
                Dim DocEntrysDestino As List(Of String) = getDocEntryDeRefOrigen(TablaDestino, RefDestino, objDocumento.TipoRefDestino, CardCode, True, True, Sociedad)
                If DocEntrysDestino Is Nothing OrElse DocEntrysDestino.Count = 0 Then Throw New Exception("No existe documento destino con referencia: " & RefDestino & " o su estado es cerrado/cancelado")

                For Each DocEntryDestino In DocEntrysDestino

                    'Objeto destino
                    oDocDestino = oCompany.GetBusinessObject(objDocumento.ObjTypeDestino)

                    If Not oDocDestino.GetByKey(DocEntryDestino) Then Throw New Exception("No puedo recuperar el documento destino con referencia: " & RefDestino & " y DocEntry:" & DocEntryDestino)

                    'DocNum
                    DocNums.Add(oDocDestino.DocNum)

                    'Contacto
                    If objDocumento.ContactoCodigo > 0 Then oDocDestino.ContactPersonCode = objDocumento.ContactoCodigo

                    'Direccion envío
                    If Not String.IsNullOrEmpty(objDocumento.DireccionEnvioCodigo) Then oDocDestino.ShipToCode = objDocumento.DireccionEnvioCodigo
                    If Not String.IsNullOrEmpty(objDocumento.DireccionEnvioDetalle) Then oDocDestino.Address2 = objDocumento.DireccionEnvioDetalle

                    'Direccion facturación
                    If Not String.IsNullOrEmpty(objDocumento.DireccionFacturaCodigo) Then oDocDestino.PayToCode = objDocumento.DireccionFacturaCodigo
                    If Not String.IsNullOrEmpty(objDocumento.DireccionFacturaDetalle) Then oDocDestino.Address = objDocumento.DireccionFacturaDetalle

                    'FinanzasRazon
                    If Not String.IsNullOrEmpty(objDocumento.FinanzasRazon) Then oDocDestino.CardName = objDocumento.FinanzasRazon

                    'FinanzasNIF
                    If Not String.IsNullOrEmpty(objDocumento.FinanzasNIF) Then oDocDestino.FederalTaxID = objDocumento.FinanzasNIF

                    'Referencia
                    If Not String.IsNullOrEmpty(objDocumento.NumAtCard) Then oDocDestino.NumAtCard = objDocumento.NumAtCard

                    'Moneda
                    If Not String.IsNullOrEmpty(objDocumento.Currency) Then oDocDestino.DocCurrency = objDocumento.Currency

                    'Sucursal
                    If objDocumento.Sucursal > 0 Then oDocDestino.BPL_IDAssignedToInvoice = objDocumento.Sucursal

                    'Proyecto
                    If Not String.IsNullOrEmpty(objDocumento.Proyecto) Then oDocDestino.Project = objDocumento.Proyecto

                    'Clase expedicion
                    If objDocumento.ClaseExpedicion > 0 Then oDocDestino.TransportationCode = objDocumento.ClaseExpedicion

                    'Responsable
                    If Not String.IsNullOrEmpty(objDocumento.Responsable) Then oDocDestino.AgentCode = objDocumento.Responsable

                    'Titular
                    If objDocumento.Titular > 0 Then oDocDestino.DocumentsOwner = objDocumento.Titular

                    'Empleado dpto compras/ventas
                    If objDocumento.EmpDptoCompraVenta > 0 Then oDocDestino.SalesPersonCode = objDocumento.EmpDptoCompraVenta

                    'Comentarios
                    If objDocumento.Comments.Length > 254 Then objDocumento.Comments = objDocumento.Comments.Substring(0, 254)
                    If Not String.IsNullOrEmpty(objDocumento.Comments) Then oDocDestino.Comments = objDocumento.Comments

                    'Entrada diario
                    If objDocumento.JournalMemo.Length > 254 Then objDocumento.JournalMemo = objDocumento.JournalMemo.Substring(0, 254)
                    If Not String.IsNullOrEmpty(objDocumento.JournalMemo) Then oDocDestino.JournalMemo = objDocumento.JournalMemo

                    'Total de documento 
                    If objDocumento.DocTotal <> 0 Then oDocDestino.DocTotal = objDocumento.DocTotal

                    'Campos de usuario
                    setCamposUsuarioCabecera(objDocumento, oDocDestino)

                    'Actualiza las líneas
                    For iLinea As Integer = 0 To oDocDestino.Lines.Count - 1

                        oDocDestino.Lines.SetCurrentLine(iLinea)

                        '20200911: Comprueba que la línea no esté cerrada
                        If oDocDestino.Lines.LineStatus <> BoStatus.bost_Close Then

                            'Traspaso parcial, búsqueda de línea
                            Dim oLinea As EntDocumentoLin = getLinCopiaParcial(objDocumento, oDocDestino)

                            If Not oLinea Is Nothing Then

                                oDocDestino.Lines.SetCurrentLine(oDocDestino.Lines.Count - 1)

                                'Descripción
                                If oLinea.Concepto.Length > 100 Then oLinea.Concepto = oLinea.Concepto.Substring(0, 100)
                                If Not String.IsNullOrEmpty(oLinea.Concepto) Then oDocDestino.Lines.ItemDescription = oLinea.Concepto

                                'Cantidad
                                If oLinea.Cantidad > 0 Then oDocDestino.Lines.Quantity = oLinea.Cantidad

                                'Precio
                                If oLinea.PrecioUnidad <> 0 Then
                                    oDocDestino.Lines.UnitPrice = oLinea.PrecioUnidad
                                    oDocDestino.Lines.Price = oLinea.PrecioUnidad
                                End If

                                'Porcentaje dto
                                If oLinea.PorcentajeDescuento <> 0 Then
                                    oDocDestino.Lines.DiscountPercent = oLinea.PorcentajeDescuento
                                End If

                                'Moneda
                                If Not String.IsNullOrEmpty(objDocumento.Currency) Then oDocDestino.Lines.Currency = objDocumento.Currency

                                'TaxOnly
                                oDocDestino.Lines.TaxOnly = BoYesNoEnum.tNO
                                If Not String.IsNullOrEmpty(oLinea.TaxOnly) AndAlso oLinea.TaxOnly = SN.Si Then oDocDestino.Lines.TaxOnly = BoYesNoEnum.tYES

                                'TaxCode
                                If Not String.IsNullOrEmpty(oLinea.TaxCode) Then oDocDestino.Lines.TaxCode = oLinea.TaxCode

                                'WTLiable
                                oDocDestino.Lines.WTLiable = BoYesNoEnum.tNO
                                If Not String.IsNullOrEmpty(oLinea.WTLiable) AndAlso oLinea.WTLiable = SN.Si Then oDocDestino.Lines.WTLiable = BoYesNoEnum.tYES

                                'VATGroup
                                If Not String.IsNullOrEmpty(oLinea.VATGroup) Then oDocDestino.Lines.VatGroup = oLinea.VATGroup

                                'Fecha entrega
                                If DateTime.TryParseExact(oLinea.ShipDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                    oDocDestino.Lines.ShipDate = Date.ParseExact(oLinea.ShipDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                                'Fecha requerida
                                If DateTime.TryParseExact(oLinea.RequiredDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                    oDocDestino.Lines.RequiredDate = Date.ParseExact(oLinea.RequiredDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                                'Proyecto
                                If Not String.IsNullOrEmpty(oLinea.Proyecto) Then oDocDestino.Lines.ProjectCode = oLinea.Proyecto

                                'Almacen 
                                If Not String.IsNullOrEmpty(oLinea.Almacen) Then oDocDestino.Lines.WarehouseCode = oLinea.Almacen

                                'Centro coste 1
                                If Not String.IsNullOrEmpty(oLinea.Centro1Coste) Then oDocDestino.Lines.CostingCode = oLinea.Centro1Coste

                                'Centro coste 2
                                If Not String.IsNullOrEmpty(oLinea.Centro2Coste) Then oDocDestino.Lines.CostingCode2 = oLinea.Centro2Coste

                                'Centro coste 3
                                If Not String.IsNullOrEmpty(oLinea.Centro3Coste) Then oDocDestino.Lines.CostingCode3 = oLinea.Centro3Coste

                                'Centro coste 4
                                If Not String.IsNullOrEmpty(oLinea.Centro4Coste) Then oDocDestino.Lines.CostingCode4 = oLinea.Centro4Coste

                                'Centro coste 5
                                If Not String.IsNullOrEmpty(oLinea.Centro5Coste) Then oDocDestino.Lines.CostingCode5 = oLinea.Centro5Coste

                                'Campos de usuario
                                setCamposUsuarioLinea(oLinea, oDocDestino)

                            End If

                        End If

                    Next

                Next

            Next

            'Actualiamos documento
            If oDocDestino.Update() = 0 Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento actualizado con éxito"
                retVal.MENSAJEAUX = String.Join("#", DocNums)

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
            oComun.LiberarObjCOM(oDocDestino)
        End Try

        Return retVal

    End Function

    Public Function ModificarDocumento(ByVal objDocumento As EntDocumentoCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oDocEntrada As Documents = Nothing
        Dim oDocDestino As Documents = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "MODIFICAR DOCUMENTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de modificación de documento para " & objDocumento.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objDocumento.UserSAP, objDocumento.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objDocumento.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocumento.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")
            If String.IsNullOrEmpty(objDocumento.RefOrigen) Then Throw New Exception("Referencia origen no suministrada")

            If Not DateTime.TryParseExact(objDocumento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha documento no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objDocumento.DocDateAux.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha documento auxiliar no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha contable no suministrada o incorrecta")

            If objDocumento.Lineas Is Nothing Then Throw New Exception("Líneas no suministradas")

            'Buscamos tabla origen por ObjectType
            Dim TablaOrigen As String = getTablaDeObjType(objDocumento.ObjTypeOrigen)
            If String.IsNullOrEmpty(TablaOrigen) Then Throw New Exception("No encuentro tabla origen con ObjectType: " & objDocumento.ObjTypeOrigen)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla origen: " & TablaOrigen)

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objDocumento.NIFTercero, objDocumento.RazonSocial, objDocumento.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumento.NIFTercero & ", Razón social: " & objDocumento.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objDocumento.RefOrigen.Split("#")

            'Origen con búsqueda indirecta (por ejemplo, albaranes de pedidos abiertos)
            getRefsOrigenDocumentoNoDirecto(objDocumento, CardCode, TablaOrigen, RefsOrigen, sLogInfo, Sociedad)

            'Recorre cada uno de los documento origen
            For Each RefOrigen In RefsOrigen

                'Comprueba que existe el documento origen y que se puede acceder a él
                'Pueden ser varios documentos origen con la misma referencia
                Dim DocEntrysOrigen As List(Of String) = getDocEntryDeRefOrigen(TablaOrigen, RefOrigen, objDocumento.TipoRefOrigen, CardCode, True, True, Sociedad)
                If DocEntrysOrigen Is Nothing OrElse DocEntrysOrigen.Count = 0 Then Throw New Exception("No existe documento origen con referencia: " & RefOrigen & " o su estado es cerrado/cancelado")

                'Líneas de RefOrigen
                Dim objRefLineas As List(Of EntDocumentoLin) = (From p In objDocumento.Lineas
                                                                Order By p.LineNum, p.VisOrder Ascending
                                                                Where p.RefOrigen = RefOrigen).Distinct.ToList()

                'Modificar cantidades
                For Each DocEntryOrigen In DocEntrysOrigen
                    retVal = ModificarDocumentoLineas(oCompany, objDocumento, objRefLineas, TablaOrigen, DocEntryOrigen, sLogInfo, Sociedad)
                Next

            Next

        'Copiar documento
        retVal = CopiaParcialDeDocumento(oCompany, objDocumento, TablaOrigen, CardCode, sLogInfo, Sociedad)

        Catch ex As Exception
        clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        retVal.CODIGO = Respuesta.Ko
        retVal.MENSAJE = ex.Message
        retVal.MENSAJEAUX = ""
        Finally
        oComun.LiberarObjCOM(oDocEntrada)
        oComun.LiberarObjCOM(oDocDestino)
        End Try

        Return retVal

    End Function

    Public Function getComprobarDocumentoID(ByVal IDDW As String,
                                            ByVal ObjType As Integer,
                                            ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado

        Dim sLogInfo As String = "DOCUMENTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar documento para IDDW: " & IDDW)

            'Obligatorios
            If String.IsNullOrEmpty(IDDW) Then Throw New Exception("ID docuware no suministrado")

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el DocNum del documento definitivo
            Dim DocNum As String = getDocNumDocumentoPorDWID(Tabla, IDDW, Sociedad)

            If Not String.IsNullOrEmpty(DocNum) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento en firme encontrado"
                retVal.MENSAJEAUX = DocNum

            Else

                retVal.CODIGO = Respuesta.Ko
                retVal.MENSAJE = "No se encuentra el documento en firme"
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

    Public Function getGenerarInforme(ByVal objInforme As EntInforme, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "GENERAR INFORME"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de generación de fichero para " & objInforme.NIFTercero)

            'Obligatorios
            If String.IsNullOrEmpty(objInforme.DOCIDDW) AndAlso String.IsNullOrEmpty(objInforme.DocNum) AndAlso String.IsNullOrEmpty(objInforme.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objInforme.NIFTercero, objInforme.RazonSocial, objInforme.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objInforme.NIFTercero & ", Razón social: " & objInforme.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objInforme.ObjType)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objInforme.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Buscamos DocEntry por DOCIDDW, DocNum o NumAtCard
            Dim DocEntry As String = oComun.getDocEntryDocumentoDefinitivo(TablaDestino, CardCode, objInforme.DOCIDDW, objInforme.DocNum, objInforme.NumAtCard, False, False, Sociedad)
            If String.IsNullOrEmpty(DocEntry) Then Throw New Exception("No encuentro documento con DOCIDDW - '" & objInforme.DOCIDDW & "', DocNum - '" & objInforme.DocNum & "' y NumAtCard - '" & objInforme.NumAtCard & "'")
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado DocEntry: " & DocEntry)

            'Buscamos DocNum por DocEntry
            Dim sDocNum As String = oComun.getDocNumDeDocEntry(TablaDestino, DocEntry, Sociedad)
            If String.IsNullOrEmpty(sDocNum) Then Throw New Exception("No encuentro documento con DocEntry: '" & DocEntry & "'")
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado DocNum: " & sDocNum)

            ''Ruta del fichero
            'Dim FicheroRuta As String = "C:\Users\mlanza\OneDrive - SEIDOR SA\Escritorio\IMAGENES\Adjunto1.pdf"

            ''Lee le fichero
            'Dim FicheroBase64 As String = ""
            'If IO.File.Exists(FicheroRuta) Then
            '    Dim FicheroBinario As Byte() = IO.File.ReadAllBytes(FicheroRuta)
            '    FicheroBase64 = Convert.ToBase64String(FicheroBinario)
            'Else
            '    Dim oInforme As New clsInforme
            '    FicheroBase64 = oInforme.GenerarFicheroBase64(DocEntry, sDocNum, Sociedad)
            'End If

            'Genera el fichero
            Dim oInforme As New clsInforme
            Dim FicheroRuta As String = ""
            Dim FicheroBase64 As String = oInforme.GenerarInformeBase64(objInforme.InfomeTipo, DocEntry, sDocNum, FicheroRuta, objInforme.UserSAP, objInforme.PassSAP, Sociedad)

            'Devuelve el fichero
            If Not String.IsNullOrEmpty(FicheroBase64) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = IO.Path.GetFileName(FicheroRuta)
                retVal.MENSAJEAUX = FicheroBase64

            Else

                retVal.CODIGO = Respuesta.Ko
                retVal.MENSAJE = "Fichero no generado"
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

    Public Function setBloqueoPago(ByVal objDocumentoBloqueo As EntDocumentoBloqueo, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oDocDestino As Documents = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "BLOQUEO PAGO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de bloqueo para " & objDocumentoBloqueo.NIFEmpresa)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoBloqueo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoBloqueo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoBloqueo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            oCompany = ConexionSAP.getCompany(objDocumentoBloqueo.UserSAP, objDocumentoBloqueo.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objDocumentoBloqueo.ObjType)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objDocumentoBloqueo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Buscamos DocEntry por DOCIDDW, DocNum o NumAtCard
            Dim DocEntry As String = oComun.getDocEntryDocumentoDefinitivo(TablaDestino, "", objDocumentoBloqueo.DOCIDDW, objDocumentoBloqueo.DocNum, objDocumentoBloqueo.NumAtCard, False, False, Sociedad)
            If String.IsNullOrEmpty(DocEntry) Then Throw New Exception("No encuentro documento con DOCIDDW - '" & objDocumentoBloqueo.DOCIDDW & "', DocNum - '" & objDocumentoBloqueo.DocNum & "' y NumAtCard - '" & objDocumentoBloqueo.NumAtCard & "'")
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado DocEntry: " & DocEntry)

            'Objeto documento
            oDocDestino = oCompany.GetBusinessObject(objDocumentoBloqueo.ObjType)
            If Not oDocDestino.GetByKey(DocEntry) Then Throw New Exception("No puedo recuperar el documento destino con DocEntry: " & DocEntry)

            'Asigna el bloqueo
            oDocDestino.PaymentBlockEntry = objDocumentoBloqueo.BloqueoPago
            oDocDestino.PaymentBlock = IIf(objDocumentoBloqueo.BloqueoPago = 0, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

            'Actualizamos documento
            If oDocDestino.Update() = 0 Then

                Dim sDocNum As String = ""
                sDocNum = oComun.getDocNumDeDocEntry(TablaDestino, DocEntry, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento actualizado con éxito"
                retVal.MENSAJEAUX = sDocNum

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
            oComun.LiberarObjCOM(oDocDestino)
        End Try

        Return retVal

    End Function

    Public Function setDocumentoTratado(ByVal objDocumentoTratado As EntDocumentoTratado, ByVal Sociedad As eSociedad) As EntResultado

        'Actualiza el campo tratado DW

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "TRATADO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de documento tratado para ObjType: " & objDocumentoTratado.ObjType & ", Ambito: " & objDocumentoTratado.Ambito & ", NIFTercero: " & objDocumentoTratado.NIFTercero & ", Razón social: " & objDocumentoTratado.RazonSocial & ", DOCIDDW: " & objDocumentoTratado.DOCIDDW & ", DocNum: " & objDocumentoTratado.DocNum & ", NumAtCard: " & objDocumentoTratado.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoTratado.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoTratado.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")

            If String.IsNullOrEmpty(objDocumentoTratado.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoTratado.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoTratado.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objDocumentoTratado.NIFTercero, objDocumentoTratado.RazonSocial, objDocumentoTratado.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoTratado.NIFTercero & ", Razón social: " & objDocumentoTratado.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoTratado.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoTratado.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Buscamos DocEntry por DOCIDDW, DocNum o NumAtCard
            Dim DocEntry As String = oComun.getDocEntryDocumentoDefinitivo(Tabla, CardCode, objDocumentoTratado.DOCIDDW, objDocumentoTratado.DocNum, objDocumentoTratado.NumAtCard, False, False, Sociedad)
            If String.IsNullOrEmpty(DocEntry) Then Throw New Exception("No encuentro documento con DOCIDDW - '" & objDocumentoTratado.DOCIDDW & "', DocNum - '" & objDocumentoTratado.DocNum & "' y NumAtCard - '" & objDocumentoTratado.NumAtCard & "'")
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado DocEntry: " & DocEntry)

            'Buscamos DocNum por DocEntry
            Dim sDocNum As String = oComun.getDocNumDeDocEntry(Tabla, DocEntry, Sociedad)
            If String.IsNullOrEmpty(sDocNum) Then Throw New Exception("No encuentro documento con DocEntry: '" & DocEntry & "'")
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado DocNum: " & sDocNum)

            'Actualiza el campo tratado DW 
            Dim Actualizado As Boolean = setDocumentoDWTratado(Tabla, DocEntry, Sociedad)

            If Actualizado Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento actualizado con éxito"
                retVal.MENSAJEAUX = sDocNum

            Else

                retVal.CODIGO = Respuesta.Ko
                retVal.MENSAJE = "Documento no actualizado"
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

    Public Function setDocumentoEstado(ByVal objDocumentoEstado As EntDocumentoEstado, ByVal Sociedad As eSociedad) As EntResultado

        'Actualiza el campo estado DW

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "ESTADO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de documento estado para ObjType: " & objDocumentoEstado.ObjType & ", Ambito: " & objDocumentoEstado.Ambito & ", NIFTercero: " & objDocumentoEstado.NIFTercero & ", Razón social: " & objDocumentoEstado.RazonSocial & ", DocNum: " & objDocumentoEstado.DocNum & ", NumAtCard: " & objDocumentoEstado.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoEstado.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoEstado.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")

            If String.IsNullOrEmpty(objDocumentoEstado.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoEstado.NumAtCard) Then _
                Throw New Exception("Número de documento o referencia origen no suministrados")

            If objDocumentoEstado.DOCESTADODW < 0 AndAlso objDocumentoEstado.DOCESTADODW > 5 Then _
                Throw New Exception("Estado docuware incorrecto. Valores válidos 0 a 5")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objDocumentoEstado.NIFTercero, objDocumentoEstado.RazonSocial, objDocumentoEstado.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoEstado.NIFTercero & ", Razón social: " & objDocumentoEstado.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoEstado.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoEstado.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Buscamos DocEntry por DOCIDDW, DocNum o NumAtCard
            Dim DocEntry As String = oComun.getDocEntryDocumentoDefinitivo(Tabla, CardCode, "", objDocumentoEstado.DocNum, objDocumentoEstado.NumAtCard, False, False, Sociedad)
            If String.IsNullOrEmpty(DocEntry) Then Throw New Exception("No encuentro documento con DocNum - '" & objDocumentoEstado.DocNum & "' y NumAtCard - '" & objDocumentoEstado.NumAtCard & "'")
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado DocEntry: " & DocEntry)

            'Buscamos DocNum por DocEntry
            Dim sDocNum As String = oComun.getDocNumDeDocEntry(Tabla, DocEntry, Sociedad)
            If String.IsNullOrEmpty(sDocNum) Then Throw New Exception("No encuentro documento con DocEntry: '" & DocEntry & "'")
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado DocNum: " & sDocNum)

            'Actualiza el campo Estado DW 
            Dim Actualizado As Boolean = setDocumentoDWEstado(Tabla, DocEntry, objDocumentoEstado.DOCIDDW, objDocumentoEstado.DOCURLDW, objDocumentoEstado.DOCESTADODW, objDocumentoEstado.DOCMOTIVODW, Sociedad)

            If Actualizado Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento actualizado con éxito"
                retVal.MENSAJEAUX = sDocNum

            Else

                retVal.CODIGO = Respuesta.Ko
                retVal.MENSAJE = "Documento no actualizado"
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

    Public Function getComprobarDocumentoOrigen(ByVal objDato As EntDato, ByVal Sociedad As eSociedad) As List(Of EntDocumento)

        'Comprueba si existe origen para un documento en firme

        Dim retVal As New List(Of EntDocumento)

        Dim oComun As New clsComun

        Dim sLogInfo As String = "ORIGEN"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar origen para ObjType: " & objDato.ObjType & ", Ambito: " & objDato.Ambito & ", NIFTercero: " & objDato.NIFTercero & ", Razón social: " & objDato.RazonSocial & ", DOCIDDW: " & objDato.DOCIDDW & ", DocNum: " & objDato.DocNum & ", NumAtCard: " & objDato.NumAtCard)

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objDato.NIFTercero, objDato.RazonSocial, objDato.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDato.NIFTercero & ", Razón social: " & objDato.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDato.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDato.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve los documentos origen
            Dim oDAODocumento As New DAODocumento(Sociedad)
            retVal = oDAODocumento.getDocumentosOrigen(Tabla, CardCode, objDato.DOCIDDW, objDato.DocNum, objDato.NumAtCard)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

#Region "Públicas: Obtener campo"

    Public Function getComprobarNumDocumento(ByVal objDocumentoCampo As EntDocumentoCampo,
                                             ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUMERO DOCUMENTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar documento para ObjType: " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) Then _
               Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el DocNum del documento definitivo
            Dim sDocNum As String = getDocNumDocumentoDefinitivo(Tabla, CardCode, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, Sociedad)

            If Not String.IsNullOrEmpty(sDocNum) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento en firme encontrado"
                retVal.MENSAJEAUX = sDocNum

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra el documento en firme"
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

    Public Function getComprobarImporte(ByVal objDocumentoCampo As EntDocumentoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba el importe de un pedido

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "IMPORTE"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar importe para ObjType: " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el DocTotal del documento definitivo
            Dim DocTotal As String = getDocTotalDocumentoDefinitivo(Tabla, CardCode, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, Sociedad)

            If Not String.IsNullOrEmpty(DocTotal) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Importe encontrado de este documento"
                retVal.MENSAJEAUX = DocTotal

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra el importe de este documento"
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

    Public Function getComprobarFechaVencimiento(ByVal objDocumentoCampo As EntDocumentoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "VENCIMIENTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar fecha vencimiento para ObjType: " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el DocNum del documento definitivo
            Dim DocDueDate As String = getDocDueDateDocumentoDefinitivo(Tabla, CardCode, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, Sociedad)

            If Not String.IsNullOrEmpty(DocDueDate) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Fecha de vencimiento de este documento encontrada"
                retVal.MENSAJEAUX = DocDueDate

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra la fecha de vencimiento de este documento"
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

    Public Function getComprobarResponsableMail(ByVal objDocumentoCampo As EntDocumentoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "RESPONSABLE"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar responsable para ObjType: " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el responsable del documento
            Dim ResponsableMail As String = getResponsableMailDocumentoDefinitivo(Tabla, CardCode, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, Sociedad)

            If Not String.IsNullOrEmpty(ResponsableMail) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Correo del responsable de este documento encontrado"
                retVal.MENSAJEAUX = ResponsableMail

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra el correo del responsable de este documento"
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

    Public Function getComprobarComentarios(ByVal objDocumentoCampo As EntDocumentoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "COMENTARIOS"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar comentarios para ObjType: " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el DocNum del documento definitivo
            Dim Comentarios As String = getComentariosDocumentoDefinitivo(Tabla, CardCode, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, Sociedad)

            If Not String.IsNullOrEmpty(Comentarios) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Comentarios encontrados de este documento"
                retVal.MENSAJEAUX = Comentarios

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentran los comentarios de este documento"
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

    Public Function getComprobarVencimientoFechaImporte(ByVal objDocumentoCampo As EntDocumentoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "VENCIMIENTO FECHA/IMPORTE"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar fecha e importe vencimiento para ObjType: " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el DocNum del documento definitivo
            Dim lDocDueDateTotal As List(Of String) = getDocDueDateTotalDocumentoDefinitivo(Tabla, CardCode, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, Sociedad)

            If Not lDocDueDateTotal Is Nothing AndAlso lDocDueDateTotal.Count > 0 Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Fecha e importe de vencimiento de este documento encontrada"
                retVal.MENSAJEAUX = String.Join("-", lDocDueDateTotal)

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra la fecha e importe de vencimiento de este documento"
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

    Public Function getComprobarVencimientoPagado(ByVal objDocumentoCampo As EntDocumentoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "VENCIMIENTO PAGADO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar vencimiento pagado para ObjType: " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DocDueDate: " & objDocumentoCampo.DocDueDate & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")
            If Not DateTime.TryParseExact(objDocumentoCampo.DocDueDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha vencimiento no suministrada o incorrecta")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el DocNum del documento definitivo
            Dim NotPaidToDate As String = getNotPaidToDateDocumentoDefinitivo(Tabla, CardCode, objDocumentoCampo.DocDueDate, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, Sociedad)

            If Not String.IsNullOrEmpty(NotPaidToDate) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Pago de vencimiento de este documento encontrado"
                retVal.MENSAJEAUX = NotPaidToDate

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra el pago de vencimiento de este documento"
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

    Public Function getComprobarImporteSinIVA(ByVal objDocumentoCampo As EntDocumentoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba el importe sin IVA de un pedido

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "IMPORTE SIN IVA"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar importe sin IVA para ObjType: " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el DocTotal del documento definitivo
            Dim DocTotal As String = getBaseTotalDocumentoDefinitivo(Tabla, CardCode, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, Sociedad)

            If Not String.IsNullOrEmpty(DocTotal) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Importe encontrado de este documento"
                retVal.MENSAJEAUX = DocTotal

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra el importe de este documento"
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

    Public Function getComprobarViaPago(ByVal objDocumentoCampo As EntDocumentoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe una vía de pago

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "VIA PAGO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar vía pago para ObjType: " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el DocNum del documento definitivo
            Dim ViaPago As String = getViaPagoDocumentoDefinitivo(Tabla, CardCode, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, Sociedad)

            If Not String.IsNullOrEmpty(ViaPago) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Vía pago encontrada de este documento"
                retVal.MENSAJEAUX = ViaPago

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentran la vía pago de este documento"
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

    Public Function getComprobarMoneda(ByVal objDocumentoCampo As EntDocumentoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe una vía de pago

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "MONEDA"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar moneda para ObjType: " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el DocNum del documento definitivo
            Dim Moneda As String = getMonedaDocumentoDefinitivo(Tabla, CardCode, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, Sociedad)

            If Not String.IsNullOrEmpty(Moneda) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Moneda encontrada de este documento"
                retVal.MENSAJEAUX = Moneda

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra la moneda de este documento"
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

    Public Function getComprobarNumTransaccion(ByVal objDocumentoCampo As EntDocumentoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUMERO TRANSACCION"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar número transacción para ObjType: " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el DocNum del documento definitivo
            Dim NumTransaccion As String = getNumTransaccionDocumentoDefinitivo(Tabla, CardCode, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, Sociedad)

            If Not String.IsNullOrEmpty(NumTransaccion) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Número de transacción encontrado de este documento"
                retVal.MENSAJEAUX = NumTransaccion

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra el número de transacción de este documento"
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

    Public Function getComprobarProyecto(ByVal objDocumentoCampo As EntDocumentoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "PROYECTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar proyecto para ObjType:  " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el DocNum del documento definitivo
            Dim sProyecto As String = getProyectoDocumentoDefinitivo(Tabla, CardCode, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, Sociedad)

            If Not String.IsNullOrEmpty(sProyecto) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Proyecto encontrado"
                retVal.MENSAJEAUX = sProyecto

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra el proyecto"
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

    Public Function getComprobarCampoUsuario(ByVal objDocumentoCampo As EntDocumentoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "CAMPO USUARIO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar campo usuario para ObjType:  " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = Utilidades.getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el DocNum del documento definitivo
            Dim CampoUsuario As String = getCampoUsuarioDocumentoDefinitivo(Tabla, CardCode, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, objDocumentoCampo.CampoUsuario, objDocumentoCampo.Nivel, Sociedad)

            If Not String.IsNullOrEmpty(CampoUsuario) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Campo de usuario encontrado"
                retVal.MENSAJEAUX = CampoUsuario

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra el campo de usuario"
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

    Public Function getComprobarSucursal(ByVal objDocumentoCampo As EntDocumentoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "SUCURSAL"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar sucursal para ObjType: " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve la sucursal del documento definitivo
            Dim Sucursal As String = getSucursalDocumentoDefinitivo(Tabla, CardCode, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, Sociedad)

            If Not String.IsNullOrEmpty(Sucursal) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Sucursal encontrada"
                retVal.MENSAJEAUX = Sucursal

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra la sucursal"
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

    Public Function getComprobarCentroCoste(ByVal objDocumentoCampo As EntDocumentoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "CENTRO COSTE"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar centro coste para ObjType:  " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ")  Encontrada tabla: " & Tabla)

            'Devuelve el centro coste del documento definitivo
            Dim CentroCoste As String = getCentroCosteDocumentoDefinitivo(Tabla, CardCode, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, objDocumentoCampo.Dimension, Sociedad)

            If Not String.IsNullOrEmpty(CentroCoste) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Centro coste encontrado"
                retVal.MENSAJEAUX = CentroCoste

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra el centro coste"
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

    Public Function getComprobarNumEnvio(ByVal objDocumentoCampo As EntDocumentoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUMERO ENVIO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar número envío para ObjType:  " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ")  Encontrada tabla: " & Tabla)

            'Devuelve el centro coste del documento definitivo
            Dim NumEnvio As String = getNumEnvioDocumentoDefinitivo(Tabla, CardCode, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, objDocumentoCampo.Dimension, Sociedad)

            If Not String.IsNullOrEmpty(NumEnvio) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Número envío encontrado"
                retVal.MENSAJEAUX = NumEnvio

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra el número envío"
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

    Public Function getComprobarNumDestino(ByVal objDocumentoCampo As EntDocumentoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUMERO DESTINO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar número destino para ObjType:  " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard & ", NumEnvio: " & objDocumentoCampo.NumEnvio)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) _
                AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumEnvio) Then _
                Throw New Exception("ID Docuware, número de documento, referencia origen o número envío no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla origen por ObjectType
            Dim TablaOrigen As String = getTablaDeObjType(objDocumentoCampo.ObjTypeOrigen)
            If String.IsNullOrEmpty(TablaOrigen) Then Throw New Exception("No encuentro tabla origen con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ")  Encontrada tabla origen: " & TablaOrigen)

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ")  Encontrada tabla destino: " & TablaDestino)

            'Devuelve el centro coste del documento definitivo
            Dim DocNums As List(Of String) = getNumDestinoDocumentoDefinitivo(TablaOrigen, TablaDestino, CardCode, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, objDocumentoCampo.NumEnvio, Sociedad)
            If DocNums Is Nothing OrElse DocNums.Count = 0 Then Throw New Exception("No existe documento relacionado o su estado es cancelado")

            retVal.CODIGO = Respuesta.Ok
            retVal.MENSAJE = "Número destino encontrado"
            retVal.MENSAJEAUX = String.Join("#", DocNums)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Public Function getComprobarTitular(ByVal objDocumentoCampo As EntDocumentoCampo, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "TITULAR"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar titular para ObjType: " & objDocumentoCampo.ObjType & ", Ambito: " & objDocumentoCampo.Ambito & ", NIFTercero: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial & ", DOCIDDW: " & objDocumentoCampo.DOCIDDW & ", DocNum: " & objDocumentoCampo.DocNum & ", NumAtCard: " & objDocumentoCampo.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objDocumentoCampo.NIFTercero) AndAlso String.IsNullOrEmpty(objDocumentoCampo.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) AndAlso String.IsNullOrEmpty(objDocumentoCampo.NumAtCard) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = ""
            If String.IsNullOrEmpty(objDocumentoCampo.DOCIDDW) AndAlso String.IsNullOrEmpty(objDocumentoCampo.DocNum) Then
                CardCode = oComun.getCardCode(objDocumentoCampo.NIFTercero, objDocumentoCampo.RazonSocial, objDocumentoCampo.Ambito, Sociedad)
                If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumentoCampo.NIFTercero & ", Razón social: " & objDocumentoCampo.RazonSocial)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objDocumentoCampo.ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objDocumentoCampo.ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el titular del documento
            Dim Titular As String = getTitularDocumentoDefinitivo(Tabla, CardCode, objDocumentoCampo.DOCIDDW, objDocumentoCampo.DocNum, objDocumentoCampo.NumAtCard, Sociedad)

            If Not String.IsNullOrEmpty(Titular) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Titular de este documento encontrado"
                retVal.MENSAJEAUX = Titular

            Else

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "No se encuentra el titular de este documento"
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

#End Region

#Region "Comunes"

    Private Function getDraftDeCardCode(ByVal CardCode As String,
                                        ByVal Sociedad As eSociedad) As String

        'Devuelve si se debe generar el documento como borrador o no

        Dim retVal As String = Draft.Borrador

        Try

            'Buscamos por CardCode
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("U_SEIDraft") & ",N'" & Draft.Borrador & "') " & vbCrLf
            SQL &= " FROM " & getDataBaseRef("OCRD", Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("CardCode") & "=N'" & CardCode & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getCuentaContableDeCardCode(ByVal CardCode As String,
                                                 ByVal Sociedad As eSociedad) As String

        'Devuelve la cuenta contable

        Dim retVal As String = ""

        Try

            'Buscamos por CardCode
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("U_SEICtaDW") & ",'') " & vbCrLf
            SQL &= " FROM " & getDataBaseRef("OCRD", Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("CardCode") & "=N'" & CardCode & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getStatusImpuestoDeCardCode(ByVal CardCode As String,
                                                           ByVal Sociedad As eSociedad) As String

        'Devuelve si el IC es intracomunitario o está exento

        Dim retVal As String = ""

        Try

            'Buscamos por CardCode
            Dim SQL As String = ""
            SQL = "  SELECT COALESCE(T0." & putQuotes("VatStatus") & ",'') " & vbCrLf
            SQL &= " FROM " & getDataBaseRef("OCRD", Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("CardCode") & "=N'" & CardCode & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getGrupoIVADeCardCode(ByVal CardCode As String,
                                           ByVal Sociedad As eSociedad) As String

        'Devuelve el grupo de IVA

        Dim retVal As String = ""

        Try

            'Buscamos por CardCode
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("ECVatGroup") & ",'') " & vbCrLf
            SQL &= " FROM " & getDataBaseRef("OCRD", Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("CardCode") & "=N'" & CardCode & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getResponsableDeCardCode(ByVal CardCode As String,
                                              ByVal Sociedad As eSociedad) As String

        'Devuelve el responsable

        Dim retVal As String = ""

        Try

            'Buscamos por CardCode
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("AgentCode") & ",'') " & vbCrLf
            SQL &= " FROM " & getDataBaseRef("OCRD", Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("CardCode") & "=N'" & CardCode & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getContactoDeEmail(ByVal CardCode As String,
                                        ByVal Email As String,
                                        ByVal Sociedad As eSociedad) As Integer

        'Devuelve el contacto

        Dim retVal As Integer = 0

        Try

            If Not String.IsNullOrEmpty(Email) Then

                'Buscamos por CardCode e Email
                Dim SQL As String = ""
                SQL = "  SELECT " & vbCrLf
                SQL &= " COALESCE(T0." & putQuotes("CntctCode") & ",0) " & vbCrLf
                SQL &= " FROM " & getDataBaseRef("OCPR", Sociedad) & " T0 " & getWithNoLock() & vbCrLf
                SQL &= " WHERE 1=1 " & vbCrLf
                SQL &= " And T0." & putQuotes("CardCode") & "=N'" & CardCode & "'" & vbCrLf
                SQL &= " And UPPER(T0." & putQuotes("E_MailL") & ")=N'" & Email.ToUpper & "'" & vbCrLf

                Dim oCon As clsConexion = New clsConexion(Sociedad)

                Dim oObj As Object = oCon.ExecuteScalar(SQL)
                If IsNumeric(oObj) AndAlso CInt(oObj) > 0 Then retVal = CInt(oObj)

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getGrupoRetencionImpuestosDeCardCode(ByVal CardCode As String,
                                                          ByVal Sociedad As eSociedad) As String

        'Devuelve el grupo de retención de impuestos

        Dim retVal As String = ""

        Try

            'Buscamos por CardCode
            Dim SQL As String = ""
            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("WTCode") & ",'') " & vbCrLf
            SQL &= " FROM " & getDataBaseRef("OCRD", Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            'SQL &= " FROM " & getDataBaseRef("CRD4", Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("CardCode") & "=N'" & CardCode & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getPeriodoAbiertoDeTaxDate(ByVal TaxDate As String,
                                                ByVal Sociedad As eSociedad) As Boolean

        'Comprueba si el periodo contable está abierto

        Dim retVal As Boolean = False

        Try

            'Buscamos por fecha contable
            Dim SQL As String = ""
            SQL = "  SELECT COUNT(*) " & vbCrLf
            SQL &= " FROM " & getDataBaseRef("OFPR", Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("PeriodStat") & "<>N'" & DocStatus.Cerrado & "' " & vbCrLf
            SQL &= " And '" & TaxDate & "' BETWEEN T0." & putQuotes("F_TaxDate") & " And T0." & putQuotes("T_TaxDate") & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso IsNumeric(oObj) Then retVal = CInt(oObj) > 0

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getVatGroupDeVatTotal(ByVal Ambito As String,
                                           ByVal LineTotal As Double,
                                           ByVal VatTotal As Double,
                                           ByVal VatPorc As Double,
                                           ByVal Intracomunitario As String,
                                           ByVal ICGrupoIVA As String,
                                           ByVal Sociedad As eSociedad) As String

        'Grupo de IVA

        Dim retVal As String = ""

        Try

            'Porcentaje de IVA (cuidado divisiones por 0)
            Dim VatPerc As Double = 0
            If LineTotal > 0 Then VatPerc = Math.Round(VatTotal / LineTotal * 100.0, 2)

            'Buscamos por porcentaje de IVA entre Sx
            Dim SQL As String = ""

            SQL = "  SELECT " & vbCrLf
            SQL &= " T." & putQuotes("CODE") & vbCrLf
            SQL &= " FROM ( " & vbCrLf
            SQL &= "    SELECT " & vbCrLf
            SQL &= "	MAX(COALESCE(T1." & putQuotes("Rate") & ", 0)) As RATE," & vbCrLf
            SQL &= "	T1." & putQuotes("Code") & " As CODE" & vbCrLf
            SQL &= "	FROM " & getDataBaseRef("OVTG", Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= "	JOIN " & getDataBaseRef("VTG1", Sociedad) & " T1 " & getWithNoLock() & " ON T1." & putQuotes("Code") & "=T0." & putQuotes("Code") & vbCrLf
            SQL &= "	WHERE 1 = 1" & vbCrLf

            'Activo
            SQL &= "    And T0." & putQuotes("Inactive") & "<>N'" & SN.Yes & "'" & vbCrLf

            'Intracomunitario
            SQL &= "    And T0." & putQuotes("IsEC") & "=N'" & IIf(Intracomunitario = SN.Si, SN.Yes, SN.No) & "'" & vbCrLf

            'IVA soportado/repercutivo
            If Ambito = Utilidades.Ambito.Ventas Then

                'Categoría
                SQL &= "    And T0." & putQuotes("Category") & "=N'" & IVA.Repercutido & "'" & vbCrLf

                'Grupo de IVA del IC
                If Not String.IsNullOrEmpty(ICGrupoIVA) Then

                    'MLD (2022/11/07): Grupo de IVA del IC cambiando los números
                    SQL &= "	And T0." & putQuotes("Code") & " IN ('x1_y2_z3'"
                    For i = 0 To 9
                        SQL &= ",N'" & ObtenerGrupoIVA(ICGrupoIVA, i.ToString) & "'"
                    Next
                    SQL &= ") " & vbCrLf

                ElseIf Intracomunitario = SN.Si Then

                    'Códigos de IVA según el nombre
                    SQL &= "    And T0." & putQuotes("Name") & " like '%intracomunitari%'" & vbCrLf

                Else

                    'Códigos de IVA por defecto
                    SQL &= "	And T0." & putQuotes("Code") & " IN ('R0','R1','R2','R3') " & vbCrLf
                    'SQL &= "    And T0." & putQuotes("Name") & " like 'IVA repercutido al%'" & vbCrLf

                End If

            Else

                'Categoría
                SQL &= "    And T0." & putQuotes("Category") & "=N'" & IVA.Soportado & "'" & vbCrLf

                'Grupo de IVA del IC
                If Not String.IsNullOrEmpty(ICGrupoIVA) Then

                    'MLD (2022/11/07): Grupo de IVA del IC cambiando los números
                    SQL &= "	And T0." & putQuotes("Code") & " IN ('x1_y2_z3'"
                    For i = 0 To 9
                        SQL &= ",N'" & ObtenerGrupoIVA(ICGrupoIVA, i.ToString) & "'"
                    Next
                    SQL &= ") " & vbCrLf

                ElseIf Intracomunitario = SN.Si Then

                    'Códigos de IVA según el nombre
                    SQL &= "    And T0." & putQuotes("Name") & " like 'Adquisiciones Intracomunitarias de Bienes%'" & vbCrLf

                Else

                    'Códigos de IVA por defecto
                    SQL &= "	And T0." & putQuotes("Code") & " IN ('S0','S1','S2','S3') " & vbCrLf
                    'SQL &= "    And T0." & putQuotes("Name") & " like 'IVA soportado al%'" & vbCrLf

                End If

            End If

            SQL &= "	GROUP BY " & vbCrLf
            SQL &= "	T1." & putQuotes("Code") & ", " & vbCrLf
            SQL &= "	T0." & putQuotes("Code") & vbCrLf

            SQL &= " ) As T " & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            'MLD (2021/02/03): Pasan el % de IVA, aplicar corrección al calculado por nosotros
            If VatPorc > 0 Then
                SQL &= " And T." & putQuotes("RATE") & "=" & VatPorc.ToString.Replace(",", ".") & vbCrLf
            Else
                SQL &= " And T." & putQuotes("RATE") & "<=" & (VatPerc * 1.03).ToString.Replace(",", ".") & vbCrLf
                SQL &= " And T." & putQuotes("RATE") & ">=" & (VatPerc * 0.97).ToString.Replace(",", ".") & vbCrLf
            End If

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)

            If Not String.IsNullOrEmpty(ICGrupoIVA) AndAlso String.IsNullOrEmpty(oObj) Then _
                Throw New Exception("No se encuentra grupo IVA al " & IIf(VatPorc > 0, VatPorc, VatPerc) & "% tomando como base " & ICGrupoIVA)

            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            If ex.Message.Contains("No se encuentra grupo IVA al ") Then Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

#End Region

#Region "Privadas"

    Private Function NuevoDocumento(ByVal objDocumento As EntDocumentoCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oDocDestino As Documents = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUEVO DOCUMENTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de documento para " & objDocumento.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objDocumento.UserSAP, objDocumento.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objDocumento.NIFTercero) Then Throw New Exception("NIF no suministrado")

            If Not DateTime.TryParseExact(objDocumento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha documento no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objDocumento.DocDateAux.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha documento auxiliar no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha contable no suministrada o incorrecta")

            If String.IsNullOrEmpty(objDocumento.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocumento.DOCURLDW) Then Throw New Exception("URL Docuware no suministrado")

            If objDocumento.Lineas Is Nothing Then Throw New Exception("Líneas no suministradas")

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objDocumento.ObjTypeDestino)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objDocumento.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objDocumento.NIFTercero, objDocumento.RazonSocial, objDocumento.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumento.NIFTercero & ", Razón social: " & objDocumento.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Buscar cuenta contable IC
            Dim CuentaContable As String = getCuentaContableDeCardCode(CardCode, Sociedad)
            If Not String.IsNullOrEmpty(CuentaContable) Then clsLog.Log.Info("(" & sLogInfo & ") Encontrada cuenta contable IC: " & CuentaContable)

            'Buscar grupo IVA IC
            Dim ICGrupoIVA As String = getGrupoIVADeCardCode(CardCode, Sociedad)
            If Not String.IsNullOrEmpty(ICGrupoIVA) Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado grupo IVA IC: " & ICGrupoIVA)

            'Buscar status impuesto IC
            Dim ICStatusImpuesto As String = getStatusImpuestoDeCardCode(CardCode, Sociedad)
            If Not String.IsNullOrEmpty(ICStatusImpuesto) Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado status impuesto IC: " & ICStatusImpuesto)

            'Buscar responsable IC
            Dim ICResponsable As String = ""
            If getEsDocumentoVenta(objDocumento.ObjTypeDestino) Then ICResponsable = getResponsableDeCardCode(CardCode, Sociedad)
            If Not String.IsNullOrEmpty(ICResponsable) Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado responsable IC: " & ICResponsable)

            'Buscar contacto IC
            Dim ICContacto As Integer = getContactoDeEmail(CardCode, objDocumento.ContactoEmail, Sociedad)
            If ICContacto > 0 Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado contacto IC: " & ICContacto)

            'Comprueba si el documento se debe generar como borrador o en firme 
            Dim Draft As String = objDocumento.Draft
            If Draft = Utilidades.Draft.Interlocutor Then
                Draft = getDraftDeCardCode(CardCode, Sociedad)
                clsLog.Log.Info("(" & sLogInfo & ") Documento borrador por IC")
            End If

            'Objeto documento
            If Draft <> Utilidades.Draft.Firme Then
                oDocDestino = oCompany.GetBusinessObject(BoObjectTypes.oDrafts)
                oDocDestino.DocObjectCode = objDocumento.ObjTypeDestino
                clsLog.Log.Info("(" & sLogInfo & ") Documento borrador")
            Else
                oDocDestino = oCompany.GetBusinessObject(objDocumento.ObjTypeDestino)
                clsLog.Log.Info("(" & sLogInfo & ") Documento en firme")
            End If

            'Serie
            If objDocumento.Serie > 0 Then oDocDestino.Series = objDocumento.Serie

            'IC
            oDocDestino.CardCode = CardCode

            'Contacto
            If objDocumento.ContactoCodigo > 0 Then
                oDocDestino.ContactPersonCode = objDocumento.ContactoCodigo
            ElseIf ICContacto > 0 Then
                oDocDestino.ContactPersonCode = ICContacto
            End If

            'Direccion envío
            If Not String.IsNullOrEmpty(objDocumento.DireccionEnvioCodigo) Then oDocDestino.ShipToCode = objDocumento.DireccionEnvioCodigo
            If Not String.IsNullOrEmpty(objDocumento.DireccionEnvioDetalle) Then oDocDestino.Address2 = objDocumento.DireccionEnvioDetalle

            'Direccion facturación
            If Not String.IsNullOrEmpty(objDocumento.DireccionFacturaCodigo) Then oDocDestino.PayToCode = objDocumento.DireccionFacturaCodigo
            If Not String.IsNullOrEmpty(objDocumento.DireccionFacturaDetalle) Then oDocDestino.Address = objDocumento.DireccionFacturaDetalle

            'FinanzasRazon
            If Not String.IsNullOrEmpty(objDocumento.FinanzasRazon) Then oDocDestino.CardName = objDocumento.FinanzasRazon

            'FinanzasNIF
            If Not String.IsNullOrEmpty(objDocumento.FinanzasNIF) Then oDocDestino.FederalTaxID = objDocumento.FinanzasNIF

            'Referencia
            oDocDestino.NumAtCard = objDocumento.NumAtCard

            'Moneda
            If Not String.IsNullOrEmpty(objDocumento.Currency) Then oDocDestino.DocCurrency = objDocumento.Currency

            'Sucursal
            If objDocumento.Sucursal > 0 Then oDocDestino.BPL_IDAssignedToInvoice = objDocumento.Sucursal

            'Proyecto
            If Not String.IsNullOrEmpty(objDocumento.Proyecto) Then oDocDestino.Project = objDocumento.Proyecto

            'Clase expedicion
            If objDocumento.ClaseExpedicion > 0 Then oDocDestino.TransportationCode = objDocumento.ClaseExpedicion

            'Responsable
            If Not String.IsNullOrEmpty(objDocumento.Responsable) Then
                oDocDestino.AgentCode = objDocumento.Responsable
            ElseIf Not String.IsNullOrEmpty(ICResponsable) Then
                oDocDestino.AgentCode = ICResponsable
            End If

            'Titular
            If objDocumento.Titular > 0 Then oDocDestino.DocumentsOwner = objDocumento.Titular

            'Empleado dpto compras/ventas
            If objDocumento.EmpDptoCompraVenta > 0 Then oDocDestino.SalesPersonCode = objDocumento.EmpDptoCompraVenta

            'Facturas anticipo
            setFacturasAnticipo(objDocumento, oDocDestino, CardCode, Sociedad)

            'Documentos relacionados
            setDocumentosRelacionados(objDocumento, oDocDestino, CardCode, sLogInfo, Sociedad)

            'Comentarios
            If objDocumento.Comments.Length > 254 Then objDocumento.Comments = objDocumento.Comments.Substring(0, 254)
            If Not String.IsNullOrEmpty(objDocumento.Comments) Then oDocDestino.Comments = objDocumento.Comments

            'Entrada diario
            If objDocumento.JournalMemo.Length > 254 Then objDocumento.JournalMemo = objDocumento.JournalMemo.Substring(0, 254)
            If Not String.IsNullOrEmpty(objDocumento.JournalMemo) Then oDocDestino.JournalMemo = objDocumento.JournalMemo

            'Bloqueo de pago
            If objDocumento.BloqueoPago > 0 Then
                oDocDestino.PaymentBlock = BoYesNoEnum.tYES
                oDocDestino.PaymentBlockEntry = objDocumento.BloqueoPago
            End If

            'Portes 
            setPortesCabecera(objDocumento, oDocDestino, CardCode, "", Nothing, sLogInfo, Sociedad)

            'Factura de reserva
            If objDocumento.Reserva = SN.Si Then
                oDocDestino.ReserveInvoice = BoYesNoEnum.tYES
                clsLog.Log.Info("(" & sLogInfo & ") Factura de reserva")
            End If

            'Fecha contable 
            oDocDestino.DocDate = Date.ParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Fecha vencimiento
            If DateTime.TryParseExact(objDocumento.DocDueDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                oDocDestino.DocDueDate = Date.ParseExact(objDocumento.DocDueDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Fecha documento 
            'Comprueba si el periodo contable está abierto para la fecha DocDate de entrada. Si no lo está, coge la fecha auxiliar DocDateAux
            Dim PeriodoAbierto As Boolean = getPeriodoAbiertoDeTaxDate(objDocumento.DocDate.ToString, Sociedad)
            If PeriodoAbierto Then
                oDocDestino.TaxDate = Date.ParseExact(objDocumento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            Else
                oDocDestino.TaxDate = Date.ParseExact(objDocumento.DocDateAux.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            End If

            'Tipo de documento. Por defecto, servicio
            If objDocumento.DocType = DocType.Articulo Then

                'Artículo
                oDocDestino.DocType = BoDocumentTypes.dDocument_Items

                'Recorre cada línea del servicio
                For Each objLinea In objDocumento.Lineas

                    'Índice
                    Dim indLinea As Integer = objDocumento.Lineas.IndexOf(objLinea) + 1

                    'Comprobar campos
                    If String.IsNullOrEmpty(objLinea.Articulo) AndAlso String.IsNullOrEmpty(objLinea.RefExt) Then _
                        Throw New Exception("Artículo o referencia externa de la línea " & indLinea.ToString & " no suministrado")

                    'Líneas especiales
                    If objLinea.Articulo = ArticuloLineaEspecialTexto OrElse objLinea.Articulo = ArticuloLineaEspecialSubtotal Then

                        'Se posiciona en la línea
                        If objLinea.Articulo = ArticuloLineaEspecialTexto AndAlso String.IsNullOrEmpty(objLinea.Concepto) Then _
                            Throw New Exception("Concepto de la línea especial " & indLinea.ToString & " no suministrado")

                        'Rellena la línea
                        oDocDestino.SpecialLines.SetCurrentLine(oDocDestino.SpecialLines.Count - 1)

                        oDocDestino.SpecialLines.LineType = IIf(objLinea.Articulo = ArticuloLineaEspecialTexto, BoDocSpecialLineType.dslt_Text, BoDocSpecialLineType.dslt_Subtotal)
                        If Not String.IsNullOrEmpty(objLinea.Concepto) Then oDocDestino.SpecialLines.LineText = objLinea.Concepto

                        oDocDestino.SpecialLines.AfterLineNumber = IIf(oDocDestino.Lines.Count = 1, -1, oDocDestino.Lines.Count - 2)
                        oDocDestino.SpecialLines.Add()

                        Continue For

                    End If

                    'Se posiciona en la línea
                    oDocDestino.Lines.SetCurrentLine(oDocDestino.Lines.Count - 1)

                    'Artículo (si no se encuentra el código se busca por la referencia de proveedor)
                    Dim ItemCode As String = oComun.getItemCode(CardCode, objLinea.Articulo, objLinea.RefExt, Sociedad)
                    If String.IsNullOrEmpty(ItemCode) Then Throw New Exception("No encuentro artículo de la línea " & indLinea.ToString & " con código: " & objLinea.Articulo & " o referencia externa: " & objLinea.RefExt & " para IC: " & CardCode)
                    clsLog.Log.Info("(" & sLogInfo & ") Encontrado artículo: " & ItemCode)

                    oDocDestino.Lines.ItemCode = ItemCode

                    'Descripción
                    If objLinea.Concepto.Length > 100 Then objLinea.Concepto = objLinea.Concepto.Substring(0, 100)
                    If Not String.IsNullOrEmpty(objLinea.Concepto) Then oDocDestino.Lines.ItemDescription = objLinea.Concepto

                    'Cantidad
                    If objLinea.Cantidad > 0 Then oDocDestino.Lines.Quantity = objLinea.Cantidad

                    'Precios especiales (para que quede plasmado precio + descuento en vez de solo precio final tras descuento)
                    Dim oPrecioSAP As ItemPriceReturnParams = Nothing
                    If getEsDocumentoVenta(objDocumento.ObjTypeDestino) Then
                        oPrecioSAP = getPrecioConDescuento(oCompany, CardCode, ItemCode, oDocDestino.DocDate, objLinea.Cantidad)
                    End If

                    'Precio
                    If objLinea.PrecioUnidad <> 0 Then
                        oDocDestino.Lines.UnitPrice = objLinea.PrecioUnidad
                        oDocDestino.Lines.Price = objLinea.PrecioUnidad
                    ElseIf Not oPrecioSAP Is Nothing AndAlso oPrecioSAP.Price <> 0 Then
                        oDocDestino.Lines.UnitPrice = oPrecioSAP.Price
                    End If

                    'Porcentaje dto
                    If objLinea.PorcentajeDescuento <> 0 Then
                        oDocDestino.Lines.DiscountPercent = objLinea.PorcentajeDescuento
                    ElseIf Not oPrecioSAP Is Nothing AndAlso oPrecioSAP.Discount <> 0 Then
                        oDocDestino.Lines.DiscountPercent = oPrecioSAP.Discount
                    End If

                    'Moneda
                    If Not String.IsNullOrEmpty(objDocumento.Currency) Then oDocDestino.Lines.Currency = objDocumento.Currency

                    'TaxOnly
                    oDocDestino.Lines.TaxOnly = BoYesNoEnum.tNO
                    If Not String.IsNullOrEmpty(objLinea.TaxOnly) AndAlso objLinea.TaxOnly = SN.Si Then oDocDestino.Lines.TaxOnly = BoYesNoEnum.tYES

                    'TaxCode
                    If Not String.IsNullOrEmpty(objLinea.TaxCode) Then oDocDestino.Lines.TaxCode = objLinea.TaxCode

                    'WTLiable
                    oDocDestino.Lines.WTLiable = BoYesNoEnum.tNO
                    If Not String.IsNullOrEmpty(objLinea.WTLiable) AndAlso objLinea.WTLiable = SN.Si Then oDocDestino.Lines.WTLiable = BoYesNoEnum.tYES

                    'VATGroup
                    If Not String.IsNullOrEmpty(objLinea.VATGroup) Then oDocDestino.Lines.VatGroup = objLinea.VATGroup

                    'Fecha entrega
                    If DateTime.TryParseExact(objLinea.ShipDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        oDocDestino.Lines.ShipDate = Date.ParseExact(objLinea.ShipDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                    'Fecha requerida
                    If DateTime.TryParseExact(objLinea.RequiredDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        oDocDestino.Lines.RequiredDate = Date.ParseExact(objLinea.RequiredDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                    'Proyecto
                    If Not String.IsNullOrEmpty(objLinea.Proyecto) Then oDocDestino.Lines.ProjectCode = objLinea.Proyecto

                    'Almacen 
                    If Not String.IsNullOrEmpty(objLinea.Almacen) Then oDocDestino.Lines.WarehouseCode = objLinea.Almacen

                    'Centro coste 1
                    If Not String.IsNullOrEmpty(objLinea.Centro1Coste) Then oDocDestino.Lines.CostingCode = objLinea.Centro1Coste

                    'Centro coste 2
                    If Not String.IsNullOrEmpty(objLinea.Centro2Coste) Then oDocDestino.Lines.CostingCode2 = objLinea.Centro2Coste

                    'Centro coste 3
                    If Not String.IsNullOrEmpty(objLinea.Centro3Coste) Then oDocDestino.Lines.CostingCode3 = objLinea.Centro3Coste

                    'Centro coste 4
                    If Not String.IsNullOrEmpty(objLinea.Centro4Coste) Then oDocDestino.Lines.CostingCode4 = objLinea.Centro4Coste

                    'Centro coste 5
                    If Not String.IsNullOrEmpty(objLinea.Centro5Coste) Then oDocDestino.Lines.CostingCode5 = objLinea.Centro5Coste

                    'Lote
                    setLotesLinea(objLinea, oDocDestino)

                    'Campos de usuario
                    setCamposUsuarioLinea(objLinea, oDocDestino)

                    'Añade línea
                    oDocDestino.Lines.Add()

                Next

            Else

                'Servicio
                oDocDestino.DocType = BoDocumentTypes.dDocument_Service

                'Comprobar que hay al menos una línea informada
                If objDocumento.Lineas Is Nothing OrElse objDocumento.Lineas.Count = 0 Then Throw New Exception("Línea 1 no suministrada")

                'Recorre cada línea del servicio
                For Each objLinea In objDocumento.Lineas

                    'Índice
                    Dim indLinea As Integer = objDocumento.Lineas.IndexOf(objLinea) + 1

                    'Líneas especiales
                    If objLinea.Articulo = "SEIDORTEXTO" Then Continue For

                    'Comprobar datos primera línea
                    If indLinea = 1 AndAlso String.IsNullOrEmpty(objLinea.Concepto) Then Throw New Exception("Concepto de la línea 1 no suministrado")

                    If indLinea = 1 AndAlso objLinea.LineTotal = 0 AndAlso objLinea.BaseTotal = 0 Then Throw New Exception("Importes de la línea 1 no suministrados")

                    'Comprobar campos del resto de líneas
                    If indLinea > 1 AndAlso String.IsNullOrEmpty(objLinea.Concepto) AndAlso Not (objLinea.LineTotal = 0 AndAlso objLinea.BaseTotal = 0) Then _
                    Throw New Exception("Concepto de la línea " & indLinea.ToString & " no suministrado")

                    'Se posiciona en la línea
                    oDocDestino.Lines.SetCurrentLine(oDocDestino.Lines.Count - 1)

                    'Si la cuenta contable viene vacía, se coge la del IC
                    If Not String.IsNullOrEmpty(objLinea.CuentaContable) Then
                        oDocDestino.Lines.AccountCode = objLinea.CuentaContable
                    Else
                        If String.IsNullOrEmpty(CuentaContable) Then Throw New Exception("Cuenta contable no definida en la ficha del IC")
                        oDocDestino.Lines.AccountCode = CuentaContable
                    End If

                    'Concepto
                    If objLinea.Concepto.Length > 100 Then objLinea.Concepto = objLinea.Concepto.Substring(0, 100)
                    oDocDestino.Lines.ItemDescription = objLinea.Concepto

                    'Si LineTotal=0, LineTotal=BaseTotal
                    If objLinea.LineTotal = 0 Then objLinea.LineTotal = objLinea.BaseTotal
                    oDocDestino.Lines.LineTotal = objLinea.LineTotal

                    'Porcentaje dto
                    If objLinea.PorcentajeDescuento <> 0 Then oDocDestino.Lines.DiscountPercent = objLinea.PorcentajeDescuento

                    'TaxOnly
                    oDocDestino.Lines.TaxOnly = BoYesNoEnum.tNO
                    If Not String.IsNullOrEmpty(objLinea.TaxOnly) AndAlso objLinea.TaxOnly = SN.Si Then oDocDestino.Lines.TaxOnly = BoYesNoEnum.tYES

                    'TaxCode
                    If Not String.IsNullOrEmpty(objLinea.TaxCode) Then oDocDestino.Lines.TaxCode = objLinea.TaxCode

                    'WTLiable
                    oDocDestino.Lines.WTLiable = BoYesNoEnum.tNO
                    If Not String.IsNullOrEmpty(objLinea.WTLiable) AndAlso objLinea.WTLiable = SN.Si Then oDocDestino.Lines.WTLiable = BoYesNoEnum.tYES

                    'Grupo impositivo
                    If Not String.IsNullOrEmpty(objLinea.VATGroup) Then
                        oDocDestino.Lines.VatGroup = objLinea.VATGroup
                    ElseIf ICStatusImpuesto <> StatusImpuesto.Obligatorio AndAlso Not String.IsNullOrEmpty(ICGrupoIVA) Then
                        oDocDestino.Lines.VatGroup = ICGrupoIVA
                    Else
                        oDocDestino.Lines.VatGroup = getVatGroupDeVatTotal(objDocumento.Ambito, objLinea.LineTotal, objLinea.VATTotal, objLinea.VATPorc, objLinea.Intracomunitario, ICGrupoIVA, Sociedad)
                    End If

                    'Moneda
                    If Not String.IsNullOrEmpty(objDocumento.Currency) Then oDocDestino.Lines.Currency = objDocumento.Currency

                    'Fecha entrega
                    If DateTime.TryParseExact(objLinea.ShipDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                    oDocDestino.Lines.ShipDate = Date.ParseExact(objLinea.ShipDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                    'Fecha requerida
                    If DateTime.TryParseExact(objLinea.RequiredDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                    oDocDestino.Lines.RequiredDate = Date.ParseExact(objLinea.RequiredDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                    'Proyecto
                    If Not String.IsNullOrEmpty(objLinea.Proyecto) Then oDocDestino.Lines.ProjectCode = objLinea.Proyecto

                    'Centro coste 1
                    If Not String.IsNullOrEmpty(objLinea.Centro1Coste) Then oDocDestino.Lines.CostingCode = objLinea.Centro1Coste

                    'Centro coste 2
                    If Not String.IsNullOrEmpty(objLinea.Centro2Coste) Then oDocDestino.Lines.CostingCode2 = objLinea.Centro2Coste

                    'Centro coste 3
                    If Not String.IsNullOrEmpty(objLinea.Centro3Coste) Then oDocDestino.Lines.CostingCode3 = objLinea.Centro3Coste

                    'Centro coste 4
                    If Not String.IsNullOrEmpty(objLinea.Centro4Coste) Then oDocDestino.Lines.CostingCode4 = objLinea.Centro4Coste

                    'Centro coste 5
                    If Not String.IsNullOrEmpty(objLinea.Centro5Coste) Then oDocDestino.Lines.CostingCode5 = objLinea.Centro5Coste

                    'Campos de usuario
                    setCamposUsuarioLinea(objLinea, oDocDestino)

                    'Añade línea
                    oDocDestino.Lines.Add()

                Next

            End If

            'Retenciones de impuestos de IC (DocTotal sin retenciones)
            setRetencionesImpuestos(objDocumento, oDocDestino, CardCode, sLogInfo, Sociedad)

            'Total de documento 
            If objDocumento.DocTotal <> 0 Then oDocDestino.DocTotal = objDocumento.DocTotal

            'ID DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIDDW").Value = objDocumento.DOCIDDW

            'URL DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIURLDW").Value = objDocumento.DOCURLDW

            'IMPORTE DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIMPDW").Value = objDocumento.DocTotal

            'REVISAR IMPORTE DOCUWARE
            'El sistema permite traspasar un borrador a definitivo, aunque el importe del documento sea diferente que al importe DW
            oDocDestino.UserFields.Fields.Item("U_SEIREVDW").Value = IIf(Draft = Utilidades.Draft.Borrador, "S", "N")

            'Campos de usuario
            setCamposUsuarioCabecera(objDocumento, oDocDestino)

            'Añadimos documento
            If oDocDestino.Add() = 0 Then

                'Obtiene el DocEntry y DocNum del documento añadido
                Dim DocEntry As String = oCompany.GetNewObjectKey

                Dim DocNum As String = ""
                If Draft = Utilidades.Draft.Firme Then DocNum = oComun.getDocNumDeDocEntry(TablaDestino, DocEntry, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento creado con éxito"
                retVal.MENSAJEAUX = DocNum

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
            oComun.LiberarObjCOM(oDocDestino)
        End Try

        Return retVal

    End Function

    Private Function CopiarDeDocumento(ByVal objDocumento As EntDocumentoCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oDocEntrada As Documents = Nothing
        Dim oDocDestino As Documents = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "COPIAR DOCUMENTO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de copia de documento para " & objDocumento.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objDocumento.UserSAP, objDocumento.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objDocumento.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocumento.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")
            If String.IsNullOrEmpty(objDocumento.RefOrigen) Then Throw New Exception("Referencia origen no suministrada")

            If Not DateTime.TryParseExact(objDocumento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha documento no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objDocumento.DocDateAux.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha documento auxiliar no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha contable no suministrada o incorrecta")

            If objDocumento.Lineas Is Nothing Then Throw New Exception("Líneas no suministradas")

            'Buscamos tabla origen por ObjectType
            Dim TablaOrigen As String = getTablaDeObjType(objDocumento.ObjTypeOrigen)
            If String.IsNullOrEmpty(TablaOrigen) Then Throw New Exception("No encuentro tabla origen con ObjectType: " & objDocumento.ObjTypeOrigen)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla origen: " & TablaOrigen)

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objDocumento.ObjTypeDestino)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objDocumento.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objDocumento.NIFTercero, objDocumento.RazonSocial, objDocumento.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumento.NIFTercero & ", Razón social: " & objDocumento.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Buscar contacto IC
            Dim ICContacto As Integer = getContactoDeEmail(CardCode, objDocumento.ContactoEmail, Sociedad)
            If ICContacto > 0 Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado contacto IC: " & ICContacto)

            'Buscar responsable IC
            Dim ICResponsable As String = ""
            If getEsDocumentoVenta(objDocumento.ObjTypeDestino) Then ICResponsable = getResponsableDeCardCode(CardCode, Sociedad)
            If Not String.IsNullOrEmpty(ICResponsable) Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado responsable IC: " & ICResponsable)

            'Comprueba si el documento se debe generar como borrador o en firme 
            Dim Draft As String = objDocumento.Draft
            If Draft = Utilidades.Draft.Interlocutor Then
                Draft = getDraftDeCardCode(CardCode, Sociedad)
                clsLog.Log.Info("(" & sLogInfo & ") Documento borrador por IC")
            End If

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objDocumento.RefOrigen.Split("#")

            'Origen con búsqueda indirecta (por ejemplo, albaranes de pedidos abiertos)
            getRefsOrigenDocumentoNoDirecto(objDocumento, CardCode, TablaOrigen, RefsOrigen, sLogInfo, Sociedad)

            'Control importes
            getControlImportes(objDocumento, CardCode, TablaOrigen, RefsOrigen, Draft, sLogInfo, Sociedad)

            'Objeto documento
            If Draft = Utilidades.Draft.Borrador Then
                oDocDestino = oCompany.GetBusinessObject(BoObjectTypes.oDrafts)
                oDocDestino.DocObjectCode = objDocumento.ObjTypeDestino
                clsLog.Log.Info("(" & sLogInfo & ") Documento borrador")
            Else
                oDocDestino = oCompany.GetBusinessObject(objDocumento.ObjTypeDestino)
                clsLog.Log.Info("(" & sLogInfo & ") Documento en firme")
            End If

            'Serie
            If objDocumento.Serie > 0 Then oDocDestino.Series = objDocumento.Serie

            'Contacto
            If objDocumento.ContactoCodigo > 0 Then
                oDocDestino.ContactPersonCode = objDocumento.ContactoCodigo
            ElseIf ICContacto > 0 Then
                oDocDestino.ContactPersonCode = ICContacto
            End If

            'Direccion envío
            If Not String.IsNullOrEmpty(objDocumento.DireccionEnvioCodigo) Then oDocDestino.ShipToCode = objDocumento.DireccionEnvioCodigo
            If Not String.IsNullOrEmpty(objDocumento.DireccionEnvioDetalle) Then oDocDestino.Address2 = objDocumento.DireccionEnvioDetalle

            'Direccion facturación
            If Not String.IsNullOrEmpty(objDocumento.DireccionFacturaCodigo) Then oDocDestino.PayToCode = objDocumento.DireccionFacturaCodigo
            If Not String.IsNullOrEmpty(objDocumento.DireccionFacturaDetalle) Then oDocDestino.Address = objDocumento.DireccionFacturaDetalle

            'Proyecto
            If Not String.IsNullOrEmpty(objDocumento.Proyecto) Then oDocDestino.Project = objDocumento.Proyecto

            'Bloqueo de pago
            If objDocumento.BloqueoPago > 0 Then
                oDocDestino.PaymentBlock = BoYesNoEnum.tYES
                oDocDestino.PaymentBlockEntry = objDocumento.BloqueoPago
            End If

            'Portes 
            setPortesCabecera(objDocumento, oDocDestino, CardCode, TablaOrigen, RefsOrigen, sLogInfo, Sociedad)

            'Factura de reserva
            If objDocumento.Reserva = SN.Si Then
                oDocDestino.ReserveInvoice = BoYesNoEnum.tYES
                clsLog.Log.Info("(" & sLogInfo & ") Factura de reserva")
            End If

            'Comprueba si la copia es parcial o no
            Dim CopiaParcial As Boolean = getEsCopiaParcial(objDocumento.Lineas)

            'Comprobaciones parciales
            If CopiaParcial Then

                clsLog.Log.Info("(" & sLogInfo & ") Copia parcial")

                'Comprueba que no lleguen líneas sin artículo/descripción en la copia parcial
                Dim bSinIdentificar As Boolean = getLinSinIdentificar(objDocumento)
                If bSinIdentificar Then Throw New Exception("Rellene el artículo/concepto a traspasar en todas las líneas")

                'Comprueba que no lleguen líneas a 0 en la copia parcial
                Dim bSinCantidad As Boolean = (From p In objDocumento.Lineas Where p.Cantidad <= 0).Distinct.ToList.Count > 0
                If bSinCantidad Then Throw New Exception("Rellene la cantidad a traspasar en todas las líneas")

                'Comprueba que el código de artículo sea ItemCode y no la referencia de proveedor
                For Each objLinea In objDocumento.Lineas

                    'Índice
                    Dim indLinea As Integer = objDocumento.Lineas.IndexOf(objLinea) + 1

                    If Not String.IsNullOrEmpty(objLinea.Articulo) OrElse Not String.IsNullOrEmpty(objLinea.RefExt) Then
                        'Artículo (si no se encuentra el código se busca por la referencia de proveedor)
                        Dim ItemCode As String = oComun.getItemCode(CardCode, objLinea.Articulo, objLinea.RefExt, Sociedad)
                        If Not String.IsNullOrEmpty(objLinea.Articulo) AndAlso objLinea.Articulo = ItemCode Then Continue For

                        If String.IsNullOrEmpty(ItemCode) Then Throw New Exception("No encuentro artículo de la línea " & indLinea.ToString & " con código: " & objLinea.Articulo & " o referencia externa: " & objLinea.RefExt & " para IC: " & CardCode)
                        clsLog.Log.Info("(" & sLogInfo & ") Encontrado artículo: " & ItemCode)

                        objLinea.Articulo = ItemCode
                    End If

                Next

            End If

            Dim ArticuloNoTraspasado As String = ""

            'Recorre cada uno de los documento origen
            For Each RefOrigen In RefsOrigen

                'Comprueba que existe el documento origen y que se puede acceder a él
                'Pueden ser varios documentos origen con la misma referencia
                Dim DocEntrysOrigen As List(Of String) = getDocEntryDeRefOrigen(TablaOrigen, RefOrigen, objDocumento.TipoRefOrigen, CardCode, True, True, Sociedad)
                If DocEntrysOrigen Is Nothing OrElse DocEntrysOrigen.Count = 0 Then Throw New Exception("No existe documento origen con referencia: " & RefOrigen & " o su estado es cerrado/cancelado")

                For Each DocEntryOrigen In DocEntrysOrigen

                    'Objeto origen
                    oDocEntrada = oCompany.GetBusinessObject(objDocumento.ObjTypeOrigen)

                    If Not oDocEntrada.GetByKey(DocEntryOrigen) Then Throw New Exception("No puedo recuperar el documento origen con referencia: " & RefOrigen & " y DocEntry:" & DocEntryOrigen)

                    'Copia los datos del documento origen al documento destino
                    oDocDestino.DocType = oDocEntrada.DocType
                    oDocDestino.CardCode = oDocEntrada.CardCode
                    oDocDestino.CardName = oDocEntrada.CardName

                    'Calcula la serie
                    Dim Serie As String = oComun.getSerieDeDocumentoDestino(oDocEntrada.Series, oDocEntrada.DocObjectCode, objDocumento.ObjTypeDestino, Sociedad)
                    If Not oDocDestino.Series > 0 AndAlso Not String.IsNullOrEmpty(Serie) Then oDocDestino.Series = CInt(Serie)

                    'Sucursal
                    If objDocumento.Sucursal > 0 AndAlso oDocEntrada.BPL_IDAssignedToInvoice > 0 AndAlso objDocumento.Sucursal <> oDocEntrada.BPL_IDAssignedToInvoice Then _
                        Throw New Exception("Sucursal destino: " & objDocumento.Sucursal & " no coincidente con sucursal origen: " & oDocEntrada.BPL_IDAssignedToInvoice)

                    If oDocEntrada.BPL_IDAssignedToInvoice > 0 AndAlso oDocDestino.BPL_IDAssignedToInvoice > 0 AndAlso oDocEntrada.BPL_IDAssignedToInvoice <> oDocDestino.BPL_IDAssignedToInvoice Then _
                        Throw New Exception("Sucursal origen: " & objDocumento.Sucursal & " no coincidente con sucursal origen: " & oDocEntrada.BPL_IDAssignedToInvoice)

                    If oDocEntrada.BPL_IDAssignedToInvoice > 0 Then oDocDestino.BPL_IDAssignedToInvoice = oDocEntrada.BPL_IDAssignedToInvoice

                    'Proyecto
                    If String.IsNullOrEmpty(oDocDestino.Project) AndAlso Not String.IsNullOrEmpty(oDocEntrada.Project) Then oDocDestino.Project = oDocEntrada.Project

                    'FinanzasRazon
                    If Not String.IsNullOrEmpty(objDocumento.FinanzasRazon) Then oDocDestino.CardName = objDocumento.FinanzasRazon

                    'FinanzasNIF
                    If Not String.IsNullOrEmpty(objDocumento.FinanzasNIF) Then oDocDestino.FederalTaxID = objDocumento.FinanzasNIF

                    'Responsable
                    If String.IsNullOrEmpty(oDocDestino.AgentCode) AndAlso Not String.IsNullOrEmpty(oDocEntrada.AgentCode) Then oDocDestino.AgentCode = oDocEntrada.AgentCode

                    'Titular
                    If Not oDocDestino.DocumentsOwner > 0 AndAlso oDocEntrada.DocumentsOwner > 0 Then oDocDestino.DocumentsOwner = oDocEntrada.DocumentsOwner

                    'Traspasa la primera línea especial que va en la posición 0 de un documento
                    setPrimeraLineaEspecial(oDocEntrada, oDocDestino)

                    'Traspasa las líneas
                    For iLinea As Integer = 0 To oDocEntrada.Lines.Count - 1

                        oDocEntrada.Lines.SetCurrentLine(iLinea)

                        '20200911: Comprueba que la línea no esté cerrada
                        If oDocEntrada.Lines.LineStatus <> BoStatus.bost_Close Then

                            'Diferencia entre traspaso total o parcial
                            If Not CopiaParcial Then

                                'Traspasa todas las líneas
                                oDocDestino.Lines.SetCurrentLine(oDocDestino.Lines.Count - 1)

                                oDocDestino.Lines.BaseEntry = oDocEntrada.DocEntry
                                oDocDestino.Lines.BaseLine = oDocEntrada.Lines.LineNum
                                oDocDestino.Lines.BaseType = oDocEntrada.DocObjectCode

                                'Portes de la línea
                                setPortesLinea(oDocEntrada, oDocDestino)

                                oDocDestino.Lines.Add()

                                'Traspasa las líneas especiales
                                setLineasEspeciales(oDocEntrada, oDocDestino)

                            Else

                                'Traspaso parcial, búsqueda de línea
                                Dim oLinea As EntDocumentoLin = getLinCopiaParcial(objDocumento, oDocEntrada)

                                If Not oLinea Is Nothing Then

                                    oDocDestino.Lines.SetCurrentLine(oDocDestino.Lines.Count - 1)

                                    oDocDestino.Lines.BaseEntry = oDocEntrada.DocEntry
                                    oDocDestino.Lines.BaseLine = oDocEntrada.Lines.LineNum
                                    oDocDestino.Lines.BaseType = oDocEntrada.DocObjectCode

                                    'Cantidad a traspasar en la línea
                                    If oDocEntrada.Lines.RemainingOpenQuantity > oLinea.Cantidad Then
                                        oDocDestino.Lines.Quantity = oLinea.Cantidad
                                    Else
                                        oDocDestino.Lines.Quantity = oDocEntrada.Lines.RemainingOpenQuantity
                                    End If

                                    'Lote
                                    setLotesLinea(oLinea, oDocDestino)

                                    'Actualiza la cantidad en la lista
                                    oLinea.Cantidad = Math.Round(oLinea.Cantidad - oDocDestino.Lines.Quantity, 6)

                                    'Portes de la línea
                                    setPortesLinea(oDocEntrada, oDocDestino)

                                    'Campos de usuario
                                    setCamposUsuarioLinea(oLinea, oDocDestino)

                                    oDocDestino.Lines.Add()

                                End If

                            End If

                        End If

                    Next

                Next

                'No continúa si no quedan líneas por traspasar
                ArticuloNoTraspasado = getLinArticuloSinTraspasar(objDocumento.DocType, CopiaParcial, objDocumento.Lineas)
                If CopiaParcial AndAlso String.IsNullOrEmpty(ArticuloNoTraspasado) Then Exit For

            Next

            'Comprueba que se hayan traspasado todas las líneas
            If CopiaParcial AndAlso Not String.IsNullOrEmpty(ArticuloNoTraspasado) Then _
                Throw New Exception("No existe el artículo/concepto " & ArticuloNoTraspasado & " o su cantidad es insuficiente en los documentos origen")

            'Fecha contable
            oDocDestino.DocDate = Date.ParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Fecha vencimiento
            If DateTime.TryParseExact(objDocumento.DocDueDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                oDocDestino.DocDueDate = Date.ParseExact(objDocumento.DocDueDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Fecha documento 
            'Comprueba si el periodo contable está abierto para la fecha DocDate de entrada. Si no lo está, coge la fecha auxiliar DocDateAux
            Dim PeriodoAbierto As Boolean = getPeriodoAbiertoDeTaxDate(objDocumento.DocDate.ToString, Sociedad)
            If PeriodoAbierto Then
                oDocDestino.TaxDate = Date.ParseExact(objDocumento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            Else
                oDocDestino.TaxDate = Date.ParseExact(objDocumento.DocDateAux.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            End If

            'Cuando se crea un documento de factura de copiar, el valor de número de referencia debe ser el proporcionado por Docuware 
            oDocDestino.NumAtCard = objDocumento.NumAtCard

            'Comentarios (si no vienen rellenos SAP pone los de por defecto)
            If objDocumento.Comments.Length > 254 Then objDocumento.Comments = objDocumento.Comments.Substring(0, 254)
            If Not String.IsNullOrEmpty(objDocumento.Comments) Then oDocDestino.Comments = objDocumento.Comments

            'Entrada diario
            If objDocumento.JournalMemo.Length > 254 Then objDocumento.JournalMemo = objDocumento.JournalMemo.Substring(0, 254)
            If Not String.IsNullOrEmpty(objDocumento.JournalMemo) Then oDocDestino.JournalMemo = objDocumento.JournalMemo

            'Sucursal
            If objDocumento.Sucursal > 0 Then oDocDestino.BPL_IDAssignedToInvoice = objDocumento.Sucursal

            'Proyecto
            If Not String.IsNullOrEmpty(objDocumento.Proyecto) Then oDocDestino.Project = objDocumento.Proyecto

            'Clase expedicion
            If objDocumento.ClaseExpedicion > 0 Then oDocDestino.TransportationCode = objDocumento.ClaseExpedicion

            'Responsable
            If Not String.IsNullOrEmpty(objDocumento.Responsable) Then
                oDocDestino.AgentCode = objDocumento.Responsable
            ElseIf String.IsNullOrEmpty(oDocDestino.AgentCode) AndAlso Not String.IsNullOrEmpty(ICResponsable) Then
                oDocDestino.AgentCode = ICResponsable
            End If

            'Titular
            If objDocumento.Titular > 0 Then oDocDestino.DocumentsOwner = objDocumento.Titular

            'Empleado dpto compras/ventas
            If objDocumento.EmpDptoCompraVenta > 0 Then oDocDestino.SalesPersonCode = objDocumento.EmpDptoCompraVenta

            'Facturas anticipo
            setFacturasAnticipo(objDocumento, oDocDestino, CardCode, Sociedad)

            'Documentos relacionados
            setDocumentosRelacionados(objDocumento, oDocDestino, CardCode, sLogInfo, Sociedad)

            'Retenciones de impuestos de IC (DocTotal sin retenciones)
            setRetencionesImpuestos(objDocumento, oDocDestino, CardCode, sLogInfo, Sociedad)

            'Total de documento 
            If objDocumento.DocTotal <> 0 Then oDocDestino.DocTotal = objDocumento.DocTotal

            'ID DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIDDW").Value = objDocumento.DOCIDDW

            'URL DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIURLDW").Value = objDocumento.DOCURLDW

            'IMPORTE DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIMPDW").Value = objDocumento.DocTotal

            'REVISAR IMPORTE DOCUWARE
            'El sistema permite traspasar un borrador a definitivo, aunque el importe del documento sea diferente que al importe DW
            oDocDestino.UserFields.Fields.Item("U_SEIREVDW").Value = IIf(Draft = Utilidades.Draft.Borrador, "S", "N")

            'Campos de usuario
            setCamposUsuarioCabecera(objDocumento, oDocDestino)

            'Añadimos documento
            If oDocDestino.Add() = 0 Then

                'Obtiene el DocEntry y DocNum del documento añadido
                Dim DocEntry As String = oCompany.GetNewObjectKey

                Dim DocNum As String = ""
                If Draft = Utilidades.Draft.Firme Then DocNum = oComun.getDocNumDeDocEntry(TablaDestino, DocEntry, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento copiado con éxito"
                retVal.MENSAJEAUX = DocNum

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
            oComun.LiberarObjCOM(oDocEntrada)
            oComun.LiberarObjCOM(oDocDestino)
        End Try

        Return retVal

    End Function

    Public Function ModificarDocumentoLineas(ByRef oCompany As Company, ByVal objCabecera As EntDocumentoCab, ByVal objLineas As List(Of EntDocumentoLin),
                                             ByVal TablaOrigen As String, ByVal DocEntryOrigen As String, ByVal sLogInfo As String,
                                             ByVal Sociedad As eSociedad) As EntResultado

        'Modificar documento Grupo La Puente
        Dim retVal As New EntResultado

        Dim oDocEntrada As Documents = Nothing

        Dim oComun As New clsComun

        Try

            'Objeto origen
            oDocEntrada = oCompany.GetBusinessObject(objCabecera.ObjTypeOrigen)
            If Not oDocEntrada.GetByKey(DocEntryOrigen) Then Throw New Exception("No puedo recuperar el documento origen con referencia: " & objCabecera.RefOrigen & " y DocEntry:" & DocEntryOrigen)

            'Recorre cada línea enviada
            For Each objLinea In objLineas

                'Comprueba que exista la línea
                If objLinea.LineNum < 0 Then Throw New Exception("Número de línea no suministrado")

                'Comprueba que haya cantidad 
                If Not objLinea.Cantidad > 0 Then Throw New Exception("Cantidad no suministrada para línea " & objLinea.LineNum)

                'Linea a insertar
                Dim bLinea As Boolean = False
                Dim objLineaNueva As New EntDocumentoLin

                'Recorre cada línea del documento
                For iLinea As Integer = 0 To oDocEntrada.Lines.Count - 1

                    oDocEntrada.Lines.SetCurrentLine(iLinea)

                    If (oDocEntrada.Lines.LineNum = objLinea.LineNum OrElse oDocEntrada.Lines.VisualOrder = objLinea.VisOrder) _
                        AndAlso oDocEntrada.Lines.LineStatus <> BoStatus.bost_Close Then

                        'Existe línea
                        bLinea = True

                        'Si la cantidad es mayor, actualiza la cantidad
                        If objLinea.Cantidad > oDocEntrada.Lines.Quantity Then

                            'Cantidad
                            oDocEntrada.Lines.Quantity = objLinea.Cantidad

                            'Lotes
                            If oDocEntrada.Lines.BatchNumbers.Count > 0 Then
                                oDocEntrada.Lines.BatchNumbers.SetCurrentLine(oDocEntrada.Lines.BatchNumbers.Count - 1)
                                If Not String.IsNullOrEmpty(oDocEntrada.Lines.BatchNumbers.BatchNumber) Then oDocEntrada.Lines.BatchNumbers.Quantity = objLinea.Cantidad
                            End If

                            Exit For

                        End If

                        'Si la cantidad es menor, modifica la cantidad y añade una línea con el restante
                        If objLinea.Cantidad < oDocEntrada.Lines.Quantity Then

                            'Nueva línea
                            objLineaNueva = New EntDocumentoLin With {.Articulo = oDocEntrada.Lines.ItemCode,
                                                                      .Concepto = oDocEntrada.Lines.ItemDescription,
                                                                      .Cantidad = oDocEntrada.Lines.Quantity - objLinea.Cantidad,
                                                                      .PrecioUnidad = oDocEntrada.Lines.Price,
                                                                      .PorcentajeDescuento = oDocEntrada.Lines.DiscountPercent,
                                                                      .WTLiable = oDocEntrada.Lines.WTLiable,
                                                                      .TaxOnly = oDocEntrada.Lines.TaxOnly,
                                                                      .TaxCode = oDocEntrada.Lines.TaxCode,
                                                                      .VATGroup = oDocEntrada.Lines.VatGroup,
                                                                      .Proyecto = oDocEntrada.Lines.ProjectCode,
                                                                      .Centro1Coste = oDocEntrada.Lines.COGSCostingCode,
                                                                      .Centro2Coste = oDocEntrada.Lines.COGSCostingCode,
                                                                      .Centro3Coste = oDocEntrada.Lines.COGSCostingCode,
                                                                      .Centro4Coste = oDocEntrada.Lines.COGSCostingCode,
                                                                      .Centro5Coste = oDocEntrada.Lines.COGSCostingCode,
                                                                      .Lotes = New List(Of EntDocumentoLote)}

                            'Lote nueva línea
                            If oDocEntrada.Lines.BatchNumbers.Count > 0 Then

                                oDocEntrada.Lines.BatchNumbers.SetCurrentLine(oDocEntrada.Lines.BatchNumbers.Count - 1)

                                If Not String.IsNullOrEmpty(oDocEntrada.Lines.BatchNumbers.BatchNumber) Then

                                    Dim objLote As New EntDocumentoLote With {.NumLote = oDocEntrada.Lines.BatchNumbers.BatchNumber,
                                                                              .Cantidad = oDocEntrada.Lines.Quantity - objLinea.Cantidad}

                                    objLineaNueva.Lotes.Add(objLote)

                                End If

                            End If

                            'Cantidad
                            oDocEntrada.Lines.Quantity = objLinea.Cantidad

                            'Añade la línea
                            If Not objLineaNueva Is Nothing AndAlso Not String.IsNullOrEmpty(objLineaNueva.Articulo) Then

                                oDocEntrada.Lines.SetCurrentLine(oDocEntrada.Lines.Count - 1)

                                If Not String.IsNullOrEmpty(oDocEntrada.Lines.ItemCode) Then
                                    oDocEntrada.Lines.Add()
                                    oDocEntrada.Lines.SetCurrentLine(oDocEntrada.Lines.Count - 1)
                                End If

                                With objLineaNueva

                                    oDocEntrada.Lines.ItemCode = .Articulo
                                    oDocEntrada.Lines.ItemDescription = .Concepto

                                    oDocEntrada.Lines.Quantity = .Cantidad
                                    oDocEntrada.Lines.Price = .PrecioUnidad
                                    oDocEntrada.Lines.DiscountPercent = .PorcentajeDescuento

                                    oDocEntrada.Lines.WTLiable = .WTLiable
                                    oDocEntrada.Lines.TaxOnly = .TaxOnly
                                    oDocEntrada.Lines.TaxCode = .TaxCode
                                    oDocEntrada.Lines.VatGroup = .VATGroup

                                    oDocEntrada.Lines.ProjectCode = .Proyecto

                                    oDocEntrada.Lines.COGSCostingCode = .Centro1Coste
                                    oDocEntrada.Lines.COGSCostingCode = .Centro2Coste
                                    oDocEntrada.Lines.COGSCostingCode = .Centro3Coste
                                    oDocEntrada.Lines.COGSCostingCode = .Centro4Coste
                                    oDocEntrada.Lines.COGSCostingCode = .Centro5Coste

                                    'Lotes
                                    If .Lotes.Count > 0 Then
                                        oDocEntrada.Lines.BatchNumbers.SetCurrentLine(oDocEntrada.Lines.BatchNumbers.Count - 1)
                                        oDocEntrada.Lines.BatchNumbers.BatchNumber = .Lotes(0).NumLote
                                        oDocEntrada.Lines.BatchNumbers.Quantity = .Lotes(0).Cantidad
                                        oDocEntrada.Lines.BatchNumbers.Add()
                                    End If

                                    oDocEntrada.Lines.Add()

                                End With

                            End If

                            Exit For

                        End If

                    End If

                Next

                'Comprueba que haya cantidad 
                If Not bLinea Then Throw New Exception("No se encuentra la línea " & objLinea.LineNum & "/" & objLinea.VisOrder & " o su estado es cerrado")

            Next

            'Añadimos documento
            If oDocEntrada.Update() = 0 Then

                'Obtiene el DocEntry y DocNum del documento añadido
                Dim DocEntry As String = DocEntryOrigen

                Dim DocNum As String = ""
                DocNum = oComun.getDocNumDeDocEntry(TablaOrigen, DocEntry, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento modificado con éxito"
                retVal.MENSAJEAUX = DocNum

                clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX)

            Else

                retVal.CODIGO = Respuesta.Ko
                retVal.MENSAJE = oCompany.GetLastErrorCode & "#" & oCompany.GetLastErrorDescription
                retVal.MENSAJEAUX = ""

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            oComun.LiberarObjCOM(oDocEntrada)
        End Try

        Return retVal

    End Function

    Public Function CopiaParcialDeDocumento(ByRef oCompany As Company, ByVal objDocumento As EntDocumentoCab, ByVal TablaOrigen As String,
                                            ByVal CardCode As String, ByVal sLogInfo As String, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oDocEntrada As Documents = Nothing
        Dim oDocDestino As Documents = Nothing

        Dim oComun As New clsComun

        Try

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objDocumento.ObjTypeDestino)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objDocumento.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Buscar contacto IC
            Dim ICContacto As Integer = getContactoDeEmail(CardCode, objDocumento.ContactoEmail, Sociedad)
            If ICContacto > 0 Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado contacto IC: " & ICContacto)

            'Buscar responsable IC
            Dim ICResponsable As String = ""
            If getEsDocumentoVenta(objDocumento.ObjTypeDestino) Then ICResponsable = getResponsableDeCardCode(CardCode, Sociedad)
            If Not String.IsNullOrEmpty(ICResponsable) Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado responsable IC: " & ICResponsable)

            'Comprueba si el documento se debe generar como borrador o en firme 
            Dim Draft As String = objDocumento.Draft
            If Draft = Utilidades.Draft.Interlocutor Then
                Draft = getDraftDeCardCode(CardCode, Sociedad)
                clsLog.Log.Info("(" & sLogInfo & ") Documento borrador por IC")
            End If

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objDocumento.RefOrigen.Split("#")

            'Control importes
            getControlImportes(objDocumento, CardCode, TablaOrigen, RefsOrigen, Draft, sLogInfo, Sociedad)

            'Objeto documento
            If Draft = Utilidades.Draft.Borrador Then
                oDocDestino = oCompany.GetBusinessObject(BoObjectTypes.oDrafts)
                oDocDestino.DocObjectCode = objDocumento.ObjTypeDestino
                clsLog.Log.Info("(" & sLogInfo & ") Documento borrador")
            Else
                oDocDestino = oCompany.GetBusinessObject(objDocumento.ObjTypeDestino)
                clsLog.Log.Info("(" & sLogInfo & ") Documento en firme")
            End If

            'Serie
            If objDocumento.Serie > 0 Then oDocDestino.Series = objDocumento.Serie

            'Contacto
            If objDocumento.ContactoCodigo > 0 Then
                oDocDestino.ContactPersonCode = objDocumento.ContactoCodigo
            ElseIf ICContacto > 0 Then
                oDocDestino.ContactPersonCode = ICContacto
            End If

            'Direccion envío
            If Not String.IsNullOrEmpty(objDocumento.DireccionEnvioCodigo) Then oDocDestino.ShipToCode = objDocumento.DireccionEnvioCodigo
            If Not String.IsNullOrEmpty(objDocumento.DireccionEnvioDetalle) Then oDocDestino.Address2 = objDocumento.DireccionEnvioDetalle

            'Direccion facturación
            If Not String.IsNullOrEmpty(objDocumento.DireccionFacturaCodigo) Then oDocDestino.PayToCode = objDocumento.DireccionFacturaCodigo
            If Not String.IsNullOrEmpty(objDocumento.DireccionFacturaDetalle) Then oDocDestino.Address = objDocumento.DireccionFacturaDetalle

            'Proyecto
            If Not String.IsNullOrEmpty(objDocumento.Proyecto) Then oDocDestino.Project = objDocumento.Proyecto

            'Bloqueo de pago
            If objDocumento.BloqueoPago > 0 Then
                oDocDestino.PaymentBlock = BoYesNoEnum.tYES
                oDocDestino.PaymentBlockEntry = objDocumento.BloqueoPago
            End If

            'Portes 
            setPortesCabecera(objDocumento, oDocDestino, CardCode, TablaOrigen, RefsOrigen, sLogInfo, Sociedad)

            'Factura de reserva
            If objDocumento.Reserva = SN.Si Then
                oDocDestino.ReserveInvoice = BoYesNoEnum.tYES
                clsLog.Log.Info("(" & sLogInfo & ") Factura de reserva")
            End If

            'Comprueba si la copia es parcial o no
            Dim CopiaParcial As Boolean = getEsCopiaParcial(objDocumento.Lineas)

            'Comprobaciones parciales
            If CopiaParcial Then

                clsLog.Log.Info("(" & sLogInfo & ") Copia parcial")

                'Comprueba que no lleguen líneas sin artículo/descripción en la copia parcial
                Dim bSinIdentificar As Boolean = getLinSinIdentificar(objDocumento)
                If bSinIdentificar Then Throw New Exception("Rellene el artículo/concepto a traspasar en todas las líneas")

                'Comprueba que no lleguen líneas a 0 en la copia parcial
                Dim bSinCantidad As Boolean = (From p In objDocumento.Lineas Where p.Cantidad <= 0).Distinct.ToList.Count > 0
                If bSinCantidad Then Throw New Exception("Rellene la cantidad a traspasar en todas las líneas")

                'Comprueba que el código de artículo sea ItemCode y no la referencia de proveedor
                For Each objLinea In objDocumento.Lineas

                    'Índice
                    Dim indLinea As Integer = objDocumento.Lineas.IndexOf(objLinea) + 1

                    If Not String.IsNullOrEmpty(objLinea.Articulo) OrElse Not String.IsNullOrEmpty(objLinea.RefExt) Then
                        'Artículo (si no se encuentra el código se busca por la referencia de proveedor)
                        Dim ItemCode As String = oComun.getItemCode(CardCode, objLinea.Articulo, objLinea.RefExt, Sociedad)
                        If objLinea.Articulo = ItemCode Then Continue For

                        If String.IsNullOrEmpty(ItemCode) Then Throw New Exception("No encuentro artículo de la línea " & indLinea.ToString & " con código: " & objLinea.Articulo & " o referencia externa: " & objLinea.RefExt & " para IC: " & CardCode)
                        clsLog.Log.Info("(" & sLogInfo & ") Encontrado artículo: " & ItemCode)

                        objLinea.Articulo = ItemCode
                    End If

                Next

            End If

            Dim ArticuloNoTraspasado As String = ""

            'Recorre cada uno de los documento origen
            For Each RefOrigen In RefsOrigen

                'Comprueba que existe el documento origen y que se puede acceder a él
                'Pueden ser varios documentos origen con la misma referencia
                Dim DocEntrysOrigen As List(Of String) = getDocEntryDeRefOrigen(TablaOrigen, RefOrigen, objDocumento.TipoRefOrigen, CardCode, True, True, Sociedad)
                If DocEntrysOrigen Is Nothing OrElse DocEntrysOrigen.Count = 0 Then Throw New Exception("No existe documento origen con referencia: " & RefOrigen & " o su estado es cerrado/cancelado")

                For Each DocEntryOrigen In DocEntrysOrigen

                    'Objeto origen
                    oDocEntrada = oCompany.GetBusinessObject(objDocumento.ObjTypeOrigen)

                    If Not oDocEntrada.GetByKey(DocEntryOrigen) Then Throw New Exception("No puedo recuperar el documento origen con referencia: " & RefOrigen & " y DocEntry:" & DocEntryOrigen)

                    'Copia los datos del documento origen al documento destino
                    oDocDestino.DocType = oDocEntrada.DocType
                    oDocDestino.CardCode = oDocEntrada.CardCode
                    oDocDestino.CardName = oDocEntrada.CardName

                    'Calcula la serie
                    Dim Serie As String = oComun.getSerieDeDocumentoDestino(oDocEntrada.Series, oDocEntrada.DocObjectCode, objDocumento.ObjTypeDestino, Sociedad)
                    If Not oDocDestino.Series > 0 AndAlso Not String.IsNullOrEmpty(Serie) Then oDocDestino.Series = CInt(Serie)

                    'Sucursal
                    If objDocumento.Sucursal > 0 AndAlso oDocEntrada.BPL_IDAssignedToInvoice > 0 AndAlso objDocumento.Sucursal <> oDocEntrada.BPL_IDAssignedToInvoice Then _
                        Throw New Exception("Sucursal destino: " & objDocumento.Sucursal & " no coincidente con sucursal origen: " & oDocEntrada.BPL_IDAssignedToInvoice)

                    If oDocEntrada.BPL_IDAssignedToInvoice > 0 AndAlso oDocDestino.BPL_IDAssignedToInvoice > 0 AndAlso oDocEntrada.BPL_IDAssignedToInvoice <> oDocDestino.BPL_IDAssignedToInvoice Then _
                        Throw New Exception("Sucursal origen: " & objDocumento.Sucursal & " no coincidente con sucursal origen: " & oDocEntrada.BPL_IDAssignedToInvoice)

                    If oDocEntrada.BPL_IDAssignedToInvoice > 0 Then oDocDestino.BPL_IDAssignedToInvoice = oDocEntrada.BPL_IDAssignedToInvoice

                    'Proyecto
                    If String.IsNullOrEmpty(oDocDestino.Project) AndAlso Not String.IsNullOrEmpty(oDocEntrada.Project) Then oDocDestino.Project = oDocEntrada.Project

                    'FinanzasRazon
                    If Not String.IsNullOrEmpty(objDocumento.FinanzasRazon) Then oDocDestino.CardName = objDocumento.FinanzasRazon

                    'FinanzasNIF
                    If Not String.IsNullOrEmpty(objDocumento.FinanzasNIF) Then oDocDestino.FederalTaxID = objDocumento.FinanzasNIF

                    'Responsable
                    If String.IsNullOrEmpty(oDocDestino.AgentCode) AndAlso Not String.IsNullOrEmpty(oDocEntrada.AgentCode) Then oDocDestino.AgentCode = oDocEntrada.AgentCode

                    'Titular
                    If Not oDocDestino.DocumentsOwner > 0 AndAlso oDocEntrada.DocumentsOwner > 0 Then oDocDestino.DocumentsOwner = oDocEntrada.DocumentsOwner

                    'Traspasa la primera línea especial que va en la posición 0 de un documento
                    setPrimeraLineaEspecial(oDocEntrada, oDocDestino)

                    'Traspasa las líneas
                    For iLinea As Integer = 0 To oDocEntrada.Lines.Count - 1

                        oDocEntrada.Lines.SetCurrentLine(iLinea)

                        '20200911: Comprueba que la línea no esté cerrada
                        If oDocEntrada.Lines.LineStatus <> BoStatus.bost_Close Then

                            'Diferencia entre traspaso total o parcial
                            If Not CopiaParcial Then

                                'Traspasa todas las líneas
                                oDocDestino.Lines.SetCurrentLine(oDocDestino.Lines.Count - 1)

                                oDocDestino.Lines.BaseEntry = oDocEntrada.DocEntry
                                oDocDestino.Lines.BaseLine = oDocEntrada.Lines.LineNum
                                oDocDestino.Lines.BaseType = oDocEntrada.DocObjectCode

                                'Portes de la línea
                                setPortesLinea(oDocEntrada, oDocDestino)

                                oDocDestino.Lines.Add()

                                'Traspasa las líneas especiales
                                setLineasEspeciales(oDocEntrada, oDocDestino)

                            Else

                                'Traspaso parcial, búsqueda de línea
                                Dim oLinea As EntDocumentoLin = getLinCopiaParcial(objDocumento, oDocEntrada)

                                If Not oLinea Is Nothing Then

                                    oDocDestino.Lines.SetCurrentLine(oDocDestino.Lines.Count - 1)

                                    oDocDestino.Lines.BaseEntry = oDocEntrada.DocEntry
                                    oDocDestino.Lines.BaseLine = oDocEntrada.Lines.LineNum
                                    oDocDestino.Lines.BaseType = oDocEntrada.DocObjectCode

                                    'Cantidad a traspasar en la línea
                                    If oDocEntrada.Lines.RemainingOpenQuantity > oLinea.Cantidad Then
                                        oDocDestino.Lines.Quantity = oLinea.Cantidad
                                    Else
                                        oDocDestino.Lines.Quantity = oDocEntrada.Lines.RemainingOpenQuantity
                                    End If

                                    'Lote
                                    setLotesLinea(oLinea, oDocDestino)

                                    'Actualiza la cantidad en la lista
                                    oLinea.Cantidad = Math.Round(oLinea.Cantidad - oDocDestino.Lines.Quantity, 6)

                                    'Portes de la línea
                                    setPortesLinea(oDocEntrada, oDocDestino)

                                    'Campos de usuario
                                    setCamposUsuarioLinea(oLinea, oDocDestino)


                                    oDocDestino.Lines.Add()

                                End If

                            End If

                        End If

                    Next

                Next

                'No continúa si no quedan líneas por traspasar
                ArticuloNoTraspasado = getLinArticuloSinTraspasar(objDocumento.DocType, CopiaParcial, objDocumento.Lineas)
                If CopiaParcial AndAlso String.IsNullOrEmpty(ArticuloNoTraspasado) Then Exit For

            Next

            'Comprueba que se hayan traspasado todas las líneas
            If CopiaParcial AndAlso Not String.IsNullOrEmpty(ArticuloNoTraspasado) Then _
                Throw New Exception("No existe el artículo/concepto " & ArticuloNoTraspasado & " o su cantidad es insuficiente en los documentos origen")

            'Fecha contable
            oDocDestino.DocDate = Date.ParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Fecha vencimiento
            If DateTime.TryParseExact(objDocumento.DocDueDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                oDocDestino.DocDueDate = Date.ParseExact(objDocumento.DocDueDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Fecha documento 
            'Comprueba si el periodo contable está abierto para la fecha DocDate de entrada. Si no lo está, coge la fecha auxiliar DocDateAux
            Dim PeriodoAbierto As Boolean = getPeriodoAbiertoDeTaxDate(objDocumento.DocDate.ToString, Sociedad)
            If PeriodoAbierto Then
                oDocDestino.TaxDate = Date.ParseExact(objDocumento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            Else
                oDocDestino.TaxDate = Date.ParseExact(objDocumento.DocDateAux.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            End If

            'Cuando se crea un documento de factura de copiar, el valor de número de referencia debe ser el proporcionado por Docuware 
            oDocDestino.NumAtCard = objDocumento.NumAtCard

            'Comentarios (si no vienen rellenos SAP pone los de por defecto)
            If objDocumento.Comments.Length > 254 Then objDocumento.Comments = objDocumento.Comments.Substring(0, 254)
            If Not String.IsNullOrEmpty(objDocumento.Comments) Then oDocDestino.Comments = objDocumento.Comments

            'Entrada diario
            If objDocumento.JournalMemo.Length > 254 Then objDocumento.JournalMemo = objDocumento.JournalMemo.Substring(0, 254)
            If Not String.IsNullOrEmpty(objDocumento.JournalMemo) Then oDocDestino.JournalMemo = objDocumento.JournalMemo

            'Sucursal
            If objDocumento.Sucursal > 0 Then oDocDestino.BPL_IDAssignedToInvoice = objDocumento.Sucursal

            'Proyecto
            If Not String.IsNullOrEmpty(objDocumento.Proyecto) Then oDocDestino.Project = objDocumento.Proyecto

            'Clase expedicion
            If objDocumento.ClaseExpedicion > 0 Then oDocDestino.TransportationCode = objDocumento.ClaseExpedicion

            'Responsable
            If Not String.IsNullOrEmpty(objDocumento.Responsable) Then
                oDocDestino.AgentCode = objDocumento.Responsable
            ElseIf String.IsNullOrEmpty(oDocDestino.AgentCode) AndAlso Not String.IsNullOrEmpty(ICResponsable) Then
                oDocDestino.AgentCode = ICResponsable
            End If

            'Titular
            If objDocumento.Titular > 0 Then oDocDestino.DocumentsOwner = objDocumento.Titular

            'Empleado dpto compras/ventas
            If objDocumento.EmpDptoCompraVenta > 0 Then oDocDestino.SalesPersonCode = objDocumento.EmpDptoCompraVenta

            'Facturas anticipo
            setFacturasAnticipo(objDocumento, oDocDestino, CardCode, Sociedad)

            'Retenciones de impuestos de IC (DocTotal sin retenciones)
            setRetencionesImpuestos(objDocumento, oDocDestino, CardCode, sLogInfo, Sociedad)

            'Total de documento 
            If objDocumento.DocTotal <> 0 Then oDocDestino.DocTotal = objDocumento.DocTotal

            'ID DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIDDW").Value = objDocumento.DOCIDDW

            'URL DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIURLDW").Value = objDocumento.DOCURLDW

            'IMPORTE DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIMPDW").Value = objDocumento.DocTotal

            'REVISAR IMPORTE DOCUWARE
            'El sistema permite traspasar un borrador a definitivo, aunque el importe del documento sea diferente que al importe DW
            oDocDestino.UserFields.Fields.Item("U_SEIREVDW").Value = IIf(Draft = Utilidades.Draft.Borrador, "S", "N")

            'Campos de usuario
            setCamposUsuarioCabecera(objDocumento, oDocDestino)

            'Añadimos documento
            If oDocDestino.Add() = 0 Then

                'Obtiene el DocEntry y DocNum del documento añadido
                Dim DocEntry As String = oCompany.GetNewObjectKey

                Dim DocNum As String = ""
                If Draft = Utilidades.Draft.Firme Then DocNum = oComun.getDocNumDeDocEntry(TablaDestino, DocEntry, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento copiado con éxito"
                retVal.MENSAJEAUX = DocNum

            Else

                retVal.CODIGO = Respuesta.Ko
                retVal.MENSAJE = oCompany.GetLastErrorCode & "#" & oCompany.GetLastErrorDescription
                retVal.MENSAJEAUX = ""

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & System.Reflection.MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        Finally
            oComun.LiberarObjCOM(oDocEntrada)
            oComun.LiberarObjCOM(oDocDestino)
        End Try

        Return retVal

    End Function

    Private Function NuevoDocumentoReciboProduccion(ByVal objDocumento As EntDocumentoCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oDocEntrada As ProductionOrders = Nothing
        Dim oDocDestino As Documents = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUEVO RECIBO PRODUCCION"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de documento para " & objDocumento.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objDocumento.UserSAP, objDocumento.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objDocumento.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocumento.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")
            If String.IsNullOrEmpty(objDocumento.RefOrigen) Then Throw New Exception("Referencia origen no suministrada")

            If Not DateTime.TryParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha contable no suministrada o incorrecta")

            If objDocumento.Lineas Is Nothing OrElse objDocumento.Lineas.Count = 0 Then Throw New Exception("Líneas no suministradas")

            'Buscamos tabla origen por ObjectType
            Dim TablaOrigen As String = getTablaDeObjType(objDocumento.ObjTypeOrigen)
            If String.IsNullOrEmpty(TablaOrigen) Then Throw New Exception("No encuentro tabla origen con ObjectType: " & objDocumento.ObjTypeOrigen)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla origen: " & TablaOrigen)

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objDocumento.ObjTypeDestino)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objDocumento.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objDocumento.NIFTercero, objDocumento.RazonSocial, objDocumento.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumento.NIFTercero & ", Razón social: " & objDocumento.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Comprueba si el documento se debe generar como borrador o en firme 
            Dim Draft As String = objDocumento.Draft
            If Draft = Utilidades.Draft.Interlocutor Then
                Draft = getDraftDeCardCode(CardCode, Sociedad)
                clsLog.Log.Info("(" & sLogInfo & ") Documento borrador por IC")
            End If

            'Objeto documento
            If Draft = Utilidades.Draft.Borrador Then
                oDocDestino = oCompany.GetBusinessObject(BoObjectTypes.oDrafts)
                oDocDestino.DocObjectCode = objDocumento.ObjTypeDestino
                clsLog.Log.Info("(" & sLogInfo & ") Documento borrador")
            Else
                oDocDestino = oCompany.GetBusinessObject(objDocumento.ObjTypeDestino)
                clsLog.Log.Info("(" & sLogInfo & ") Documento en firme")
            End If

            'Comprueba que existe el documento origen y que se puede acceder a él
            Dim DocEntryOrigen As String = getDocEntryDeRefOrigenUnica(TablaOrigen, objDocumento.RefOrigen, objDocumento.TipoRefOrigen, CardCode, True, Sociedad)
            If String.IsNullOrEmpty(DocEntryOrigen) Then Throw New Exception("No existe documento origen con referencia: " & objDocumento.RefOrigen & " o su estado es cerrado/cancelado")

            'Objeto origen
            oDocEntrada = oCompany.GetBusinessObject(objDocumento.ObjTypeOrigen)
            If Not oDocEntrada.GetByKey(DocEntryOrigen) Then Throw New Exception("No puedo recuperar el documento origen con referencia: " & objDocumento.RefOrigen & " y DocEntry:" & DocEntryOrigen)

            'Fecha contable 
            oDocDestino.DocDate = Date.ParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            oDocDestino.Lines.SetCurrentLine(oDocDestino.Lines.Count - 1)
            oDocDestino.Lines.BaseType = objDocumento.ObjTypeOrigen
            oDocDestino.Lines.BaseEntry = DocEntryOrigen

            If objDocumento.Lineas(0).CantidadOrigen = SN.Si Then
                Dim Cantidad As Double = objDocumento.Lineas(0).Cantidad - getCantidadCompletadaOrden(DocEntryOrigen, Sociedad)
                If Cantidad <= 0 Then Throw New Exception("La cantidad del documento no puede ser cero o negativa")
                oDocDestino.Lines.Quantity = Cantidad
            Else
                oDocDestino.Lines.Quantity = objDocumento.Lineas(0).Cantidad
            End If

            'Cuando se crea un documento de factura de copiar una entrada de mercancía, el valor de número de referencia debe ser el proporcionado por Docuware 
            oDocDestino.Reference2 = objDocumento.NumAtCard

            'Total de documento 
            If objDocumento.DocTotal <> 0 Then oDocDestino.DocTotal = objDocumento.DocTotal

            'ID DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIDDW").Value = objDocumento.DOCIDDW

            'URL DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIURLDW").Value = objDocumento.DOCURLDW

            'IMPORTE DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIMPDW").Value = objDocumento.DocTotal

            'REVISAR IMPORTE DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIREVDW").Value = IIf(Draft = Utilidades.Draft.Borrador, "S", "N")

            'Campos de usuario
            setCamposUsuarioCabecera(objDocumento, oDocDestino)

            'Añadimos documento
            If oDocDestino.Add() = 0 Then

                'Obtiene el DocEntry y DocNum del documento añadido
                Dim DocEntry As String = oCompany.GetNewObjectKey

                Dim DocNum As String = ""
                If Draft = Utilidades.Draft.Firme Then DocNum = oComun.getDocNumDeDocEntry(TablaDestino, DocEntry, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento copiado de ordenes de fabricación con éxito"
                retVal.MENSAJEAUX = DocNum

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
            oComun.LiberarObjCOM(oDocEntrada)
            oComun.LiberarObjCOM(oDocDestino)
        End Try

        Return retVal

    End Function

    Private Function NuevoDocumentoSolicitud(ByVal objDocumento As EntDocumentoCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oDocDestino As Documents = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUEVA SOLICITUD"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de documento solicitud para " & objDocumento.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objDocumento.UserSAP, objDocumento.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objDocumento.NIFTercero) Then Throw New Exception("NIF no suministrado")

            If Not DateTime.TryParseExact(objDocumento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha documento no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objDocumento.DocDateAux.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha documento auxiliar no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha contable no suministrada o incorrecta")

            If String.IsNullOrEmpty(objDocumento.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocumento.DOCURLDW) Then Throw New Exception("URL Docuware no suministrado")

            If objDocumento.Lineas Is Nothing Then Throw New Exception("Líneas no suministradas")

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objDocumento.ObjTypeDestino)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType:  " & objDocumento.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objDocumento.NIFTercero, objDocumento.RazonSocial, objDocumento.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumento.NIFTercero & ", Razón social: " & objDocumento.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Buscar cuenta contable IC
            Dim CuentaContable As String = getCuentaContableDeCardCode(CardCode, Sociedad)
            If Not String.IsNullOrEmpty(CuentaContable) Then clsLog.Log.Info("(" & sLogInfo & ") Encontrada cuenta contable IC: " & CuentaContable)

            'Buscar status impuesto IC
            Dim ICStatusImpuesto As String = getStatusImpuestoDeCardCode(CardCode, Sociedad)
            If Not String.IsNullOrEmpty(ICStatusImpuesto) Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado status impuesto IC: " & ICStatusImpuesto)

            'Buscar grupo IVA IC
            Dim ICGrupoIVA As String = getGrupoIVADeCardCode(CardCode, Sociedad)
            If Not String.IsNullOrEmpty(ICGrupoIVA) Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado grupo IVA IC: " & ICGrupoIVA)

            'Buscar responsable IC
            Dim ICResponsable As String = ""
            If getEsDocumentoVenta(objDocumento.ObjTypeDestino) Then ICResponsable = getResponsableDeCardCode(CardCode, Sociedad)
            If Not String.IsNullOrEmpty(ICResponsable) Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado responsable IC: " & ICResponsable)

            'Comprueba si el documento se debe generar como borrador o en firme 
            Dim Draft As String = objDocumento.Draft
            If Draft = Utilidades.Draft.Interlocutor Then
                Draft = getDraftDeCardCode(CardCode, Sociedad)
                clsLog.Log.Info("(" & sLogInfo & ") Documento borrador por IC")
            End If

            'Objeto documento
            If Draft <> Utilidades.Draft.Firme Then
                oDocDestino = oCompany.GetBusinessObject(BoObjectTypes.oDrafts)
                oDocDestino.DocObjectCode = objDocumento.ObjTypeDestino
                clsLog.Log.Info("(" & sLogInfo & ") Documento borrador")
            Else
                oDocDestino = CType(oCompany.GetBusinessObject(objDocumento.ObjTypeDestino), Documents)
                clsLog.Log.Info("(" & sLogInfo & ") Documento en firme")
            End If

            'Moneda
            If Not String.IsNullOrEmpty(objDocumento.Currency) Then oDocDestino.DocCurrency = objDocumento.Currency

            'Sucursal
            If objDocumento.Sucursal > 0 Then oDocDestino.BPL_IDAssignedToInvoice = objDocumento.Sucursal

            'Proyecto
            If Not String.IsNullOrEmpty(objDocumento.Proyecto) Then oDocDestino.Project = objDocumento.Proyecto

            'Clase expedicion
            If objDocumento.ClaseExpedicion > 0 Then oDocDestino.TransportationCode = objDocumento.ClaseExpedicion

            'Responsable
            If Not String.IsNullOrEmpty(objDocumento.Responsable) Then
                oDocDestino.AgentCode = objDocumento.Responsable
            ElseIf Not String.IsNullOrEmpty(ICResponsable) Then
                oDocDestino.AgentCode = ICResponsable
            End If

            'Titular
            If objDocumento.Titular > 0 Then oDocDestino.DocumentsOwner = objDocumento.Titular

            'Empleado dpto compras/ventas
            If objDocumento.EmpDptoCompraVenta > 0 Then oDocDestino.SalesPersonCode = objDocumento.EmpDptoCompraVenta

            'Comentarios
            If objDocumento.Comments.Length > 254 Then objDocumento.Comments = objDocumento.Comments.Substring(0, 254)
            If Not String.IsNullOrEmpty(objDocumento.Comments) Then oDocDestino.Comments = objDocumento.Comments

            'Entrada diario
            If objDocumento.JournalMemo.Length > 254 Then objDocumento.JournalMemo = objDocumento.JournalMemo.Substring(0, 254)
            If Not String.IsNullOrEmpty(objDocumento.JournalMemo) Then oDocDestino.JournalMemo = objDocumento.JournalMemo

            'Solicitante
            If Not String.IsNullOrEmpty(objDocumento.Solicitante) Then
                oDocDestino.ReqType = 171
                oDocDestino.Requester = objDocumento.Solicitante
            End If

            If Not String.IsNullOrEmpty(objDocumento.Email) Then oDocDestino.RequesterEmail = objDocumento.Email

            'Fecha contable
            oDocDestino.DocDate = Date.ParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Fecha requerida
            oDocDestino.RequriedDate = Now.Date.AddYears(1)

            'Fecha documento 
            'Comprueba si el periodo contable está abierto para la fecha DocDate de entrada. Si no lo está, coge la fecha auxiliar DocDateAux
            Dim PeriodoAbierto As Boolean = getPeriodoAbiertoDeTaxDate(objDocumento.DocDate.ToString, Sociedad)
            If PeriodoAbierto Then
                oDocDestino.TaxDate = Date.ParseExact(objDocumento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            Else
                oDocDestino.TaxDate = Date.ParseExact(objDocumento.DocDateAux.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            End If

            'Tipo de documento. Por defecto, servicio
            If objDocumento.DocType = DocType.Articulo Then

                'Artículo
                oDocDestino.DocType = BoDocumentTypes.dDocument_Items

                'Recorre cada línea del servicio
                For Each objLinea In objDocumento.Lineas

                    'Índice
                    Dim indLinea As Integer = objDocumento.Lineas.IndexOf(objLinea) + 1

                    'Comprobar campos
                    If String.IsNullOrEmpty(objLinea.Articulo) AndAlso String.IsNullOrEmpty(objLinea.RefExt) Then _
                        Throw New Exception("Artículo o referencia externa de la línea " & indLinea.ToString & " no suministrado")

                    'Se posiciona en la línea
                    oDocDestino.Lines.SetCurrentLine(oDocDestino.Lines.Count - 1)

                    'Proveedor
                    If objDocumento.Ambito = Ambito.Compras Then oDocDestino.Lines.LineVendor = CardCode

                    'Fecha requerida
                    oDocDestino.Lines.RequiredDate = oDocDestino.RequriedDate

                    'Artículo (si no se encuentra el código se busca por la referencia de proveedor)
                    Dim ItemCode As String = oComun.getItemCode(CardCode, objLinea.Articulo, objLinea.RefExt, Sociedad)
                    If String.IsNullOrEmpty(ItemCode) Then Throw New Exception("No encuentro artículo de la línea " & indLinea.ToString & " con código: " & objLinea.Articulo & " o referencia externa: " & objLinea.RefExt & " para IC: " & CardCode)
                    clsLog.Log.Info("(" & sLogInfo & ") Encontrado artículo: " & ItemCode)

                    oDocDestino.Lines.ItemCode = ItemCode

                    'Descripción
                    If objLinea.Concepto.Length > 100 Then objLinea.Concepto = objLinea.Concepto.Substring(0, 100)
                    If Not String.IsNullOrEmpty(objLinea.Concepto) Then oDocDestino.Lines.ItemDescription = objLinea.Concepto

                    'Cantidad
                    If objLinea.Cantidad > 0 Then oDocDestino.Lines.Quantity = objLinea.Cantidad

                    'Precio
                    If objLinea.PrecioUnidad <> 0 Then
                        oDocDestino.Lines.UnitPrice = objLinea.PrecioUnidad
                        oDocDestino.Lines.Price = objLinea.PrecioUnidad
                    End If

                    'Porcentaje dto
                    If objLinea.PorcentajeDescuento <> 0 Then oDocDestino.Lines.DiscountPercent = objLinea.PorcentajeDescuento

                    'Moneda
                    If Not String.IsNullOrEmpty(objDocumento.Currency) Then oDocDestino.Lines.Currency = objDocumento.Currency

                    'TaxOnly
                    oDocDestino.Lines.TaxOnly = BoYesNoEnum.tNO
                    If Not String.IsNullOrEmpty(objLinea.TaxOnly) AndAlso objLinea.TaxOnly = SN.Si Then oDocDestino.Lines.TaxOnly = BoYesNoEnum.tYES

                    'TaxCode
                    If Not String.IsNullOrEmpty(objLinea.TaxCode) Then oDocDestino.Lines.TaxCode = objLinea.TaxCode

                    'WTLiable
                    oDocDestino.Lines.WTLiable = BoYesNoEnum.tNO
                    If Not String.IsNullOrEmpty(objLinea.WTLiable) AndAlso objLinea.WTLiable = SN.Si Then oDocDestino.Lines.WTLiable = BoYesNoEnum.tYES

                    'VATGroup
                    If Not String.IsNullOrEmpty(objLinea.VATGroup) Then oDocDestino.Lines.VatGroup = objLinea.VATGroup

                    'Fecha entrega
                    If DateTime.TryParseExact(objLinea.ShipDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        oDocDestino.Lines.ShipDate = Date.ParseExact(objLinea.ShipDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                    'Fecha requerida
                    If DateTime.TryParseExact(objLinea.RequiredDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        oDocDestino.Lines.RequiredDate = Date.ParseExact(objLinea.RequiredDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                    'Proyecto
                    If Not String.IsNullOrEmpty(objLinea.Proyecto) Then oDocDestino.Lines.ProjectCode = objLinea.Proyecto

                    'Almacen 
                    If Not String.IsNullOrEmpty(objLinea.Almacen) Then oDocDestino.Lines.WarehouseCode = objLinea.Almacen

                    'Centro coste 1
                    If Not String.IsNullOrEmpty(objLinea.Centro1Coste) Then oDocDestino.Lines.CostingCode = objLinea.Centro1Coste

                    'Centro coste 2
                    If Not String.IsNullOrEmpty(objLinea.Centro2Coste) Then oDocDestino.Lines.CostingCode2 = objLinea.Centro2Coste

                    'Centro coste 3
                    If Not String.IsNullOrEmpty(objLinea.Centro3Coste) Then oDocDestino.Lines.CostingCode3 = objLinea.Centro3Coste

                    'Centro coste 4
                    If Not String.IsNullOrEmpty(objLinea.Centro4Coste) Then oDocDestino.Lines.CostingCode4 = objLinea.Centro4Coste

                    'Centro coste 5
                    If Not String.IsNullOrEmpty(objLinea.Centro5Coste) Then oDocDestino.Lines.CostingCode5 = objLinea.Centro5Coste

                    'Campos de usuario
                    setCamposUsuarioLinea(objLinea, oDocDestino)

                    'Añade línea
                    oDocDestino.Lines.Add()

                Next

            Else

                'Servicio
                oDocDestino.DocType = BoDocumentTypes.dDocument_Service

                'Comprobar que hay al menos una línea informada
                If objDocumento.Lineas Is Nothing OrElse objDocumento.Lineas.Count = 0 Then Throw New Exception("Línea 1 no suministrada")

                'Recorre cada línea del servicio
                For Each objLinea In objDocumento.Lineas

                    'Índice
                    Dim indLinea As Integer = objDocumento.Lineas.IndexOf(objLinea) + 1

                    'Comprobar datos primera línea
                    If indLinea = 1 AndAlso String.IsNullOrEmpty(objLinea.Concepto) Then Throw New Exception("Concepto de la línea 1 no suministrado")

                    'Comprobar campos del resto de líneas
                    If indLinea > 1 AndAlso String.IsNullOrEmpty(objLinea.Concepto) AndAlso Not (objLinea.LineTotal = 0 AndAlso objLinea.BaseTotal = 0) Then _
                        Throw New Exception("Concepto de la línea " & indLinea.ToString & " no suministrado")

                    'Se posiciona en la línea
                    oDocDestino.Lines.SetCurrentLine(oDocDestino.Lines.Count - 1)

                    'Proveedor
                    If objDocumento.Ambito = Ambito.Compras Then oDocDestino.Lines.LineVendor = CardCode

                    'Fecha requerida
                    oDocDestino.Lines.RequiredDate = oDocDestino.RequriedDate

                    'Si la cuenta contable viene vacía, se coge la del IC
                    If Not String.IsNullOrEmpty(objLinea.CuentaContable) Then
                        oDocDestino.Lines.AccountCode = objLinea.CuentaContable
                    Else
                        If String.IsNullOrEmpty(CuentaContable) Then Throw New Exception("Cuenta contable no definida en la ficha del IC")
                        oDocDestino.Lines.AccountCode = CuentaContable
                    End If

                    'Concepto
                    If objLinea.Concepto.Length > 100 Then objLinea.Concepto = objLinea.Concepto.Substring(0, 100)
                    oDocDestino.Lines.ItemDescription = objLinea.Concepto

                    'Si LineTotal=0, LineTotal=BaseTotal
                    If objLinea.LineTotal = 0 Then objLinea.LineTotal = objLinea.BaseTotal
                    oDocDestino.Lines.LineTotal = objLinea.LineTotal

                    'Porcentaje dto
                    If objLinea.PorcentajeDescuento <> 0 Then oDocDestino.Lines.DiscountPercent = objLinea.PorcentajeDescuento

                    'TaxOnly
                    oDocDestino.Lines.TaxOnly = BoYesNoEnum.tNO
                    If Not String.IsNullOrEmpty(objLinea.TaxOnly) AndAlso objLinea.TaxOnly = SN.Si Then oDocDestino.Lines.TaxOnly = BoYesNoEnum.tYES

                    'TaxCode
                    If Not String.IsNullOrEmpty(objLinea.TaxCode) Then oDocDestino.Lines.TaxCode = objLinea.TaxCode

                    'WTLiable
                    oDocDestino.Lines.WTLiable = BoYesNoEnum.tNO
                    If Not String.IsNullOrEmpty(objLinea.WTLiable) AndAlso objLinea.WTLiable = SN.Si Then oDocDestino.Lines.WTLiable = BoYesNoEnum.tYES

                    'Grupo impositivo
                    If Not String.IsNullOrEmpty(objLinea.VATGroup) Then
                        oDocDestino.Lines.VatGroup = objLinea.VATGroup
                    ElseIf ICStatusImpuesto <> StatusImpuesto.Obligatorio AndAlso Not String.IsNullOrEmpty(ICGrupoIVA) Then
                        oDocDestino.Lines.VatGroup = ICGrupoIVA
                    Else
                        oDocDestino.Lines.VatGroup = getVatGroupDeVatTotal(objDocumento.Ambito, objLinea.LineTotal, objLinea.VATTotal, objLinea.VATPorc, objLinea.Intracomunitario, ICGrupoIVA, Sociedad)
                    End If

                    'Moneda
                    If Not String.IsNullOrEmpty(objDocumento.Currency) Then oDocDestino.Lines.Currency = objDocumento.Currency

                    'Fecha entrega
                    If DateTime.TryParseExact(objLinea.ShipDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        oDocDestino.Lines.ShipDate = Date.ParseExact(objLinea.ShipDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                    'Fecha requerida
                    If DateTime.TryParseExact(objLinea.RequiredDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        oDocDestino.Lines.RequiredDate = Date.ParseExact(objLinea.RequiredDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                    'Proyecto
                    If Not String.IsNullOrEmpty(objLinea.Proyecto) Then oDocDestino.Lines.ProjectCode = objLinea.Proyecto

                    'Centro coste 1
                    If Not String.IsNullOrEmpty(objLinea.Centro1Coste) Then oDocDestino.Lines.CostingCode = objLinea.Centro1Coste

                    'Centro coste 2
                    If Not String.IsNullOrEmpty(objLinea.Centro2Coste) Then oDocDestino.Lines.CostingCode2 = objLinea.Centro2Coste

                    'Centro coste 3
                    If Not String.IsNullOrEmpty(objLinea.Centro3Coste) Then oDocDestino.Lines.CostingCode3 = objLinea.Centro3Coste

                    'Centro coste 4
                    If Not String.IsNullOrEmpty(objLinea.Centro4Coste) Then oDocDestino.Lines.CostingCode4 = objLinea.Centro4Coste

                    'Centro coste 5
                    If Not String.IsNullOrEmpty(objLinea.Centro5Coste) Then oDocDestino.Lines.CostingCode5 = objLinea.Centro5Coste

                    'Campos de usuario
                    setCamposUsuarioLinea(objLinea, oDocDestino)

                    'Añade línea
                    oDocDestino.Lines.Add()

                Next

            End If

            'ID DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIDDW").Value = objDocumento.DOCIDDW

            'URL DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIURLDW").Value = objDocumento.DOCURLDW

            'IMPORTE DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIMPDW").Value = objDocumento.DocTotal

            'REVISAR IMPORTE DOCUWARE
            'MAIL 20200919 (David Sánchez): El sistema permite traspasar un borrador a definitivo, aunque el importe del documento sea diferente que al importe DW
            oDocDestino.UserFields.Fields.Item("U_SEIREVDW").Value = IIf(Draft = Utilidades.Draft.Borrador, "S", "N")

            'Campos de usuario
            setCamposUsuarioCabecera(objDocumento, oDocDestino)

            'Añadimos documento
            If oDocDestino.Add() = 0 Then

                'Obtiene el DocEntry y DocNum del documento añadido
                Dim DocEntry As String = oCompany.GetNewObjectKey

                Dim DocNum As String = ""
                If Draft = Utilidades.Draft.Firme Then DocNum = oComun.getDocNumDeDocEntry(TablaDestino, DocEntry, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento creado con éxito"
                retVal.MENSAJEAUX = DocNum

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
            oComun.LiberarObjCOM(oDocDestino)
        End Try

        Return retVal

    End Function

    Private Function NuevoDocumentoCobroPago(ByVal objDocumento As EntDocumentoCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oDocDestino As Payments = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUEVO COBRO/PAGO IC"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de documento para " & objDocumento.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objDocumento.UserSAP, objDocumento.PassSAP, Sociedad)

            If oCompany Is Nothing OrElse Not oCompany.Connected Then
                Throw New Exception("No se puede conectar a SAP")
            End If

            'Obligatorios
            If String.IsNullOrEmpty(objDocumento.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocumento.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")

            If Not DateTime.TryParseExact(objDocumento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha documento no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objDocumento.DocDateAux.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha documento auxiliar no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha contable no suministrada o incorrecta")

            If String.IsNullOrEmpty(objDocumento.CobroPagoImporte) Then Throw New Exception("Importe cobro/pago no suministrado")
            If String.IsNullOrEmpty(objDocumento.CobroPagoCuenta) Then Throw New Exception("Cuenta cobro/pago no suministrada")

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objDocumento.ObjTypeDestino)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objDocumento.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objDocumento.NIFTercero, objDocumento.RazonSocial, objDocumento.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumento.NIFTercero & ", Razón social: " & objDocumento.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Buscar responsable IC
            Dim ICResponsable As String = ""
            If getEsDocumentoVenta(objDocumento.ObjTypeDestino) Then ICResponsable = getResponsableDeCardCode(CardCode, Sociedad)
            If Not String.IsNullOrEmpty(ICResponsable) Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado responsable IC: " & ICResponsable)

            'Buscar contacto IC
            Dim ICContacto As Integer = getContactoDeEmail(CardCode, objDocumento.ContactoEmail, Sociedad)
            If ICContacto > 0 Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado contacto IC: " & ICContacto)

            'Comprueba si el documento se debe generar como borrador o en firme 
            Dim Draft As String = objDocumento.Draft
            If Draft = Utilidades.Draft.Interlocutor Then
                Draft = getDraftDeCardCode(CardCode, Sociedad)
                clsLog.Log.Info("(" & sLogInfo & ") Documento borrador por IC")
            End If

            'Objeto documento
            If Draft <> Utilidades.Draft.Firme Then
                oDocDestino = oCompany.GetBusinessObject(BoObjectTypes.oPaymentsDrafts)

                'Tipo de borrador
                If objDocumento.ObjTypeDestino = ObjType.Cobro Then
                    oDocDestino.DocObjectCode = BoPaymentsObjectType.bopot_IncomingPayments
                Else
                    oDocDestino.DocObjectCode = BoPaymentsObjectType.bopot_OutgoingPayments
                End If

                clsLog.Log.Info("(" & sLogInfo & ") Documento borrador")
            Else
                oDocDestino = oCompany.GetBusinessObject(objDocumento.ObjTypeDestino)
                clsLog.Log.Info("(" & sLogInfo & ") Documento en firme")
            End If

            'Tipo de pago
            If objDocumento.ObjTypeDestino = ObjType.Cobro Then
                oDocDestino.DocTypte = BoRcptTypes.rCustomer
            Else
                oDocDestino.DocTypte = BoRcptTypes.rSupplier
            End If

            'Serie
            If objDocumento.Serie > 0 Then oDocDestino.Series = objDocumento.Serie

            'IC
            oDocDestino.CardCode = CardCode

            'Contacto
            If objDocumento.ContactoCodigo > 0 Then
                oDocDestino.ContactPersonCode = objDocumento.ContactoCodigo
            ElseIf ICContacto > 0 Then
                oDocDestino.ContactPersonCode = ICContacto
            End If

            'Referencia
            If Not String.IsNullOrEmpty(objDocumento.NumAtCard) Then oDocDestino.CounterReference = objDocumento.NumAtCard

            'FinanzasRazon
            If Not String.IsNullOrEmpty(objDocumento.FinanzasRazon) Then oDocDestino.CardName = objDocumento.FinanzasRazon

            'Moneda
            If Not String.IsNullOrEmpty(objDocumento.Currency) Then oDocDestino.DocCurrency = objDocumento.Currency

            'Responsable
            If Not String.IsNullOrEmpty(objDocumento.Responsable) Then
                oDocDestino.BillOfExchangeAgent = objDocumento.Responsable
            ElseIf Not String.IsNullOrEmpty(ICResponsable) Then
                oDocDestino.BillOfExchangeAgent = ICResponsable
            End If

            'Comentarios
            If objDocumento.Comments.Length > 254 Then objDocumento.Comments = objDocumento.Comments.Substring(0, 254)
            If Not String.IsNullOrEmpty(objDocumento.Comments) Then oDocDestino.Remarks = objDocumento.Comments

            'Fecha contable 
            oDocDestino.DocDate = Date.ParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Fecha documento 
            'Comprueba si el periodo contable está abierto para la fecha DocDate de entrada. Si no lo está, coge la fecha auxiliar DocDateAux
            Dim PeriodoAbierto As Boolean = getPeriodoAbiertoDeTaxDate(objDocumento.DocDate.ToString, Sociedad)
            If PeriodoAbierto Then
                oDocDestino.TaxDate = Date.ParseExact(objDocumento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            Else
                oDocDestino.TaxDate = Date.ParseExact(objDocumento.DocDateAux.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            End If

            'Copia los datos del documento origen al documento destino
            oDocDestino.CardCode = CardCode

            'FinanzasRazon
            If Not String.IsNullOrEmpty(objDocumento.FinanzasRazon) Then oDocDestino.CardName = objDocumento.FinanzasRazon

            'FinanzasNIF
            If Not String.IsNullOrEmpty(objDocumento.FinanzasNIF) Then oDocDestino.FederalTaxID = objDocumento.FinanzasNIF

            'Medio pago
            setMedioCobroPago(objDocumento, oDocDestino, Sociedad)

            'Proyecto
            If Not String.IsNullOrEmpty(objDocumento.CobroPagoProyecto) Then oDocDestino.ProjectCode = objDocumento.CobroPagoProyecto

            'ID DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIDDW").Value = objDocumento.DOCIDDW

            'URL DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIURLDW").Value = objDocumento.DOCURLDW

            'IMPORTE DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIMPDW").Value = objDocumento.DocTotal

            'REVISAR IMPORTE DOCUWARE
            'El sistema permite traspasar un borrador a definitivo, aunque el importe del documento sea diferente que al importe DW
            oDocDestino.UserFields.Fields.Item("U_SEIREVDW").Value = IIf(Draft = Utilidades.Draft.Borrador, "S", "N")

            'Documentos relacionados
            setDocumentosRelacionadosCobroPago(objDocumento, oDocDestino, CardCode, sLogInfo, Sociedad)

            'Campos de usuario
            setCamposUsuarioCabeceraCobroPago(objDocumento, oDocDestino)

            'Añadimos documento
            If oDocDestino.Add() = 0 Then

                'Obtiene el DocEntry y DocNum del documento añadido
                Dim DocEntry As String = oCompany.GetNewObjectKey

                Dim DocNum As String = ""
                If Draft = Utilidades.Draft.Firme Then DocNum = oComun.getDocNumDeDocEntry(TablaDestino, DocEntry, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento creado con éxito"
                retVal.MENSAJEAUX = DocNum

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
            oComun.LiberarObjCOM(oDocDestino)
        End Try

        Return retVal

    End Function

    Private Function NuevoDocumentoCobroPagoACuenta(ByVal objDocumento As EntDocumentoCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oDocDestino As Payments = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUEVO COBRO/PAGO CUENTA"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de documento para " & objDocumento.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objDocumento.UserSAP, objDocumento.PassSAP, Sociedad)

            If oCompany Is Nothing OrElse Not oCompany.Connected Then
                Throw New Exception("No se puede conectar a SAP")
            End If

            'Obligatorios
            If String.IsNullOrEmpty(objDocumento.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocumento.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")

            If Not DateTime.TryParseExact(objDocumento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha documento no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objDocumento.DocDateAux.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha documento auxiliar no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha contable no suministrada o incorrecta")

            If String.IsNullOrEmpty(objDocumento.CobroPagoImporte) Then Throw New Exception("Importe cobro/pago no suministrado")
            If String.IsNullOrEmpty(objDocumento.CobroPagoCuenta) Then Throw New Exception("Cuenta cobro/pago no suministrada")

            If objDocumento.Lineas Is Nothing Then Throw New Exception("Líneas no suministradas")

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objDocumento.ObjTypeDestino)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objDocumento.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Objeto documento
            If objDocumento.Draft <> Utilidades.Draft.Firme Then
                oDocDestino = oCompany.GetBusinessObject(BoObjectTypes.oPaymentsDrafts)

                'Tipo de borrador
                If objDocumento.ObjTypeDestino = ObjType.Cobro Then
                    oDocDestino.DocObjectCode = BoPaymentsObjectType.bopot_IncomingPayments
                Else
                    oDocDestino.DocObjectCode = BoPaymentsObjectType.bopot_OutgoingPayments
                End If

                clsLog.Log.Info("(" & sLogInfo & ") Documento borrador")
            Else
                oDocDestino = oCompany.GetBusinessObject(objDocumento.ObjTypeDestino)
                clsLog.Log.Info("(" & sLogInfo & ") Documento en firme")
            End If

            'Tipo de pago a cuenta
            oDocDestino.DocTypte = BoRcptTypes.rAccount

            'Serie
            If objDocumento.Serie > 0 Then oDocDestino.Series = objDocumento.Serie

            'Razón social
            If Not String.IsNullOrEmpty(objDocumento.RazonSocial) Then oDocDestino.CardName = objDocumento.RazonSocial

            'Referencia
            If Not String.IsNullOrEmpty(objDocumento.NumAtCard) Then oDocDestino.CounterReference = objDocumento.NumAtCard

            'Moneda
            If Not String.IsNullOrEmpty(objDocumento.Currency) Then oDocDestino.DocCurrency = objDocumento.Currency

            'Comentarios
            If objDocumento.Comments.Length > 254 Then objDocumento.Comments = objDocumento.Comments.Substring(0, 254)
            If Not String.IsNullOrEmpty(objDocumento.Comments) Then oDocDestino.Remarks = objDocumento.Comments

            'Fecha contable 
            oDocDestino.DocDate = Date.ParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Fecha documento 
            'Comprueba si el periodo contable está abierto para la fecha DocDate de entrada. Si no lo está, coge la fecha auxiliar DocDateAux
            Dim PeriodoAbierto As Boolean = getPeriodoAbiertoDeTaxDate(objDocumento.DocDate.ToString, Sociedad)
            If PeriodoAbierto Then
                oDocDestino.TaxDate = Date.ParseExact(objDocumento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            Else
                oDocDestino.TaxDate = Date.ParseExact(objDocumento.DocDateAux.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            End If

            'FinanzasRazon
            If Not String.IsNullOrEmpty(objDocumento.FinanzasRazon) Then oDocDestino.CardName = objDocumento.FinanzasRazon

            'FinanzasNIF
            If Not String.IsNullOrEmpty(objDocumento.FinanzasNIF) Then oDocDestino.FederalTaxID = objDocumento.FinanzasNIF

            'Medio pago
            setMedioCobroPago(objDocumento, oDocDestino, Sociedad)

            'Proyecto
            If Not String.IsNullOrEmpty(objDocumento.CobroPagoProyecto) Then oDocDestino.ProjectCode = objDocumento.CobroPagoProyecto

            'LINEAS

            'Comprobar que hay al menos una línea informada
            If objDocumento.Lineas Is Nothing OrElse objDocumento.Lineas.Count = 0 Then Throw New Exception("Línea 1 no suministrada")

            'Recorre cada línea del servicio
            For Each objLinea In objDocumento.Lineas

                'Índice
                Dim indLinea As Integer = objDocumento.Lineas.IndexOf(objLinea) + 1

                'Comprobar datos primera línea
                If indLinea = 1 AndAlso String.IsNullOrEmpty(objLinea.CuentaContable) Then Throw New Exception("Cuenta contable de la línea " & indLinea.ToString & " no suministrada")

                'Comprobar campos del resto de líneas
                If indLinea > 1 AndAlso String.IsNullOrEmpty(objLinea.CuentaContable) AndAlso Not (objLinea.LineTotal = 0 AndAlso objLinea.BaseTotal = 0) Then _
                    Throw New Exception("Cuenta contable de la línea " & indLinea.ToString & " no suministrada")

                If objLinea.LineTotal = 0 AndAlso objLinea.BaseTotal = 0 Then Throw New Exception("Importes de la línea " & indLinea.ToString & " no suministrados")

                'Se posiciona en la línea
                oDocDestino.AccountPayments.SetCurrentLine(oDocDestino.AccountPayments.Count - 1)

                'Cuenta contable
                oDocDestino.AccountPayments.AccountCode = objLinea.CuentaContable

                'Descripción
                If objLinea.Concepto.Length > 100 Then objLinea.Concepto = objLinea.Concepto.Substring(0, 100)
                If Not String.IsNullOrEmpty(objLinea.Concepto) Then oDocDestino.AccountPayments.Decription = objLinea.Concepto

                'Si LineTotal=0, LineTotal=BaseTotal
                If objLinea.LineTotal = 0 Then objLinea.LineTotal = objLinea.BaseTotal
                oDocDestino.AccountPayments.SumPaid = IIf(objLinea.BaseTotal > 0, objLinea.BaseTotal, objLinea.LineTotal)

                'Grupo impositivo
                If Not String.IsNullOrEmpty(objLinea.VATGroup) Then
                    oDocDestino.AccountPayments.VatGroup = objLinea.VATGroup
                ElseIf objLinea.VATTotal > 0 OrElse objLinea.VATPorc > 0 Then
                    oDocDestino.AccountPayments.VatGroup = getVatGroupDeVatTotal(objDocumento.Ambito, objLinea.LineTotal, objLinea.VATTotal, objLinea.VATPorc, objLinea.Intracomunitario, "", Sociedad)
                End If

                'Proyecto
                If Not String.IsNullOrEmpty(objLinea.Proyecto) Then oDocDestino.AccountPayments.ProjectCode = objLinea.Proyecto

                'Centro coste 1
                If Not String.IsNullOrEmpty(objLinea.Centro1Coste) Then oDocDestino.AccountPayments.ProfitCenter = objLinea.Centro1Coste

                'Centro coste 2
                If Not String.IsNullOrEmpty(objLinea.Centro2Coste) Then oDocDestino.AccountPayments.ProfitCenter2 = objLinea.Centro2Coste

                'Centro coste 3
                If Not String.IsNullOrEmpty(objLinea.Centro3Coste) Then oDocDestino.AccountPayments.ProfitCenter3 = objLinea.Centro3Coste

                'Centro coste 4
                If Not String.IsNullOrEmpty(objLinea.Centro4Coste) Then oDocDestino.AccountPayments.ProfitCenter4 = objLinea.Centro4Coste

                'Centro coste 5
                If Not String.IsNullOrEmpty(objLinea.Centro5Coste) Then oDocDestino.AccountPayments.ProfitCenter5 = objLinea.Centro5Coste

                'Campos de usuario
                setCamposUsuarioLineaCobroPago(objLinea, oDocDestino)

                'Añade línea
                oDocDestino.AccountPayments.Add()

            Next

            'ID DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIDDW").Value = objDocumento.DOCIDDW

            'URL DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIURLDW").Value = objDocumento.DOCURLDW

            'IMPORTE DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIMPDW").Value = objDocumento.DocTotal

            'REVISAR IMPORTE DOCUWARE
            'El sistema permite traspasar un borrador a definitivo, aunque el importe del documento sea diferente que al importe DW
            oDocDestino.UserFields.Fields.Item("U_SEIREVDW").Value = IIf(objDocumento.Draft = Utilidades.Draft.Borrador, "S", "N")

            'Campos de usuario
            setCamposUsuarioCabeceraCobroPago(objDocumento, oDocDestino)

            'Añadimos documento
            If oDocDestino.Add() = 0 Then

                'Obtiene el DocEntry y DocNum del documento añadido
                Dim DocEntry As String = oCompany.GetNewObjectKey

                Dim DocNum As String = ""
                If objDocumento.Draft = Utilidades.Draft.Firme Then DocNum = oComun.getDocNumDeDocEntry(TablaDestino, DocEntry, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento creado con éxito"
                retVal.MENSAJEAUX = DocNum

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
            oComun.LiberarObjCOM(oDocDestino)
        End Try

        Return retVal

    End Function

    Private Function CopiarADocumentoCobroPago(ByVal objDocumento As EntDocumentoCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oDocEntrada As Documents = Nothing
        Dim oDocDestino As Payments = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "COPIA COBRO/PAGO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de documento para " & objDocumento.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objDocumento.UserSAP, objDocumento.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objDocumento.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocumento.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")
            If String.IsNullOrEmpty(objDocumento.RefOrigen) Then Throw New Exception("Referencia origen no suministrada")

            If Not DateTime.TryParseExact(objDocumento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha documento no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objDocumento.DocDateAux.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha documento auxiliar no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha contable no suministrada o incorrecta")

            If String.IsNullOrEmpty(objDocumento.CobroPagoImporte) Then Throw New Exception("Importe cobro/pago no suministrado")
            If String.IsNullOrEmpty(objDocumento.CobroPagoCuenta) Then Throw New Exception("Cuenta cobro/pago no suministrada")

            'Buscamos tabla origen por ObjectType
            Dim TablaOrigen As String = getTablaDeObjType(objDocumento.ObjTypeOrigen)
            If String.IsNullOrEmpty(TablaOrigen) Then Throw New Exception("No encuentro tabla origen con ObjectType: " & objDocumento.ObjTypeOrigen)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla origen: " & TablaOrigen)

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objDocumento.ObjTypeDestino)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objDocumento.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objDocumento.NIFTercero, objDocumento.RazonSocial, objDocumento.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumento.NIFTercero & ", Razón social: " & objDocumento.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Buscar responsable IC
            Dim ICResponsable As String = ""
            If getEsDocumentoVenta(objDocumento.ObjTypeDestino) Then ICResponsable = getResponsableDeCardCode(CardCode, Sociedad)
            If Not String.IsNullOrEmpty(ICResponsable) Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado responsable IC: " & ICResponsable)

            'Buscar contacto IC
            Dim ICContacto As Integer = getContactoDeEmail(CardCode, objDocumento.ContactoEmail, Sociedad)
            If ICContacto > 0 Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado contacto IC: " & ICContacto)

            'Comprueba si el documento se debe generar como borrador o en firme 
            Dim Draft As String = objDocumento.Draft
            If Draft = Utilidades.Draft.Interlocutor Then
                Draft = getDraftDeCardCode(CardCode, Sociedad)
                clsLog.Log.Info("(" & sLogInfo & ") Documento borrador por IC")
            End If

            'Objeto documento
            If Draft <> Utilidades.Draft.Firme Then
                oDocDestino = oCompany.GetBusinessObject(BoObjectTypes.oPaymentsDrafts)

                'Tipo de borrador
                If objDocumento.ObjTypeDestino = ObjType.Cobro Then
                    oDocDestino.DocObjectCode = BoPaymentsObjectType.bopot_IncomingPayments
                Else
                    oDocDestino.DocObjectCode = BoPaymentsObjectType.bopot_OutgoingPayments
                End If

                clsLog.Log.Info("(" & sLogInfo & ") Documento borrador")
            Else
                oDocDestino = oCompany.GetBusinessObject(objDocumento.ObjTypeDestino)
                clsLog.Log.Info("(" & sLogInfo & ") Documento en firme")
            End If

            'Tipo de pago
            If objDocumento.ObjTypeDestino = ObjType.Cobro Then
                oDocDestino.DocTypte = BoRcptTypes.rCustomer
            Else
                oDocDestino.DocTypte = BoRcptTypes.rSupplier
            End If

            'Serie
            If objDocumento.Serie > 0 Then oDocDestino.Series = objDocumento.Serie

            'IC
            oDocDestino.CardCode = CardCode

            'Contacto
            If objDocumento.ContactoCodigo > 0 Then
                oDocDestino.ContactPersonCode = objDocumento.ContactoCodigo
            ElseIf ICContacto > 0 Then
                oDocDestino.ContactPersonCode = ICContacto
            End If

            'Referencia
            If Not String.IsNullOrEmpty(objDocumento.NumAtCard) Then oDocDestino.CounterReference = objDocumento.NumAtCard

            'Moneda
            If Not String.IsNullOrEmpty(objDocumento.Currency) Then
                oDocDestino.DocCurrency = objDocumento.Currency
            End If

            'Comentarios
            If objDocumento.Comments.Length > 254 Then objDocumento.Comments = objDocumento.Comments.Substring(0, 254)
            If Not String.IsNullOrEmpty(objDocumento.Comments) Then oDocDestino.Remarks = objDocumento.Comments

            'Fecha contable 
            oDocDestino.DocDate = Date.ParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Fecha documento
            'Comprueba si el periodo contable está abierto para la fecha DocDate de entrada. Si no lo está, coge la fecha auxiliar DocDateAux
            Dim PeriodoAbierto As Boolean = getPeriodoAbiertoDeTaxDate(objDocumento.DocDate.ToString, Sociedad)
            If PeriodoAbierto Then
                oDocDestino.TaxDate = Date.ParseExact(objDocumento.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            Else
                oDocDestino.TaxDate = Date.ParseExact(objDocumento.DocDateAux.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            End If

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objDocumento.RefOrigen.Split("#")
            If RefsOrigen.Count > 1 Then Throw New Exception("No se puede crear un cobro de varios documentos origen: " & objDocumento.RefOrigen)

            'Comprueba que existe el documento origen y que se puede acceder a él
            'Pueden ser varios documentos origen con la misma referencia
            Dim DocEntrysOrigen As List(Of String) = getDocEntryDeRefOrigen(TablaOrigen, RefsOrigen(0), objDocumento.TipoRefOrigen, CardCode, True, True, Sociedad)

            If DocEntrysOrigen Is Nothing OrElse DocEntrysOrigen.Count = 0 Then
                Throw New Exception("No existe documento origen con referencia: " & RefsOrigen(0) & " o su estado es cerrado/cancelado")
            ElseIf DocEntrysOrigen.Count > 1 Then
                Throw New Exception("No se puede crear un cobro de varios documentos origen: " & RefsOrigen(0))
            End If

            'Objeto origen
            oDocEntrada = oCompany.GetBusinessObject(objDocumento.ObjTypeOrigen)

            If Not oDocEntrada.GetByKey(DocEntrysOrigen(0)) Then Throw New Exception("No puedo recuperar el documento origen con referencia: " & RefsOrigen(0) & " y DocEntry:" & DocEntrysOrigen(0))

            'Copia los datos del documento origen al documento destino
            oDocDestino.CardCode = oDocEntrada.CardCode
            oDocDestino.CardName = oDocEntrada.CardName

            'Calcula la serie
            Dim Serie As String = oComun.getSerieDeDocumentoDestino(oDocEntrada.Series, oDocEntrada.DocObjectCode, objDocumento.ObjTypeDestino, Sociedad)
            If Not oDocDestino.Series > 0 AndAlso Not String.IsNullOrEmpty(Serie) Then oDocDestino.Series = CInt(Serie)

            'FinanzasRazon
            If Not String.IsNullOrEmpty(objDocumento.FinanzasRazon) Then oDocDestino.CardName = objDocumento.FinanzasRazon

            'Responsable
            If Not String.IsNullOrEmpty(objDocumento.Responsable) Then
                oDocDestino.BillOfExchangeAgent = objDocumento.Responsable
            ElseIf Not String.IsNullOrEmpty(oDocEntrada.AgentCode) Then
                oDocDestino.BillOfExchangeAgent = oDocEntrada.AgentCode
            ElseIf Not String.IsNullOrEmpty(ICResponsable) Then
                oDocDestino.BillOfExchangeAgent = ICResponsable
            End If

            'Para que no haya problemas de redondeos... A veces se queda un par de centimos sin cobrar...
            If Math.Abs(oDocEntrada.DocTotal - objDocumento.CobroPagoImporte) > 0 Then
                'Hay diferencia...
                If Math.Abs(oDocEntrada.DocTotal - objDocumento.CobroPagoImporte) <= 0.03 Then
                    clsLog.Log.Info("(" & sLogInfo & ") Redondeo de diferencia minima Doc.:" & oDocEntrada.DocTotal & " A cobrar:" & objDocumento.CobroPagoImporte)
                    objDocumento.CobroPagoImporte = oDocEntrada.DocTotal
                End If
            End If

            'Solo 1 vencimiento
            'oDocDestino.Invoices.SumApplied = objDocumento.CobroPagoImporte
            'oDocDestino.Invoices.DocEntry = DocEntrysOrigen(0)
            'oDocDestino.Invoices.InvoiceType = objDocumento.ObjTypeOrigen

            'Vencimientos
            If getComprobarVencimientos(objDocumento.ObjTypeOrigen) Then

                For i As Integer = 1 To oDocEntrada.Installments.Count  'Cuotas de cada vencimiento

                    If i > 1 Then oDocDestino.Invoices.Add()

                    oDocEntrada.Installments.SetCurrentLine(i - 1)
                    oDocDestino.Invoices.SetCurrentLine(i - 1)

                    oDocDestino.Invoices.SumApplied = objDocumento.CobroPagoImporte
                    oDocDestino.Invoices.DocEntry = DocEntrysOrigen(0)
                    oDocDestino.Invoices.InvoiceType = objDocumento.ObjTypeOrigen
                    oDocDestino.Invoices.InstallmentId = i

                Next

            End If

            'Medio pago
            setMedioCobroPago(objDocumento, oDocDestino, Sociedad)

            'Proyecto
            If Not String.IsNullOrEmpty(objDocumento.CobroPagoProyecto) Then oDocDestino.ProjectCode = objDocumento.CobroPagoProyecto

            'ID DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIDDW").Value = objDocumento.DOCIDDW

            'URL DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIURLDW").Value = objDocumento.DOCURLDW

            'IMPORTE DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIMPDW").Value = objDocumento.DocTotal

            'REVISAR IMPORTE DOCUWARE
            'El sistema permite traspasar un borrador a definitivo, aunque el importe del documento sea diferente que al importe DW
            oDocDestino.UserFields.Fields.Item("U_SEIREVDW").Value = IIf(Draft = Utilidades.Draft.Borrador, "S", "N")

            'Documentos relacionados
            setDocumentosRelacionadosCobroPago(objDocumento, oDocDestino, CardCode, sLogInfo, Sociedad)

            'Campos de usuario
            setCamposUsuarioCabeceraCobroPago(objDocumento, oDocDestino)

            'Añadimos documento
            If oDocDestino.Add() = 0 Then

                'Obtiene el DocEntry y DocNum del documento añadido
                Dim DocEntry As String = oCompany.GetNewObjectKey

                Dim DocNum As String = ""
                If Draft = Utilidades.Draft.Firme Then DocNum = oComun.getDocNumDeDocEntry(TablaDestino, DocEntry, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento creado con éxito"
                retVal.MENSAJEAUX = DocNum

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
            oComun.LiberarObjCOM(oDocDestino)
            oComun.LiberarObjCOM(oDocEntrada)
        End Try

        Return retVal

    End Function

    Private Function CopiarADocumentoPrecioEntrega(ByVal objDocumento As EntDocumentoCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oService As LandedCostsService = Nothing
        Dim oDocEntrada As Documents = Nothing
        Dim oDocDestino As LandedCost = Nothing
        Dim oDocParams As LandedCostParams = Nothing
        Dim oDocLinea As LandedCost_ItemLine = Nothing
        Dim oDocCoste As LandedCost_CostLines = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "COPIA PRECIO ENTREGA"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de documento para " & objDocumento.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objDocumento.UserSAP, objDocumento.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objDocumento.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objDocumento.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")
            If String.IsNullOrEmpty(objDocumento.RefOrigen) Then Throw New Exception("Referencia origen no suministrada")

            If Not DateTime.TryParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha contable no suministrada o incorrecta")

            'Buscamos tabla origen por ObjectType
            Dim TablaOrigen As String = getTablaDeObjType(objDocumento.ObjTypeOrigen)
            If String.IsNullOrEmpty(TablaOrigen) Then Throw New Exception("No encuentro tabla origen con ObjectType: " & objDocumento.ObjTypeOrigen)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla origen: " & TablaOrigen)

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objDocumento.ObjTypeDestino)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objDocumento.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objDocumento.NIFTercero, objDocumento.RazonSocial, objDocumento.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objDocumento.NIFTercero & ", Razón social: " & objDocumento.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Objeto documento
            oService = oCompany.GetCompanyService().GetBusinessService(objDocumento.ObjTypeDestino)
            oDocDestino = oService.GetDataInterface(LandedCostsServiceDataInterfaces.lcsLandedCost)
            clsLog.Log.Info("(" & sLogInfo & ") Documento en firme")

            'Serie
            If objDocumento.Serie > 0 Then oDocDestino.Series = objDocumento.Serie

            'Fecha contable 
            oDocDestino.PostingDate = Date.ParseExact(objDocumento.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Fecha vencimiento
            If DateTime.TryParseExact(objDocumento.DocDueDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                oDocDestino.DueDate = Date.ParseExact(objDocumento.DocDueDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Moneda
            If Not String.IsNullOrEmpty(objDocumento.Currency) Then oDocDestino.DocumentCurrency = objDocumento.Currency

            'Referencia
            If Not String.IsNullOrEmpty(objDocumento.NumAtCard) Then oDocDestino.Reference = objDocumento.NumAtCard

            'Comentarios
            If objDocumento.Comments.Length > 254 Then objDocumento.Comments = objDocumento.Comments.Substring(0, 254)
            If Not String.IsNullOrEmpty(objDocumento.Comments) Then oDocDestino.Remarks = objDocumento.Comments

            'Entrada diario
            If objDocumento.JournalMemo.Length > 254 Then objDocumento.JournalMemo = objDocumento.JournalMemo.Substring(0, 254)
            If Not String.IsNullOrEmpty(objDocumento.JournalMemo) Then oDocDestino.JournalRemarks = objDocumento.JournalMemo

            'RefsOrigen contiene los NumAtCard
            Dim RefsOrigen As String() = objDocumento.RefOrigen.Split("#")
            If RefsOrigen.Count > 1 Then Throw New Exception("No se puede crear un cobro de varios documentos origen: " & objDocumento.RefOrigen)

            'Comprueba que existe el documento origen y que se puede acceder a él
            'Pueden ser varios documentos origen con la misma referencia
            Dim DocEntrysOrigen As List(Of String) = getDocEntryDeRefOrigen(TablaOrigen, RefsOrigen(0), objDocumento.TipoRefOrigen, CardCode, False, True, Sociedad)

            If DocEntrysOrigen Is Nothing OrElse DocEntrysOrigen.Count = 0 Then
                Throw New Exception("No existe documento origen con referencia: " & RefsOrigen(0) & " o su estado es cerrado/cancelado")
            ElseIf DocEntrysOrigen.Count > 1 Then
                Throw New Exception("No se puede crear un documento de varios documentos origen: " & RefsOrigen(0))
            End If

            'Objeto origen
            oDocEntrada = oCompany.GetBusinessObject(objDocumento.ObjTypeOrigen)

            If Not oDocEntrada.GetByKey(DocEntrysOrigen(0)) Then Throw New Exception("No puedo recuperar el documento origen con referencia: " & RefsOrigen(0) & " y DocEntry:" & DocEntrysOrigen(0))

            'Líneas
            For iLinea As Integer = 0 To oDocEntrada.Lines.Count - 1

                oDocEntrada.Lines.SetCurrentLine(iLinea)

                'Traspasa todas las líneas
                oDocLinea = oDocDestino.LandedCost_ItemLines.Add

                oDocLinea.BaseEntry = oDocEntrada.DocEntry
                oDocLinea.BaseLine = oDocEntrada.Lines.LineNum
                oDocLinea.BaseDocumentType = getBaseObjTypePrecioEntrega(oDocEntrada.DocObjectCode)

            Next

            'Aduana
            If oDocEntrada.AduanaImporte > 0 Then oDocDestino.ActualCustoms = oDocEntrada.AduanaImporte

            'ID DOCUWARE
            oDocDestino.UserFields.Item("U_SEIIDDW").Value = objDocumento.DOCIDDW

            'URL DOCUWARE
            oDocDestino.UserFields.Item("U_SEIURLDW").Value = objDocumento.DOCURLDW

            'Añadimos documento
            oDocParams = oService.AddLandedCost(oDocDestino)

            'Obtiene el DocEntry y DocNum del documento añadido
            Dim DocEntry As String = oDocParams.LandedCostNumber

            Dim DocNum As String = ""
            DocNum = oComun.getDocNumDeDocEntry(TablaDestino, DocEntry, Sociedad)

            retVal.CODIGO = Respuesta.Ok
            retVal.MENSAJE = "Documento creado con éxito"
            retVal.MENSAJEAUX = DocNum

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        Finally
            oComun.LiberarObjCOM(oDocDestino)
            oComun.LiberarObjCOM(oDocEntrada)
            oComun.LiberarObjCOM(oDocLinea)
            oComun.LiberarObjCOM(oDocCoste)
            oComun.LiberarObjCOM(oDocParams)
            oComun.LiberarObjCOM(oService)
        End Try

        Return retVal

    End Function

    Private Function CerrarDocumentoPrecioEntrega(ByVal objDocumento As EntDocumentoCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oCompany As Company
        Dim oService As LandedCostsService = Nothing
        Dim oDocParams As LandedCostParams = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "CIERRE PRECIO ENTREGA"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de cierre de documento para " & objDocumento.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objDocumento.UserSAP, objDocumento.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Objeto origen
            oService = oCompany.GetCompanyService().GetBusinessService(69)

            oDocParams = oService.GetDataInterface(LandedCostsServiceDataInterfaces.lcsLandedCostParams)
            oDocParams.LandedCostNumber = 1

            oService.CloseLandedCost(oDocParams)

            retVal.CODIGO = Respuesta.Ok
            retVal.MENSAJE = "Documento cerrado con éxito"
            retVal.MENSAJEAUX = "1"

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        Finally
            oComun.LiberarObjCOM(oDocParams)
            oComun.LiberarObjCOM(oService)
        End Try

        Return retVal

    End Function

    Private Function CopiarPrecioEntregaADocumento(ByVal objDocumento As EntDocumentoCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oService As LandedCostsService = Nothing
        Dim oDocEntrada As LandedCost = Nothing
        Dim oDocDestino As Documents = Nothing
        Dim oDocLinea As LandedCost_ItemLine = Nothing
        Dim oDocCoste As LandedCost_CostLines = Nothing
        Dim oDocParams As LandedCostParams = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "COPIA PRECIO ENTREGA"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de documento para " & objDocumento.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objDocumento.UserSAP, objDocumento.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Objeto origen
            oService = oCompany.GetCompanyService().GetBusinessService(69)

            oDocParams = oService.GetDataInterface(LandedCostsServiceDataInterfaces.lcsLandedCostParams)
            oDocParams.LandedCostNumber = 4

            oDocEntrada = oService.GetDataInterface(LandedCostsServiceDataInterfaces.lcsLandedCost)
            oDocEntrada = oService.GetLandedCost(oDocParams)

            'Objeto documento
            oDocDestino = oCompany.GetBusinessObject(BoObjectTypes.oDrafts)
            oDocDestino.DocObjectCode = 18

            'DocType
            oDocDestino.DocType = BoDocumentTypes.dDocument_Service

            'CardCode
            oDocDestino.CardCode = "P99999"

            'Fecha contable 
            oDocDestino.DocDate = Now.Date

            'Fecha documento
            oDocDestino.DocDate = Now.Date

            'Fecha vencimiento
            oDocDestino.DocDueDate = Now.Date

            'Referencia
            objDocumento.NumAtCard = "DUA MACYS"


            'If Not oDocDestino.GetByKey(4) Then Throw New Exception("No puedo recuperar el documento origen con referencia: 4 y DocEntry: 4")

            'Línea 1 
            oDocDestino.Lines.SetCurrentLine(oDocDestino.Lines.Count - 1)

            oDocDestino.Lines.BaseType = 69
            oDocDestino.Lines.BaseEntry = 4
            oDocDestino.Lines.BaseLinea = 0

            oDocDestino.Lines.ItemDescription = "LINEA 1"
            oDocDestino.Lines.AccountCode = "410900"
            oDocDestino.Lines.Price = 200
            oDocDestino.Lines.VatGroup = "S3"

            oDocDestino.Lines.Add()

            'Línea 2
            oDocDestino.Lines.SetCurrentLine(oDocDestino.Lines.Count - 1)

            oDocDestino.Lines.BaseType = 69
            oDocDestino.Lines.BaseEntry = 4
            oDocDestino.Lines.BaseLinea = 1

            oDocDestino.Lines.ItemDescription = "LINEA 2"
            oDocDestino.Lines.AccountCode = "410900"
            oDocDestino.Lines.Price = 150
            oDocDestino.Lines.VatGroup = "S3"

            oDocDestino.Lines.Add()

            'Añadimos documento
            If oDocDestino.Add() = 0 Then

                'Obtiene el DocEntry y DocNum del documento añadido
                Dim DocEntry As String = oCompany.GetNewObjectKey

                Dim DocNum As String = ""
                DocNum = oComun.getDocNumDeDocEntry("OPOR", DocEntry, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento copiado con éxito"
                retVal.MENSAJEAUX = DocNum

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
            oComun.LiberarObjCOM(oDocDestino)
            oComun.LiberarObjCOM(oDocEntrada)
            oComun.LiberarObjCOM(oDocLinea)
            oComun.LiberarObjCOM(oDocParams)
            oComun.LiberarObjCOM(oDocCoste)
            oComun.LiberarObjCOM(oService)
        End Try

        Return retVal

    End Function

    Private Function getPrecioConDescuento(ByRef oCompamy As Company, ByVal CardCode As String, ByVal ItemCode As String,
                                           ByVal DocDate As Date, ByVal Quantity As Double) As ItemPriceReturnParams

        Dim oPrecios As ItemPriceParams = Nothing
        Dim oPreciosReturn As ItemPriceReturnParams = Nothing

        Try

            oPrecios = oCompamy.GetCompanyService().GetDataInterface(CompanyServiceDataInterfaces.csdiItemPriceParams)
            oPreciosReturn = oCompamy.GetCompanyService().GetDataInterface(CompanyServiceDataInterfaces.csdiItemPriceReturnParams)

            With oPrecios
                .CardCode = CardCode
                .ItemCode = ItemCode
                .Date = DocDate
                .UoMQuantity = Quantity
            End With

            oPreciosReturn = oCompamy.GetCompanyService().GetItemPrice(oPrecios)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return oPreciosReturn

    End Function

    Private Function getDocNumDocumentoPorDWID(ByVal Tabla As String,
                                               ByVal IDDW As String,
                                               ByVal Sociedad As eSociedad) As String

        'Devuelve el DocNum del documento en firme

        Dim retVal As String = ""

        Try

            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("DocNum") & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And COALESCE(T0." & putQuotes("U_SEIIDDW") & ",'')  = N'" & IDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Sub setDatosDWDocumento(ByVal Tabla As String,
                                    ByVal CardCode As String,
                                    ByVal DocDate As Integer,
                                    ByVal DocNums As List(Of String),
                                    ByVal IDDW As String,
                                    ByVal URLDW As String,
                                    ByVal Sociedad As eSociedad)

        'Actualiza los campos de DW en el documento

        Try

            'Actualizamos por CardCode, DocDate y DocNum
            Dim SQL As String = ""
            SQL = "  UPDATE " & getDataBaseRef(Tabla, Sociedad) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEIIDDW") & " =N'" & IDDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIURLDW") & " =N'" & URLDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIIMPDW") & " = " & putQuotes("DocTotal") & " " & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And COALESCE(" & putQuotes("CardCode") & ",'') = N'" & CardCode & "'" & vbCrLf
            'SQL &= " And COALESCE(" & putQuotes("DocDate") & "," & getDefaultDate & ") = N'" & DocDate & "'" & vbCrLf

            SQL &= " And COALESCE(" & putQuotes("DocNum") & ",-1)  IN ('-1' " & vbCrLf
            For Each DocNum In DocNums
                SQL &= ", N'" & DocNum & "'" & vbCrLf
            Next
            SQL &= ") " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("No se ha podido actualizar el documento: " & putQuotes(Tabla) & " para CardCode: " & CardCode & " y DocNum: " & String.Join("#", DocNums))

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Function setDocumentoDWTratado(ByVal Tabla As String,
                                           ByVal DocEntry As String,
                                           ByVal Sociedad As eSociedad) As Boolean

        'Actualiza los campos de DW en el documento

        Dim retval As Boolean = False

        Try

            'Actualizamos por CardCode, DocDate y DocNum
            Dim SQL As String = ""
            SQL = "  UPDATE " & getDataBaseRef(Tabla, Sociedad) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEITRADW") & " =N'" & SN.Si & "' " & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And COALESCE(" & putQuotes("DocEntry") & ",'-1') = N'" & DocEntry & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("Actualización nula al ejecutar sentencia " & SQL)

            retval = True

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retval

    End Function

    Private Function setDocumentoDWEstado(ByVal Tabla As String,
                                          ByVal DocEntry As String,
                                          ByVal IDDW As String,
                                          ByVal URLDW As String,
                                          ByVal ESTADODW As Integer,
                                          ByVal MOTIVODW As String,
                                          ByVal Sociedad As eSociedad) As Boolean

        'Actualiza los campos de DW en el documento

        Dim retval As Boolean = False

        Try

            'Actualizamos por CardCode, DocDate y DocNum
            Dim SQL As String = ""
            SQL = "  UPDATE " & getDataBaseRef(Tabla, Sociedad) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEIIDDW") & " =N'" & IDDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIURLDW") & " =N'" & URLDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIESTDW") & " =N'" & ESTADODW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIMOTDW") & " =N'" & MOTIVODW & "' " & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And COALESCE(" & putQuotes("DocEntry") & ",'-1') = N'" & DocEntry & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("Actualización nula al ejecutar sentencia " & SQL)

            retval = True

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retval

    End Function

    Private Function getDocEntryDeRefOrigen(ByVal Tabla As String,
                                            ByVal RefOrigen As String,
                                            ByVal TipoRefOrigen As String,
                                            ByVal CardCode As String,
                                            ByVal Cerrado As Boolean,
                                            ByVal Cancelado As Boolean,
                                            ByVal Sociedad As eSociedad) As List(Of String)

        'Devuelve el DocEntry del documento

        Dim retVal As New List(Of String)

        Try

            'Buscamos por DocNum, DocEntry o NumAtCard
            'El documento no esté cerrado o cancelado para que no cree borradores sin líneas
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("DocEntry") & " As DocEntry " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
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
            If Cerrado Then SQL &= " And T0." & putQuotes("DocStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf
            If Cancelado Then SQL &= " And T0." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(DR.Item("DocEntry").ToString)
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getDocNumOrigenDeRefOrigen(ByVal TablaOrigen As String,
                                                ByVal TablaDestino As String,
                                                ByVal RefsOrigen As String(),
                                                ByVal TipoRefOrigen As String,
                                                ByVal CardCode As String,
                                                ByVal Sociedad As eSociedad) As List(Of String)

        Dim retval As New List(Of String)

        Try

            Dim SQL As String = ""

            SQL = "  SELECT " & vbCrLf
            SQL &= " T3." & putQuotes("DocNum") & vbCrLf
            SQL &= " FROM " & putQuotes(TablaOrigen) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " INNER JOIN " & putQuotes(TablaOrigen.Substring(1, 3) & "1") & " T1 " & getWithNoLock() & " ON T0." & putQuotes("DocEntry") & " = T1." & putQuotes("DocEntry") & vbCrLf
            SQL &= " INNER JOIN " & putQuotes(TablaDestino.Substring(1, 3) & "1") & " T2 " & getWithNoLock() & " ON T2." & putQuotes("BaseType") & " = T0." & putQuotes("ObjType") & vbCrLf
            SQL &= "                                                                                                                  AND T2." & putQuotes("BaseEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf
            SQL &= "                                                                                                                  AND T2." & putQuotes("BaseLine") & " = T1." & putQuotes("LineNum") & vbCrLf
            SQL &= " INNER JOIN " & putQuotes(TablaDestino) & " T3 " & getWithNoLock() & " ON T3." & putQuotes("DocEntry") & " = T2." & putQuotes("DocEntry") & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            Select Case TipoRefOrigen
                Case RefOrigen.DocNum
                    SQL &= " And T0." & putQuotes("DocNum")
                    SQL &= " IN ('-1'"
                    For Each Ref In RefsOrigen
                        SQL &= " ,N'" & Ref & "'"
                    Next
                    SQL &= " )" & vbCrLf
                Case RefOrigen.DocEntry
                    SQL &= " And T0." & putQuotes("DocEntry")
                    SQL &= " IN ('-1'"
                    For Each Ref In RefsOrigen
                        SQL &= " ,N'" & Ref & "'"
                    Next
                    SQL &= " )" & vbCrLf
                Case Else
                    '20220531: Buscar documentos por NumAtCard por startwith
                    SQL &= " And (T0." & putQuotes("NumAtCard") & " LIKE '_x1y2z3_'"
                    For Each Ref In RefsOrigen
                        SQL &= " OR T0." & putQuotes("NumAtCard") & " LIKE N'" & Ref & "%' "
                    Next
                    SQL &= " )" & vbCrLf
            End Select

            SQL &= " And T3." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            SQL &= " And T3." & putQuotes("DocStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf
            SQL &= " And T3." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf
            SQL &= " And T2." & putQuotes("LineStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf

            SQL &= " GROUP BY " & vbCrLf
            SQL &= " T3." & putQuotes("DocNum") & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retval.Add(DR.Item("DocNum"))
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retval

    End Function

    Private Function getDocNumDeRefOrigen(ByVal Tabla As String,
                                            ByVal RefOrigen As String,
                                            ByVal TipoRefOrigen As String,
                                            ByVal CardCode As String,
                                            ByVal Sociedad As eSociedad) As List(Of String)

        'Devuelve el DocEntry del documento

        Dim retVal As New List(Of String)

        Try

            'Buscamos por DocNum, DocEntry o NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("DocNum") & " As DocNum " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
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
            'SQL &= " And T0." & putQuotes("DocStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf
            SQL &= " And T0." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(DR.Item("DocNum").ToString)
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getDocEntryDeRefOrigenUnica(ByVal Tabla As String,
                                                 ByVal RefOrigen As String,
                                                 ByVal TipoRefOrigen As String,
                                                 ByVal CardCode As String,
                                                 ByVal bStatus As String,
                                                 ByVal Sociedad As eSociedad) As String

        'Devuelve el DocEntry del documento

        Dim retVal As String = ""

        Try

            'Buscamos por DocNum, DocEntry o NumAtCard
            'El documento no esté cerrado o cancelado para que no cree borradores sin líneas
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("DocEntry") & " As DocEntry " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
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

            If bStatus Then SQL &= " And T0." & putQuotes("Status") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso IsNumeric(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getDocEntryPedidoVenta(ByVal Tabla As String,
                                            ByVal DocDate As Integer,
                                            ByVal CardCode As String,
                                            ByVal DocType As String,
                                            ByVal Sociedad As eSociedad) As List(Of String)

        'Devuelve el DocEntry del documento

        Dim retVal As New List(Of String)

        Dim auxDocDate = DateTime.ParseExact(DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
        Dim yearDocDate = auxDocDate.Year
        Dim monthDocDate = auxDocDate.Month

        Try

            Dim SQL As String = ""
            SQL = "  SELECT * " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0  " & getWithNoLock() & vbCrLf
            SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " ON T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf
            SQL &= " JOIN " & getDataBaseRef("OPRJ", Sociedad) & " T2 " & getWithNoLock() & " ON T2." & putQuotes("PrjCode") & " = T1." & putQuotes("Project") & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " AND T0." & putQuotes("DocStatus") & " = N'" & DocStatus.Abierto & "'" & vbCrLf
            SQL &= " AND T0." & putQuotes("CardCode") & " = N'" & CardCode & "' " & vbCrLf
            SQL &= " AND T0." & putQuotes("DocType") & " = N'" & DocType & "' " & vbCrLf
            SQL &= " AND YEAR(T0." & putQuotes("DocDate") & ") = " & yearDocDate & vbCrLf
            SQL &= " AND MONTH(T0." & putQuotes("DocDate") & ") = " & monthDocDate & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(DR.Item("DocEntry").ToString)
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getDocNumDeCabecera(ByVal Tabla As String,
                                        ByVal CardCode As String,
                                        ByVal DocNum As String,
                                        ByVal NumAtCard As String,
                                        ByVal Sociedad As eSociedad) As String

        'Devuelve el DocNum del documento en firme

        Dim retVal As String = ""

        Try

            'Buscamos por CardCode, TaxDate y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("DocNum") & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And COALESCE(T0." & putQuotes("CardCode") & ",'') = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And COALESCE(T0." & putQuotes("DocNum") & ",-1) = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And COALESCE(T0." & putQuotes("NumAtCard") & ",'') = N'" & NumAtCard & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getDocTotalDeReferencias(ByVal Tabla As String,
                                             ByVal CardCode As String,
                                             ByVal Referencias As String(),
                                             ByVal TipoRefOrigen As String,
                                             ByVal Sociedad As eSociedad) As Double

        'Devuelve el importe total de los documentos 

        Dim retVal As Double = 0

        Try

            'Buscamos por CardCode, DocDate y DocNum
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            'SQL &= " SUM(T0." & putQuotes("DocTotal") & " - T0." & putQuotes("PaidToDate") & ") " & vbCrLf
            SQL &= " CASE " & vbCrLf
            SQL &= "    WHEN SUM(T0." & putQuotes("DocTotalFC") & ") = 0 THEN SUM(T0." & putQuotes("DocTotal") & " - T0." & putQuotes("PaidToDate") & ") " & vbCrLf
            SQL &= "    ELSE SUM(T0." & putQuotes("DocTotalFC") & " - T0." & putQuotes("PaidFC") & ") " & vbCrLf
            SQL &= " END " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            SQL &= " And COALESCE(T0." & putQuotes("CardCode") & ",'') = N'" & CardCode & "'" & vbCrLf
            SQL &= " And T0." & putQuotes("DocStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf
            SQL &= " And T0." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf

            Select Case TipoRefOrigen
                Case RefOrigen.DocNum
                    SQL &= " And T0." & putQuotes("DocNum")
                    SQL &= " IN ('-1'"
                    For Each Ref In Referencias
                        SQL &= " ,'" & Ref & "'"
                    Next
                    SQL &= " )" & vbCrLf
                Case RefOrigen.DocEntry
                    SQL &= " And T0." & putQuotes("DocEntry")
                    SQL &= " IN ('-1'"
                    For Each Ref In Referencias
                        SQL &= " ,'" & Ref & "'"
                    Next
                    SQL &= " )" & vbCrLf
                Case Else
                    'Buscar documentos por NumAtCard por startwith
                    SQL &= " And (T0." & putQuotes("NumAtCard") & " LIKE '_x1y2z3_'"
                    For Each Ref In Referencias
                        SQL &= " OR T0." & putQuotes("NumAtCard") & " LIKE N'" & Ref & "%' "
                    Next
                    SQL &= " )" & vbCrLf
            End Select

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso IsNumeric(oObj) Then retVal = CDbl(oObj)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getPortesDeReferencias(ByVal Tabla As String,
                                           ByVal CardCode As String,
                                           ByVal Referencias As String(),
                                           ByVal TipoRefOrigen As String,
                                           ByVal Sociedad As eSociedad) As List(Of EntOrigen)

        'Devuelvelas referencias origen de los portes

        Dim retVal As New List(Of EntOrigen)

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por CardCode, DocDate y DocNum
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T3." & putQuotes("ObjType") & ", " & vbCrLf
            SQL &= " T3." & putQuotes("DocEntry") & ", " & vbCrLf
            SQL &= " T3." & putQuotes("LineNum") & " " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "3", Sociedad) & " T3 " & getWithNoLock() & " On T3." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            If bICLinea Then _
                SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " On T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            SQL &= " And T0." & putQuotes("DocStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf
            SQL &= " And T0." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf

            SQL &= " And T3." & putQuotes("LineTotal") & " > 0 " & vbCrLf

            Select Case TipoRefOrigen
                Case RefOrigen.DocNum
                    SQL &= " And T0." & putQuotes("DocNum")
                    SQL &= " IN ('-1'"
                    For Each Ref In Referencias
                        SQL &= " ,'" & Ref & "'"
                    Next
                    SQL &= " )" & vbCrLf
                Case RefOrigen.DocEntry
                    SQL &= " And T0." & putQuotes("DocEntry")
                    SQL &= " IN ('-1'"
                    For Each Ref In Referencias
                        SQL &= " ,'" & Ref & "'"
                    Next
                    SQL &= " )" & vbCrLf
                Case Else
                    'Buscar documentos por NumAtCard por startwith
                    SQL &= " And (T0." & putQuotes("NumAtCard") & " LIKE '_x1y2z3_'"
                    For Each Ref In Referencias
                        SQL &= " OR T0." & putQuotes("NumAtCard") & " LIKE N'" & Ref & "%' "
                    Next
                    SQL &= " )" & vbCrLf
            End Select

            SQL &= "GROUP BY " & vbCrLf

            SQL &= " T3." & putQuotes("ObjType") & ", " & vbCrLf
            SQL &= " T3." & putQuotes("DocEntry") & ", " & vbCrLf
            SQL &= " T3." & putQuotes("LineNum") & " " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            Dim DT As DataTable = oCon.ObtenerDT(SQL)

            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(New EntOrigen With {.ObjType = CInt(DR.Item(0)),
                                                   .DocEntry = CInt(DR.Item(1)),
                                                   .LineNum = CInt(DR.Item(2))})
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getDocRetencionTotalDeReferencias(ByVal Tabla As String,
                                                       ByVal CardCode As String,
                                                       ByVal Referencias As String(),
                                                       ByVal TipoRefOrigen As String,
                                                       ByVal Sociedad As eSociedad) As Double

        'Devuelve el importe total sujeto a retención de los documentos 

        Dim retVal As Double = 0

        Try

            'Buscamos por CardCode, DocDate y DocNum
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " SUM(T1." & putQuotes("LineTotal") & ") " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " On T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            SQL &= " And COALESCE(T0." & putQuotes("CardCode") & ",'') = N'" & CardCode & "'" & vbCrLf
            SQL &= " And T0." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf
            SQL &= " And T1." & putQuotes("WtLiable") & " = N'" & SN.Yes & "'" & vbCrLf
            SQL &= " And T1." & putQuotes("LineStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf

            Select Case TipoRefOrigen
                Case RefOrigen.DocNum
                    SQL &= " And T0." & putQuotes("DocNum")
                    SQL &= " IN ('-1'"
                    For Each Ref In Referencias
                        SQL &= " ,'" & Ref & "'"
                    Next
                    SQL &= " )" & vbCrLf
                Case RefOrigen.DocEntry
                    SQL &= " And T0." & putQuotes("DocEntry")
                    SQL &= " IN ('-1'"
                    For Each Ref In Referencias
                        SQL &= " ,'" & Ref & "'"
                    Next
                    SQL &= " )" & vbCrLf
                Case Else
                    'Buscar documentos por NumAtCard por startwith
                    SQL &= " And (T0." & putQuotes("NumAtCard") & " LIKE '_x1y2z3_'"
                    For Each Ref In Referencias
                        SQL &= " OR T0." & putQuotes("NumAtCard") & " LIKE N'" & Ref & "%' "
                    Next
                    SQL &= " )" & vbCrLf
            End Select

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso IsNumeric(oObj) Then retVal = CDbl(oObj)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getArticuloDeReferenciaExterna(ByVal ReferenciaExterna As String, ByVal CardCode As String, ByVal Sociedad As eSociedad) As String

        'Obtener el Articulo cuando no esta el campo

        Dim retVal As String = ""

        Try

            'Buscamos por Numero de Catálogo del Articulo
            Dim SQL As String = ""
            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("ItemCode") & ",'') "
            SQL &= " FROM " & getDataBaseRef("OSCN", Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("CardCode") & "=N'" & CardCode & "'" & vbCrLf
            SQL &= " And T0." & putQuotes("Substitute") & "=N'" & ReferenciaExterna & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getDocumentoOrigen(ByVal sArticulo As String,
                                        ByVal RefOrigen As String,
                                        ByVal TablaOrigen As String,
                                        ByVal Cantidad As Double,
                                        ByVal Sociedad As eSociedad) As KeyValuePair(Of Integer, Integer)

        Dim retVal As New KeyValuePair(Of Integer, Integer)

        Try

            Dim SQL As String = ""
            SQL = "  SELECT  " & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("DocEntry") & ",0) As DocEntry, " & vbCrLf
            SQL &= " COALESCE(T1." & putQuotes("LineNum") & ",0) As LineNum " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(TablaOrigen, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " INNER JOIN " & getDataBaseRef(TablaOrigen.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " ON T0." & putQuotes("DocEntry") & " = T1." & putQuotes("DocEntry") & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("DocNum") & "=N'" & RefOrigen & "'" & vbCrLf
            SQL &= " And T1." & putQuotes("ItemCode") & "=N'" & sArticulo & "'" & vbCrLf
            SQL &= " And T1." & putQuotes("OpenQty") & ">=N'" & Cantidad.ToString.Replace(",", ".") & "'" & vbCrLf
            SQL &= " And T0." & putQuotes("DocStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf
            SQL &= " And T1." & putQuotes("LineStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf
            SQL &= " And T0." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then retVal = New KeyValuePair(Of Integer, Integer)(DT.Rows(0).Item("DocEntry"), DT.Rows(0).Item("LineNum"))

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal
    End Function

    Private Function getCantidadCompletadaOrden(ByVal DocEntryOrdenFabricacion As Long,
                                                ByVal Sociedad As eSociedad) As Double

        'Comprueba si existe un documento definitivo

        Dim retVal As Double = 0

        Try

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("CmpltQty") & ",0) " & vbCrLf
            SQL &= " FROM " & getDataBaseRef("OWOR", Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("DocEntry") & " = " & DocEntryOrdenFabricacion & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso IsNumeric(oObj) Then retVal = Double.Parse(oObj)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getPortes(ByVal PortesTipo As Integer, ByVal Sociedad As eSociedad) As Integer

        Dim retVal As Integer = 0

        Try

            'Buscamos un porte por defecto
            Dim SQL As String = ""

            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("ExpnsCode") & ",0) " & vbCrLf
            SQL &= " FROM " & getDataBaseRef("OEXD", Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            If PortesTipo > 0 Then SQL &= " And T0." & putQuotes("ExpnsCode") & "=N'" & PortesTipo & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso IsNumeric(oObj) AndAlso CInt(oObj) > 0 Then retVal = CInt(oObj)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getCobroPagoEfectoNumero(ByVal Sociedad As eSociedad) As Integer

        Dim retVal As Integer = 0

        Try

            'Buscamos un porte por defecto
            Dim SQL As String = ""

            SQL = "  SELECT  " & vbCrLf
            SQL &= " COALESCE(MAX(T0." & putQuotes("BoeNum") & "),0)+1 " & vbCrLf
            SQL &= " FROM " & getDataBaseRef("OBOE", Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso IsNumeric(oObj) AndAlso CInt(oObj) > 0 Then retVal = CInt(oObj)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getEsCopiaParcial(ByVal objLineas As List(Of EntDocumentoLin)) As Boolean

        'Comprueba si la copia es parcial o no
        Dim retVal As Boolean = False

        Try

            retVal = (From p In objLineas
                      Where p.LineNum >= 0 _
                      OrElse p.VisOrder >= 0 _
                      OrElse Not String.IsNullOrEmpty(p.Articulo.Trim) _
                      OrElse Not String.IsNullOrEmpty(p.RefExt.Trim) _
                      OrElse Not String.IsNullOrEmpty(p.Concepto.Trim)).Distinct.ToList.Count > 0

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getLinSinIdentificar(ByVal objDocumento As EntDocumentoCab) As Boolean

        'Comprueba si la copia es parcial o no
        Dim retVal As Boolean = False

        Try

            'Comprueba que no lleguen líneas sin artículo/descripción en la copia parcial
            Dim bSinIdentificar As Boolean = (From p In objDocumento.Lineas
                                              Where p.LineNum < 0 _
                                              AndAlso p.VisOrder < 0 _
                                              AndAlso (((String.IsNullOrEmpty(p.Articulo.Trim) AndAlso String.IsNullOrEmpty(p.RefExt.Trim) AndAlso objDocumento.DocType = DocType.Articulo) _
                                                        OrElse
                                                        (String.IsNullOrEmpty(p.Concepto.Trim) AndAlso objDocumento.DocType = DocType.Servicio)))).Distinct.ToList.Count > 0

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getLinCopiaParcial(ByVal objDocumento As EntDocumentoCab, ByVal oDocEntrada As Documents) As EntDocumentoLin

        'Traspaso parcial
        'Comprueba que el artículo/servicio sea el que corresponde. Si el número de línea viene relleno, se busca por ese campo 
        'Si el tipo de referencia de la línea viene relleno se busca por ese campo
        Dim objLinea As New EntDocumentoLin

        Try

            If Not String.IsNullOrEmpty(objDocumento.Lineas(0).TipoRefOrigen) AndAlso Not String.IsNullOrEmpty(objDocumento.Lineas(0).RefOrigen) _
                AndAlso objDocumento.Lineas(0).VisOrder >= 0 Then
                'Búsqueda por documento origen con VisOrder
                objLinea = (From p In objDocumento.Lineas
                            Where p.VisOrder = oDocEntrada.Lines.VisualOrder _
                            AndAlso p.Cantidad > 0 _
                            AndAlso ((p.TipoRefOrigen = Utilidades.RefOrigen.DocEntry AndAlso p.RefOrigen = oDocEntrada.DocEntry.ToString) _
                                    OrElse
                                    (p.TipoRefOrigen = Utilidades.RefOrigen.DocNum AndAlso p.RefOrigen = oDocEntrada.DocNum.ToString) _
                                    OrElse
                                    (p.TipoRefOrigen = Utilidades.RefOrigen.NumAtCard AndAlso oDocEntrada.NumAtCard.StartsWith(p.RefOrigen)))).FirstOrDefault
            ElseIf Not String.IsNullOrEmpty(objDocumento.Lineas(0).TipoRefOrigen) AndAlso Not String.IsNullOrEmpty(objDocumento.Lineas(0).RefOrigen) _
                AndAlso objDocumento.Lineas(0).LineNum >= 0 Then
                'Búsqueda por documento origen con LineNum
                objLinea = (From p In objDocumento.Lineas
                            Where p.LineNum = oDocEntrada.Lines.LineNum _
                            AndAlso p.Cantidad > 0 _
                            AndAlso ((p.TipoRefOrigen = Utilidades.RefOrigen.DocEntry AndAlso p.RefOrigen = oDocEntrada.DocEntry.ToString) _
                                    OrElse
                                    (p.TipoRefOrigen = Utilidades.RefOrigen.DocNum AndAlso p.RefOrigen = oDocEntrada.DocNum.ToString) _
                                    OrElse
                                    (p.TipoRefOrigen = Utilidades.RefOrigen.NumAtCard AndAlso oDocEntrada.NumAtCard.StartsWith(p.RefOrigen)))).FirstOrDefault
            ElseIf Not String.IsNullOrEmpty(objDocumento.Lineas(0).TipoRefOrigen) AndAlso Not String.IsNullOrEmpty(objDocumento.Lineas(0).RefOrigen) Then
                'Búsqueda por documento origen
                objLinea = (From p In objDocumento.Lineas
                            Where p.Cantidad > 0 _
                            AndAlso ((p.TipoRefOrigen = Utilidades.RefOrigen.DocEntry AndAlso p.RefOrigen = oDocEntrada.DocEntry.ToString) _
                                    OrElse
                                    (p.TipoRefOrigen = Utilidades.RefOrigen.DocNum AndAlso p.RefOrigen = oDocEntrada.DocNum.ToString) _
                                    OrElse
                                    (p.TipoRefOrigen = Utilidades.RefOrigen.NumAtCard AndAlso oDocEntrada.NumAtCard.StartsWith(p.RefOrigen)))).FirstOrDefault
            ElseIf objDocumento.Lineas(0).VisOrder >= 0 Then
                'Búsqueda por VisOrder
                objLinea = (From p In objDocumento.Lineas
                            Where p.VisOrder = oDocEntrada.Lines.VisualOrder AndAlso p.Cantidad > 0).FirstOrDefault
            ElseIf objDocumento.Lineas(0).LineNum >= 0 Then
                'Búsqueda por LineNum
                objLinea = (From p In objDocumento.Lineas
                            Where p.LineNum = oDocEntrada.Lines.LineNum AndAlso p.Cantidad > 0).FirstOrDefault
            Else
                'Búsqueda por artículo/concepto 
                objLinea = (From p In objDocumento.Lineas
                            Where ((p.Articulo.Trim.ToUpper = oDocEntrada.Lines.ItemCode.Trim.ToUpper AndAlso p.Cantidad > 0 AndAlso objDocumento.DocType = DocType.Articulo) _
                                  OrElse
                                  (p.Concepto.Trim.ToUpper = oDocEntrada.Lines.ItemDescription.Trim.ToUpper AndAlso p.Cantidad > 0 AndAlso objDocumento.DocType = DocType.Servicio))).FirstOrDefault
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return objLinea

    End Function

    Private Function getLinArticuloSinTraspasar(ByVal DocType As String,
                                                ByVal CopiaParcial As Boolean,
                                                ByVal objLineas As List(Of EntDocumentoLin)) As String

        'Devuelve el artítulo/concepto de la línea sin traspasar
        Dim retVal As String = ""

        Try

            If CopiaParcial Then

                Dim LinNoTraspasada As EntDocumentoLin = (From p In objLineas Where p.Cantidad > 0).FirstOrDefault

                If Not LinNoTraspasada Is Nothing Then retVal = IIf(DocType = Utilidades.DocType.Articulo, LinNoTraspasada.Articulo.Trim, LinNoTraspasada.Concepto.Trim)

            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Sub getControlImportes(ByRef objDocumento As EntDocumentoCab,
                                   ByVal CardCode As String,
                                   ByRef TablaOrigen As String,
                                   ByRef RefsOrigen As String(),
                                   ByRef Draft As String,
                                   ByVal sLogInfo As String,
                                   ByVal Sociedad As eSociedad)

        'Control importes

        Try

            'Si el total del documento destino no coincide con la suma de las líneas de los documentos origen o es cero, se hace un borrador
            Dim DocTotalOrigen As Double = getDocTotalDeReferencias(TablaOrigen, CardCode, RefsOrigen, objDocumento.TipoRefOrigen, Sociedad)
            If DocTotalOrigen > 0 AndAlso objDocumento.IRPFImporte > 0 Then DocTotalOrigen -= objDocumento.IRPFImporte

            'No hay referencia origen
            If DocTotalOrigen = 0 Then Throw New Exception("La suma del importe origen es 0€")

            'Control de importes
            If objDocumento.ControlDiferencia = 0 Then
                'Debe coincidir el importe
                If DocTotalOrigen <> objDocumento.DocTotal Then
                    Draft = Utilidades.Draft.Borrador
                    clsLog.Log.Info("(" & sLogInfo & ") Documento borrador por importe")
                End If
            Else
                'Calcular la diferencia entre importe origen e importe DW
                If Math.Abs(DocTotalOrigen - objDocumento.DocTotal) > objDocumento.ControlDiferencia Then _
                    Throw New Exception("La diferencia entre el importe del documento origen y el importe de Docuware es mayor a " & objDocumento.ControlDiferencia & "€")
            End If

        Catch ex As Exception
            'clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub getRefsOrigenDocumentoNoDirecto(ByRef objDocumento As EntDocumentoCab,
                                                ByVal CardCode As String,
                                                ByRef TablaOrigen As String,
                                                ByRef RefsOrigen As String(),
                                                ByVal sLogInfo As String,
                                                ByVal Sociedad As eSociedad)

        'Origen con búsqueda indirecta (por ejemplo, albaranes de pedidos abiertos)

        Try

            If getEsDocumentoNoDirecto(objDocumento.ObjTypeOrigen) Then

                'ObtTypeOrigen: Documentos abiertos con origen X
                Dim ObjTypeOrigen As KeyValuePair(Of Integer, Integer) = getOrigenDeObjTypeNoDirecto(objDocumento.ObjTypeOrigen)
                Dim DocBaseOrigen As List(Of String) = getDocNumOrigenDeRefOrigen(getTablaDeObjType(ObjTypeOrigen.Value), TablaOrigen, RefsOrigen, objDocumento.TipoRefOrigen, CardCode, Sociedad)
                If DocBaseOrigen Is Nothing OrElse DocBaseOrigen.Count = 0 Then Throw New Exception("No existen documentos no directos con referencias: " & objDocumento.RefOrigen & " o su estado es cerrado/cancelado")

                'Nuevas referencias
                objDocumento.ObjTypeOrigen = ObjTypeOrigen.Key
                objDocumento.TipoRefOrigen = RefOrigen.DocNum
                objDocumento.RefOrigen = String.Join("#", DocBaseOrigen.ToArray)
                clsLog.Log.Info("(" & sLogInfo & ") Encontradas referencias documentos no directos: " & objDocumento.RefOrigen)

                RefsOrigen = DocBaseOrigen.ToArray

            End If

        Catch ex As Exception
            'clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub setFacturasAnticipo(ByRef objDocumento As EntDocumentoCab,
                                    ByRef oDocDestino As Documents,
                                    ByVal CardCode As String,
                                    ByVal Sociedad As eSociedad)

        'Facturas anticipo

        Try

            If getEsFactura(objDocumento.ObjTypeDestino) _
                AndAlso Not String.IsNullOrEmpty(objDocumento.AnticipoRefOrigen) _
                AndAlso Not String.IsNullOrEmpty(objDocumento.AnticipoTipoRefOrigen) Then

                'Objeto
                Dim oDownPaymentsToDraw As DownPaymentsToDraw = oDocDestino.DownPaymentsToDraw

                'Anticipo tabla origen
                Dim AnticipoTablaOrigen As String = getTablaAnticipoDeObjType(objDocumento.ObjTypeDestino)

                'AnticipoRefsOrigen contiene las referencias de las facturas anticipo
                Dim AnticipoRefsOrigen As String() = objDocumento.AnticipoRefOrigen.Split("#")

                'Recorre cada uno de los documento origen
                For Each AnticipoRefOrigen In AnticipoRefsOrigen

                    'Comprueba que existe la factura de anticipo y que se puede acceder a ella
                    Dim AnticipoDocEntryOrigen As String = getDocEntryDeRefOrigenUnica(AnticipoTablaOrigen, AnticipoRefOrigen, objDocumento.AnticipoTipoRefOrigen, CardCode, False, Sociedad)

                    If String.IsNullOrEmpty(AnticipoDocEntryOrigen) Then Throw New Exception("No existe factura anticipo con referencia: " & AnticipoRefOrigen)

                    oDownPaymentsToDraw.DocEntry = CInt(AnticipoDocEntryOrigen)
                    oDownPaymentsToDraw.Add()

                Next

            End If

            'Porcentaje anticipo
            If objDocumento.ObjTypeDestino = ObjType.FacturaCompraAnticipo Then
                oDocDestino.DownPaymentType = DownPaymentTypeEnum.dptInvoice
                If objDocumento.PorcentajeAnticipo > 0 Then oDocDestino.DownPaymentPercentage = objDocumento.PorcentajeAnticipo
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub setDocumentosRelacionados(ByRef objDocumento As EntDocumentoCab,
                                          ByRef oDocDestino As Documents,
                                          ByVal CardCode As String,
                                          ByVal sLogInfo As String,
                                          ByVal Sociedad As eSociedad)

        'Documentos relacionados

        Try

            If Not objDocumento.DocRelacionados Is Nothing AndAlso objDocumento.DocRelacionados.Count > 0 Then

                'Recorre cada uno de los documentos relacionados
                For Each DocRelacionado In objDocumento.DocRelacionados

                    'Índice
                    Dim indDocRelacionado As Integer = objDocumento.DocRelacionados.IndexOf(DocRelacionado) + 1

                    'Obligatorios
                    If String.IsNullOrEmpty(DocRelacionado.RefOrigen) Then Throw New Exception("Referencia origen documento relacionado " & indDocRelacionado.ToString & " no suministrada")

                    'Buscamos tabla documento relacionado por ObjectType
                    Dim TablaDocRelacionado As String = getTablaDeObjType(DocRelacionado.ObjType)
                    If String.IsNullOrEmpty(TablaDocRelacionado) Then Throw New Exception("No encuentro tabla documento relacionado " & indDocRelacionado.ToString & " con ObjectType: " & DocRelacionado.ObjType)
                    clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla documento relacionado: " & TablaDocRelacionado)

                    'Pueden ser varios documentos origen con la misma referencia
                    Dim DocEntrysDocRelacionado As List(Of String) = getDocEntryDeRefOrigen(TablaDocRelacionado, DocRelacionado.RefOrigen, DocRelacionado.TipoRefOrigen, CardCode, False, False, Sociedad)
                    If DocEntrysDocRelacionado Is Nothing OrElse DocEntrysDocRelacionado.Count = 0 Then Throw New Exception("No existe documento relacionado " & indDocRelacionado.ToString & " con referencia: " & DocRelacionado.RefOrigen)

                    'Relación
                    For Each DocEntryDocRelacionado In DocEntrysDocRelacionado
                        'oDocDestino.DocumentReferences.ReferencedObjectType = ReferencedObjectTypeEnum.rot_ExternalDocument
                        oDocDestino.DocumentReferences.ReferencedObjectType = DocRelacionado.ObjType
                        oDocDestino.DocumentReferences.ReferencedDocEntry = DocEntryDocRelacionado
                        If Not String.IsNullOrEmpty(DocRelacionado.Comentarios) Then oDocDestino.DocumentReferences.Remark = DocRelacionado.Comentarios

                        oDocDestino.DocumentReferences.Add()
                    Next

                Next

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub setDocumentosRelacionadosCobroPago(ByRef objDocumento As EntDocumentoCab,
                                                   ByRef oDocDestino As Payments,
                                                   ByVal CardCode As String,
                                                   ByVal sLogInfo As String,
                                                   ByVal Sociedad As eSociedad)

        'Documentos relacionados

        Try

            If Not objDocumento.DocRelacionados Is Nothing AndAlso objDocumento.DocRelacionados.Count > 0 Then

                'Recorre cada uno de los documentos relacionados
                For Each DocRelacionado In objDocumento.DocRelacionados

                    'Índice
                    Dim indDocRelacionado As Integer = objDocumento.DocRelacionados.IndexOf(DocRelacionado) + 1

                    'Obligatorios
                    If String.IsNullOrEmpty(DocRelacionado.RefOrigen) Then Throw New Exception("Referencia origen documento relacionado " & indDocRelacionado.ToString & " no suministrada")

                    'Buscamos tabla documento relacionado por ObjectType
                    Dim TablaDocRelacionado As String = getTablaDeObjType(DocRelacionado.ObjType)
                    If String.IsNullOrEmpty(TablaDocRelacionado) Then Throw New Exception("No encuentro tabla documento relacionado " & indDocRelacionado.ToString & " con ObjectType: " & DocRelacionado.ObjType)
                    clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla documento relacionado: " & TablaDocRelacionado)

                    'Pueden ser varios documentos origen con la misma referencia
                    Dim DocEntrysDocRelacionado As List(Of String) = getDocEntryDeRefOrigen(TablaDocRelacionado, DocRelacionado.RefOrigen, DocRelacionado.TipoRefOrigen, CardCode, False, False, Sociedad)
                    If DocEntrysDocRelacionado Is Nothing OrElse DocEntrysDocRelacionado.Count = 0 Then Throw New Exception("No existe documento relacionado " & indDocRelacionado.ToString & " con referencia: " & DocRelacionado.RefOrigen)

                    'Relación
                    For Each DocEntryDocRelacionado In DocEntrysDocRelacionado
                        'oDocDestino.DocumentReferences.ReferencedObjectType = ReferencedObjectTypeEnum.rot_ExternalDocument
                        oDocDestino.DocumentReferences.ReferencedObjectType = DocRelacionado.ObjType
                        oDocDestino.DocumentReferences.ReferencedDocEntry = DocEntryDocRelacionado
                        If Not String.IsNullOrEmpty(DocRelacionado.Comentarios) Then oDocDestino.DocumentReferences.Remark = DocRelacionado.Comentarios

                        oDocDestino.DocumentReferences.Add()
                    Next

                Next

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub setPortesCabecera(ByRef objDocumento As EntDocumentoCab,
                                  ByRef oDocDestino As Documents,
                                  ByVal CardCode As String,
                                  ByVal TablaOrigen As String,
                                  ByVal RefsOrigen As String(),
                                  ByVal sLogInfo As String,
                                  ByVal Sociedad As eSociedad)

        'Portes

        Try

            'Código portes
            Dim PortesCodigo As Integer = getPortes(0, Sociedad)
            If PortesCodigo > 0 Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado portes: " & PortesCodigo)

            'Asignación portes
            If Not objDocumento.PortesDetalle Is Nothing AndAlso objDocumento.PortesDetalle.Count > 0 Then

                For Each objPorte In objDocumento.PortesDetalle
                    oDocDestino.Expenses.SetCurrentLine(oDocDestino.Expenses.Count - 1)

                    If objPorte.Codigo > 0 Then
                        oDocDestino.Expenses.ExpenseCode = objPorte.Codigo
                    ElseIf PortesCodigo > 0 Then
                        oDocDestino.Expenses.ExpenseCode = PortesCodigo
                    Else
                        Throw New Exception("Código de portes no definido")
                    End If

                    oDocDestino.Expenses.LineTotal = objPorte.Importe

                    If Not String.IsNullOrEmpty(objPorte.VATGroup) Then oDocDestino.Expenses.VatGroup = objPorte.VATGroup

                    oDocDestino.Expenses.Add()
                Next

            ElseIf Not String.IsNullOrEmpty(TablaOrigen) AndAlso Not RefsOrigen Is Nothing AndAlso RefsOrigen.Count > 0 Then

                Dim Portes As List(Of EntOrigen) = getPortesDeReferencias(TablaOrigen, CardCode, RefsOrigen, objDocumento.TipoRefOrigen, Sociedad)

                For Each porte In Portes
                    oDocDestino.Expenses.SetCurrentLine(oDocDestino.Expenses.Count - 1)

                    oDocDestino.Expenses.BaseDocType = porte.ObjType
                    oDocDestino.Expenses.BaseDocEntry = porte.DocEntry
                    oDocDestino.Expenses.BaseDocLine = porte.LineNum

                    oDocDestino.Expenses.Add()
                Next

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub setPortesLinea(ByRef oDocEntrada As Documents,
                               ByRef oDocDestino As Documents)

        'Portes de la línea

        Try

            For i As Integer = 0 To oDocEntrada.Lines.Expenses.Count - 1

                'Se posiciona en el porte origen y destino
                oDocEntrada.Lines.Expenses.SetCurrentLine(i)
                oDocDestino.Lines.Expenses.SetCurrentLine(oDocDestino.Lines.Expenses.Count - 1)

                If oDocEntrada.Lines.Expenses.ExpenseCode > 0 Then
                    oDocDestino.Lines.Expenses.ExpenseCode = oDocEntrada.Lines.Expenses.ExpenseCode
                    oDocDestino.Lines.Expenses.LineTotal = oDocEntrada.Lines.Expenses.LineTotal
                    oDocDestino.Lines.Expenses.Add()
                End If

            Next

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub setRetencionesImpuestos(ByRef objDocumento As EntDocumentoCab,
                                        ByRef oDocDestino As Documents,
                                        ByVal CardCode As String,
                                        ByVal sLogInfo As String,
                                        ByVal Sociedad As eSociedad)

        'Retenciones de impuestos de IC (DocTotal sin retenciones)

        Try

            If objDocumento.IRPFImporte > 0 Then

                'Buscar grupo de retención de impuestos
                Dim GrupoRetencionImpuestos As String = getGrupoRetencionImpuestosDeCardCode(CardCode, Sociedad)
                If String.IsNullOrEmpty(GrupoRetencionImpuestos) Then Throw New Exception("Grupo retención de impuestos no definido en la ficha del IC")
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado grupo retención impuestos IC: " & GrupoRetencionImpuestos)

                'Base imponible de líneas con retención impuestos
                Dim BaseImponibleImpuestos As Double = (From p In objDocumento.Lineas
                                                        Where p.WTLiable = SN.Si
                                                        Select p.BaseTotal).ToList.Sum

                If Not BaseImponibleImpuestos > 0 Then Throw New Exception("Base imponible 0 de líneas con retención de impuestos")

                oDocDestino.WithholdingTaxData.SetCurrentLine(oDocDestino.WithholdingTaxData.Count - 1)
                oDocDestino.WithholdingTaxData.WTCode = GrupoRetencionImpuestos
                oDocDestino.WithholdingTaxData.WTAmount = objDocumento.IRPFImporte
                oDocDestino.WithholdingTaxData.TaxableAmount = BaseImponibleImpuestos
                oDocDestino.WithholdingTaxData.Add()

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub setCamposUsuarioCabecera(ByRef objDocumento As EntDocumentoCab,
                                         ByRef oDocDestino As Documents)

        'Campos de usuario

        Try

            If Not objDocumento.CamposUsuario Is Nothing AndAlso objDocumento.CamposUsuario.Count > 0 Then

                For Each oCampoUsuario In objDocumento.CamposUsuario

                    Dim oUserField As Field = oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo)

                    Select Case oUserField.Type

                        Case BoFieldTypes.db_Numeric
                            If oUserField.SubType = BoFldSubTypes.st_Time Then
                                If DateTime.TryParseExact(oCampoUsuario.Valor, "HHmm", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                    oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "HHmm", CultureInfo.CurrentCulture)
                            Else
                                If IsNumeric(oCampoUsuario.Valor) Then _
                                    oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CInt(oCampoUsuario.Valor)
                            End If

                        Case BoFieldTypes.db_Float
                            If IsNumeric(oCampoUsuario.Valor) Then _
                                oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CDbl(oCampoUsuario.Valor.Replace(".", ","))

                        Case BoFieldTypes.db_Date
                            If DateTime.TryParseExact(oCampoUsuario.Valor, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                        Case Else
                            If Not String.IsNullOrEmpty(oCampoUsuario.Valor) Then _
                                oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo).Value = oCampoUsuario.Valor

                    End Select

                Next

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub setCamposUsuarioCabeceraCobroPago(ByRef objDocumento As EntDocumentoCab,
                                                  ByRef oDocDestino As Payments)

        'Campos de usuario

        Try

            If Not objDocumento.CamposUsuario Is Nothing AndAlso objDocumento.CamposUsuario.Count > 0 Then

                For Each oCampoUsuario In objDocumento.CamposUsuario

                    Dim oUserField As Field = oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo)

                    Select Case oUserField.Type

                        Case BoFieldTypes.db_Numeric
                            If oUserField.SubType = BoFldSubTypes.st_Time Then
                                If DateTime.TryParseExact(oCampoUsuario.Valor, "HHmm", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                    oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "HHmm", CultureInfo.CurrentCulture)
                            Else
                                If IsNumeric(oCampoUsuario.Valor) Then _
                                    oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CInt(oCampoUsuario.Valor)
                            End If

                        Case BoFieldTypes.db_Float
                            If IsNumeric(oCampoUsuario.Valor) Then _
                                oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CDbl(oCampoUsuario.Valor.Replace(".", ","))

                        Case BoFieldTypes.db_Date
                            If DateTime.TryParseExact(oCampoUsuario.Valor, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                        Case Else
                            If Not String.IsNullOrEmpty(oCampoUsuario.Valor) Then _
                                oDocDestino.UserFields.Fields.Item(oCampoUsuario.Campo).Value = oCampoUsuario.Valor

                    End Select

                Next

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub setCamposUsuarioLinea(ByRef objLinea As EntDocumentoLin,
                                      ByRef oDocDestino As Documents)

        'Campos de usuario

        Try

            If Not objLinea.CamposUsuario Is Nothing AndAlso objLinea.CamposUsuario.Count > 0 Then

                For Each oCampoUsuario In objLinea.CamposUsuario

                    Dim oUserField As Field = oDocDestino.Lines.UserFields.Fields.Item(oCampoUsuario.Campo)

                    Select Case oUserField.Type

                        Case BoFieldTypes.db_Numeric
                            If oUserField.SubType = BoFldSubTypes.st_Time Then
                                If DateTime.TryParseExact(oCampoUsuario.Valor, "HHmm", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                            oDocDestino.Lines.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "HHmm", CultureInfo.CurrentCulture)
                            Else
                                If IsNumeric(oCampoUsuario.Valor) Then _
                                            oDocDestino.Lines.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CInt(oCampoUsuario.Valor)
                            End If

                        Case BoFieldTypes.db_Float
                            If IsNumeric(oCampoUsuario.Valor) Then _
                                        oDocDestino.Lines.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CDbl(oCampoUsuario.Valor.Replace(".", ","))

                        Case BoFieldTypes.db_Date
                            If DateTime.TryParseExact(oCampoUsuario.Valor, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                        oDocDestino.Lines.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                        Case Else
                            If Not String.IsNullOrEmpty(oCampoUsuario.Valor) Then _
                                        oDocDestino.Lines.UserFields.Fields.Item(oCampoUsuario.Campo).Value = oCampoUsuario.Valor

                    End Select

                Next

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub setCamposUsuarioLote(ByRef objLote As EntDocumentoLote,
                                     ByRef oDocDestino As Documents)

        'Campos de usuario

        Try

            If Not objLote.CamposUsuario Is Nothing AndAlso objLote.CamposUsuario.Count > 0 Then

                For Each oCampoUsuario In objLote.CamposUsuario

                    Dim oUserField As Field = oDocDestino.Lines.BatchNumbers.UserFields.Fields.Item(oCampoUsuario.Campo)

                    Select Case oUserField.Type

                        Case BoFieldTypes.db_Numeric
                            If oUserField.SubType = BoFldSubTypes.st_Time Then
                                If DateTime.TryParseExact(oCampoUsuario.Valor, "HHmm", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                    oDocDestino.Lines.BatchNumbers.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "HHmm", CultureInfo.CurrentCulture)
                            Else
                                If IsNumeric(oCampoUsuario.Valor) Then _
                                    oDocDestino.Lines.BatchNumbers.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CInt(oCampoUsuario.Valor)
                            End If

                        Case BoFieldTypes.db_Float
                            If IsNumeric(oCampoUsuario.Valor) Then _
                                oDocDestino.Lines.BatchNumbers.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CDbl(oCampoUsuario.Valor.Replace(".", ","))

                        Case BoFieldTypes.db_Date
                            If DateTime.TryParseExact(oCampoUsuario.Valor, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                oDocDestino.Lines.BatchNumbers.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                        Case Else
                            If Not String.IsNullOrEmpty(oCampoUsuario.Valor) Then _
                                oDocDestino.Lines.BatchNumbers.UserFields.Fields.Item(oCampoUsuario.Campo).Value = oCampoUsuario.Valor

                    End Select

                Next

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub setCamposUsuarioLineaCobroPago(ByRef objLinea As EntDocumentoLin,
                                               ByRef oDocDestino As Payments)

        'Campos de usuario

        Try

            If Not objLinea.CamposUsuario Is Nothing AndAlso objLinea.CamposUsuario.Count > 0 Then

                For Each oCampoUsuario In objLinea.CamposUsuario

                    Dim oUserField As Field = oDocDestino.AccountPayments.UserFields.Fields.Item(oCampoUsuario.Campo)

                    Select Case oUserField.Type

                        Case BoFieldTypes.db_Numeric
                            If oUserField.SubType = BoFldSubTypes.st_Time Then
                                If DateTime.TryParseExact(oCampoUsuario.Valor, "HHmm", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                        oDocDestino.AccountPayments.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "HHmm", CultureInfo.CurrentCulture)
                            Else
                                If IsNumeric(oCampoUsuario.Valor) Then _
                                        oDocDestino.AccountPayments.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CInt(oCampoUsuario.Valor)
                            End If

                        Case BoFieldTypes.db_Float
                            If IsNumeric(oCampoUsuario.Valor) Then _
                                    oDocDestino.AccountPayments.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CDbl(oCampoUsuario.Valor.Replace(".", ","))

                        Case BoFieldTypes.db_Date
                            If DateTime.TryParseExact(oCampoUsuario.Valor, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                    oDocDestino.AccountPayments.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                        Case Else
                            If Not String.IsNullOrEmpty(oCampoUsuario.Valor) Then _
                                    oDocDestino.AccountPayments.UserFields.Fields.Item(oCampoUsuario.Campo).Value = oCampoUsuario.Valor

                    End Select

                Next

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub setLotesLinea(ByRef objLinea As EntDocumentoLin,
                              ByRef oDocDestino As Documents)

        'Lote

        Try

            If Not objLinea.Lotes Is Nothing AndAlso objLinea.Lotes.Count > 0 Then

                For Each oLote In objLinea.Lotes

                    oDocDestino.Lines.BatchNumbers.SetCurrentLine(oDocDestino.Lines.BatchNumbers.Count - 1)

                    If Not String.IsNullOrEmpty(oLote.NumLote) Then oDocDestino.Lines.BatchNumbers.BatchNumber = oLote.NumLote

                    If oLote.Cantidad > 0 Then
                        oDocDestino.Lines.BatchNumbers.Quantity = oLote.Cantidad
                    Else
                        oDocDestino.Lines.BatchNumbers.Quantity = oDocDestino.Lines.Quantity
                    End If

                    If Not String.IsNullOrEmpty(oLote.Atributo1) Then oDocDestino.Lines.BatchNumbers.ManufacturerSerialNumber = oLote.Atributo1

                    If Not String.IsNullOrEmpty(oLote.Atributo2) Then oDocDestino.Lines.BatchNumbers.InternalSerialNumber = oLote.Atributo2

                    If DateTime.TryParseExact(oLote.FechaAdmision.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                                oDocDestino.Lines.BatchNumbers.AddmisionDate = Date.ParseExact(oLote.FechaAdmision.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                    If DateTime.TryParseExact(oLote.FechaFabricacion.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                              oDocDestino.Lines.BatchNumbers.ManufacturingDate = Date.ParseExact(oLote.FechaFabricacion.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                    If DateTime.TryParseExact(oLote.FechaVencimiento.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                              oDocDestino.Lines.BatchNumbers.ExpiryDate = Date.ParseExact(oLote.FechaVencimiento.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                    'Campos usuario 
                    setCamposUsuarioLote(oLote, oDocDestino)

                    oDocDestino.Lines.BatchNumbers.Add()

                Next

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub setPrimeraLineaEspecial(ByRef oDocEntrada As Documents,
                                        ByRef oDocDestino As Documents)

        'Traspasa la primera línea especial que va en la posición 0 de un documento

        Try

            For iLineaEspecial As Integer = 0 To oDocEntrada.SpecialLines.Count - 1

                oDocEntrada.SpecialLines.SetCurrentLine(iLineaEspecial)

                'Comprueba que la línea especial esté rellena
                If (oDocEntrada.SpecialLines.LineType = BoDocSpecialLineType.dslt_Text AndAlso Not String.IsNullOrEmpty(oDocEntrada.SpecialLines.LineText)) _
                    OrElse oDocEntrada.SpecialLines.LineType = BoDocSpecialLineType.dslt_Subtotal Then

                    If oDocEntrada.SpecialLines.AfterLineNumber < 0 Then

                        'Traspasa la línea especial
                        oDocDestino.SpecialLines.SetCurrentLine(oDocDestino.SpecialLines.Count - 1)
                        oDocDestino.SpecialLines.LineType = oDocEntrada.SpecialLines.LineType
                        oDocDestino.SpecialLines.AfterLineNumber = oDocDestino.Lines.Count - 2

                        If Not String.IsNullOrEmpty(oDocEntrada.SpecialLines.LineText) Then oDocDestino.SpecialLines.LineText = oDocEntrada.SpecialLines.LineText

                        oDocDestino.SpecialLines.Add()

                    End If

                End If

            Next

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub setLineasEspeciales(ByRef oDocEntrada As Documents,
                                    ByRef oDocDestino As Documents)

        'Traspasa las líneas especiales

        Try

            For iLineaEspecial As Integer = 0 To oDocEntrada.SpecialLines.Count - 1

                oDocEntrada.SpecialLines.SetCurrentLine(iLineaEspecial)

                'Comprueba que la línea especial esté rellena
                If (oDocEntrada.SpecialLines.LineType = BoDocSpecialLineType.dslt_Text AndAlso Not String.IsNullOrEmpty(oDocEntrada.SpecialLines.LineText)) _
                    OrElse oDocEntrada.SpecialLines.LineType = BoDocSpecialLineType.dslt_Subtotal Then

                    'NOTA: Usar VisualOrder o LineNum?
                    If oDocEntrada.SpecialLines.AfterLineNumber = oDocEntrada.Lines.LineNum Then

                        'Traspasa la línea especial
                        oDocDestino.SpecialLines.SetCurrentLine(oDocDestino.SpecialLines.Count - 1)
                        oDocDestino.SpecialLines.LineType = oDocEntrada.SpecialLines.LineType
                        oDocDestino.SpecialLines.AfterLineNumber = oDocDestino.Lines.Count - 2

                        If Not String.IsNullOrEmpty(oDocEntrada.SpecialLines.LineText) Then oDocDestino.SpecialLines.LineText = oDocEntrada.SpecialLines.LineText

                        oDocDestino.SpecialLines.Add()

                    End If

                End If

            Next

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Sub setMedioCobroPago(ByRef objDocumento As Documents,
                                  ByRef oDocDestino As Payments,
                                  ByVal Sociedad As eSociedad)

        'Medio pago

        Try

            Select Case objDocumento.CobroPagoTipo

                Case CobroPago.Efectivo

                    oDocDestino.CashAccount = objDocumento.CobroPagoCuenta

                    oDocDestino.CashSum = objDocumento.CobroPagoImporte

                Case CobroPago.Efecto

                    If Not DateTime.TryParseExact(objDocumento.DocDueDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        Throw New Exception("Fecha vencimiento efecto no suministrada o incorrecta")

                    If String.IsNullOrEmpty(objDocumento.CobroPagoViaPago) Then Throw New Exception("Vía pago efecto no suministrada")
                    If String.IsNullOrEmpty(objDocumento.CobroPagoBanco) Then Throw New Exception("Banco efecto no suministrado")
                    If String.IsNullOrEmpty(objDocumento.CobroPagoPais) Then Throw New Exception("País efecto no suministrado")

                    oDocDestino.BillOfExchange.BillOfExchangeNo = getCobroPagoEfectoNumero(Sociedad)

                    oDocDestino.BillOfExchangeAmount = objDocumento.CobroPagoImporte

                    oDocDestino.BillOfExchange.BillOfExchangeDueDate = Date.ParseExact(objDocumento.DocDueDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                    If Not String.IsNullOrEmpty(objDocumento.CobroPagoReferencia) Then oDocDestino.BillOfExchange.ReferenceNo = objDocumento.CobroPagoReferencia

                    oDocDestino.BillOfExchange.PaymentMethodCode = objDocumento.CobroPagoViaPago

                    oDocDestino.BillOfExchange.BPBankAct = objDocumento.CobroPagoCuenta

                    oDocDestino.BillOfExchange.BPBankCode = objDocumento.CobroPagoBanco

                    oDocDestino.BillOfExchange.BPBankCountry = objDocumento.CobroPagoPais

                Case CobroPago.Transferencia

                    oDocDestino.TransferAccount = objDocumento.CobroPagoCuenta

                    oDocDestino.TransferSum = objDocumento.CobroPagoImporte

                Case CobroPago.Cheque

                    If String.IsNullOrEmpty(objDocumento.CobroPagoBanco) Then Throw New Exception("Banco cheque no suministrado")
                    If Not objDocumento.CobroPagoNumCheque > 0 Then Throw New Exception("Número cheque no suministrado")

                    If Not DateTime.TryParseExact(objDocumento.DocDueDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        Throw New Exception("Fecha vencimiento cheque no suministrada o incorrecta")

                    oDocDestino.Checks.CheckSum = objDocumento.CobroPagoImporte

                    oDocDestino.Checks.AccounttNum = objDocumento.CobroPagoCuenta

                    oDocDestino.Checks.BankCode = objDocumento.CobroPagoBanco

                    oDocDestino.Checks.CheckNumber = objDocumento.CobroPagoNumCheque

                    oDocDestino.Checks.DueDate = Date.ParseExact(objDocumento.DocDueDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                Case CobroPago.Tarjeta

                    If Not objDocumento.CobroPagoTarjeta > 0 Then Throw New Exception("Código tarjeta no suministrado")

                    If Not DateTime.TryParseExact(objDocumento.CobroPagoValidezTarjeta.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        Throw New Exception("Fecha validez tarjeta no suministrada o incorrecta")

                    'If String.IsNullOrEmpty(objDocumento.CobroPagoNumTarjeta) Then Throw New Exception("Número tarjeta no suministrado")
                    If String.IsNullOrEmpty(objDocumento.CobroPagoComprobanteTarjeta) Then Throw New Exception("Comprobante tarjeta no suministrado")

                    oDocDestino.CreditCards.CreditCard = objDocumento.CobroPagoTarjeta

                    oDocDestino.CreditCards.CreditAcct = objDocumento.CobroPagoCuenta

                    oDocDestino.CreditCards.CreditSum = objDocumento.CobroPagoImporte

                    'oDocDestino.CreditCards.CreditCardNumber = objDocumento.CobroPagoNumTarjeta

                    oDocDestino.CreditCards.CardValidUntil = Date.ParseExact(objDocumento.CobroPagoValidezTarjeta.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                    oDocDestino.CreditCards.VoucherNum = objDocumento.CobroPagoComprobanteTarjeta

                Case Else

                    Throw New Exception("Tipo cobro/pago no permitido: " & objDocumento.CobroPagoTipo)

            End Select

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub



#Region "Privadas: Obtener campo"

    Private Function getDocNumDocumentoDefinitivo(ByVal Tabla As String,
                                                  ByVal CardCode As String,
                                                  ByVal DOCIDDW As String,
                                                  ByVal DocNum As String,
                                                  ByVal NumAtCard As String,
                                                  ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un documento definitivo

        Dim retVal As String = ""

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("DocNum") & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            If bICLinea Then _
                SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " ON T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getDocTotalDocumentoDefinitivo(ByVal Tabla As String,
                                                    ByVal CardCode As String,
                                                    ByVal DOCIDDW As String,
                                                    ByVal DocNum As String,
                                                    ByVal NumAtCard As String,
                                                    ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un documento definitivo

        Dim retVal As String = ""

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("DocTotal") & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            If bICLinea Then _
                SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " On T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf


            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString.Replace(",", ".")

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getDocDueDateDocumentoDefinitivo(ByVal Tabla As String,
                                                      ByVal CardCode As String,
                                                      ByVal DOCIDDW As String,
                                                      ByVal DocNum As String,
                                                      ByVal NumAtCard As String,
                                                      ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un documento definitivo

        Dim retVal As String = ""

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""

            SQL = "  SELECT " & vbCrLf
            SQL &= " " & getDateAsString_yyyyMMdd("T0.", "DocDueDate") & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            If bICLinea Then _
                SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " On T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getResponsableMailDocumentoDefinitivo(ByVal Tabla As String,
                                                           ByVal CardCode As String,
                                                           ByVal DOCIDDW As String,
                                                           ByVal DocNum As String,
                                                           ByVal NumAtCard As String,
                                                           ByVal Sociedad As eSociedad) As String

        'Busca y devuelve el correo del responsable del documento

        Dim retVal As String = ""

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por DocNum, DocEntry o NumAtCard
            Dim SQL As String = ""

            SQL = "  SELECT " & vbCrLf
            SQL &= " COALESCE(T2." & putQuotes("email") & ",'') " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            If bICLinea Then
                SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " ON T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf
                SQL &= " LEFT JOIN " & getDataBaseRef("OHEM", Sociedad) & " T2 " & getWithNoLock() & " ON T1." & putQuotes("SlpCode") & " = T2." & putQuotes("salesPrson") & vbCrLf
            Else
                SQL &= " LEFT JOIN " & getDataBaseRef("OHEM", Sociedad) & " T2 " & getWithNoLock() & " ON T0." & putQuotes("SlpCode") & " = T2." & putQuotes("salesPrson") & vbCrLf
            End If

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getComentariosDocumentoDefinitivo(ByVal Tabla As String,
                                                       ByVal CardCode As String,
                                                       ByVal DOCIDDW As String,
                                                       ByVal DocNum As String,
                                                       ByVal NumAtCard As String,
                                                       ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un documento definitivo

        Dim retVal As String = ""

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("Comments") & ",'') " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            If bICLinea Then _
                SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " ON T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getDocDueDateTotalDocumentoDefinitivo(ByVal Tabla As String,
                                                           ByVal CardCode As String,
                                                           ByVal DOCIDDW As String,
                                                           ByVal DocNum As String,
                                                           ByVal NumAtCard As String,
                                                           ByVal Sociedad As eSociedad) As List(Of String)

        'Comprueba si existe un documento definitivo

        Dim retVal As New List(Of String)

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""

            SQL = "  SELECT " & vbCrLf
            SQL &= " COALESCE(" & getDateAsString_yyyyMMdd("T6.", "DueDate") & ", " & getDateAsString_yyyyMMdd("T0.", "DocDueDate") & ") AS DocDueDate," & vbCrLf
            SQL &= " COALESCE(T6." & putQuotes("InsTotal") & ", T0." & putQuotes("DocTotal") & ") As DocDueTotal " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "6", Sociedad) & " T6 " & getWithNoLock() & " On T6." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            If bICLinea Then _
                SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " On T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            SQL &= "GROUP BY " & vbCrLf

            If bICLinea Then SQL &= "T1." & putQuotes("DocEntry") & ", " & vbCrLf

            SQL &= "T6." & putQuotes("DueDate") & ", " & vbCrLf
            SQL &= "T0." & putQuotes("DocDueDate") & ", " & vbCrLf
            SQL &= "T6." & putQuotes("InsTotal") & ", " & vbCrLf
            SQL &= "T0." & putQuotes("DocTotal") & ", " & vbCrLf
            SQL &= "T6." & putQuotes("DocEntry") & ", " & vbCrLf
            SQL &= "T0." & putQuotes("DocEntry") & vbCrLf

            SQL &= "ORDER BY " & vbCrLf

            SQL &= "T6." & putQuotes("DueDate") & ", " & vbCrLf
            SQL &= "T6." & putQuotes("InsTotal") & ", " & vbCrLf
            SQL &= "T0." & putQuotes("DocDueDate") & ", " & vbCrLf
            SQL &= "T0." & putQuotes("DocTotal") & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            Dim DT As DataTable = oCon.ObtenerDT(SQL)

            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(DT.Rows(0).Item("DocDueDate").ToString & "-" & DT.Rows(0).Item("DocDueTotal").ToString.Replace(",", "."))
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getNotPaidToDateDocumentoDefinitivo(ByVal Tabla As String,
                                                         ByVal CardCode As String,
                                                         ByVal DocDueDate As Integer,
                                                         ByVal DOCIDDW As String,
                                                         ByVal DocNum As String,
                                                         ByVal NumAtCard As String,
                                                         ByVal Sociedad As eSociedad) As String

        'Comprueba si se ha pagado una factura
        Dim retVal As String = ""

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String

            SQL = "  SELECT " & vbCrLf
            SQL &= " COALESCE(T0." & putQuotes("DocTotal") & ",0) - COALESCE(T6." & putQuotes("PaidToDate") & ",0) As NotPaidToDate " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3), Sociedad) & "6 T6 " & getWithNoLock() & " On T6." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            If bICLinea Then _
                SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " On T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            SQL &= " And T6." & putQuotes("DueDate") & "  = N'" & DocDueDate & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            SQL &= " GROUP BY "

            If bICLinea Then SQL &= "T1." & putQuotes("DocEntry") & ", " & vbCrLf

            SQL &= "T0." & putQuotes("DocEntry") & ", " & vbCrLf
            SQL &= "T6." & putQuotes("DocEntry") & ", " & vbCrLf
            SQL &= "T0." & putQuotes("DocTotal") & ", " & vbCrLf
            SQL &= "T6." & putQuotes("PaidToDate") & " " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            Dim oObj As Object = oCon.ExecuteScalar(SQL)

            If Not oObj Is Nothing AndAlso IsNumeric(oObj) Then
                retVal = oObj.ToString.Replace(",", ".")
            Else
                Throw New Exception("No encuentro vencimiento para el documento con IDDW: " & DOCIDDW & ", DocNum: " & DocNum & ", NumAtCard: " & NumAtCard & " y DocDueDate: " & DocDueDate)
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getBaseTotalDocumentoDefinitivo(ByVal Tabla As String,
                                                     ByVal CardCode As String,
                                                     ByVal DOCIDDW As String,
                                                     ByVal DocNum As String,
                                                     ByVal NumAtCard As String,
                                                     ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un documento definitivo

        Dim retVal As String = ""

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("DocTotal") & " - T0." & putQuotes("VatSum") & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            If bICLinea Then _
                SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " On T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString.Replace(",", ".")

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getViaPagoDocumentoDefinitivo(ByVal Tabla As String,
                                                   ByVal CardCode As String,
                                                   ByVal DOCIDDW As String,
                                                   ByVal DocNum As String,
                                                   ByVal NumAtCard As String,
                                                   ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un documento definitivo
        Dim retVal As String = ""

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String
            SQL = "  SELECT "
            SQL &= " CASE "
            SQL &= "    WHEN COALESCE(T2." & putQuotes("Descript") & ",'')<>N'' THEN COALESCE(T2." & putQuotes("Descript") & ",'') "
            SQL &= "    WHEN COALESCE(T0." & putQuotes("GroupNum") & ",-1)=-1 AND COALESCE(T0." & putQuotes("PaidSum") & ",0)>0  THEN COALESCE(T3." & putQuotes("PymntGroup") & ",'') "
            SQL &= "    ELSE '' "
            SQL &= " END "
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & " "

            If bICLinea Then
                SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " ON T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & " "
            End If

            SQL &= " LEFT JOIN " & getDataBaseRef("OPYM", Sociedad) & " T2 " & getWithNoLock() & " ON T2." & putQuotes("PayMethCod") & " = T0." & putQuotes("PeyMethod") & " "
            SQL &= " LEFT JOIN " & getDataBaseRef("OCTG", Sociedad) & " T3 " & getWithNoLock() & " ON T3." & putQuotes("GroupNum") & " = T0." & putQuotes("GroupNum") & " "

            SQL &= " WHERE 1=1 "

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'"

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'"

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'"

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getMonedaDocumentoDefinitivo(ByVal Tabla As String,
                                                  ByVal CardCode As String,
                                                  ByVal DOCIDDW As String,
                                                  ByVal DocNum As String,
                                                  ByVal NumAtCard As String,
                                                  ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un documento definitivo
        Dim retVal As String = ""

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String
            SQL = "  SELECT "
            SQL &= " COALESCE(T0." & putQuotes("DocCur") & ",'') "
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & " "

            If bICLinea Then _
                SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " ON T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & " "

            SQL &= " WHERE 1=1 "

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'"

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'"

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'"

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getNumTransaccionDocumentoDefinitivo(ByVal Tabla As String,
                                                          ByVal CardCode As String,
                                                          ByVal DOCIDDW As String,
                                                          ByVal DocNum As String,
                                                          ByVal NumAtCard As String,
                                                          ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un documento definitivo

        Dim retVal As String = ""

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("TransId") & vbCrLf
            SQL &= " FROM " & Utilidades.getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            If bICLinea Then _
                SQL &= " JOIN " & Utilidades.getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " ON T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getProyectoDocumentoDefinitivo(ByVal Tabla As String,
                                                    ByVal CardCode As String,
                                                    ByVal DOCIDDW As String,
                                                    ByVal DocNum As String,
                                                    ByVal NumAtCard As String,
                                                    ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un documento definitivo

        Dim retVal As String = ""

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " T1." & putQuotes("Project") & vbCrLf
            SQL &= " FROM " & Utilidades.getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " JOIN " & Utilidades.getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " ON T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            SQL &= " ORDER BY T1." & putQuotes("VisOrder") & " ASC" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            Dim oObj As Object = oCon.ExecuteScalar(SQL)

            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then
                retVal = oObj.ToString
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getCampoUsuarioDocumentoDefinitivo(ByVal Tabla As String,
                                                        ByVal CardCode As String,
                                                        ByVal DOCIDDW As String,
                                                        ByVal DocNum As String,
                                                        ByVal NumAtCard As String,
                                                        ByVal CampoUsuario As String,
                                                        ByVal Nivel As Integer,
                                                        ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un documento definitivo

        Dim retVal As String = ""

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT TOP 1 " & vbCrLf

            If Nivel = 0 Then _
                SQL &= " T0." & putQuotes(CampoUsuario) & vbCrLf

            If Nivel > 0 Then _
                SQL &= " T2." & putQuotes(CampoUsuario) & vbCrLf

            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            If bICLinea Then _
                SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " ON T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            If Nivel > 0 Then _
                SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & Nivel, Sociedad) & " T2 " & getWithNoLock() & " ON T2." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            Dim oObj As Object = oCon.ExecuteScalar(SQL)

            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj.ToString) Then
                retVal = oObj.ToString
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & System.Reflection.MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getSucursalDocumentoDefinitivo(ByVal Tabla As String,
                                                    ByVal CardCode As String,
                                                    ByVal DOCIDDW As String,
                                                    ByVal DocNum As String,
                                                    ByVal NumAtCard As String,
                                                    ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un documento definitivo

        Dim retVal As String = ""

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("BPLId") & vbCrLf
            SQL &= " FROM " & Utilidades.getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            If bICLinea Then _
                SQL &= " JOIN " & Utilidades.getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " ON T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso IsNumeric(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getCentroCosteDocumentoDefinitivo(ByVal Tabla As String,
                                                       ByVal CardCode As String,
                                                       ByVal DOCIDDW As String,
                                                       ByVal DocNum As String,
                                                       ByVal NumAtCard As String,
                                                       ByVal Dimension As Integer,
                                                       ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un documento definitivo

        Dim retVal As String = ""

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT TOP 1 " & vbCrLf

            Select Case Dimension
                Case 5
                    SQL &= " T1." & putQuotes("OcrCode5") & vbCrLf
                Case 4
                    SQL &= " T1." & putQuotes("OcrCode4") & vbCrLf
                Case 3
                    SQL &= " T1." & putQuotes("OcrCode3") & vbCrLf
                Case 2
                    SQL &= " T1." & putQuotes("OcrCode2") & vbCrLf
                Case Else
                    SQL &= " T1." & putQuotes("OcrCode") & vbCrLf
            End Select

            SQL &= " FROM " & Utilidades.getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " JOIN " & Utilidades.getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " ON T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            SQL &= " ORDER BY T1." & putQuotes("VisOrder") & " ASC" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getNumEnvioDocumentoDefinitivo(ByVal Tabla As String,
                                                    ByVal CardCode As String,
                                                    ByVal DOCIDDW As String,
                                                    ByVal DocNum As String,
                                                    ByVal NumAtCard As String,
                                                    ByVal Dimension As Integer,
                                                    ByVal Sociedad As eSociedad) As String

        'Comprueba si existe un documento definitivo

        Dim retVal As String = ""

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por CardCode y NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT TOP 1" & vbCrLf
            SQL &= " T0." & putQuotes("TrackNo") & vbCrLf
            SQL &= " FROM " & Utilidades.getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getNumDestinoDocumentoDefinitivo(ByVal TablaOrigen As String,
                                                      ByVal TablaDestino As String,
                                                      ByVal CardCode As String,
                                                      ByVal DOCIDDW As String,
                                                      ByVal DocNum As String,
                                                      ByVal NumAtCard As String,
                                                      ByVal NumEnvio As String,
                                                      ByVal Sociedad As eSociedad) As List(Of String)

        'Comprueba si existe un documento definitivo
        Dim retVal As New List(Of String)

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(TablaOrigen = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por CardCode, NumAtCard, IDDW y TrackNo
            Dim SQL As String
            SQL = "  SELECT " & vbCrLf
            SQL &= " COALESCE(T3." & putQuotes("DocNum") & ",'') As DocNum " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(TablaOrigen, Sociedad) & " T0 " & getWithNoLock() & " " & vbCrLf

            If bICLinea Then
                SQL &= " JOIN " & getDataBaseRef(TablaOrigen.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " On T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & " " & vbCrLf
            Else
                SQL &= " JOIN " & getDataBaseRef(TablaOrigen.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " On T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & " " & vbCrLf
            End If

            SQL &= " JOIN " & getDataBaseRef(TablaDestino.Substring(1, 3) & "1", Sociedad) & " T2 " & getWithNoLock() & " On T2." & putQuotes("BaseType") & " = T0." & putQuotes("ObjType")
            SQL &= "                                                                                And T2." & putQuotes("BaseEntry") & " = T0." & putQuotes("DocEntry")
            SQL &= "                                                                                And T2." & putQuotes("BaseLine") & " = T1." & putQuotes("LineNum") & vbCrLf

            SQL &= " JOIN " & getDataBaseRef(TablaDestino, Sociedad) & " T3 " & getWithNoLock() & " On T3." & putQuotes("DocEntry") & " = T2." & putQuotes("DocEntry") & " " & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumEnvio) Then SQL &= " And T0." & putQuotes("TrackNo") & " = N'" & NumEnvio & "'" & vbCrLf

            'SQL &= " And T3." & putQuotes("DocStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf
            SQL &= " And T3." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim DT As DataTable = oCon.ObtenerDT(SQL)
            If Not DT Is Nothing And DT.Rows.Count > 0 Then
                For Each DR In DT.Rows
                    retVal.Add(DR.Item("DocNum").ToString)
                Next
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function

    Private Function getTitularDocumentoDefinitivo(ByVal Tabla As String,
                                                   ByVal CardCode As String,
                                                   ByVal DOCIDDW As String,
                                                   ByVal DocNum As String,
                                                   ByVal NumAtCard As String,
                                                   ByVal Sociedad As eSociedad) As String

        'Busca y devuelve el titular del documento

        Dim retVal As String = ""

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por DocNum, DocEntry o NumAtCard
            Dim SQL As String = ""

            SQL = "  SELECT " & vbCrLf
            SQL &= " COALESCE(T2." & putQuotes("empID") & ",-1) " & vbCrLf
            SQL &= " FROM " & getDataBaseRef(Tabla, Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            If bICLinea Then
                SQL &= " JOIN " & getDataBaseRef(Tabla.Substring(1, 3) & "1", Sociedad) & " T1 " & getWithNoLock() & " ON T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf
                SQL &= " LEFT JOIN " & getDataBaseRef("OHEM", Sociedad) & " T2 " & getWithNoLock() & " ON T1." & putQuotes("OwnerCode") & " = T2." & putQuotes("empID") & vbCrLf
            Else
                SQL &= " LEFT JOIN " & getDataBaseRef("OHEM", Sociedad) & " T2 " & getWithNoLock() & " ON T0." & putQuotes("OwnerCode") & " = T2." & putQuotes("empID") & vbCrLf
            End If

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If IsNumeric(oObj) AndAlso CInt(oObj) > 0 Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

        Return retVal

    End Function


#End Region

#End Region

End Class
