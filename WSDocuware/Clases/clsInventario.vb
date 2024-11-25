Imports SAPbobsCOM
Imports System.Reflection
Imports System.Globalization
Imports WSDocuware.Utilidades

Public Class clsInventario

#Region "Públicas"

    Public Function CrearInventario(ByVal objInventario As EntInventarioCab, ByVal Sociedad As eSociedad) As EntResultado

        'Crea un documento como en borrador o en firme en base a otros o directamente 
        Dim retVal As New EntResultado

        Try

            If objInventario.ObjTypeDestino = ObjType.Traslado Then
                'Crea traslado
                retVal = NuevoTraslado(objInventario, Sociedad)
            Else
                'Crea entrd/salida mercancías
                retVal = NuevoInventario(objInventario, Sociedad)
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Public Function ActualizarInventario(ByVal objInventario As EntInventarioCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "ACTUALIZAR INVENTARIO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de documento: " & objInventario.Ref2 & " para IDDW " & objInventario.DOCIDDW)

            'Obligatorios
            If String.IsNullOrEmpty(objInventario.Ref2) AndAlso String.IsNullOrEmpty(objInventario.DocNum) Then Throw New Exception("Referencia y/o número de documento no suministrados")
            If String.IsNullOrEmpty(objInventario.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objInventario.DOCURLDW) Then Throw New Exception("URL Docuware no suministrada")

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objInventario.ObjTypeDestino)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objInventario.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'RefsOrigen contiene los Ref2
            Dim RefsOrigen As String() = objInventario.RefOrigen.Split("#")

            'Recorre cada uno de los documento origen
            For Each RefOrigen In RefsOrigen

                'Comprueba que existe el documento origen y que se puede acceder a él
                'Pueden ser varios documentos origen con la misma referencia
                Dim DocNums As List(Of String) = getDocNumDeRefOrigen(Tabla, RefOrigen, objInventario.TipoRefOrigen, Sociedad)
                If DocNums Is Nothing OrElse DocNums.Count = 0 Then Throw New Exception("No existe documento con referencia: " & RefOrigen & " o su estado es cancelado")

                'Se actualizan los campos de usuario de DW
                setDatosDWdocumento(Tabla, objInventario.DocDate, DocNums, objInventario.DOCIDDW, objInventario.DOCURLDW, Sociedad)

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento actualizado con éxito"
                retVal.MENSAJEAUX = String.Join("#", DocNums)

                clsLog.Log.Info(retVal.CODIGO & " - " & retVal.MENSAJE & " - " & retVal.MENSAJEAUX & " - " & retVal.MENSAJEAUX & " - " & retVal.MENSAJEAUX)

            Next

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Public Function getComprobarInventarioID(ByVal IDDW As String,
                                             ByVal ObjType As Integer,
                                             ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado

        Dim sLogInfo As String = "INVENTARIO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar documento para IDDW: " & IDDW)

            'Obligatorios
            If String.IsNullOrEmpty(IDDW) Then Throw New Exception("ID docuware no suministrado")

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el DocNum del inventario definitivo
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

    Public Function getComprobarInventario(ByVal ObjType As Integer,
                                           ByVal DOCIDDW As String,
                                           ByVal DocNum As String,
                                           ByVal Ref2 As String,
                                           ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un documento en firme

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "INVENTARIO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Comprobar documento para ObjType: " & ObjType & ", DOCIDDW: " & DOCIDDW & ", DocNum: " & DocNum & ", Ref2: " & Ref2)

            'Obligatorios
            If String.IsNullOrEmpty(DOCIDDW) AndAlso String.IsNullOrEmpty(DocNum) AndAlso String.IsNullOrEmpty(Ref2) Then _
                Throw New Exception("ID Docuware, número de documento o referencia origen no suministrados")

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(ObjType)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & ObjType)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Devuelve el DocNum del documento definitivo
            Dim sDocNum As String = getDocNumDocumentoDefinitivo(Tabla, DOCIDDW, DocNum, Ref2, Sociedad)

            If Not String.IsNullOrEmpty(sDocNum) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Documento en firme encontrado"
                retVal.MENSAJEAUX = sDocNum

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

#End Region

#Region "Privadas"

    Private Function NuevoTraslado(ByVal objInventario As EntInventarioCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oDocDestino As StockTransfer = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUEVO TRASLADO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de documento para " & objInventario.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objInventario.UserSAP, objInventario.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objInventario.NIFTercero) Then Throw New Exception("NIF no suministrado")

            If Not DateTime.TryParseExact(objInventario.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha documento no suministrada o incorrecta")

            If String.IsNullOrEmpty(objInventario.AlmacenOrigen) Then Throw New Exception("Almacén origen no suministrado")
            If String.IsNullOrEmpty(objInventario.AlmacenDestino) Then Throw New Exception("Almacén destino no suministrado")
            If String.IsNullOrEmpty(objInventario.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objInventario.DOCURLDW) Then Throw New Exception("URL Docuware no suministrado")

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objInventario.ObjTypeDestino)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objInventario.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objInventario.NIFTercero, objInventario.RazonSocial, objInventario.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objInventario.NIFTercero & ", Razón social: " & objInventario.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Objeto documento
            If objInventario.Draft <> Draft.Firme Then
                oDocDestino = oCompany.GetBusinessObject(BoObjectTypes.oStockTransferDraft)
                oDocDestino.DocObjectCode = objInventario.ObjTypeDestino
                clsLog.Log.Info("(" & sLogInfo & ") Documento borrador")
            Else
                oDocDestino = oCompany.GetBusinessObject(BoObjectTypes.oStockTransfer)
                oDocDestino.DocObjectCode = objInventario.ObjTypeDestino
                clsLog.Log.Info("(" & sLogInfo & ") Documento en firme")
            End If

            'CardCode
            oDocDestino.CardCode = CardCode

            'Fecha documento
            oDocDestino.DocDate = Date.ParseExact(objInventario.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Comentarios
            If objInventario.Comments.Length > 254 Then objInventario.Comments = objInventario.Comments.Substring(0, 254)
            oDocDestino.Comments = objInventario.Comments

            'Entrada diario
            If objInventario.JournalMemo.Length > 254 Then objInventario.JournalMemo = objInventario.JournalMemo.Substring(0, 254)
            oDocDestino.JournalMemo = objInventario.JournalMemo

            'Referencia
            If objInventario.Ref2.Length > 11 Then objInventario.Ref2 = objInventario.Ref2.Substring(0, 11)
            oDocDestino.Reference2 = objInventario.Ref2

            'Almacenes
            oDocDestino.FromWarehouse = objInventario.AlmacenOrigen
            oDocDestino.ToWarehouse = objInventario.AlmacenDestino

            'Recorre cada línea 
            For Each objLinea In objInventario.Lineas

                'Índice
                Dim indLinea As Integer = objInventario.Lineas.IndexOf(objLinea) + 1

                'Comprobar campos
                If String.IsNullOrEmpty(objLinea.Articulo) AndAlso String.IsNullOrEmpty(objLinea.RefExt) Then _
                    Throw New Exception("Artículo o referencia externa de la línea " & indLinea.ToString & " no suministrado")

                If Not objLinea.Cantidad > 0 Then Throw New Exception("Cantidad de la línea " & indLinea.ToString & " no suministrado")

                'Se posiciona en la línea
                oDocDestino.Lines.SetCurrentLine(oDocDestino.Lines.Count - 1)

                'Artículo (si no se encuentra el código se busca por la referencia de proveedor)
                Dim ItemCode As String = oComun.getItemCode(CardCode, objLinea.Articulo, objLinea.RefExt, Sociedad)
                If String.IsNullOrEmpty(ItemCode) Then Throw New Exception("No encuentro artículo de la línea " & indLinea.ToString & " con código: " & objLinea.Articulo & " o referencia externa: " & objLinea.RefExt & " para IC: " & CardCode)
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado artículo: " & ItemCode)

                oDocDestino.Lines.ItemCode = ItemCode

                'Descripción
                If objLinea.Concepto.Length > 100 Then objLinea.Concepto = objLinea.Concepto.Substring(0, 100)
                oDocDestino.Lines.ItemDescription = objLinea.Concepto

                'Cantidad
                If objLinea.Cantidad > 0 Then oDocDestino.Lines.Quantity = objLinea.Cantidad

                'Precio
                If objLinea.PrecioUnidad <> 0 Then
                    oDocDestino.Lines.UnitPrice = objLinea.PrecioUnidad
                    oDocDestino.Lines.Price = objLinea.PrecioUnidad
                End If

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

                'Lote
                If Not objLinea.Lotes Is Nothing AndAlso objLinea.Lotes.Count > 0 Then

                    For Each oLote In objLinea.Lotes

                        oDocDestino.Lines.BatchNumbers.SetCurrentLine(oDocDestino.Lines.BatchNumbers.Count - 1)

                        If Not String.IsNullOrEmpty(oLote.NumLote) Then oDocDestino.Lines.BatchNumbers.BatchNumber = oLote.NumLote

                        If oLote.Cantidad > 0 Then oDocDestino.Lines.BatchNumbers.Quantity = oLote.Cantidad

                        If Not String.IsNullOrEmpty(oLote.Atributo1) Then oDocDestino.Lines.BatchNumbers.ManufacturerSerialNumber = oLote.Atributo1

                        If Not String.IsNullOrEmpty(oLote.Atributo2) Then oDocDestino.Lines.BatchNumbers.InternalSerialNumber = oLote.Atributo2

                        If DateTime.TryParseExact(oLote.FechaAdmision.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                            oDocDestino.Lines.BatchNumbers.AddmisionDate = Date.ParseExact(oLote.FechaAdmision.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                        If DateTime.TryParseExact(oLote.FechaFabricacion.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                            oDocDestino.Lines.BatchNumbers.ManufacturingDate = Date.ParseExact(oLote.FechaFabricacion.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                        If DateTime.TryParseExact(oLote.FechaVencimiento.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                            oDocDestino.Lines.BatchNumbers.ExpiryDate = Date.ParseExact(oLote.FechaVencimiento.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                        oDocDestino.Lines.BatchNumbers.Add()

                    Next

                End If

                'Añade línea
                oDocDestino.Lines.Add()

            Next

            'ID DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIDDW").Value = objInventario.DOCIDDW

            'URL DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIURLDW").Value = objInventario.DOCURLDW

            'Añadimos documento
            If oDocDestino.Add() = 0 Then

                'Obtiene el DocEntry y DocNum del documento añadido
                Dim DocEntry As String = oCompany.GetNewObjectKey

                Dim DocNum As String = ""
                If objInventario.Draft = Draft.Firme Then DocNum = oComun.getDocNumDeDocEntry(TablaDestino, DocEntry, Sociedad)

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

    Private Function NuevoInventario(ByVal objInventario As EntInventarioCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oDocDestino As Documents = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUEVO INVENTARIO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de documento para " & objInventario.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objInventario.UserSAP, objInventario.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If Not DateTime.TryParseExact(objInventario.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha documento no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objInventario.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha contable no suministrada o incorrecta")

            If String.IsNullOrEmpty(objInventario.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objInventario.DOCURLDW) Then Throw New Exception("URL Docuware no suministrado")

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objInventario.ObjTypeDestino)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objInventario.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objInventario.NIFTercero, objInventario.RazonSocial, objInventario.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then
                clsLog.Log.Info("(" & sLogInfo & ") No encuentro IC con NIF: " & objInventario.NIFTercero & ", Razón social: " & objInventario.RazonSocial)
                'Throw New Exception("No encuentro IC con NIF: " & objInventario.NIFTercero & ", Razón social: " & objInventario.RazonSocial)
            Else
                clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)
            End If

            'Objeto documento
            If objInventario.Draft <> Draft.Firme Then
                oDocDestino = oCompany.GetBusinessObject(BoObjectTypes.oDrafts)
                oDocDestino.DocObjectCode = objInventario.ObjTypeDestino
                clsLog.Log.Info("(" & sLogInfo & ") Documento borrador")
            Else
                oDocDestino = oCompany.GetBusinessObject(objInventario.ObjTypeDestino)
                clsLog.Log.Info("(" & sLogInfo & ") Documento en firme")
            End If

            'Fecha documento y contable
            oDocDestino.DocDate = Date.ParseExact(objInventario.DocDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            oDocDestino.TaxDate = Date.ParseExact(objInventario.AccountDate.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'Comentarios
            If objInventario.Comments.Length > 254 Then objInventario.Comments = objInventario.Comments.Substring(0, 254)
            oDocDestino.Comments = objInventario.Comments

            'Entrada diario
            If objInventario.JournalMemo.Length > 254 Then objInventario.JournalMemo = objInventario.JournalMemo.Substring(0, 254)
            oDocDestino.JournalMemo = objInventario.JournalMemo

            'Referencia
            If objInventario.Ref2.Length > 11 Then objInventario.Ref2 = objInventario.Ref2.Substring(0, 11)
            oDocDestino.Reference2 = objInventario.Ref2

            'Recorre cada línea 
            For Each objLinea In objInventario.Lineas

                'Índice
                Dim indLinea As Integer = objInventario.Lineas.IndexOf(objLinea) + 1

                'Comprobar campos
                If String.IsNullOrEmpty(objLinea.Articulo) AndAlso String.IsNullOrEmpty(objLinea.RefExt) Then _
                    Throw New Exception("Artículo o referencia externa de la línea " & indLinea.ToString & " no suministrado")

                If Not objLinea.Cantidad > 0 Then Throw New Exception("Cantidad de la línea " & indLinea.ToString & " no suministrado")
                If String.IsNullOrEmpty(objLinea.Almacen) Then Throw New Exception("Almacén de la línea " & indLinea.ToString & " no suministrado")

                'Se posiciona en la línea
                oDocDestino.Lines.SetCurrentLine(oDocDestino.Lines.Count - 1)

                'Artículo (si no se encuentra el código se busca por la referencia de proveedor)
                If Not String.IsNullOrEmpty(CardCode) Then
                    Dim ItemCode As String = oComun.getItemCode(CardCode, objLinea.Articulo, objLinea.RefExt, Sociedad)
                    If String.IsNullOrEmpty(ItemCode) Then Throw New Exception("No encuentro artículo de la línea " & indLinea.ToString & " con código: " & objLinea.Articulo & " o referencia externa: " & objLinea.RefExt & " para IC: " & CardCode)
                    clsLog.Log.Info("(" & sLogInfo & ") Encontrado artículo: " & ItemCode)
                    oDocDestino.Lines.ItemCode = ItemCode
                Else
                    oDocDestino.Lines.ItemCode = objLinea.Articulo
                End If

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

                'Almacenes
                oDocDestino.Lines.WarehouseCode = objLinea.Almacen

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

                'Lote
                If Not objLinea.Lotes Is Nothing AndAlso objLinea.Lotes.Count > 0 Then

                    For Each oLote In objLinea.Lotes

                        oDocDestino.Lines.BatchNumbers.SetCurrentLine(oDocDestino.Lines.BatchNumbers.Count - 1)

                        If Not String.IsNullOrEmpty(oLote.NumLote) Then oDocDestino.Lines.BatchNumbers.BatchNumber = oLote.NumLote

                        If oLote.Cantidad > 0 Then oDocDestino.Lines.BatchNumbers.Quantity = oLote.Cantidad

                        If Not String.IsNullOrEmpty(oLote.Atributo1) Then oDocDestino.Lines.BatchNumbers.ManufacturerSerialNumber = oLote.Atributo1

                        If Not String.IsNullOrEmpty(oLote.Atributo2) Then oDocDestino.Lines.BatchNumbers.InternalSerialNumber = oLote.Atributo2

                        If DateTime.TryParseExact(oLote.FechaAdmision.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                            oDocDestino.Lines.BatchNumbers.AddmisionDate = Date.ParseExact(oLote.FechaAdmision.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                        If DateTime.TryParseExact(oLote.FechaFabricacion.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                            oDocDestino.Lines.BatchNumbers.ManufacturingDate = Date.ParseExact(oLote.FechaFabricacion.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                        If DateTime.TryParseExact(oLote.FechaVencimiento.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                            oDocDestino.Lines.BatchNumbers.ExpiryDate = Date.ParseExact(oLote.FechaVencimiento.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                        oDocDestino.Lines.BatchNumbers.Add()

                    Next

                End If

                'Añade línea
                oDocDestino.Lines.Add()

            Next

            'ID DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIIDDW").Value = objInventario.DOCIDDW

            'URL DOCUWARE
            oDocDestino.UserFields.Fields.Item("U_SEIURLDW").Value = objInventario.DOCURLDW

            'Añadimos documento
            If oDocDestino.Add() = 0 Then

                'Obtiene el DocEntry y DocNum del documento añadido
                Dim DocEntry As String = oCompany.GetNewObjectKey

                Dim DocNum As String = ""
                If objInventario.Draft = Draft.Firme Then DocNum = oComun.getDocNumDeDocEntry(TablaDestino, DocEntry, Sociedad)

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

    Private Sub setDatosDWdocumento(ByVal Tabla As String,
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
            SQL &= " ," & putQuotes("U_SEIURLDW") & " =N'" & URLDW & " " & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            'SQL &= " And COALESCE(" & putQuotes("DocDate") & "," & getDefaultDate & ") = N'" & DocDate & "'" & vbCrLf

            SQL &= " And COALESCE(" & putQuotes("DocNum") & ",-1)  IN ('-1' " & vbCrLf
            For Each DocNum In DocNums
                SQL &= ", N'" & DocNum & "'" & vbCrLf
            Next
            SQL &= ") " & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("No se ha podido actualizar el documento: " & putQuotes(Tabla) & " para DocNum: " & String.Join("#", DocNums))

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Private Function getDocNumDeRefOrigen(ByVal Tabla As String,
                                          ByVal RefOrigen As String,
                                          ByVal TipoRefOrigen As String,
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
                    SQL &= " And T0." & putQuotes("Ref2") & " LIKE N'" & RefOrigen & "%'" & vbCrLf
            End Select

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

    Private Function getDocNumDocumentoDefinitivo(ByVal Tabla As String,
                                                  ByVal DOCIDDW As String,
                                                  ByVal DocNum As String,
                                                  ByVal Ref2 As String,
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

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & " = N'" & DocNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(Ref2) Then SQL &= " And T0." & putQuotes("Ref2") & " = N'" & Ref2 & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

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
