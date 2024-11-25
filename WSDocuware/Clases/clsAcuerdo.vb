Imports SAPbobsCOM
Imports System.Reflection
Imports System.Globalization
Imports WSDocuware.Utilidades

Public Class clsAcuerdo

#Region "Públicas"

    Public Function NuevoAcuerdo(ByVal objAcuerdo As EntAcuerdoCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company

        Dim oService As BlanketAgreementsService = Nothing
        Dim oAgreement As BlanketAgreement = Nothing
        Dim oParams As BlanketAgreementParams = Nothing
        Dim oItem As BlanketAgreements_ItemsLine = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUEVO ACUERDO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de acuerdo global para " & objAcuerdo.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objAcuerdo.UserSAP, objAcuerdo.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objAcuerdo.NIFTercero) Then Throw New Exception("NIF no suministrado")
            If String.IsNullOrEmpty(objAcuerdo.Descripcion) Then Throw New Exception("Descripcion no suministrada")

            If Not DateTime.TryParseExact(objAcuerdo.FechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha inicio no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objAcuerdo.FechaFin.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha fin no suministrada o incorrecta")
            If Not DateTime.TryParseExact(objAcuerdo.FechaFirma.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha firma no suministrada o incorrecta")

            If String.IsNullOrEmpty(objAcuerdo.DOCIDDW) Then Throw New Exception("ID Docuware no suministrado")
            If String.IsNullOrEmpty(objAcuerdo.DOCURLDW) Then Throw New Exception("URL Docuware no suministrado")

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objAcuerdo.ObjTypeDestino)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objAcuerdo.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objAcuerdo.NIFTercero, objAcuerdo.RazonSocial, objAcuerdo.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objAcuerdo.NIFTercero & ", Razón social: " & objAcuerdo.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Comprueba que no existe el acuerdo
            Dim NumAcuerdo As String = getNumberDeAcuerdo(TablaDestino, CardCode, "", objAcuerdo.NumAtCard, -1, objAcuerdo.Descripcion, Sociedad)
            If Not String.IsNullOrEmpty(NumAcuerdo) Then Throw New Exception("Ya existe acuerdo " & NumAcuerdo & " para NIF: " & objAcuerdo.NIFTercero & ", Razón social: " & objAcuerdo.RazonSocial & ", Descripción: " & objAcuerdo.Descripcion)

            'Objeto acuerdo
            oService = oCompany.GetCompanyService().GetBusinessService(ServiceTypes.BlanketAgreementsService)
            oAgreement = oService.GetDataInterface(BlanketAgreementsServiceDataInterfaces.basBlanketAgreement)
            oParams = oService.GetDataInterface(BlanketAgreementsServiceDataInterfaces.basBlanketAgreementParams)

            'Cabecera
            oAgreement.BPCode = CardCode
            If objAcuerdo.PersonaContacto > 0 Then oAgreement.ContactPersonCode = objAcuerdo.PersonaContacto

            If Not String.IsNullOrEmpty(objAcuerdo.NumAtCard) Then oAgreement.NumAtCard = objAcuerdo.NumAtCard
            If Not String.IsNullOrEmpty(objAcuerdo.Proyecto) Then oAgreement.Project = objAcuerdo.Proyecto

            oAgreement.StartDate = Date.ParseExact(objAcuerdo.FechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            oAgreement.EndDate = Date.ParseExact(objAcuerdo.FechaFin.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            oAgreement.SigningDate = Date.ParseExact(objAcuerdo.FechaFirma.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            oAgreement.Description = objAcuerdo.Descripcion

            oAgreement.AgreementType = objAcuerdo.Tipo      'BlanketAgreementTypeEnum.atGeneral
            oAgreement.AgreementMethod = objAcuerdo.Metodo  'BlanketAgreementMethodEnum.amItem
            oAgreement.Status = objAcuerdo.Status           'BlanketAgreementStatusEnum.asDraft

            'Campos Docuware
            oAgreement.UserFields.Item("U_SEIIDDW").Value = objAcuerdo.DOCIDDW
            oAgreement.UserFields.Item("U_SEIURLDW").Value = objAcuerdo.DOCURLDW

            'Líneas
            If objAcuerdo.Metodo = BlanketAgreementMethodEnum.amItem Then

                '---------------------------
                ' Artículos
                '---------------------------

                'Controlar que no se asigne más de una línea
                If objAcuerdo.Lineas.Count = 0 Then Throw New Exception("No se han indicado las líneas del acuerdo global de tipo artículo")

                'Recorre cada línea
                For Each objLinea In objAcuerdo.Lineas

                    'Añade la línea
                    oItem = oAgreement.BlanketAgreements_ItemsLines.Add()

                    'Índice
                    Dim indLinea As Integer = objAcuerdo.Lineas.IndexOf(objLinea) + 1

                    'Comprobar campos
                    If String.IsNullOrEmpty(objLinea.Articulo) AndAlso String.IsNullOrEmpty(objLinea.RefExt) AndAlso objAcuerdo.Lineas.Count > 1 Then _
                        Throw New Exception("Artículo o referencia externa de la línea " & indLinea.ToString & " no suministrado")

                    'Artículo (si no se encuentra el código se busca por la referencia de proveedor)
                    Dim ItemCode As String = oComun.getItemCode(CardCode, objLinea.Articulo, objLinea.RefExt, Sociedad)
                    If String.IsNullOrEmpty(ItemCode) Then Throw New Exception("No encuentro artículo de la línea " & indLinea.ToString & " con código: " & objLinea.Articulo & " o referencia externa: " & objLinea.RefExt & " para IC: " & CardCode)
                    clsLog.Log.Info("(" & sLogInfo & ") Encontrado artículo: " & ItemCode)

                    oItem.ItemNo = ItemCode

                    ''Artículo 
                    'If Not String.IsNullOrEmpty(ItemCode) Then
                    '    oItem.ItemNo = ItemCode
                    'Else
                    '    oItem.ItemNo = ConfigurationManager.AppSettings.Item("ArticuloGenerico").ToString
                    'End If

                    'Descripción
                    If objLinea.Concepto.Length > 100 Then objLinea.Concepto = objLinea.Concepto.Substring(0, 100)

                    If Not String.IsNullOrEmpty(objLinea.Concepto) Then oItem.ItemDescription = objLinea.Concepto

                    'Cantidad
                    If objLinea.Cantidad > 0 Then oItem.PlannedQuantity = objLinea.Cantidad

                    'Precio
                    If objLinea.PrecioUnidad > 0 Then oItem.UnitPrice = objLinea.PrecioUnidad

                    'Porcentaje dto
                    If objLinea.PorcentajeDescuento > 0 Then oItem.LineDiscount = objLinea.PorcentajeDescuento

                    'Moneda
                    If Not String.IsNullOrEmpty(objLinea.Moneda) Then oItem.PriceCurrency = objLinea.Moneda

                    'Proyecto
                    If Not String.IsNullOrEmpty(objLinea.Proyecto) Then oItem.Project = objLinea.Proyecto

                    'Porcentaje devolución
                    If objLinea.PorcentajeDevolucion > 0 Then oItem.PortionOfReturns = objLinea.PorcentajeDevolucion

                    'Fecha garantia
                    If DateTime.TryParseExact(objLinea.FechaGarantia.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        oItem.EndOfWarranty = Date.ParseExact(objLinea.FechaGarantia.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                    'Texto libre
                    If objLinea.Comentarios.Length > 100 Then objLinea.Comentarios = objLinea.Comentarios.Substring(0, 100)
                    If Not String.IsNullOrEmpty(objLinea.Comentarios) Then oItem.FreeText = objLinea.Comentarios

                Next

            Else

                '---------------------------
                ' Monetario
                '---------------------------

                'Controlar que no se asigne más de una línea
                If objAcuerdo.Lineas.Count > 1 Then Throw New Exception("No se pueden asignar varias líneas a un acuerdo global de tipo monetario")

                For Each objLinea In objAcuerdo.Lineas

                    'Añade la línea
                    oItem = oAgreement.BlanketAgreements_ItemsLines.Add()

                    'Índice
                    Dim indLinea As Integer = objAcuerdo.Lineas.IndexOf(objLinea) + 1

                    'Comprobar campos
                    If Not objLinea.ImportePlanificado > 0 Then Throw New Exception("Importe planificado de la línea " & indLinea.ToString & " no suministrado")

                    'ImportePlanificado
                    If objLinea.ImportePlanificado > 0 Then oItem.PlannedAmountLC = objLinea.ImportePlanificado

                    'Porcentaje dto
                    If objLinea.PorcentajeDescuento > 0 Then oItem.LineDiscount = objLinea.PorcentajeDescuento

                    'Moneda
                    If Not String.IsNullOrEmpty(objLinea.Moneda) Then oItem.PriceCurrency = objLinea.Moneda

                    'Proyecto
                    If Not String.IsNullOrEmpty(objLinea.Proyecto) Then oItem.Project = objLinea.Proyecto

                    'Porcentaje devolución
                    If objLinea.PorcentajeDevolucion > 0 Then oItem.PortionOfReturns = objLinea.PorcentajeDevolucion

                    'Fecha garantia
                    If DateTime.TryParseExact(objLinea.FechaGarantia.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then
                        oItem.EndOfWarranty = Date.ParseExact(objLinea.FechaGarantia.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
                    End If

                    'Texto libre
                    If objLinea.Comentarios.Length > 100 Then objLinea.Comentarios = objLinea.Comentarios.Substring(0, 100)
                    If Not String.IsNullOrEmpty(objLinea.Comentarios) Then oItem.FreeText = objLinea.Comentarios

                Next

            End If

            'Añadimos acuerdo
            oParams = oService.AddBlanketAgreement(oAgreement)

            'Obtiene el Number del documento añadido
            Dim Number As String = getNumberDeAcuerdo(TablaDestino, CardCode, "", objAcuerdo.NumAtCard, -1, objAcuerdo.Descripcion, Sociedad)

            retVal.CODIGO = Respuesta.Ok
            retVal.MENSAJE = "Acuerdo creado con éxito"
            retVal.MENSAJEAUX = Number

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        Finally
            oComun.LiberarObjCOM(oItem)
            oComun.LiberarObjCOM(oParams)
            oComun.LiberarObjCOM(oAgreement)
            oComun.LiberarObjCOM(oService)
        End Try

        Return retVal

    End Function

    Public Function ActualizarAcuerdo(ByVal objAcuerdo As EntAcuerdoCab, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company

        Dim oService As BlanketAgreementsService = Nothing
        Dim oAgreement As BlanketAgreement = Nothing
        Dim oParams As BlanketAgreementParams = Nothing
        Dim oItem As BlanketAgreements_ItemsLine = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "ACTUALIZAR ACUERDO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de acuerdo global: " & objAcuerdo.Numero)

            oCompany = ConexionSAP.getCompany(objAcuerdo.UserSAP, objAcuerdo.PassSAP, Sociedad)

            If oCompany Is Nothing OrElse Not oCompany.Connected Then
                Throw New Exception("No se puede conectar a SAP")
            End If

            'Obligatorios
            If String.IsNullOrEmpty(objAcuerdo.Numero) Then Throw New Exception("Número no suministrado")

            If objAcuerdo.FechaInicio > 0 AndAlso Not DateTime.TryParseExact(objAcuerdo.FechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha inicio no suministrada o incorrecta")
            If objAcuerdo.FechaFin > 0 AndAlso Not DateTime.TryParseExact(objAcuerdo.FechaFin.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha fin no suministrada o incorrecta")
            If objAcuerdo.FechaFirma > 0 AndAlso Not DateTime.TryParseExact(objAcuerdo.FechaFirma.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                Throw New Exception("Fecha firma no suministrada o incorrecta")

            'Buscamos tabla destino por ObjectType
            Dim TablaDestino As String = getTablaDeObjType(objAcuerdo.ObjTypeDestino)
            If String.IsNullOrEmpty(TablaDestino) Then Throw New Exception("No encuentro tabla destino con ObjectType: " & objAcuerdo.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla destino: " & TablaDestino)

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objAcuerdo.NIFTercero, objAcuerdo.RazonSocial, objAcuerdo.Ambito, Sociedad)
            If Not String.IsNullOrEmpty(CardCode) Then clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Objeto acuerdo
            oService = oCompany.GetCompanyService().GetBusinessService(ServiceTypes.BlanketAgreementsService)
            oParams = oService.GetDataInterface(BlanketAgreementsServiceDataInterfaces.basBlanketAgreementParams)
            oParams.AgreementNo = CInt(objAcuerdo.Numero)
            oAgreement = oService.GetBlanketAgreement(oParams)

            'Cabecera
            If Not String.IsNullOrEmpty(CardCode) Then oAgreement.BPCode = CardCode

            If objAcuerdo.PersonaContacto > 0 Then oAgreement.ContactPersonCode = objAcuerdo.PersonaContacto

            If Not String.IsNullOrEmpty(objAcuerdo.NumAtCard) Then oAgreement.NumAtCard = objAcuerdo.NumAtCard
            If Not String.IsNullOrEmpty(objAcuerdo.Proyecto) Then oAgreement.Project = objAcuerdo.Proyecto

            If objAcuerdo.FechaInicio > 0 Then oAgreement.StartDate = Date.ParseExact(objAcuerdo.FechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            If objAcuerdo.FechaFin > 0 Then oAgreement.EndDate = Date.ParseExact(objAcuerdo.FechaFin.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)
            If objAcuerdo.FechaFirma > 0 Then oAgreement.SigningDate = Date.ParseExact(objAcuerdo.FechaFirma.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            If Not String.IsNullOrEmpty(objAcuerdo.Descripcion) Then oAgreement.Description = objAcuerdo.Descripcion

            If objAcuerdo.Tipo > -1 AndAlso objAcuerdo.Tipo < 2 Then oAgreement.AgreementType = objAcuerdo.Tipo          'BlanketAgreementTypeEnum.atGeneral
            If objAcuerdo.Metodo > -1 AndAlso objAcuerdo.Metodo < 2 Then oAgreement.AgreementMethod = objAcuerdo.Metodo  'BlanketAgreementMethodEnum.amItem
            If objAcuerdo.Status > -1 AndAlso objAcuerdo.Status < 4 Then oAgreement.Status = objAcuerdo.Status           'BlanketAgreementStatusEnum.asDraft

            'Campos Docuware
            If Not String.IsNullOrEmpty(objAcuerdo.DOCIDDW) Then oAgreement.UserFields.Item("U_SEIIDDW").Value = objAcuerdo.DOCIDDW
            If Not String.IsNullOrEmpty(objAcuerdo.DOCURLDW) Then oAgreement.UserFields.Item("U_SEIURLDW").Value = objAcuerdo.DOCURLDW

            'Líneas (sólo se actualizan si se reciben)
            If objAcuerdo.Lineas.Count > 0 Then

                If objAcuerdo.Metodo = BlanketAgreementMethodEnum.amItem Then

                    '---------------------------
                    ' Artículos
                    '---------------------------

                    '20221129: Comprueba que el código de artículo sea ItemCode y no la referencia de proveedor
                    For Each objLinea In objAcuerdo.Lineas

                        'Índice
                        Dim indLinea As Integer = objAcuerdo.Lineas.IndexOf(objLinea) + 1

                        If Not String.IsNullOrEmpty(objLinea.Articulo) OrElse Not String.IsNullOrEmpty(objLinea.RefExt) Then
                            'Artículo (si no se encuentra el código se busca por la referencia de proveedor)
                            Dim ItemCode As String = oComun.getItemCode(oAgreement.BPCode, objLinea.Articulo, objLinea.RefExt, Sociedad)
                            If objLinea.Articulo = ItemCode Then Continue For

                            If String.IsNullOrEmpty(ItemCode) Then _
                                Throw New Exception("No encuentro artículo de la línea " & indLinea.ToString & " con código: " & objLinea.Articulo & " o referencia externa: " & objLinea.RefExt & " para IC: " & oAgreement.BPCode)
                            clsLog.Log.Info("(" & sLogInfo & ") Encontrado artículo: " & ItemCode)

                            objLinea.Articulo = ItemCode
                        End If

                    Next

                    'Recorre cada línea
                    For i = 0 To oAgreement.BlanketAgreements_ItemsLines.Count - 1

                        'Se posiciona en la línea
                        oItem = oAgreement.BlanketAgreements_ItemsLines.Item(i)

                        'Busca el objeto que debe sustituir la línea
                        Dim objLinea As EntAcuerdoLin = (From p In objAcuerdo.Lineas
                                                         Where p.LineNum = oItem.AgreementRowNumber OrElse p.Articulo = oItem.ItemNo).FirstOrDefault

                        If objLinea Is Nothing Then Throw New Exception("No encuentro interlocutor con LineNum: " & oItem.AgreementRowNumber & ", Articulo: " & oItem.ItemNo)

                        'Artículo 
                        If Not String.IsNullOrEmpty(objLinea.Articulo) Then oItem.ItemNo = objLinea.Articulo

                        'Descripción
                        If objLinea.Concepto.Length > 100 Then objLinea.Concepto = objLinea.Concepto.Substring(0, 100)
                        If Not String.IsNullOrEmpty(objLinea.Concepto) Then oItem.ItemDescription = objLinea.Concepto

                        'Cantidad
                        If objLinea.Cantidad > 0 Then oItem.PlannedQuantity = objLinea.Cantidad

                        'Precio
                        If objLinea.PrecioUnidad > 0 Then oItem.UnitPrice = objLinea.PrecioUnidad

                        'Porcentaje dto
                        If objLinea.PorcentajeDescuento > 0 Then oItem.LineDiscount = objLinea.PorcentajeDescuento

                        'Moneda
                        If Not String.IsNullOrEmpty(objLinea.Moneda) Then oItem.PriceCurrency = objLinea.Moneda

                        'Proyecto
                        If Not String.IsNullOrEmpty(objLinea.Proyecto) Then oItem.Project = objLinea.Proyecto

                        'Porcentaje devolución
                        If objLinea.PorcentajeDevolucion > 0 Then oItem.PortionOfReturns = objLinea.PorcentajeDevolucion

                        'Fecha garantia
                        If DateTime.TryParseExact(objLinea.FechaGarantia.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                            oItem.EndOfWarranty = Date.ParseExact(objLinea.FechaGarantia.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                        'Texto libre
                        If objLinea.Comentarios.Length > 100 Then objLinea.Comentarios = objLinea.Comentarios.Substring(0, 100)
                        If Not String.IsNullOrEmpty(objLinea.Comentarios) Then oItem.FreeText = objLinea.Comentarios

                    Next

                Else

                    '---------------------------
                    ' Monetario
                    '---------------------------

                    'Se posiciona en la línea
                    oItem = oAgreement.BlanketAgreements_ItemsLines.Item(0)

                    'Busca el objeto que debe sustituir la línea
                    Dim objLinea As EntAcuerdoLin = objAcuerdo.Lineas(0)

                    'ImportePlanificado
                    If objLinea.ImportePlanificado > 0 Then oItem.PlannedAmountLC = objLinea.ImportePlanificado

                    'Porcentaje dto
                    If objLinea.PorcentajeDescuento > 0 Then oItem.LineDiscount = objLinea.PorcentajeDescuento

                    'Moneda
                    If Not String.IsNullOrEmpty(objLinea.Moneda) Then oItem.PriceCurrency = objLinea.Moneda

                    'Proyecto
                    If Not String.IsNullOrEmpty(objLinea.Proyecto) Then oItem.Project = objLinea.Proyecto

                    'Porcentaje devolución
                    If objLinea.PorcentajeDevolucion > 0 Then oItem.PortionOfReturns = objLinea.PorcentajeDevolucion

                    'Fecha garantia
                    If DateTime.TryParseExact(objLinea.FechaGarantia.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                        oItem.EndOfWarranty = Date.ParseExact(objLinea.FechaGarantia.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                    'Texto libre
                    If objLinea.Comentarios.Length > 100 Then objLinea.Comentarios = objLinea.Comentarios.Substring(0, 100)
                    If Not String.IsNullOrEmpty(objLinea.Comentarios) Then oItem.FreeText = objLinea.Comentarios

                End If

            End If

            'Actualizamos acuerdo
            oService.UpdateBlanketAgreement(oAgreement)

            retVal.CODIGO = Respuesta.Ok
            retVal.MENSAJE = "Acuerdo actualizado con éxito"
            retVal.MENSAJEAUX = objAcuerdo.Numero

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        Finally
            oComun.LiberarObjCOM(oItem)
            oComun.LiberarObjCOM(oParams)
            oComun.LiberarObjCOM(oAgreement)
            oComun.LiberarObjCOM(oService)
        End Try

        Return retVal

    End Function

    Public Function getComprobarAcuerdo(ByVal objAcuerdo As EntAcuerdoCab, ByVal Sociedad As eSociedad) As EntResultado

        'Comprueba si existe un acuerdo

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "ACUERDO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ")  Comprobar acuerdo para ObjType: " & objAcuerdo.ObjTypeDestino & ", Ambito: " & objAcuerdo.Ambito & ", NIFTercero: " & objAcuerdo.NIFTercero & ", Razón social: " & objAcuerdo.RazonSocial & ", DOCIDDW: " & objAcuerdo.DOCIDDW & ", Descripcion: " & objAcuerdo.Descripcion)

            'Obligatorios
            If String.IsNullOrEmpty(objAcuerdo.NIFTercero) AndAlso String.IsNullOrEmpty(objAcuerdo.RazonSocial) Then Throw New Exception("NIF o razón social no suministrado")
            If String.IsNullOrEmpty(objAcuerdo.DOCIDDW) AndAlso String.IsNullOrEmpty(objAcuerdo.Descripcion) Then Throw New Exception("ID Docuware o descripción no suministrados")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objAcuerdo.NIFTercero, objAcuerdo.RazonSocial, objAcuerdo.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objAcuerdo.NIFTercero & ", Razón social: " & objAcuerdo.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objAcuerdo.ObjTypeDestino)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objAcuerdo.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ")  Encontrada tabla: " & Tabla)

            'Devuelve el Number del acuerdo
            Dim DocNum As String = getNumberDeAcuerdo(Tabla, CardCode, objAcuerdo.DOCIDDW, objAcuerdo.NumAtCard, objAcuerdo.Numero, objAcuerdo.Descripcion, Sociedad)

            If Not String.IsNullOrEmpty(DocNum) Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Acuerdo encontrado"
                retVal.MENSAJEAUX = DocNum

            Else

                retVal.CODIGO = Respuesta.Ko
                retVal.MENSAJE = "No se encuentra el acuerdo"
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

    Public Function setAcuerdoEstado(ByVal objAcuerdoEstado As EntAcuerdoEstado, ByVal Sociedad As eSociedad) As EntResultado

        'Actualiza el campo estado DW

        Dim retVal As New EntResultado

        Dim oComun As New clsComun

        Dim sLogInfo As String = "ACUERDO ESTADO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de acuerdo estado para ObjType: " & objAcuerdoEstado.ObjTypeDestino & ", Ambito: " & objAcuerdoEstado.Ambito & ", NIFTercero: " & objAcuerdoEstado.NIFTercero & ", Razón social: " & objAcuerdoEstado.RazonSocial & ", Number: " & objAcuerdoEstado.Numero & ", NumAtCard: " & objAcuerdoEstado.NumAtCard)

            'Obligatorios
            If String.IsNullOrEmpty(objAcuerdoEstado.NIFTercero) AndAlso String.IsNullOrEmpty(objAcuerdoEstado.RazonSocial) Then _
                Throw New Exception("NIF o razón social no suministrado")

            If String.IsNullOrEmpty(objAcuerdoEstado.Numero) AndAlso String.IsNullOrEmpty(objAcuerdoEstado.NumAtCard) Then _
                Throw New Exception("Número de acuerdo o referencia origen no suministrados")

            If objAcuerdoEstado.DOCESTADODW < 0 AndAlso objAcuerdoEstado.DOCESTADODW > 5 Then _
                Throw New Exception("Estado docuware incorrecto. Valores válidos 0 a 5")

            'Buscamos IC por NIF, U_SEIICDW o razón social
            Dim CardCode As String = oComun.getCardCode(objAcuerdoEstado.NIFTercero, objAcuerdoEstado.RazonSocial, objAcuerdoEstado.Ambito, Sociedad)
            If String.IsNullOrEmpty(CardCode) Then Throw New Exception("No encuentro IC con NIF: " & objAcuerdoEstado.NIFTercero & ", Razón social: " & objAcuerdoEstado.RazonSocial)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado IC: " & CardCode)

            'Buscamos tabla por ObjectType
            Dim Tabla As String = getTablaDeObjType(objAcuerdoEstado.ObjTypeDestino)
            If String.IsNullOrEmpty(Tabla) Then Throw New Exception("No encuentro tabla con el ObjectType: " & objAcuerdoEstado.ObjTypeDestino)
            clsLog.Log.Info("(" & sLogInfo & ") Encontrada tabla: " & Tabla)

            'Buscamos AbsID por Number o NumAtCard
            Dim AbsID As String = getAbsIDDeAcuerdo(Tabla, CardCode, objAcuerdoEstado.Numero, objAcuerdoEstado.NumAtCard, Sociedad)
            If String.IsNullOrEmpty(AbsID) Then Throw New Exception("No encuentro acuerdo con Number - '" & objAcuerdoEstado.Numero & "' y NumAtCard - '" & objAcuerdoEstado.NumAtCard & "'")
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado AbsID: " & AbsID)

            'Buscamos Number por AbsID
            Dim Number As String = getNumberDeAbsID(Tabla, AbsID, Sociedad)
            If String.IsNullOrEmpty(Number) Then Throw New Exception("No encuentro acuerdo con AbsID: '" & AbsID & "'")
            clsLog.Log.Info("(" & sLogInfo & ") Encontrado Number: " & AbsID)

            'Actualiza el campo Estado DW 
            Dim Actualizado As Boolean = setAcuerdoEstadoDW(Tabla, AbsID, objAcuerdoEstado.DOCIDDW, objAcuerdoEstado.DOCURLDW, objAcuerdoEstado.DOCESTADODW, objAcuerdoEstado.DOCMOTIVODW, Sociedad)

            If Actualizado Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Acuerdo actualizado con éxito"
                retVal.MENSAJEAUX = Number

            Else

                retVal.CODIGO = Respuesta.Ko
                retVal.MENSAJE = "Acuerdo no actualizado"
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

    Private Function getNumberAcuerdoPorDWID(ByVal Tabla As String,
                                             ByVal IDDW As String,
                                             ByVal Sociedad As eSociedad) As String

        'Devuelve el Number del acuerdo

        Dim retVal As String = ""

        Try

            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("Number") & vbCrLf
            SQL &= " FROM " & putQuotes(Tabla) & " T0 " & getWithNoLock() & vbCrLf
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

    Private Sub setAcuerdoDatosDW(ByVal Tabla As String,
                                  ByVal BPCode As String,
                                  ByVal Number As String,
                                  ByVal IDDW As String,
                                  ByVal URLDW As String,
                                  ByVal Sociedad As eSociedad)

        'Actualiza los campos de DW en el documento

        Try

            'Actualizamos por CardCode, DocDate y DocNum
            Dim SQL As String = ""
            SQL = "  UPDATE " & putQuotes(Tabla) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEIIDDW") & " =N'" & IDDW & "', " & vbCrLf
            SQL &= " " & putQuotes("U_SEIURLDW") & " =N'" & URLDW & "' " & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And COALESCE(" & putQuotes("BpCode") & ",'') = N'" & BPCode & "'" & vbCrLf
            SQL &= " And COALESCE(" & putQuotes("Number") & ",-1) = N'" & Number & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("Actualización nula al ejecutar sentencia " & SQL)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

    End Sub

    Private Function setAcuerdoEstadoDW(ByVal Tabla As String,
                                        ByVal AbsID As String,
                                        ByVal IDDW As String,
                                        ByVal URLDW As String,
                                        ByVal ESTADODW As Integer,
                                        ByVal MOTIVODW As String,
                                        ByVal Sociedad As eSociedad) As Boolean

        'Actualiza los campos de DW en el documento

        Dim retval As Boolean = False

        Try

            'Actualizamos por CardCode y Number
            Dim SQL As String = ""
            SQL = "  UPDATE " & getDataBaseRef(Tabla, Sociedad) & " SET  " & vbCrLf
            SQL &= " " & putQuotes("U_SEIIDDW") & " =N'" & IDDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIURLDW") & " =N'" & URLDW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIESTDW") & " =N'" & ESTADODW & "' " & vbCrLf
            SQL &= " ," & putQuotes("U_SEIMOTDW") & " =N'" & MOTIVODW & "' " & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And COALESCE(" & putQuotes("AbsID") & ",-1) = N'" & AbsID & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)
            If oCon.ExecuteQuery(SQL) = 0 Then Throw New Exception("Actualización nula al ejecutar sentencia " & SQL)

            retval = True

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retval

    End Function

    Private Function getNumberDeAcuerdo(ByVal Tabla As String,
                                        ByVal BPCode As String,
                                        ByVal DOCIDDW As String,
                                        ByVal NumAtCard As String,
                                        ByVal Numero As String,
                                        ByVal Descripcion As String,
                                        ByVal Sociedad As eSociedad) As String

        'Devuelve el Number del acuerdo

        Dim retVal As String = ""

        Try

            'Buscamos por Docnum
            Dim SQL As String = ""
            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " T0." & putQuotes("Number") & vbCrLf
            SQL &= " FROM " & putQuotes(Tabla) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("BpCode") & " = N'" & BPCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then
                SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf
            End If

            If Not String.IsNullOrEmpty(NumAtCard) Then
                SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf
            End If

            If Not String.IsNullOrEmpty(Numero) Then
                SQL &= " And T0." & putQuotes("Number") & " = N'" & Numero & "'" & vbCrLf
            End If

            If Not String.IsNullOrEmpty(Descripcion) Then
                SQL &= " And T0." & putQuotes("Descript") & " = N'" & Descripcion & "'" & vbCrLf
            End If

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getNumberDeAbsID(ByVal Tabla As String,
                                      ByVal AbsID As String,
                                      ByVal Sociedad As eSociedad) As String

        'Devuelve el Number del acuerdo

        Dim retVal As String = ""

        Try

            'Buscamos por Docnum
            Dim SQL As String = ""
            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " T0." & putQuotes("Number") & vbCrLf
            SQL &= " FROM " & putQuotes(Tabla) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("AbsID") & " = N'" & AbsID & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getAbsIDDeAcuerdo(ByVal Tabla As String,
                                       ByVal BPCode As String,
                                       ByVal Number As String,
                                       ByVal NumAtCard As String,
                                       ByVal Sociedad As eSociedad) As String

        'Devuelve el AbsID del acuerdo

        Dim retVal As String = ""

        Try

            'Buscamos por Number o NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " T0." & putQuotes("AbsID") & vbCrLf
            SQL &= " FROM " & putQuotes(Tabla) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            SQL &= " And T0." & putQuotes("BpCode") & " = N'" & BPCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(Number) Then
                SQL &= " And T0." & putQuotes("Number") & " = N'" & Number & "'" & vbCrLf
            End If

            If Not String.IsNullOrEmpty(NumAtCard) Then
                SQL &= " And T0." & putQuotes("NumAtCard") & " = N'" & NumAtCard & "'" & vbCrLf
            End If

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
