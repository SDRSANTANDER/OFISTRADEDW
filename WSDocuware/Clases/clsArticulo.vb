Imports SAPbobsCOM
Imports System.Reflection
Imports System.Globalization
Imports WSDocuware.Utilidades

Public Class clsArticulo

#Region "Públicas"

    Public Function NuevoArticulo(ByVal objArticulo As EntArticuloSAP, ByVal Sociedad As eSociedad) As EntResultado

        'Crea artículo
        Dim retVal As New EntResultado

        Try

            Select Case objArticulo.Tipo

                Case ItemType.ActivoFijo
                    'Activo fijo
                    retVal = NuevoArticuloActivoFijo(objArticulo, Sociedad)

                Case Else
                    'Normal
                    retVal = NuevoArticuloNormal(objArticulo, Sociedad)

            End Select

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
            retVal.CODIGO = Respuesta.Ko
            retVal.MENSAJE = ex.Message
            retVal.MENSAJEAUX = ""
        End Try

        Return retVal

    End Function

    Public Function ActualizarArticulo(ByVal objArticulo As EntArticuloSAP, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oItem As Items = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "ACTUALIZAR ARTICULO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de artículo: " & objArticulo.Codigo)

            oCompany = ConexionSAP.getCompany(objArticulo.UserSAP, objArticulo.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If String.IsNullOrEmpty(objArticulo.Codigo) Then Throw New Exception("Código no suministrado")

            'Objeto artículo
            oItem = oCompany.GetBusinessObject(BoObjectTypes.oItems)

            'Comprueba si existe el artículo
            If Not oItem.GetByKey(objArticulo.Codigo) Then

                Throw New Exception("No se encuentra el artículo con ItemCode: " & objArticulo.Codigo)

            Else

                'ItemName
                If Not String.IsNullOrEmpty(objArticulo.Nombre) Then oItem.ItemName = objArticulo.Nombre

                'FrgnName
                If Not String.IsNullOrEmpty(objArticulo.NombreExtranjero) Then oItem.ForeignName = objArticulo.NombreExtranjero

                'ItemType
                If IsNumeric(objArticulo.Tipo) Then oItem.ItemType = CInt(objArticulo.Tipo)

                'ItmsGrpCod
                If IsNumeric(objArticulo.Grupo) Then oItem.ItemsGroupCode = CInt(objArticulo.Grupo)

                'UgpEntry
                If IsNumeric(objArticulo.GrupoUnidadMedida) Then oItem.UoMGroupEntry = CInt(objArticulo.GrupoUnidadMedida)

                'CodeBars
                If Not String.IsNullOrEmpty(objArticulo.CodigoBarras) Then oItem.BarCode = objArticulo.CodigoBarras

                'PrchseItem
                oItem.PurchaseItem = IIf(String.IsNullOrEmpty(objArticulo.Compra) OrElse objArticulo.Compra = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

                'SellItem
                oItem.SalesItem = IIf(String.IsNullOrEmpty(objArticulo.Venta) OrElse objArticulo.Venta = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

                'InvntItem
                oItem.InventoryItem = IIf(String.IsNullOrEmpty(objArticulo.Inventario) OrElse objArticulo.Inventario = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

                '---------------------------------
                ' Pestaña general
                '---------------------------------

                'WTLiable
                oItem.WTLiable = IIf(String.IsNullOrEmpty(objArticulo.GeneralSujetoRetencion) OrElse objArticulo.GeneralSujetoRetencion = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

                'FirmCode
                If IsNumeric(objArticulo.GeneralFabricante) Then oItem.Manufacturer = CInt(objArticulo.GeneralFabricante)

                'ShipType
                If IsNumeric(objArticulo.GeneralClaseExpedicion) Then oItem.ShipType = CInt(objArticulo.GeneralClaseExpedicion)

                'ManBtchNum
                oItem.ManageBatchNumbers = IIf(String.IsNullOrEmpty(objArticulo.GeneralGestionarArticuloPorLotes) OrElse objArticulo.GeneralGestionarArticuloPorLotes = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

                'ISRelevant
                oItem.IntrastatExtension.IntrastatRelevant = IIf(String.IsNullOrEmpty(objArticulo.GeneralRelevanteInstratat) OrElse objArticulo.GeneralRelevanteInstratat = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

                'frozenFor
                oItem.Frozen = IIf(objArticulo.GeneralActivo = SN.Si, BoYesNoEnum.tYES, BoYesNoEnum.tNO)

                '---------------------------------
                ' Pestaña compras
                '---------------------------------

                'CardCode
                If Not String.IsNullOrEmpty(objArticulo.CompraProveedorDefecto) Then oItem.Mainsupplier = objArticulo.CompraProveedorDefecto

                'CardCode
                If Not String.IsNullOrEmpty(objArticulo.CompraNumeroCatalogoFabricante) Then oItem.SupplierCatalogNo = objArticulo.CompraNumeroCatalogoFabricante

                'BuyUnitMsr
                If Not String.IsNullOrEmpty(objArticulo.CompraUnidadMedida) Then oItem.PurchaseUnit = objArticulo.CompraUnidadMedida

                'NumInBuy
                If objArticulo.CompraArticulosUnidad > 0 Then oItem.PurchaseItemsPerUnit = CDbl(objArticulo.CompraArticulosUnidad)

                'PurPackMsr
                If Not String.IsNullOrEmpty(objArticulo.CompraUnidadMedidaEmbalaje) Then oItem.PurchasePackagingUnit = objArticulo.CompraUnidadMedidaEmbalaje

                'PurPackUn
                If objArticulo.CompraCantidadEmbalaje > 0 Then oItem.PurchaseQtyPerPackUnit = CDbl(objArticulo.CompraCantidadEmbalaje)

                'VatGroupPu
                If Not String.IsNullOrEmpty(objArticulo.CompraGrupoImpositivo) Then oItem.PurchaseVATGroup = objArticulo.CompraGrupoImpositivo

                '---------------------------------
                ' Pestaña ventas
                '---------------------------------

                'SalUnitMsr
                If Not String.IsNullOrEmpty(objArticulo.VentaUnidadMedida) Then oItem.SalesUnit = objArticulo.VentaUnidadMedida

                'NumInSale
                If objArticulo.VentaArticulosUnidad > 0 Then oItem.SalesItemsPerUnit = CDbl(objArticulo.VentaArticulosUnidad)

                'SalPackMsr
                If Not String.IsNullOrEmpty(objArticulo.VentaUnidadMedidaEmbalaje) Then oItem.SalesPackagingUnit = objArticulo.VentaUnidadMedidaEmbalaje

                'SalPackUn
                If objArticulo.VentaCantidadEmbalaje > 0 Then oItem.SalesQtyPerPackUnit = CDbl(objArticulo.VentaCantidadEmbalaje)

                'VatGroupSa
                If Not String.IsNullOrEmpty(objArticulo.VentaGrupoImpositivo) Then oItem.SalesVATGroup = objArticulo.VentaGrupoImpositivo

                '---------------------------------
                ' Pestaña inventario
                '---------------------------------

                'GLMethod
                If Not String.IsNullOrEmpty(objArticulo.InventarioCuentasDeMayorPor) Then oItem.GLMethod = CDbl(objArticulo.InventarioCuentasDeMayorPor)

                'InvntryUom
                If Not String.IsNullOrEmpty(objArticulo.InventarioUnidadMedida) Then oItem.InventoryUOM = objArticulo.InventarioUnidadMedida

                'IWeight1
                If objArticulo.InventarioPeso > 0 Then oItem.InventoryWeight = CDbl(objArticulo.InventarioPeso)

                'ByWh
                oItem.ManageStockByWarehouse = IIf(String.IsNullOrEmpty(objArticulo.InventarioGestionStockAlmacen) OrElse objArticulo.InventarioGestionStockAlmacen = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

                '---------------------------------
                ' Pestaña activos fijos
                '---------------------------------

                'AssetClass
                If Not String.IsNullOrEmpty(objArticulo.ActivoFijoClase) Then oItem.AssetClass = objArticulo.ActivoFijoClase

                'AssetGroup
                If Not String.IsNullOrEmpty(objArticulo.ActivoFijoGrupo) Then oItem.AssetGroup = objArticulo.ActivoFijoGrupo

                'DeprGroup
                If Not String.IsNullOrEmpty(objArticulo.ActivoFijoGrupoAmortizacion) Then oItem.DepreciationGroup = objArticulo.ActivoFijoGrupoAmortizacion

                'InventryNo
                If Not String.IsNullOrEmpty(objArticulo.ActivoFijoNumeroInventario) Then oItem.InventoryNumber = objArticulo.ActivoFijoNumeroInventario

                'AssetSerNo
                If Not String.IsNullOrEmpty(objArticulo.ActivoFijoNumeroSerie) Then oItem.AssetSerialNumber = objArticulo.ActivoFijoNumeroSerie

                'Location
                If objArticulo.ActivoFijoEmplazamiento > 0 Then oItem.Location = objArticulo.ActivoFijoEmplazamiento

                'Technician
                If objArticulo.ActivoFijoTecnico > 0 Then oItem.Technician = objArticulo.ActivoFijoTecnico

                'Employee
                If objArticulo.ActivoFijoEmpleado > 0 Then oItem.Employee = objArticulo.ActivoFijoEmpleado

                'CapDate
                If DateTime.TryParseExact(objArticulo.ActivoFijoFechaCapitalizacion.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                    oItem.CapitalizationDate = Date.ParseExact(objArticulo.ActivoFijoFechaCapitalizacion.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                'StatAsset
                oItem.StatisticalAsset = IIf(String.IsNullOrEmpty(objArticulo.ActivoFijoEstadistico) OrElse objArticulo.ActivoFijoEstadistico = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

                'Cession
                oItem.Cession = IIf(String.IsNullOrEmpty(objArticulo.ActivoFijoCesion) OrElse objArticulo.ActivoFijoCesion = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

                'DepreciationArea
                If Not String.IsNullOrEmpty(objArticulo.ActivoFijoAreaValoracion) Then oItem.DepreciationParameters.DepreciationArea = objArticulo.ActivoFijoAreaValoracion

                'DepreciationStartDate
                If DateTime.TryParseExact(objArticulo.ActivoFijoFechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                    oItem.DepreciationParameters.DepreciationStartDate = Date.ParseExact(objArticulo.ActivoFijoFechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                'DepreciationType
                If Not String.IsNullOrEmpty(objArticulo.ActivoFijoTipo) Then oItem.DepreciationParameters.DepreciationType = objArticulo.ActivoFijoTipo

                'FiscalYear
                If Not String.IsNullOrEmpty(objArticulo.ActivoFijoAnyoFiscal) Then oItem.DepreciationParameters.FiscalYear = objArticulo.ActivoFijoAnyoFiscal

                'UsefulLife
                If objArticulo.ActivoFijoVidaUtil > 0 Then oItem.DepreciationParameters.UsefulLife = objArticulo.ActivoFijoVidaUtil

                'TotalUnitsInUsefulLife
                If objArticulo.ActivoFijoVidaUtilUnidades > 0 Then oItem.DepreciationParameters.TotalUnitsInUsefulLife = objArticulo.ActivoFijoVidaUtilUnidades

                'CapitalGoodsOnHoldLimit
                If objArticulo.ActivoFijoCAPHistorico > 0 Then oItem.CapitalGoodsOnHoldLimit = objArticulo.ActivoFijoCAPHistorico

                'CapitalGoodsOnHoldLimit
                If objArticulo.ActivoFijoCAPHistorico > 0 Then oItem.CapitalGoodsOnHoldLimit = objArticulo.ActivoFijoCAPHistorico

                '---------------------------------
                ' Pestaña comentarios
                '---------------------------------

                'UserText
                If Not String.IsNullOrEmpty(objArticulo.Comentarios) Then oItem.User_Text = objArticulo.Comentarios

                '---------------------------------
                ' Campos de usuario
                '---------------------------------

                'Campos de usuario
                If Not objArticulo.CamposUsuario Is Nothing AndAlso objArticulo.CamposUsuario.Count > 0 Then

                    For Each oCampoUsuario In objArticulo.CamposUsuario

                        Dim oUserField As Field = oItem.UserFields.Fields.Item(oCampoUsuario.Campo)

                        Select Case oUserField.Type

                            Case BoFieldTypes.db_Numeric
                                If oUserField.SubType = BoFldSubTypes.st_Time Then
                                    If DateTime.TryParseExact(oCampoUsuario.Valor, "HHmm", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                    oItem.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "HHmm", CultureInfo.CurrentCulture)
                                Else
                                    If IsNumeric(oCampoUsuario.Valor) Then _
                                    oItem.Lines.UserFields.Item(oCampoUsuario.Campo).Value = CInt(oCampoUsuario.Valor)
                                End If

                            Case BoFieldTypes.db_Float
                                If IsNumeric(oCampoUsuario.Valor) Then _
                                oItem.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CDbl(oCampoUsuario.Valor.Replace(".", ","))

                            Case BoFieldTypes.db_Date
                                If DateTime.TryParseExact(oCampoUsuario.Valor, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                oItem.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                            Case Else
                                If Not String.IsNullOrEmpty(oCampoUsuario.Valor) Then _
                                oItem.UserFields.Fields.Item(oCampoUsuario.Campo).Value = oCampoUsuario.Valor

                        End Select

                    Next

                End If

                '---------------------------------
                ' Actualización
                '---------------------------------

                'Actualizamos artículo
                If oItem.Update() = 0 Then

                    retVal.CODIGO = Respuesta.Ok
                    retVal.MENSAJE = "Artículo actualizado con éxito"
                    retVal.MENSAJEAUX = objArticulo.Codigo

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
            oComun.LiberarObjCOM(oItem)
        End Try

        Return retVal

    End Function

#End Region

#Region "Privadas"

    Private Function NuevoArticuloNormal(ByVal objArticulo As EntArticuloSAP, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oItem As Items = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUEVO ARTICULO"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de artículo para " & objArticulo.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objArticulo.UserSAP, objArticulo.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Obligatorios
            If objArticulo.Serie <= 0 AndAlso String.IsNullOrEmpty(objArticulo.Codigo) Then Throw New Exception("Serie o código no suministrado")
            If String.IsNullOrEmpty(objArticulo.Nombre) Then Throw New Exception("Nombre no suministrado")

            'Objeto artículo
            oItem = oCompany.GetBusinessObject(BoObjectTypes.oItems)

            'Comprueba si existe el artículo
            If objArticulo.Serie <= 0 AndAlso Not String.IsNullOrEmpty(objArticulo.Codigo) Then

                If oItem.GetByKey(objArticulo.Codigo) Then Throw New Exception("Ya existe el artículo con ItemCode: " & objArticulo.Codigo)

            End If

            'El código depende de la serie

            'Series
            If objArticulo.Serie > 0 Then oItem.Series = objArticulo.Serie

            'ItemCode
            If Not String.IsNullOrEmpty(objArticulo.Codigo) Then oItem.ItemCode = objArticulo.Codigo

            'ItemName
            If Not String.IsNullOrEmpty(objArticulo.Nombre) Then oItem.ItemName = objArticulo.Nombre

            'FrgnName
            If Not String.IsNullOrEmpty(objArticulo.NombreExtranjero) Then oItem.ForeignName = objArticulo.NombreExtranjero

            'ItemType
            If IsNumeric(objArticulo.Tipo) Then oItem.ItemType = CInt(objArticulo.Tipo)

            'ItmsGrpCod
            If IsNumeric(objArticulo.Grupo) Then oItem.ItemsGroupCode = CInt(objArticulo.Grupo)

            'UgpEntry
            If IsNumeric(objArticulo.GrupoUnidadMedida) Then oItem.UoMGroupEntry = CInt(objArticulo.GrupoUnidadMedida)

            'CodeBars
            If Not String.IsNullOrEmpty(objArticulo.CodigoBarras) Then oItem.BarCode = objArticulo.CodigoBarras

            'PrchseItem
            oItem.PurchaseItem = IIf(String.IsNullOrEmpty(objArticulo.Compra) OrElse objArticulo.Compra = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

            'SellItem
            oItem.SalesItem = IIf(String.IsNullOrEmpty(objArticulo.Venta) OrElse objArticulo.Venta = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

            'InvntItem
            oItem.InventoryItem = IIf(String.IsNullOrEmpty(objArticulo.Inventario) OrElse objArticulo.Inventario = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

            '---------------------------------
            ' Pestaña general
            '---------------------------------

            'WTLiable
            oItem.WTLiable = IIf(String.IsNullOrEmpty(objArticulo.GeneralSujetoRetencion) OrElse objArticulo.GeneralSujetoRetencion = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

            'FirmCode
            If IsNumeric(objArticulo.GeneralFabricante) Then oItem.Manufacturer = CInt(objArticulo.GeneralFabricante)

            'ShipType
            If IsNumeric(objArticulo.GeneralClaseExpedicion) Then oItem.ShipType = CInt(objArticulo.GeneralClaseExpedicion)

            'ManBtchNum
            oItem.ManageBatchNumbers = IIf(String.IsNullOrEmpty(objArticulo.GeneralGestionarArticuloPorLotes) OrElse objArticulo.GeneralGestionarArticuloPorLotes = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

            'ISRelevant
            oItem.IntrastatExtension.IntrastatRelevant = IIf(String.IsNullOrEmpty(objArticulo.GeneralRelevanteInstratat) OrElse objArticulo.GeneralRelevanteInstratat = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

            'frozenFor
            oItem.Frozen = IIf(objArticulo.GeneralActivo = SN.Si, BoYesNoEnum.tYES, BoYesNoEnum.tNO)

            '---------------------------------
            ' Pestaña compras
            '---------------------------------

            'CardCode
            If Not String.IsNullOrEmpty(objArticulo.CompraProveedorDefecto) Then oItem.Mainsupplier = objArticulo.CompraProveedorDefecto

            'CardCode
            If Not String.IsNullOrEmpty(objArticulo.CompraNumeroCatalogoFabricante) Then oItem.SupplierCatalogNo = objArticulo.CompraNumeroCatalogoFabricante

            'BuyUnitMsr
            If Not String.IsNullOrEmpty(objArticulo.CompraUnidadMedida) Then oItem.PurchaseUnit = objArticulo.CompraUnidadMedida

            'NumInBuy
            If objArticulo.CompraArticulosUnidad > 0 Then oItem.PurchaseItemsPerUnit = CDbl(objArticulo.CompraArticulosUnidad)

            'PurPackMsr
            If Not String.IsNullOrEmpty(objArticulo.CompraUnidadMedidaEmbalaje) Then oItem.PurchasePackagingUnit = objArticulo.CompraUnidadMedidaEmbalaje

            'PurPackUn
            If objArticulo.CompraCantidadEmbalaje > 0 Then oItem.PurchaseQtyPerPackUnit = CDbl(objArticulo.CompraCantidadEmbalaje)

            'VatGroupPu
            If Not String.IsNullOrEmpty(objArticulo.CompraGrupoImpositivo) Then oItem.PurchaseVATGroup = objArticulo.CompraGrupoImpositivo

            '---------------------------------
            ' Pestaña ventas
            '---------------------------------

            'SalUnitMsr
            If Not String.IsNullOrEmpty(objArticulo.VentaUnidadMedida) Then oItem.SalesUnit = objArticulo.VentaUnidadMedida

            'NumInSale
            If objArticulo.VentaArticulosUnidad > 0 Then oItem.SalesItemsPerUnit = CDbl(objArticulo.VentaArticulosUnidad)

            'SalPackMsr
            If Not String.IsNullOrEmpty(objArticulo.VentaUnidadMedidaEmbalaje) Then oItem.SalesPackagingUnit = objArticulo.VentaUnidadMedidaEmbalaje

            'SalPackUn
            If objArticulo.VentaCantidadEmbalaje > 0 Then oItem.SalesQtyPerPackUnit = CDbl(objArticulo.VentaCantidadEmbalaje)

            'VatGroupSa
            If Not String.IsNullOrEmpty(objArticulo.VentaGrupoImpositivo) Then oItem.SalesVATGroup = objArticulo.VentaGrupoImpositivo

            '---------------------------------
            ' Pestaña inventario
            '---------------------------------

            'GLMethod
            If Not String.IsNullOrEmpty(objArticulo.InventarioCuentasDeMayorPor) Then oItem.GLMethod = CDbl(objArticulo.InventarioCuentasDeMayorPor)

            'InvntryUom
            If Not String.IsNullOrEmpty(objArticulo.InventarioUnidadMedida) Then oItem.InventoryUOM = objArticulo.InventarioUnidadMedida

            'IWeight1
            If objArticulo.InventarioPeso > 0 Then oItem.InventoryWeight = CDbl(objArticulo.InventarioPeso)

            'ByWh
            oItem.ManageStockByWarehouse = IIf(String.IsNullOrEmpty(objArticulo.InventarioGestionStockAlmacen) OrElse objArticulo.InventarioGestionStockAlmacen = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

            '---------------------------------
            ' Pestaña comentarios
            '---------------------------------

            'UserText
            If Not String.IsNullOrEmpty(objArticulo.Comentarios) Then oItem.User_Text = objArticulo.Comentarios

            '---------------------------------
            ' Campos de usuario
            '---------------------------------

            'Campos de usuario
            If Not objArticulo.CamposUsuario Is Nothing AndAlso objArticulo.CamposUsuario.Count > 0 Then

                For Each oCampoUsuario In objArticulo.CamposUsuario

                    Dim oUserField As Field = oItem.UserFields.Fields.Item(oCampoUsuario.Campo)

                    Select Case oUserField.Type

                        Case BoFieldTypes.db_Numeric
                            If oUserField.SubType = BoFldSubTypes.st_Time Then
                                If DateTime.TryParseExact(oCampoUsuario.Valor, "HHmm", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                    oItem.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "HHmm", CultureInfo.CurrentCulture)
                            Else
                                If IsNumeric(oCampoUsuario.Valor) Then _
                                    oItem.Lines.UserFields.Item(oCampoUsuario.Campo).Value = CInt(oCampoUsuario.Valor)
                            End If

                        Case BoFieldTypes.db_Float
                            If IsNumeric(oCampoUsuario.Valor) Then _
                                oItem.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CDbl(oCampoUsuario.Valor.Replace(".", ","))

                        Case BoFieldTypes.db_Date
                            If DateTime.TryParseExact(oCampoUsuario.Valor, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                oItem.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                        Case Else
                            If Not String.IsNullOrEmpty(oCampoUsuario.Valor) Then _
                                oItem.UserFields.Fields.Item(oCampoUsuario.Campo).Value = oCampoUsuario.Valor

                    End Select

                Next

            End If

            '---------------------------------
            ' Creación
            '---------------------------------

            'Añadimos artículo
            If oItem.Add() = 0 Then

                retVal.CODIGO = Respuesta.Ok
                retVal.MENSAJE = "Artículo creado con éxito"
                retVal.MENSAJEAUX = getItemCodeDeItemName(objArticulo.Nombre, Sociedad)

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
            oComun.LiberarObjCOM(oItem)
        End Try

        Return retVal

    End Function

    Private Function NuevoArticuloActivoFijo(ByVal objArticulo As EntArticuloSAP, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado
        Dim oCompany As Company
        Dim oItem As Items = Nothing

        Dim oComun As New clsComun

        Dim sLogInfo As String = "NUEVO ARTICULO AF"

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de creación de artículo para " & objArticulo.NIFEmpresa)

            oCompany = ConexionSAP.getCompany(objArticulo.UserSAP, objArticulo.PassSAP, Sociedad)
            If oCompany Is Nothing OrElse Not oCompany.Connected Then Throw New Exception("No se puede conectar a SAP")

            'Transacción
            If Not oCompany.InTransaction Then oCompany.StartTransaction() Else Throw New Exception("Transacción en curso")

            'Obligatorios
            If objArticulo.Serie <= 0 AndAlso String.IsNullOrEmpty(objArticulo.Codigo) Then Throw New Exception("Serie o código no suministrado")
            If String.IsNullOrEmpty(objArticulo.Nombre) Then Throw New Exception("Nombre no suministrado")

            'Objeto artículo
            oItem = oCompany.GetBusinessObject(BoObjectTypes.oItems)

            'Comprueba si existe el artículo
            If objArticulo.Serie <= 0 AndAlso Not String.IsNullOrEmpty(objArticulo.Codigo) Then

                If oItem.GetByKey(objArticulo.Codigo) Then Throw New Exception("Ya existe el artículo con ItemCode: " & objArticulo.Codigo)

            End If

            'El código depende de la serie

            'Series
            If objArticulo.Serie > 0 Then oItem.Series = objArticulo.Serie

            'ItemCode
            If Not String.IsNullOrEmpty(objArticulo.Codigo) Then oItem.ItemCode = objArticulo.Codigo

            'ItemName
            If Not String.IsNullOrEmpty(objArticulo.Nombre) Then oItem.ItemName = objArticulo.Nombre

            'FrgnName
            If Not String.IsNullOrEmpty(objArticulo.NombreExtranjero) Then oItem.ForeignName = objArticulo.NombreExtranjero

            'ItemType
            If IsNumeric(objArticulo.Tipo) Then oItem.ItemType = CInt(objArticulo.Tipo)

            'ItmsGrpCod
            If IsNumeric(objArticulo.Grupo) Then oItem.ItemsGroupCode = CInt(objArticulo.Grupo)

            'UgpEntry
            If IsNumeric(objArticulo.GrupoUnidadMedida) Then oItem.UoMGroupEntry = CInt(objArticulo.GrupoUnidadMedida)

            'CodeBars
            If Not String.IsNullOrEmpty(objArticulo.CodigoBarras) Then oItem.BarCode = objArticulo.CodigoBarras

            'PrchseItem
            oItem.PurchaseItem = IIf(String.IsNullOrEmpty(objArticulo.Compra) OrElse objArticulo.Compra = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

            'SellItem
            oItem.SalesItem = IIf(String.IsNullOrEmpty(objArticulo.Venta) OrElse objArticulo.Venta = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

            'InvntItem
            oItem.InventoryItem = IIf(String.IsNullOrEmpty(objArticulo.Inventario) OrElse objArticulo.Inventario = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

            '---------------------------------
            ' Pestaña general
            '---------------------------------

            'WTLiable
            oItem.WTLiable = IIf(String.IsNullOrEmpty(objArticulo.GeneralSujetoRetencion) OrElse objArticulo.GeneralSujetoRetencion = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

            'FirmCode
            If IsNumeric(objArticulo.GeneralFabricante) Then oItem.Manufacturer = CInt(objArticulo.GeneralFabricante)

            'ShipType
            If IsNumeric(objArticulo.GeneralClaseExpedicion) Then oItem.ShipType = CInt(objArticulo.GeneralClaseExpedicion)

            'ManBtchNum
            oItem.ManageBatchNumbers = IIf(String.IsNullOrEmpty(objArticulo.GeneralGestionarArticuloPorLotes) OrElse objArticulo.GeneralGestionarArticuloPorLotes = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

            'ISRelevant
            oItem.IntrastatExtension.IntrastatRelevant = IIf(String.IsNullOrEmpty(objArticulo.GeneralRelevanteInstratat) OrElse objArticulo.GeneralRelevanteInstratat = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

            'frozenFor
            oItem.Frozen = IIf(objArticulo.GeneralActivo = SN.Si, BoYesNoEnum.tYES, BoYesNoEnum.tNO)

            '---------------------------------
            ' Pestaña compras
            '---------------------------------

            'CardCode
            If Not String.IsNullOrEmpty(objArticulo.CompraProveedorDefecto) Then oItem.Mainsupplier = objArticulo.CompraProveedorDefecto

            'CardCode
            If Not String.IsNullOrEmpty(objArticulo.CompraNumeroCatalogoFabricante) Then oItem.SupplierCatalogNo = objArticulo.CompraNumeroCatalogoFabricante

            'BuyUnitMsr
            If Not String.IsNullOrEmpty(objArticulo.CompraUnidadMedida) Then oItem.PurchaseUnit = objArticulo.CompraUnidadMedida

            'NumInBuy
            If objArticulo.CompraArticulosUnidad > 0 Then oItem.PurchaseItemsPerUnit = CDbl(objArticulo.CompraArticulosUnidad)

            'PurPackMsr
            If Not String.IsNullOrEmpty(objArticulo.CompraUnidadMedidaEmbalaje) Then oItem.PurchasePackagingUnit = objArticulo.CompraUnidadMedidaEmbalaje

            'PurPackUn
            If objArticulo.CompraCantidadEmbalaje > 0 Then oItem.PurchaseQtyPerPackUnit = CDbl(objArticulo.CompraCantidadEmbalaje)

            'VatGroupPu
            If Not String.IsNullOrEmpty(objArticulo.CompraGrupoImpositivo) Then oItem.PurchaseVATGroup = objArticulo.CompraGrupoImpositivo

            '---------------------------------
            ' Pestaña ventas
            '---------------------------------

            'SalUnitMsr
            If Not String.IsNullOrEmpty(objArticulo.VentaUnidadMedida) Then oItem.SalesUnit = objArticulo.VentaUnidadMedida

            'NumInSale
            If objArticulo.VentaArticulosUnidad > 0 Then oItem.SalesItemsPerUnit = CDbl(objArticulo.VentaArticulosUnidad)

            'SalPackMsr
            If Not String.IsNullOrEmpty(objArticulo.VentaUnidadMedidaEmbalaje) Then oItem.SalesPackagingUnit = objArticulo.VentaUnidadMedidaEmbalaje

            'SalPackUn
            If objArticulo.VentaCantidadEmbalaje > 0 Then oItem.SalesQtyPerPackUnit = CDbl(objArticulo.VentaCantidadEmbalaje)

            'VatGroupSa
            If Not String.IsNullOrEmpty(objArticulo.VentaGrupoImpositivo) Then oItem.SalesVATGroup = objArticulo.VentaGrupoImpositivo

            '---------------------------------
            ' Pestaña inventario
            '---------------------------------

            'GLMethod
            If Not String.IsNullOrEmpty(objArticulo.InventarioCuentasDeMayorPor) Then oItem.GLMethod = CDbl(objArticulo.InventarioCuentasDeMayorPor)

            'InvntryUom
            If Not String.IsNullOrEmpty(objArticulo.InventarioUnidadMedida) Then oItem.InventoryUOM = objArticulo.InventarioUnidadMedida

            'IWeight1
            If objArticulo.InventarioPeso > 0 Then oItem.InventoryWeight = CDbl(objArticulo.InventarioPeso)

            'ByWh
            oItem.ManageStockByWarehouse = IIf(String.IsNullOrEmpty(objArticulo.InventarioGestionStockAlmacen) OrElse objArticulo.InventarioGestionStockAlmacen = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

            '---------------------------------
            ' Pestaña activos fijos
            '---------------------------------

            'AssetClass
            If Not String.IsNullOrEmpty(objArticulo.ActivoFijoClase) Then oItem.AssetClass = objArticulo.ActivoFijoClase

            'AssetGroup
            If Not String.IsNullOrEmpty(objArticulo.ActivoFijoGrupo) Then oItem.AssetGroup = objArticulo.ActivoFijoGrupo

            'DeprGroup
            If Not String.IsNullOrEmpty(objArticulo.ActivoFijoGrupoAmortizacion) Then oItem.DepreciationGroup = objArticulo.ActivoFijoGrupoAmortizacion

            'InventryNo
            If Not String.IsNullOrEmpty(objArticulo.ActivoFijoNumeroInventario) Then oItem.InventoryNumber = objArticulo.ActivoFijoNumeroInventario

            'AssetSerNo
            If Not String.IsNullOrEmpty(objArticulo.ActivoFijoNumeroSerie) Then oItem.AssetSerialNumber = objArticulo.ActivoFijoNumeroSerie

            'Location
            If objArticulo.ActivoFijoEmplazamiento > 0 Then oItem.Location = objArticulo.ActivoFijoEmplazamiento

            'Technician
            If objArticulo.ActivoFijoTecnico > 0 Then oItem.Technician = objArticulo.ActivoFijoTecnico

            'Employee
            If objArticulo.ActivoFijoEmpleado > 0 Then oItem.Employee = objArticulo.ActivoFijoEmpleado

            'CapDate
            If DateTime.TryParseExact(objArticulo.ActivoFijoFechaCapitalizacion.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                oItem.CapitalizationDate = Date.ParseExact(objArticulo.ActivoFijoFechaCapitalizacion.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            'StatAsset
            oItem.StatisticalAsset = IIf(String.IsNullOrEmpty(objArticulo.ActivoFijoEstadistico) OrElse objArticulo.ActivoFijoEstadistico = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

            'Cession
            oItem.Cession = IIf(String.IsNullOrEmpty(objArticulo.ActivoFijoCesion) OrElse objArticulo.ActivoFijoCesion = SN.No, BoYesNoEnum.tNO, BoYesNoEnum.tYES)

            ''DepreciationArea
            'If Not String.IsNullOrEmpty(objArticulo.ActivoFijoAreaValoracion) Then oItem.DepreciationParameters.DepreciationArea = objArticulo.ActivoFijoAreaValoracion

            ''DepreciationStartDate
            'If DateTime.TryParseExact(objArticulo.ActivoFijoFechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
            '    oItem.DepreciationParameters.DepreciationStartDate = Date.ParseExact(objArticulo.ActivoFijoFechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

            ''DepreciationType
            'If Not String.IsNullOrEmpty(objArticulo.ActivoFijoTipo) Then oItem.DepreciationParameters.DepreciationType = objArticulo.ActivoFijoTipo

            ''FiscalYear
            'If Not String.IsNullOrEmpty(objArticulo.ActivoFijoAnyoFiscal) Then oItem.DepreciationParameters.FiscalYear = objArticulo.ActivoFijoAnyoFiscal

            ''UsefulLife
            'If objArticulo.ActivoFijoVidaUtil > 0 Then oItem.DepreciationParameters.UsefulLife = objArticulo.ActivoFijoVidaUtil

            ''TotalUnitsInUsefulLife
            'If objArticulo.ActivoFijoVidaUtilUnidades > 0 Then oItem.DepreciationParameters.TotalUnitsInUsefulLife = objArticulo.ActivoFijoVidaUtilUnidades

            'CapitalGoodsOnHoldLimit
            If objArticulo.ActivoFijoCAPHistorico > 0 Then oItem.CapitalGoodsOnHoldLimit = objArticulo.ActivoFijoCAPHistorico

            '---------------------------------
            ' Pestaña comentarios
            '---------------------------------

            'UserText
            If Not String.IsNullOrEmpty(objArticulo.Comentarios) Then oItem.User_Text = objArticulo.Comentarios

            '---------------------------------
            ' Campos de usuario
            '---------------------------------

            'Campos de usuario
            If Not objArticulo.CamposUsuario Is Nothing AndAlso objArticulo.CamposUsuario.Count > 0 Then

                For Each oCampoUsuario In objArticulo.CamposUsuario

                    Dim oUserField As Field = oItem.UserFields.Fields.Item(oCampoUsuario.Campo)

                    Select Case oUserField.Type

                        Case BoFieldTypes.db_Numeric
                            If oUserField.SubType = BoFldSubTypes.st_Time Then
                                If DateTime.TryParseExact(oCampoUsuario.Valor, "HHmm", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                    oItem.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "HHmm", CultureInfo.CurrentCulture)
                            Else
                                If IsNumeric(oCampoUsuario.Valor) Then _
                                    oItem.Lines.UserFields.Item(oCampoUsuario.Campo).Value = CInt(oCampoUsuario.Valor)
                            End If

                        Case BoFieldTypes.db_Float
                            If IsNumeric(oCampoUsuario.Valor) Then _
                                oItem.UserFields.Fields.Item(oCampoUsuario.Campo).Value = CDbl(oCampoUsuario.Valor.Replace(".", ","))

                        Case BoFieldTypes.db_Date
                            If DateTime.TryParseExact(oCampoUsuario.Valor, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                                oItem.UserFields.Fields.Item(oCampoUsuario.Campo).Value = Date.ParseExact(oCampoUsuario.Valor.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                        Case Else
                            If Not String.IsNullOrEmpty(oCampoUsuario.Valor) Then _
                                oItem.UserFields.Fields.Item(oCampoUsuario.Campo).Value = oCampoUsuario.Valor

                    End Select

                Next

            End If

            '---------------------------------
            ' Creación
            '---------------------------------

            'Añadimos artículo
            If oItem.Add() = 0 Then

                'Si es activo fijo, actualiza parámetros de amortización
                retVal = ActualizarParametrosAmortizacionActivoFijo(oCompany, oItem, objArticulo, sLogInfo, Sociedad)

                'Todo correcto
                If Not oCompany Is Nothing AndAlso oCompany.InTransaction AndAlso retVal.CODIGO = Respuesta.Ok Then oCompany.EndTransaction(BoWfTransOpt.wf_Commit)

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
            'Cerrar transaccion
            If Not oCompany Is Nothing AndAlso oCompany.InTransaction Then oCompany.EndTransaction(BoWfTransOpt.wf_RollBack)

            'Liberar
            oComun.LiberarObjCOM(oItem)
        End Try

        Return retVal

    End Function

    Public Function ActualizarParametrosAmortizacionActivoFijo(ByRef oCompany As Company, ByRef oItem As Items, ByVal objArticulo As EntArticuloSAP, ByVal sLogInfo As String, ByVal Sociedad As eSociedad) As EntResultado

        Dim retVal As New EntResultado

        Try

            clsLog.Log.Info("(" & sLogInfo & " - " & NOMBRESOCIEDAD(Sociedad) & ") Inicio de actualización de amortización artículo: " & objArticulo.Codigo)

            'Comprueba si existe el artículo
            If Not oItem.GetByKey(objArticulo.Codigo) Then

                Throw New Exception("No se encuentra el artículo con ItemCode: " & objArticulo.Codigo)

            Else

                'DepreciationArea
                If Not String.IsNullOrEmpty(objArticulo.ActivoFijoAreaValoracion) Then oItem.DepreciationParameters.DepreciationArea = objArticulo.ActivoFijoAreaValoracion

                'DepreciationStartDate
                If DateTime.TryParseExact(objArticulo.ActivoFijoFechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture, DateTimeStyles.None, New Date) Then _
                    oItem.DepreciationParameters.DepreciationStartDate = Date.ParseExact(objArticulo.ActivoFijoFechaInicio.ToString, "yyyyMMdd", CultureInfo.CurrentCulture)

                'DepreciationType
                If Not String.IsNullOrEmpty(objArticulo.ActivoFijoTipo) Then oItem.DepreciationParameters.DepreciationType = objArticulo.ActivoFijoTipo

                'FiscalYear
                If Not String.IsNullOrEmpty(objArticulo.ActivoFijoAnyoFiscal) Then oItem.DepreciationParameters.FiscalYear = objArticulo.ActivoFijoAnyoFiscal

                'UsefulLife
                If objArticulo.ActivoFijoVidaUtil > 0 Then oItem.DepreciationParameters.UsefulLife = objArticulo.ActivoFijoVidaUtil

                'TotalUnitsInUsefulLife
                If objArticulo.ActivoFijoVidaUtilUnidades > 0 Then oItem.DepreciationParameters.TotalUnitsInUsefulLife = objArticulo.ActivoFijoVidaUtilUnidades

                'Actualizamos artículo
                If oItem.Update() = 0 Then

                    retVal.CODIGO = Respuesta.Ok
                    retVal.MENSAJE = "Artículo creado con éxito"
                    retVal.MENSAJEAUX = getItemCodeDeItemName(objArticulo.Nombre, Sociedad)

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
        End Try

        Return retVal

    End Function

    Private Function getItemCodeDeItemName(ByVal ItemName As String,
                                           ByVal Sociedad As eSociedad) As String

        'Devuelve el ItemCode del artículo

        Dim retVal As String = ""

        Try

            'Buscamos por ItemName
            Dim SQL As String = ""

            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " T0." & putQuotes("ItemCode") & " As ItemCode " & vbCrLf
            SQL &= " FROM " & getDataBaseRef("OITM", Sociedad) & " T0 " & getWithNoLock() & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("ItemName") & "  = N'" & ItemName & "'" & vbCrLf

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
