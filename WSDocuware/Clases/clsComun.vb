Imports System.Reflection
Imports WSDocuware.Utilidades

Public Class clsComun

#Region "IC"

    Public Function getCardCode(ByVal NIFTercero As String,
                                 ByVal RazonSocial As String,
                                 ByVal Ambito As String,
                                 ByVal Sociedad As eSociedad)

        'Devuelve el CardCode

        Dim CardCode As String = ""

        Dim oComun As New clsComun

        Try

            'Buscamos IC por CardCode, NIF, U_SEIICDW o razón social

            If Not String.IsNullOrEmpty(NIFTercero) AndAlso Not String.IsNullOrEmpty(RazonSocial) Then _
                CardCode = oComun.getCardCodeDeNIFRazonSocial(NIFTercero, RazonSocial, Ambito, BusquedaIC.CardCode, Sociedad)

            If String.IsNullOrEmpty(CardCode) AndAlso Not String.IsNullOrEmpty(NIFTercero) AndAlso Not String.IsNullOrEmpty(RazonSocial) Then _
                CardCode = oComun.getCardCodeDeNIFRazonSocial(NIFTercero, RazonSocial, Ambito, BusquedaIC.NIF, Sociedad)

            If String.IsNullOrEmpty(CardCode) AndAlso Not String.IsNullOrEmpty(NIFTercero) AndAlso Not String.IsNullOrEmpty(RazonSocial) Then _
                CardCode = oComun.getCardCodeDeNIFRazonSocial(NIFTercero, RazonSocial, Ambito, BusquedaIC.U_SEIICDW, Sociedad)

            If String.IsNullOrEmpty(CardCode) AndAlso Not String.IsNullOrEmpty(NIFTercero) Then _
                CardCode = oComun.getCardCodeDeNIF(NIFTercero, Ambito, BusquedaIC.CardCode, Sociedad)

            If String.IsNullOrEmpty(CardCode) AndAlso Not String.IsNullOrEmpty(NIFTercero) Then _
                CardCode = oComun.getCardCodeDeNIF(NIFTercero, Ambito, BusquedaIC.NIF, Sociedad)

            If String.IsNullOrEmpty(CardCode) AndAlso Not String.IsNullOrEmpty(NIFTercero) Then _
                CardCode = oComun.getCardCodeDeNIF(NIFTercero, Ambito, BusquedaIC.U_SEIICDW, Sociedad)

            If String.IsNullOrEmpty(CardCode) AndAlso Not String.IsNullOrEmpty(RazonSocial) Then _
                CardCode = oComun.getCardCodeDeRazonSocial(RazonSocial, Ambito, Sociedad)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return CardCode

    End Function

    Private Function getCardCodeDeNIFRazonSocial(ByVal NIF As String,
                                                 ByVal RazonSocial As String,
                                                 ByVal Ambito As String,
                                                 ByVal sBusquedaIC As String,
                                                 ByVal Sociedad As eSociedad) As String

        'Devuelve el cardcode de un NIF, U_SEIICDW o razón social

        Dim retVal As String = ""

        Try

            'Buscamos por razón social y NIF o U_SEIICCW
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("CardCode") & vbCrLf
            SQL &= " FROM " & putQuotes("OCRD") & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            SQL &= " And T0." & putQuotes("CardName") & " LIKE '%" & RazonSocial & "%'" & vbCrLf

            Select Case sBusquedaIC
                Case BusquedaIC.CardCode
                    SQL &= " And T0." & putQuotes("CardCode") & " = N'" & NIF & "'" & vbCrLf
                Case BusquedaIC.NIF
                    SQL &= " And T0." & putQuotes("LicTradNum") & " LIKE '%" & NIF & "%'" & vbCrLf
                Case Else
                    SQL &= " And COALESCE(T0." & putQuotes("U_SEIICDW") & ",'') LIKE '%" & NIF & "%'" & vbCrLf
            End Select

            Select Case Ambito
                Case Utilidades.Ambito.Compras
                    SQL &= " And T0." & putQuotes("CardType") & "=N'" & CardType.Proveedor & "'" & vbCrLf
                Case Utilidades.Ambito.Ventas
                    SQL &= " And T0." & putQuotes("CardType") & "=N'" & CardType.Cliente & "'" & vbCrLf
            End Select

            SQL &= " And T0." & putQuotes("frozenFor") & "<>N'" & SN.Yes & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getCardCodeDeNIF(ByVal NIF As String,
                                      ByVal Ambito As String,
                                      ByVal sBusquedaIC As String,
                                      ByVal Sociedad As eSociedad) As String

        'Devuelve el cardcode de un NIF

        Dim retVal As String = ""

        Try

            'Buscamos por NIF o U_SEIICCW
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("CardCode") & vbCrLf
            SQL &= " FROM " & putQuotes("OCRD") & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            Select Case sBusquedaIC
                Case BusquedaIC.CardCode
                    SQL &= " And T0." & putQuotes("CardCode") & " = N'" & NIF & "'" & vbCrLf
                Case BusquedaIC.NIF
                    SQL &= " And T0." & putQuotes("LicTradNum") & " LIKE '%" & NIF & "%'" & vbCrLf
                Case Else
                    SQL &= " And COALESCE(T0." & putQuotes("U_SEIICDW") & ",'') LIKE '%" & NIF & "%'" & vbCrLf
            End Select

            Select Case Ambito
                Case Utilidades.Ambito.Compras
                    SQL &= " And T0." & putQuotes("CardType") & "=N'" & CardType.Proveedor & "'" & vbCrLf
                Case Utilidades.Ambito.Ventas
                    SQL &= " And T0." & putQuotes("CardType") & "=N'" & CardType.Cliente & "'" & vbCrLf
            End Select

            SQL &= " And T0." & putQuotes("frozenFor") & "<>N'" & SN.Yes & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getCardCodeDeRazonSocial(ByVal RazonSocial As String,
                                              ByVal Ambito As String,
                                              ByVal Sociedad As eSociedad) As String

        'Devuelve el cardcode de un nombre

        Dim retVal As String = ""

        Try

            'Buscamos por CardName
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("CardCode") & vbCrLf
            SQL &= " FROM " & putQuotes("OCRD") & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("CardName") & " LIKE '%" & RazonSocial & "%'" & vbCrLf

            Select Case Ambito
                Case Utilidades.Ambito.Compras
                    SQL &= " And T0." & putQuotes("CardType") & "=N'" & CardType.Proveedor & "'" & vbCrLf
                Case Utilidades.Ambito.Ventas
                    SQL &= " And T0." & putQuotes("CardType") & "=N'" & CardType.Cliente & "'" & vbCrLf
            End Select

            SQL &= " And T0." & putQuotes("frozenFor") & "<>N'" & SN.Yes & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

#End Region

#Region "Artículo"

    Public Function getItemCode(ByVal CardCode As String,
                                ByVal Articulo As String,
                                ByVal RefExt As String,
                                ByVal Sociedad As eSociedad)

        'Devuelve el ItemCode

        Dim ItemCode As String = ""

        Dim oComun As New clsComun

        Try

            'Buscamos ItemCode por CardCode, ItemCode o Substitute

            If Not String.IsNullOrEmpty(Articulo) Then _
                ItemCode = oComun.getItemCodeDeArticulo(Articulo, Sociedad)

            If String.IsNullOrEmpty(ItemCode) AndAlso Not String.IsNullOrEmpty(CardCode) AndAlso Not String.IsNullOrEmpty(Articulo) Then _
                ItemCode = oComun.getItemCodeDeReferenciaProveedor(CardCode, Articulo, Sociedad)

            If String.IsNullOrEmpty(ItemCode) AndAlso Not String.IsNullOrEmpty(RefExt) Then _
                ItemCode = oComun.getItemCodeDeArticulo(RefExt, Sociedad)

            If String.IsNullOrEmpty(ItemCode) AndAlso Not String.IsNullOrEmpty(CardCode) AndAlso Not String.IsNullOrEmpty(RefExt) Then _
                ItemCode = oComun.getItemCodeDeReferenciaProveedor(CardCode, RefExt, Sociedad)

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return ItemCode

    End Function

    Private Function getItemCodeDeArticulo(ByVal Articulo As String,
                                           ByVal Sociedad As eSociedad) As String

        'Devuelve el ItemCode 

        Dim retVal As String = ""

        Try

            'Buscamos por código
            Dim SQL As String = ""
            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " T0." & putQuotes("ItemCode") & vbCrLf
            SQL &= " FROM " & putQuotes("OITM") & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            SQL &= " And T0." & putQuotes("ItemCode") & " = N'" & Articulo & "'" & vbCrLf
            SQL &= " And T0." & putQuotes("frozenFor") & " = N'" & SN.No & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Private Function getItemCodeDeReferenciaProveedor(ByVal CardCode As String,
                                                      ByVal Articulo As String,
                                                      ByVal Sociedad As eSociedad) As String

        'Devuelve el ItemCode 

        Dim retVal As String = ""

        Try

            'Buscamos por CardCode y referencia
            Dim SQL As String = ""
            SQL = "  SELECT TOP 1 " & vbCrLf
            SQL &= " T0." & putQuotes("ItemCode") & vbCrLf
            SQL &= " FROM " & putQuotes("OITM") & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " JOIN " & putQuotes("OSCN") & " T1 " & getWithNoLock() & " ON T1." & putQuotes("ItemCode") & " = T0." & putQuotes("ItemCode") & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            SQL &= " And T1." & putQuotes("CardCode") & " = N'" & CardCode & "'" & vbCrLf
            SQL &= " And T1." & putQuotes("Substitute") & " = N'" & Articulo & "'" & vbCrLf
            SQL &= " And T0." & putQuotes("frozenFor") & " = N'" & SN.No & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function


#End Region

#Region "Documentos"

    Public Function getSerieDeDocumentoDestino(ByVal SerieOrigen As Integer,
                                               ByVal ObjTypeOrigen As Integer,
                                               ByVal ObjTypeDestino As Integer,
                                               ByVal Sociedad As eSociedad) As String

        'Devuelve la serie del documento destino a partir del documento origen

        Dim retVal As String = ""

        Try


            Dim SQL As String
            SQL = "  SELECT "
            SQL &= " COALESCE(" & setNumberAsString("T1.", "Series") & ",'') " & vbCrLf
            SQL &= " FROM " & putQuotes("NNM1") & " T0 " & getWithNoLock() & " "
            SQL &= " JOIN " & putQuotes("NNM1") & " T1 " & getWithNoLock() & " ON T1." & putQuotes("SeriesName") & " = T0." & putQuotes("SeriesName") & " "

            SQL &= " WHERE 1=1 "

            SQL &= " And T0." & putQuotes("Series") & "  = N'" & SerieOrigen & "'"
            SQL &= " And T0." & putQuotes("ObjectCode") & "  = N'" & ObjTypeOrigen & "'"
            SQL &= " And T1." & putQuotes("ObjectCode") & " = N'" & ObjTypeDestino & "'"

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getDocEntryDocumentoDefinitivo(ByVal Tabla As String,
                                                   ByVal CardCode As String,
                                                   ByVal DOCIDDW As String,
                                                   ByVal DocNum As String,
                                                   ByVal NumAtCard As String,
                                                   ByVal bAbierto As Boolean,
                                                   ByVal bCancelado As Boolean,
                                                   ByVal Sociedad As eSociedad) As String

        'Devuelve el DocEntry del documento

        Dim retVal As String = ""

        Try

            'Tablas en donde el CardCode va a nivel de línea
            Dim bICLinea As Boolean = CBool(Tabla = getTablaDeObjType(ObjType.SolicitudCompra))

            'Buscamos por DOCIDDW, DocNum o NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("DocEntry") & " As DocEntry " & vbCrLf
            SQL &= " FROM " & putQuotes(Tabla) & " T0 " & getWithNoLock() & vbCrLf

            If bICLinea Then SQL &= " JOIN " & putQuotes(Tabla.Substring(1, 3) & "1") & " T1 " & getWithNoLock() & " ON T1." & putQuotes("DocEntry") & " = T0." & putQuotes("DocEntry") & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(CardCode) AndAlso Not bICLinea Then SQL &= " And T0." & putQuotes("CardCode") & "  = N'" & CardCode & "'" & vbCrLf
            If Not String.IsNullOrEmpty(CardCode) AndAlso bICLinea Then SQL &= " And T1." & putQuotes("LineVendor") & "  = N'" & CardCode & "'" & vbCrLf

            If Not String.IsNullOrEmpty(NumAtCard) Then SQL &= " And T0." & putQuotes("NumAtCard") & "  = N'" & NumAtCard & "'" & vbCrLf
            If Not String.IsNullOrEmpty(DocNum) Then SQL &= " And T0." & putQuotes("DocNum") & "  = N'" & DocNum & "'" & vbCrLf
            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & "  = N'" & DOCIDDW & "'" & vbCrLf

            If bAbierto Then SQL &= " And T0." & putQuotes("DocStatus") & " <> N'" & DocStatus.Cerrado & "'" & vbCrLf
            If bCancelado Then SQL &= " And T0." & putQuotes("CANCELED") & " = N'" & SN.No & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getDocNumDeDocEntry(ByVal Tabla As String,
                                         ByVal DocEntry As String,
                                         ByVal Sociedad As eSociedad) As String

        'Devuelve el docNum del documento

        Dim retVal As String = ""

        Try

            'Buscamos por DocEntry
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("DocNum") & vbCrLf
            SQL &= " FROM " & putQuotes(Tabla) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("DocEntry") & " = N'" & DocEntry & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getNextCode(ByVal Tabla As String, ByVal Sociedad As eSociedad) As String

        'Devuelve el siguiente code 

        Dim retVal As String = ""

        Try

            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " COALESCE(" & getStringAsNumber("T0.", "Code") & ",0)+1 " & vbCrLf
            SQL &= " FROM " & putQuotes(Tabla) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf

            Dim oCon As clsConexion =New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not String.IsNullOrEmpty(oObj) AndAlso IsNumeric(oObj.ToString) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Sub EliminarFichero(ByVal sFileName As String)

        'Elimina un fichero
        Try

            If IO.File.Exists(sFileName) Then
                IO.File.Delete(sFileName)
            End If

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

    End Sub

#End Region

#Region "Documentos"

    Public Function getTransIdAsientoDefinitivo(ByVal Tabla As String,
                                                ByVal DOCIDDW As String,
                                                ByVal TransNum As String,
                                                ByVal Ref1 As String,
                                                ByVal Ref2 As String,
                                                ByVal Ref3 As String,
                                                ByVal Sociedad As eSociedad) As String

        'Devuelve el TransId del Asiento

        Dim retVal As String = ""

        Try

            'Buscamos por DOCIDDW, DocNum o NumAtCard
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("TransId") & " As TransId " & vbCrLf
            SQL &= " FROM " & putQuotes(Tabla) & " T0 " & getWithNoLock() & vbCrLf

            SQL &= " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrEmpty(TransNum) Then SQL &= " And T0." & putQuotes("Number") & " = N'" & TransNum & "'" & vbCrLf

            If Not String.IsNullOrEmpty(Ref1) Then SQL &= " And T0." & putQuotes("Ref1") & " = N'" & Ref1 & "'" & vbCrLf
            If Not String.IsNullOrEmpty(Ref2) Then SQL &= " And T0." & putQuotes("Ref2") & " = N'" & Ref2 & "'" & vbCrLf
            If Not String.IsNullOrEmpty(Ref3) Then SQL &= " And T0." & putQuotes("Ref3") & " = N'" & Ref3 & "'" & vbCrLf

            If Not String.IsNullOrEmpty(DOCIDDW) Then SQL &= " And T0." & putQuotes("U_SEIIDDW") & " = N'" & DOCIDDW & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function

    Public Function getTransNumDeTransId(ByVal Tabla As String,
                                         ByVal TransId As String,
                                         ByVal Sociedad As eSociedad) As String

        'Devuelve el TransNum del Transumento

        Dim retVal As String = ""

        Try

            'Buscamos por TransEntry
            Dim SQL As String = ""
            SQL = "  SELECT " & vbCrLf
            SQL &= " T0." & putQuotes("Number") & vbCrLf
            SQL &= " FROM " & putQuotes(Tabla) & " T0 " & getWithNoLock() & vbCrLf
            SQL &= " WHERE 1=1 " & vbCrLf
            SQL &= " And T0." & putQuotes("TransId") & " = N'" & TransId & "'" & vbCrLf

            Dim oCon As clsConexion = New clsConexion(Sociedad)

            Dim oObj As Object = oCon.ExecuteScalar(SQL)
            If Not oObj Is Nothing AndAlso Not String.IsNullOrEmpty(oObj) Then retVal = oObj.ToString

        Catch ex As Exception
            clsLog.Log.Fatal(ex.Message & " en " & MethodBase.GetCurrentMethod().Name)
        End Try

        Return retVal

    End Function


#End Region

    Public Sub LiberarObjCOM(ByRef oObjCOM As Object, Optional ByVal bCollect As Boolean = False)
        '
        'Liberar y destruir Objecto com 
        ' En los UDO'S es necesario utilizar GC.Collect  para eliminarlos de la memoria
        If Not IsNothing(oObjCOM) Then
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oObjCOM)
            oObjCOM = Nothing
            If bCollect Then
                GC.Collect()
            End If
        End If

    End Sub

End Class
