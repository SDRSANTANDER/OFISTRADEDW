-------------------------------------------------------------------------------------------------------------------------
-- ICs
-- ACTUALIZAR CAMPO U_SEIIDDW CON VALOR CardCode [2023/09/22] (SEIDOR) 
-------------------------------------------------------------------------------------------------------------------------
IF @object_type='2' AND (@transaction_type= N'A' or  @transaction_type= N'U')

	BEGIN
		UPDATE OCRD SET U_SEIICDW=CardCode
		WHERE 1=1
		AND COALESCE(CardCode,'')  = @list_of_cols_val_tab_del 
	END


-------------------------------------------------------------------------------------------------------------------------
-- FACTURAS DE COMPRA
-- CONTROLAR IMPORTE DW FACTURAS DE PROVEEDOR [2020/09/21] (SEIDOR) 
-------------------------------------------------------------------------------------------------------------------------
IF @object_type = '18' AND (@transaction_type = N'A') 
BEGIN
	IF (				
		SELECT COUNT(T0."DocEntry")
		FROM OPCH T0
		WHERE 1=1
		AND T0."DocEntry" = @list_of_cols_val_tab_del 
		AND COALESCE(T0."U_SEIREVDW",'N')='S'
		AND COALESCE(T0."U_SEIIMPDW",0)<>COALESCE(T0."DocTotal",0)
		) > 0 

		BEGIN
			SELECT @error = 1, @error_message = 'TN: Importe DW! El importe de documento no coincide con el importe DW' 
		END
END


-------------------------------------------------------------------------------------------------------------------------
-- FACTURAS DE COMPRA
-- ACTUALIZAR ID Y URL DW FACTURAS DE PROVEEDOR EN FIRME CON CAMPOS DE BORRADOR [2023/11/30] (SEIDOR) 
-------------------------------------------------------------------------------------------------------------------------
IF @object_type='18' AND (@transaction_type= N'A')
BEGIN

	--Comprueba si hay que actualizar los campos
	IF (SELECT COUNT(*) FROM OPCH T0
	JOIN ODRF T1 ON T1.DocEntry=T0.draftKey AND T1.ObjType=T0.ObjType
	WHERE 1=1
	AND (COALESCE(T0.U_SEIIDDW,'')<>COALESCE(T1.U_SEIIDDW,'') OR COALESCE(Cast(T0.U_SEIURLDW as nvarchar(max)),'')<>COALESCE(Cast(T1.U_SEIURLDW as nvarchar(max)),''))
	AND COALESCE(Cast(T1.U_SEIURLDW as nvarchar(max)),'')<>''
	AND COALESCE(T0.DocEntry,'')=@list_of_cols_val_tab_del
	)>0

	BEGIN 

		--Actualiza los campos
		UPDATE T0 SET
		T0.U_SEIIDDW=T1.U_SEIIDDW,
		T0.U_SEIURLDW=T1.U_SEIURLDW
		FROM OPCH T0
		JOIN ODRF T1 ON T1.DocEntry=T0.draftKey AND T1.ObjType=T0.ObjType
		WHERE 1=1
		AND (COALESCE(T0.U_SEIIDDW,'')<>COALESCE(T1.U_SEIIDDW,'') OR COALESCE(Cast(T0.U_SEIURLDW as nvarchar(max)),'')<>COALESCE(Cast(T1.U_SEIURLDW as nvarchar(max)),''))
		AND COALESCE(Cast(T1.U_SEIURLDW as nvarchar(max)),'')<>''
		AND COALESCE(T0.DocEntry,'')=@list_of_cols_val_tab_del
				
		--Comprueba los campos actualizados
		IF (SELECT COUNT(*) FROM OPCH T0
		JOIN ODRF T1 ON T1.DocEntry=T0.draftKey AND T1.ObjType=T0.ObjType
		WHERE 1=1
		AND (COALESCE(T0.U_SEIIDDW,'')<>COALESCE(T1.U_SEIIDDW,'') OR COALESCE(Cast(T0.U_SEIURLDW as nvarchar(max)),'')<>COALESCE(Cast(T1.U_SEIURLDW as nvarchar(max)),''))
		AND COALESCE(Cast(T1.U_SEIURLDW as nvarchar(max)),'')<>''
		AND COALESCE(T0.DocEntry,'')=@list_of_cols_val_tab_del
		)>0
			BEGIN
				SELECT @error = 1
				SELECT @error_message=N'DocuWare: No se ha podido actualizar el ID y URL de la factura de compra en firme'
			END
	END

END


-------------------------------------------------------------------------------------------------------------------------
-- DOCUMENTOS DE COMPRA
-- CONTROLAR NumAtCard NO REPETIDO PARA MISMO TIPO DE DOCUMENTO DE COMPRA Y PROVEEDOR [2022/11/30] (SEIDOR) 
-------------------------------------------------------------------------------------------------------------------------

-- Pedido compra
IF @object_type='22' AND (@transaction_type= N'A' or  @transaction_type= N'U')
BEGIN
	IF (SELECT COUNT(*) FROM OPOR T0 
	WHERE 1=1
	AND COALESCE(T0.U_SEIIDDW,'')<>''
	AND COALESCE(T0.CardCode,'')  = (SELECT COALESCE(X0.CardCode,'') FROM OPOR X0 WHERE X0.Docentry=@list_of_cols_val_tab_del AND X0.CANCELED='N')
	AND COALESCE(T0.NumAtCard,'') = (SELECT COALESCE(X0.NumAtCard,'') FROM OPOR X0 WHERE X0.Docentry=@list_of_cols_val_tab_del AND X0.CANCELED='N')
	)>1
		BEGIN
			SELECT @error = 1
			SELECT @error_message=N'DocuWare: Ya existe la referencia de proveedor en otro pedido de compra'
		END
END

-- Entrega compra
IF @object_type='20' AND (@transaction_type= N'A' or  @transaction_type= N'U')
BEGIN
	IF (SELECT COUNT(*) FROM OPDN T0 
	WHERE 1=1
	AND COALESCE(T0.U_SEIIDDW,'')<>''
	AND COALESCE(T0.CardCode,'')  = (SELECT COALESCE(X0.CardCode,'') FROM OPDN X0 WHERE X0.Docentry=@list_of_cols_val_tab_del AND X0.CANCELED='N')
	AND COALESCE(T0.NumAtCard,'') = (SELECT COALESCE(X0.NumAtCard,'') FROM OPDN X0 WHERE X0.Docentry=@list_of_cols_val_tab_del AND X0.CANCELED='N')
	)>1
		BEGIN
			SELECT @error = 1
			SELECT @error_message=N'DocuWare: Ya existe la referencia de proveedor en otra entrega de compra'
		END
END

-- Devolucion compra
IF @object_type='21' AND (@transaction_type= N'A' or  @transaction_type= N'U')
BEGIN
	IF (SELECT COUNT(*) FROM ORPD T0 
	WHERE 1=1
	AND COALESCE(T0.U_SEIIDDW,'')<>''
	AND COALESCE(T0.CardCode,'')  = (SELECT COALESCE(X0.CardCode,'') FROM ORPD X0 WHERE X0.Docentry=@list_of_cols_val_tab_del AND X0.CANCELED='N')
	AND COALESCE(T0.NumAtCard,'') = (SELECT COALESCE(X0.NumAtCard,'') FROM ORPD X0 WHERE X0.Docentry=@list_of_cols_val_tab_del AND X0.CANCELED='N')
	)>1
		BEGIN
			SELECT @error = 1
			SELECT @error_message=N'DocuWare: Ya existe la referencia de proveedor en otra devolución de compra'
		END
END

-- Factura compra
IF @object_type='18' AND (@transaction_type= N'A' or  @transaction_type= N'U')
BEGIN
	IF (SELECT COUNT(*) FROM OPCH T0 
	WHERE 1=1
	AND COALESCE(T0.U_SEIIDDW,'')<>''
	AND COALESCE(T0.CardCode,'')  = (SELECT COALESCE(X0.CardCode,'') FROM OPCH X0 WHERE X0.Docentry=@list_of_cols_val_tab_del AND X0.CANCELED='N')
	AND COALESCE(T0.NumAtCard,'') = (SELECT COALESCE(X0.NumAtCard,'') FROM OPCH X0 WHERE X0.Docentry=@list_of_cols_val_tab_del AND X0.CANCELED='N')
	)>1
		BEGIN
			SELECT @error = 1
			SELECT @error_message=N'DocuWare: Ya existe la referencia de proveedor en otra factura de compra'
		END
END

-- Abono compra
IF @object_type='19' AND (@transaction_type= N'A' or  @transaction_type= N'U')
BEGIN
	IF (SELECT COUNT(*) FROM ORPC T0 
	WHERE 1=1
	AND COALESCE(T0.U_SEIIDDW,'')<>''
	AND COALESCE(T0.CardCode,'')  = (SELECT COALESCE(X0.CardCode,'') FROM ORPC X0 WHERE X0.Docentry=@list_of_cols_val_tab_del AND X0.CANCELED='N')
	AND COALESCE(T0.NumAtCard,'') = (SELECT COALESCE(X0.NumAtCard,'') FROM ORPC X0 WHERE X0.Docentry=@list_of_cols_val_tab_del AND X0.CANCELED='N')
	)>1
		BEGIN
			SELECT @error = 1
			SELECT @error_message=N'DocuWare: Ya existe la referencia de proveedor en otro abono de compra'
		END
END