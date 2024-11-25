-------------------------------------------------------------------------------------------------------------------------
-- FACTURAS DE COMPRA
-- CONTROLAR ID, URL E IMPORTE DW FACTURAS DE PROVEEDOR [2024/01/17] (SEIDOR) 
-------------------------------------------------------------------------------------------------------------------------
IF @object_type = '18' AND (@transaction_type = N'A') 

	BEGIN
		UPDATE T0 SET 
		T0.U_SEIIDDW=T1.U_SEIIDDW,
		T0.U_SEIURLDW=T1.U_SEIURLDW,
		T0.U_SEIIMPDW=T1.U_SEIIMPDW
		FROM OPCH T0 
		LEFT JOIN ODRF T1 ON T1.DocEntry=T0.draftKey
		WHERE 1=1	
		And COALESCE(T0.U_SEIIDDW,'')=''
		And COALESCE(T1.U_SEIIDDW,'')<>''
	END
