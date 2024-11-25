CREATE VIEW SEI_VIEW_DW_RELACIONES_DOCUMENTOS_COMPRA AS
SELECT 
-1 as 'SolicitudCompra', 
COALESCE(T0.[DocNum],-1) as 'PedidoVenta', 
-1 as 'PedidoCompra', 
COALESCE(T0.[CardCode],'') As 'InterlocutorId', 
COALESCE(T0.[CardName],'') As 'InterlocutorRazonSocial', 
COALESCE(T0.[DocDate],GETDATE()) As 'FechaContable',  
COALESCE(T1.[LineNum],-1) As 'LineaNumero', 
COALESCE(T1.[ItemCode],'') As 'ArticuloId', 
COALESCE(T1.[Dscription],'') As 'ArticuloDescripcion', 
COALESCE(T3.[QryGroup3],'N') As 'ArticuloTipoContrato',
COALESCE(T1.[Quantity],0) As 'Cantidad',  
COALESCE(T1.[Price],0) As 'Precio'
FROM ORDR T0 WITH(NOLOCK) 
INNER JOIN RDR1 T1 WITH(NOLOCK) ON T0.[DocEntry] = T1.[DocEntry]
LEFT JOIN OITM T3 WITH(NOLOCK) ON T3.[ItemCode]=T1.[ItemCode]
WHERE 1=1
AND COALESCE(T0.[U_SEITRADW],'N')='N'
AND T1.TrgetEntry IS NULL

UNION ALL

SELECT
COALESCE(T0.[DocNum],-1) As 'SolicitudCompra',
-1 As 'PedidoVenta', 
-1 as 'PedidoCompra', 
COALESCE(T0.[CardCode],'') As 'InterlocutorId', 
COALESCE(T0.[CardName],'') As 'InterlocutorRazonSocial', 
COALESCE(T0.[DocDate],GETDATE()) As 'FechaContable',  
COALESCE(T1.[LineNum],-1) As 'LineaNumero', 
COALESCE(T1.[ItemCode],'') As 'ArticuloId', 
COALESCE(T1.[Dscription],'') As 'ArticuloDescripcion', 
COALESCE(T3.[QryGroup3],'N') As 'ArticuloTipoContrato',
COALESCE(T1.[Quantity],0) As 'Cantidad',  
COALESCE(T1.[Price],0) As 'Precio'
FROM OPQT T0 WITH(NOLOCK) 
INNER JOIN PQT1 T1 WITH(NOLOCK) ON T0.[DocEntry] = T1.[DocEntry]
LEFT JOIN OITM T3 WITH(NOLOCK) ON T3.[ItemCode]=T1.[ItemCode]
WHERE 1=1
AND COALESCE(T0.[U_SEITRADW],'N')='N'

UNION ALL

SELECT 
-1 as 'SolicitudCompra', 
-1 as 'PedidoVenta', 
COALESCE(T0.[DocNum],-1) as 'PedidoCompra', 
COALESCE(T0.[CardCode],'') As 'InterlocutorId', 
COALESCE(T0.[CardName],'') As 'InterlocutorRazonSocial', 
COALESCE(T0.[DocDate],GETDATE()) As 'FechaContable',  
COALESCE(T1.[LineNum],-1) As 'LineaNumero', 
COALESCE(T1.[ItemCode],'') As 'ArticuloId', 
COALESCE(T1.[Dscription],'') As 'ArticuloDescripcion', 
COALESCE(T3.[QryGroup3],'N') As 'ArticuloTipoContrato',
COALESCE(T1.[Quantity],0) As 'Cantidad',  
COALESCE(T1.[Price],0) As 'Precio'
FROM OPOR T0 WITH(NOLOCK) 
INNER JOIN POR1 T1 WITH(NOLOCK) ON T0.[DocEntry] = T1.[DocEntry]
LEFT JOIN OITM T3 WITH(NOLOCK) ON T3.[ItemCode]=T1.[ItemCode]
WHERE 1=1
AND COALESCE(T0.[U_SEITRADW],'N')='N'
AND T1.TrgetEntry IS NULL





