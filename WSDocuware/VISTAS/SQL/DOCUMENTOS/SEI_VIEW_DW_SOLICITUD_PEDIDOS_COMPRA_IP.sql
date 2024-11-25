CREATE VIEW SEI_VIEW_DW_SOLICITUD_PEDIDOS_COMPRA_IP AS
SELECT 
COALESCE(T3.[empID],-1) As 'InvestigadorPrincipalId', 
COALESCE(T3.[firstName],'') As 'InvestigadorPrincipalNombre', 
COALESCE(T3.[lastName],'') As 'InvestigadorPrincipalApellido', 
COALESCE(T3.[email],'') As 'InvestigadorPrincipalMail', 
COALESCE(T0.[DocEntry],'') As 'Numerador', 
COALESCE(T0.[DocNum],-1) As 'DocumentoNumero', 
COALESCE(T0.[CardCode],'') As 'InterlocutorId', 
COALESCE(T0.[CardName],'') As 'InterLocutorRazonSocial', 
COALESCE(T0.[DocDate],GETDATE()) As 'FechaContable', 
COALESCE(T0.[DocDueDate],GETDATE()) As 'FechaEntrega', 
COALESCE(T0.[NumAtCard],'') As 'InterlocutorNumeroReferencia', 
COALESCE(T3.[empID],-1) As 'InvestigadorId', 
COALESCE(T3.[firstName],'') As 'InvestigadorNombre', 
COALESCE(T3.[lastName],'') As 'InvestigadorApellido', 
COALESCE(T3.[email],'') As 'InvestigadorMail', 
COALESCE(T1.[LineNum]+1,-1) As 'LineaNumero',
COALESCE(T1.[LineStatus],'') As 'LineaEstado', 
COALESCE(T1.[ItemCode],'') As 'ArticuloId', 
COALESCE(T1.[Dscription],'') As 'ArticuloDescripcion', 
COALESCE(T1.[Quantity],0) As 'Cantidad', 
COALESCE(T1.[Price],0) As 'Precio', 
COALESCE(T1.[LineTotal],0) As 'LineaTotal', 
COALESCE(T1.[Project],'') As 'ProyectoId', 
'' As 'ProyectoDescripcion', 
'' As 'Hito', 
COALESCE(T1.[OcrCode],'') As 'NormaReparto', 
COALESCE(T1.[FreeTxt],'') As 'TextoLibre', 
'' As 'CASS', 
'' As 'ArticuloReferencia', 
'' As 'ArticuloGenericoLaboratorio',
COALESCE(T5.[QryGroup3],'N') As 'ArticuloTipoContrato'
FROM [dbo].[OPQT] T0 WITH(NOLOCK) 
INNER JOIN [dbo].[PQT1] T1 WITH(NOLOCK) ON T0.[DocEntry] = T1.[DocEntry]
LEFT JOIN [dbo].[OHEM] T3 WITH(NOLOCK) ON T0.[OwnerCode] = T3.[empID]
LEFT JOIN [dbo].[OITM] T5 WITH(NOLOCK) ON T5.[ItemCode]=T1.[ItemCode]
WHERE 1=1 
AND COALESCE(T0.[U_SEITRADW],'N')='N'

