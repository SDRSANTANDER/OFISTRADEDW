CREATE VIEW SEI_VIEW_DW_GRUPOS_ARTICULO AS
SELECT
COALESCE(T0."ItmsGrpCod",-1) As ID,
COALESCE(T0."ItmsGrpNam",'') As NOMBRE
FROM OITB T0 WITH(NOLOCK)
WHERE 1=1 

 
