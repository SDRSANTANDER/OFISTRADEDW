CREATE VIEW SEI_VIEW_DW_ACTIVOS_FIJOS_EMPLAZAMIENTOS AS
SELECT
COALESCE(T0."Code",-1) As ID,
COALESCE(T0."Location",'') As NOMBRE
FROM OLCT T0 WITH(NOLOCK)
WHERE 1=1 

