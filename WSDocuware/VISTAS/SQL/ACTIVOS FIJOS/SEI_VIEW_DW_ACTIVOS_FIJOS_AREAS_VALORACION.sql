CREATE VIEW SEI_VIEW_DW_ACTIVOS_FIJOS_AREAS_VALORACION AS
SELECT
COALESCE(T0."Code",'') As ID,
COALESCE(T0."Descr",'') As NOMBRE,
COALESCE(T0."AreaType",'') As TIPO
FROM ODPA T0 WITH(NOLOCK)
WHERE 1=1 


