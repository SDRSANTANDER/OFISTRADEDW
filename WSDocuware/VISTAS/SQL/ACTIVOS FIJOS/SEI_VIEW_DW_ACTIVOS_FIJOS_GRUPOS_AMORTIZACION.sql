CREATE VIEW SEI_VIEW_DW_ACTIVOS_FIJOS_GRUPOS_AMORTIZACION AS
SELECT
COALESCE(T0."Code",'') As ID,
COALESCE(T0."Descr",'') As NOMBRE,
COALESCE(T0."Group",'') As GRUPO
FROM OADG T0 WITH(NOLOCK)
WHERE 1=1 

