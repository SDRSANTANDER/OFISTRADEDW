CREATE VIEW SEI_VIEW_DW_LLAMADAS_PROBLEMA_TIPO AS 
SELECT
COALESCE(T0."prblmTypID", -1) As ID,
COALESCE(T0."Name",'') As NOMBRE,
COALESCE(T0."Descriptio",'') As DESCRIPCION
FROM OSCP T0 WITH(NOLOCK)
WHERE 1=1 
And T0."Active"='Y'
