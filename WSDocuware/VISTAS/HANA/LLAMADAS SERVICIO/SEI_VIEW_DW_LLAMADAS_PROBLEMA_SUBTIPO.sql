CREATE VIEW "BASEDATOS_HANA"."SEI_VIEW_DW_LLAMADAS_PROBLEMA_SUBTIPO" AS 
SELECT
COALESCE(T0."ProSubTyId", -1) As ID,
COALESCE(T0."Name",'') As NOMBRE,
COALESCE(T0."Descriptio",'') As DESCRIPCION
FROM OPST T0
WHERE 1=1 
And T0."Active"='Y'
