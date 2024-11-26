CREATE VIEW SEI_VIEW_DW_EMPLEADOS AS 
SELECT
COALESCE(T0."empID",-1) As ID,
COALESCE(T0."firstName",'') + ' ' + COALESCE(T0."lastName",'') As NOMBRE,
COALESCE(T0."email",'') As CORREOE,
COALESCE(T0."CostCenter",'') As CENTROCOSTE 
FROM OHEM T0 WITH(NOLOCK)
WHERE 1=1 
And T0."Active"='Y'