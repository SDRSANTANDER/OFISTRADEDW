CREATE VIEW SEI_VIEW_DW_CLASES_EXPEDICION AS
SELECT
COALESCE(T0."TrnspCode",-1) As ID,
COALESCE(T0."TrnspName",'') As NOMBRE
FROM OSHP T0 WITH(NOLOCK)
WHERE 1=1 
And T0."Active"='Y' 

