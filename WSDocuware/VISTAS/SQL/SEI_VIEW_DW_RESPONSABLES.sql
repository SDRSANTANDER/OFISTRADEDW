CREATE VIEW SEI_VIEW_DW_RESPONSABLES AS
SELECT
COALESCE(T0."AgentCode",'') As ID,
COALESCE(T0."AgentName",'') As NOMBRE
FROM OAGP T0 WITH(NOLOCK)
WHERE 1=1 
