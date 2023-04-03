FUNCION  MAX

-- MAX (Columna)
--Cuando fue la ultima fecha y la primera fecha de la orden?
SELECT 
MAX(OrderDate) [LasDate],
MIN(OrderDate) [FirstDate]
FROM Sales.SalesOrderHeader
--Codigo que te permite tener la ultima fecha del registro de la orden .Asi mismo la primera orden de todos los registro--