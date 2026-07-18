SELECT 
	o.order_date,
    o.order_id,
    c.first_name,
    sh.name AS 'shipper',
    os.name AS 'status'
FROM orders o
LEFT JOIN customers c
	ON c.customer_id = o.customer_id
LEFT JOIN shippers sh
# KAKS KORDA LEFT JOIN, KUNA TAHAME, ET customers oleks baas ning selle alusel
#oleksid kõik kliendid kuvatud
	ON sh.shipper_id = o.shipper_id 
JOIN order_statuses os
	ON os.order_status_id = o.status
