# ridade juurdeliitmine kasutades UNION funktsiooni
SELECT 
	order_id,
	order_date,
'active' AS 'status' #saab juurde liita veeru vastava väärtusega, kirjutades selle 
#SELECT funktsiooni alla jutumärkidesse
FROM orders
# iga tellimuse juurde panna nimetus
# 2019 aasta tellimus aktiivne, varasem arhiivis
WHERE order_date >= '2019-01-01'
UNION #hetkel kaks erinevat tingimust, mis rakendatakse korraga UNION funktsiooni abil
SELECT 
	order_id,
    order_date,
    'archived' AS 'status'
FROM orders
WHERE order_date < '2019-01-01'
