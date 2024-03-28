USE xilnex_rpt;
--- MS_TARGET
SELECT * 
FROM MS_TARGET where tanggal = '20240201' 
ORDER BY Target DESC;

SELECT DISTINCT DAY(tanggal), MONTH(tanggal), YEAR(tanggal) 
FROM MS_TARGET where tanggal = '20240102' limit 100;

--- MS_TARGET_DAILY
SELECT * 
FROM MS_TARGET_DAILY limit 100;

--- MS_PROMO
SELECT * 
FROM ms_promo limit 100;

--- tr_tbl_customer 
SELECT * 
FROM tr_tbl_customer limit 100;

SELECT * FROM tr_tbl_purchase_order limit 100;

SELECT * FROM tr_tbl_fast_moving limit 100;

SELECT * FROM tr_tbl_sales limit 100;

SELECT * FROM tr_tbl_stock limit 100;

SELECT * FROM tr_tbl_sales_by_item limit 100;

SELECT SUM(total_price_after_tax) vv 
FROM tr_tbl_sales_by_item 
WHERE SALES_DATE BETWEEN '20240201' AND '20240202' 
GROUP BY sales_date;

SELECT DISTINCT TOTAL_NET_MARGIN_ACCURATE 
FROM tr_tbl_sales_by_item 
WHERE SALES_DATE BETWEEN '20240201' AND '20240202' 
AND IFNULL(TOTAL_NET_MARGIN_ACCURATE,0) > 0;
SELECT @@VERSION
SELECT * 
FROM tr_tbl_sales_stock_item 
WHERE ITEM_NAME like '%Yrins%'
AND BRANCH like '%rempoa%'

limit 100;

SELECT * FROM tr_tbl_sales_summary limit 100;

SELECT * FROM tr_tbl_stock_history limit 100;

SELECT * FROM tr_tbl_stock_konsol_eci limit 100;

SELECT * FROM tr_tbl_stock_recon limit 100;

SELECT * FROM tr_tbl_transfer_note limit 100;


