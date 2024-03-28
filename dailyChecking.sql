USE xilnex_rpt
SELECT 	BRANCH, VENDOR_NAME, PRINCIPAL_CODE, ITEM_CODE, ITEM_NAME, CATEGORY, SUB_CATEGORY, 
		`Consignment/Outright`, PARETO_TYPE, `%`, All_Sales, L1y, Ytd, L3mo, L2mo, LM, MtD, 
		avg, OH_QTY, Incoming_PO, INCOMING_STOCK, MIN, MAX, ALLOW_TO_PURCHASE, `AVG/DAY`, 
		DOS, DOS_Outstanding, ALLOCATE_QTY, QTY_AFTER_ALLOC, DOS_AFTER_ALLOC, BELLOW_MIN, 
		MIN_ALLOCATE_QTY, QTY_AFTER_MIN_ALLOC, MIN_DOS_AFTER_ALLOC, BELOW_MAX, FINAL_ALLOCATE_QTY, 
		QTY_AFTER_FINAL_ALLOC, FINAL_DOS_AFTER_QTY, TIMESTAMP, Price_Mtd, Value_On_Hand_Stock, 
		Item_Department, ITEM_DIVISION, GROUP_NAME
FROM tr_tbl_sales_stock_item 
limit 100;

-- SELECT STOCK_DATE, STORE, ITEM_CODE_OLD, ITEM_CODE, ITEM_NAME, CATEGORY, ITEM_DIVISION, ITEM_GROUP, PIC_CM, BRAND, SUBCATEGORY, `TYPE`, QTY_OH, STOCK_VALUE, NON_FUNCTIONAL_25

SELECT MIN(STOCK_DATE) maxstockDate 


select SUM(STOCK_VALUE) stockvalue 
FROM tr_tbl_stock_history ttsh 
WHERE STOCK_DATE = '20240218';

SELECT SUM(VALUE_ON_HAND_STOCK) stockvalueonhand
FROM tr_tbl_stock ;


AND NON_FUNCTIONAL_25 <> 'INACTIVE'


SELECT * FROM tr_tbl_stock_history   limit 100

SELECT SUM(VALUE_ON_HAND_STOCK)
FROM tr_tbl_stock 

SELECT * FROM tr_tbl_stock limit 100
SELECT * FROM tr_tbl_stock_history limit 100



SELECT DISTINCT `TIMESTAMP`  
FROM xilnex_rpt.tr_tbl_sales_stock_item


SELECT * 
FROM  tr_tbl_stock tts 
-- where ITEM_CODE = '1010240061'
limit 100;


SELECT * 
FROM tr_tbl_customer ttc 
LIMIT 100;


SELECT * 
FROM  tr_tbl_stock_history ttsh  
where ITEM_CODE like '%1010240003%'
limit 100;

SELECT * 
FROM  tr_tbl_stock ttsh  
where ITEM_CODE like '%1010240003%'
limit 100;


select * 
from tr_tbl_sales
limit 100;


SELECT * 
FROM tr_tbl_stock_history ttsh 
WHERE ITEM_CODE = '8100074868'
AND STORE = 'Apotek Wellings Bintaro'
LIMIT 100;


-- === check last sales date
SELECT MAX(CASE WHEN DATEDIFF(SALES_DATE, CURDATE()) = 0 THEN DATE(DATE_ADD(SALES_DATE, INTERVAL -1 DAY)) ELSE SALES_DATE END) MAXSALESDATE
FROM tr_tbl_sales_by_item 

SELECT COUNT(*) FROM tr_tbl_sales_by_item 

SELECT SUM(IFNULL(TOTAL_PRICE_BEFORE_TAX,0)) ss  
FROM tr_tbl_sales_by_item 
WHERE SALES_DATE between '20240301' and '20240303'

SELECT SUM(IFNULL(REBATE,0) + IFNULL(TOTAL_NET_MARGIN_ACCURATE ,0))
FROM tr_tbl_sales_by_item 
WHERE SALES_DATE between '20240301' and '20240303'


SELECT 72000.0000/507409974.4400

507481974.4400
select 507409974.4400/507481974.4400

IFNULL([REBATE],0)+IFNULL([Total Net Margin Accurate],0)





SELECT SUM(TOTAL_PRICE_AFTER_TAX) Sales
FROM tr_tbl_sales_by_item
WHERE SALES_DATE = '20240205'

SELECT STORE_LOCATION, SUM(TOTAL_PRICE_AFTER_TAX) Sales
FROM tr_tbl_sales_by_item
WHERE SALES_DATE = '20240205'
AND SALES_STATUS NOT LIKE 'CANCEL%'
GROUP BY STORE_LOCATION 

SELECT SALES_DATE, STORE_LOCATION, TOTAL_PRICE_AFTER_TAX 
FROM tr_tbl_sales_by_item ttsbi  
where STORE_LOCATION = 'Apotek Wellings Citra Garden' 
and SALES_DATE  = '20240205'
ORDER BY 3


SELECT *
FROM tr_tbl_sales_by_item ttsbi  
where STORE_LOCATION = 'Apotek Wellings Citra Garden' 
and SALES_DATE  = '20240205'
AND TOTAL_PRICE_AFTER_TAX  = 666000

-- === Check Margin update
SELECT MAX(SALES_DATE) MAXSALESDATE
FROM tr_tbl_sales_by_item 
WHERE TOTAL_NET_MARGIN_ACCURATE IS NOT NULL



SELECT ITEM_DIVISION, CATEGORY_CM, PIC_CM
FROM xilnex.ms_category_pic_cm
order by ;


SELECT StoreLocation, AM
FROM xilnex_rpt.ip_store_am
order by AM
;



