SELECT * FROM tr_tbl_sales_by_item limit 100;

;WITH AM AS (
SELECT 'Apotek Wellings Cirendeu' StoreLocation , 'Efraim'  AM
UNION select 'Apotek Wellings Bintaro' , 'Efraim'
UNION SELECT 'Apotek Wellings Gandaria' , 'Efraim'
UNION SELECT 'Apotek Wellings Citra Garden' , 'Ristia'
UNION SELECT 'Apotek Wellings Erajaya Plaza' , 'Other'
UNION SELECT 'Apotek Wellings Green Lake' , 'Ristia'
UNION SELECT 'Apotek Wellings Greenville' , 'Ristia'
UNION SELECT 'Apotek Wellings Harapan Indah' , 'Efraim'
UNION SELECT 'Apotek Wellings Jatiwaringin' , 'Other'
UNION SELECT 'Apotek Wellings Kelapa Gading' , 'Ristia'
UNION SELECT 'Apotek Wellings PIK' , 'Ristia'
UNION SELECT 'Apotek Wellings PIK 2' , 'Ristia'
UNION SELECT 'Apotek Wellings Pondok Kelapa' , 'Efraim'
UNION SELECT 'Apotek Wellings Rempoa' , 'Efraim'
UNION SELECT 'Apotek Wellings Sunter' , 'Ristia'
UNION SELECT 'Apotek Wellings Tebet' , 'Efraim'
UNION SELECT 'Apotek Wellings Teluk Gong' , 'Ristia'
UNION SELECT 'Apotek Wellings Veteran' , 'Efraim'
UNION SELECT 'Apotek Wellings Greenville (Ecom)' , 'Hakim'
UNION SELECT 'Marketplace Greenville' , 'Hakim'
)

SELECT * FROM AM

Group category
IF [PIC_CM] = "Nida" THEN "Pharma"
ELSEIF [PIC_CM] = "Irma" THEN "Beauty & GMS"
ELSEIF [PIC_CM] = "Winey" THEN "Health Nutrition & Support"
ELSEIF [PIC_CM] = "Putri" THEN "Health Vit. Supplement"
ELSE "Unknown"
END

SELECT 'Nida' PIC_CM, 'Pharma' GCAT1
UNION SELECT 'Irma' PIC_CM, 'Beauty & GMS' GCAT1
UNION SELECT 'Winey' PIC_CM, 'Health Nutrition & Support' GCAT1
UNION SELECT 'Putri' PIC_CM, 'Health Vit. Supplement' GCAT1



IF [Member Type_] = "Member"
AND DATETRUNC('month',[Client Creation Date]) = DATETRUNC('month',[Sales Date])
AND DATETRUNC('year',[Client Creation Date]) = DATETRUNC('year',[Sales Date])
AND [Client Store Registration] = [Store Location]
THEN "New Member"
ELSEIF [Member Type_] = "Other" THEN "Other"
ELSEIF [Member Type_] = "Non Member" THEN "Non Member"
ELSE "Existing"
END


-- ==== QUERY
select sal.* 
FROM (
SELECT 	a.STORE_LOCATION, IFNULL(am.AM,'Other') AreaManager, a.SALES_NUMBER,
		a.SALES_DATE, a.SALES_TIME, a.CLIENT_ID, a.CLIENT_NAME,
		ttc.PHONE, ttc.GENDER, 
		CASE WHEN a.MEMBER_TYPE = 'member'
				 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
				 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
				 AND a.CLIENT_STORE_REGISTRATION = a.STORE_LOCATION 
			 THEN 'New Member'
			 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
			 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
		ELSE 'Existing'
		END AS MemberClassification,
		a.ITEM_CODE, a.ITEM_NAME, IFNULL(gcat.GCAT1, 'Unknown') GroupCategory1,
		a.ITEM_DIVISION , a.ITEM_GROUP, a.ITEM_MODEL, a.NON_FUNCTIONAL_25, a.CHANNEL_TYPE, a.PROMOTION_TYPE, 
		SUM(a.PARENT_QTY) PARENT_QTY, SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
FROM tr_tbl_sales_by_item a
LEFT JOIN tr_tbl_customer ttc ON ttc.CODE = a.CLIENT_ID 
LEFT JOIN (
	SELECT 'Apotek Wellings Cirendeu' StoreLocation , 'Efraim'  AM
	UNION select 'Apotek Wellings Bintaro' , 'Efraim'
	UNION SELECT 'Apotek Wellings Gandaria' , 'Efraim'
	UNION SELECT 'Apotek Wellings Citra Garden' , 'Ristia'
	UNION SELECT 'Apotek Wellings Erajaya Plaza' , 'Other'
	UNION SELECT 'Apotek Wellings Green Lake' , 'Ristia'
	UNION SELECT 'Apotek Wellings Greenville' , 'Ristia'
	UNION SELECT 'Apotek Wellings Harapan Indah' , 'Efraim'
	UNION SELECT 'Apotek Wellings Jatiwaringin' , 'Other'
	UNION SELECT 'Apotek Wellings Kelapa Gading' , 'Ristia'
	UNION SELECT 'Apotek Wellings PIK' , 'Ristia'
	UNION SELECT 'Apotek Wellings PIK 2' , 'Ristia'
	UNION SELECT 'Apotek Wellings Pondok Kelapa' , 'Efraim'
	UNION SELECT 'Apotek Wellings Rempoa' , 'Efraim'
	UNION SELECT 'Apotek Wellings Sunter' , 'Ristia'
	UNION SELECT 'Apotek Wellings Tebet' , 'Efraim'
	UNION SELECT 'Apotek Wellings Teluk Gong' , 'Ristia'
	UNION SELECT 'Apotek Wellings Veteran' , 'Efraim'
	UNION SELECT 'Apotek Wellings Greenville (Ecom)' , 'Hakim'
	UNION SELECT 'Marketplace Greenville' , 'Hakim'
	) am on am.StoreLocation = a.STORE_LOCATION 
LEFT JOIN (
	SELECT 'Nida' PIC_CM, 'Pharma' GCAT1
	UNION SELECT 'Irma' PIC_CM, 'Beauty & GMS' GCAT1
	UNION SELECT 'Winey' PIC_CM, 'Health Nutrition & Support' GCAT1
	UNION SELECT 'Putri' PIC_CM, 'Health Vit. Supplement' GCAT1
	) gcat ON a.PIC_CM = gcat.PIC_CM
WHERE a.SALES_DATE BETWEEN '20231201' AND '20231231'
GROUP BY a.STORE_LOCATION, IFNULL(am.AM,'Other'), a.SALES_NUMBER,
		a.SALES_DATE, a.SALES_TIME, a.CLIENT_ID, a.CLIENT_NAME,
		ttc.PHONE, ttc.GENDER, 
		CASE WHEN a.MEMBER_TYPE = 'member'
				 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
				 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
				 AND a.CLIENT_STORE_REGISTRATION = a.STORE_LOCATION 
			 THEN 'New Member'
			 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
			 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
		ELSE 'Existing'
		END,
		a.ITEM_CODE, a.ITEM_NAME, IFNULL(gcat.GCAT1, 'Unknown'),
		a.ITEM_DIVISION , a.ITEM_GROUP, a.ITEM_MODEL, a.NON_FUNCTIONAL_25, a.CHANNEL_TYPE, a.PROMOTION_TYPE
) sal
LEFT JOIN tr_tbl_stock tts 
	ON sal.ITEM_CODE = tts.ITEM_CODE  
	AND sal.STORE_LOCATION = tts.STORE


	
	
	

select * from tr_tbl_sales_stock_item ttssi LIMIT 100;
select * from tr_tbl_stock tts LIMIT 100;
select * from tr_tbl_sales_by_item ttsbi limit 100;



select sal.*, tts.AVAILABLE_STOCK, tts.VALUE_AVAILABLE_STOCK
FROM (
SELECT 	a.STORE_LOCATION, 
		CASE WHEN Store_location = 'Apotek Wellings Cirendeu'  THEN 'Efraim' 
		WHEN Store_location = 'Apotek Wellings Bintaro' THEN 'Efraim'
		WHEN Store_location = 'Apotek Wellings Gandaria' THEN 'Efraim'
		WHEN Store_location = 'Apotek Wellings Citra Garden' THEN 'Ristia'
		WHEN Store_location = 'Apotek Wellings Erajaya Plaza' THEN 'Other'
		WHEN Store_location = 'Apotek Wellings Green Lake' THEN 'Ristia'
		WHEN Store_location = 'Apotek Wellings Greenville' THEN 'Ristia'
		WHEN Store_location = 'Apotek Wellings Harapan Indah' THEN 'Efraim'
		WHEN Store_location = 'Apotek Wellings Jatiwaringin' THEN 'Other'
		WHEN Store_location = 'Apotek Wellings Kelapa Gading' THEN 'Ristia'
		WHEN Store_location = 'Apotek Wellings PIK' THEN 'Ristia'
		WHEN Store_location = 'Apotek Wellings PIK 2' THEN 'Ristia'
		WHEN Store_location = 'Apotek Wellings Pondok Kelapa' THEN 'Efraim'
		WHEN Store_location = 'Apotek Wellings Rempoa' THEN 'Efraim'
		WHEN Store_location = 'Apotek Wellings Sunter' THEN 'Ristia'
		WHEN Store_location = 'Apotek Wellings Tebet' THEN 'Efraim'
		WHEN Store_location = 'Apotek Wellings Teluk Gong' THEN 'Ristia'
		WHEN Store_location = 'Apotek Wellings Veteran' THEN 'Efraim'
		WHEN Store_location = 'Apotek Wellings Greenville (Ecom)' THEN 'Hakim'
		WHEN Store_location = 'Marketplace Greenville' THEN 'Hakim'
		ELSE 'Other' END AreaManager, 
			a.SALES_NUMBER,
		a.SALES_DATE, a.SALES_TIME, a.CLIENT_ID, a.CLIENT_NAME,
		ttc.PHONE, ttc.GENDER, 
		CASE WHEN a.MEMBER_TYPE = 'member'
				 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
				 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
				 AND a.CLIENT_STORE_REGISTRATION = a.STORE_LOCATION 
			 THEN 'New Member'
			 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
			 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
		ELSE 'Existing'
		END AS MemberClassification,
		a.ITEM_CODE, a.ITEM_NAME, 
		CASE WHEN a.PIC_CM = 'Nida' THEN 'Pharma'
		WHEN a.PIC_CM = 'Irma'  THEN 'Beauty & GMS' 
		WHEN a.PIC_CM = 'Winey'  THEN 'Health Nutrition & Support' 
		WHEN a.PIC_CM = 'Putri'  THEN 'Health Vit. Supplement' 
		ELSE 'Unknown' END GroupCategory1,
		a.ITEM_DIVISION , a.ITEM_GROUP, a.ITEM_MODEL, a.NON_FUNCTIONAL_25, a.CHANNEL_TYPE, a.PROMOTION_TYPE, 
		SUM(a.PARENT_QTY) PARENT_QTY, SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
FROM tr_tbl_sales_by_item a
 LEFT JOIN tr_tbl_customer ttc ON ttc.CODE = a.CLIENT_ID 
WHERE a.SALES_DATE BETWEEN '20231201' AND '20231231'
AND a.Store_location <> 'Apotek Wellings Erajaya Plaza'
GROUP BY a.STORE_LOCATION,
		CASE WHEN Store_location = 'Apotek Wellings Cirendeu'  THEN 'Efraim' 
		WHEN Store_location = 'Apotek Wellings Bintaro' THEN 'Efraim'
		WHEN Store_location = 'Apotek Wellings Gandaria' THEN 'Efraim'
		WHEN Store_location = 'Apotek Wellings Citra Garden' THEN 'Ristia'
		WHEN Store_location = 'Apotek Wellings Erajaya Plaza' THEN 'Other'
		WHEN Store_location = 'Apotek Wellings Green Lake' THEN 'Ristia'
		WHEN Store_location = 'Apotek Wellings Greenville' THEN 'Ristia'
		WHEN Store_location = 'Apotek Wellings Harapan Indah' THEN 'Efraim'
		WHEN Store_location = 'Apotek Wellings Jatiwaringin' THEN 'Other'
		WHEN Store_location = 'Apotek Wellings Kelapa Gading' THEN 'Ristia'
		WHEN Store_location = 'Apotek Wellings PIK' THEN 'Ristia'
		WHEN Store_location = 'Apotek Wellings PIK 2' THEN 'Ristia'
		WHEN Store_location = 'Apotek Wellings Pondok Kelapa' THEN 'Efraim'
		WHEN Store_location = 'Apotek Wellings Rempoa' THEN 'Efraim'
		WHEN Store_location = 'Apotek Wellings Sunter' THEN 'Ristia'
		WHEN Store_location = 'Apotek Wellings Tebet' THEN 'Efraim'
		WHEN Store_location = 'Apotek Wellings Teluk Gong' THEN 'Ristia'
		WHEN Store_location = 'Apotek Wellings Veteran' THEN 'Efraim'
		WHEN Store_location = 'Apotek Wellings Greenville (Ecom)' THEN 'Hakim'
		WHEN Store_location = 'Marketplace Greenville' THEN 'Hakim'
		ELSE 'Other' END ,
		a.SALES_NUMBER,
		a.SALES_DATE, a.SALES_TIME, a.CLIENT_ID, a.CLIENT_NAME,
		ttc.PHONE, ttc.GENDER, 
		CASE WHEN a.MEMBER_TYPE = 'member'
				 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
				 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
				 AND a.CLIENT_STORE_REGISTRATION = a.STORE_LOCATION 
			 THEN 'New Member'
			 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
			 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
		ELSE 'Existing' END,
		a.ITEM_CODE, a.ITEM_NAME, 
		CASE WHEN a.PIC_CM = 'Nida' THEN 'Pharma'
		WHEN a.PIC_CM = 'Irma'  THEN 'Beauty & GMS' 
		WHEN a.PIC_CM = 'Winey'  THEN 'Health Nutrition & Support' 
		WHEN a.PIC_CM = 'Putri'  THEN 'Health Vit. Supplement' 
		ELSE 'Unknown' END,
		a.ITEM_DIVISION , a.ITEM_GROUP, a.ITEM_MODEL, a.NON_FUNCTIONAL_25, 
		a.CHANNEL_TYPE, a.PROMOTION_TYPE
) sal
LEFT JOIN tr_tbl_stock tts 
	ON sal.ITEM_CODE = tts.ITEM_CODE  
	AND sal.STORE_LOCATION = tts.STORE




-- Query lagi
SELECT 	a.STORE_LOCATION, 
		IFNULL(am.AM,'Other') AreaManager, 
		a.SALES_NUMBER,
		a.SALES_DATE, a.SALES_TIME, a.CLIENT_ID, a.CLIENT_NAME,
		ttc.PHONE, ttc.GENDER, 
		CASE WHEN a.MEMBER_TYPE = 'member'
				 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
				 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
				 AND a.CLIENT_STORE_REGISTRATION = a.STORE_LOCATION 
			 THEN 'New Member'
			 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
			 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
		ELSE 'Existing'
		END AS MemberClassification,
		a.ITEM_CODE, a.ITEM_NAME, 
		IFNULL(gcat.GCAT1, 'Unknown') GroupCategory1,
		a.ITEM_DIVISION , a.ITEM_GROUP, a.ITEM_MODEL, a.NON_FUNCTIONAL_25, a.CHANNEL_TYPE, a.PROMOTION_TYPE, 
		SUM(a.PARENT_QTY) PARENT_QTY, SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
FROM tr_tbl_sales_by_item a
 LEFT JOIN tr_tbl_customer ttc ON ttc.CODE = a.CLIENT_ID 
LEFT JOIN (
	SELECT 'Apotek Wellings Cirendeu' StoreLocation , 'Efraim'  AM
	UNION select 'Apotek Wellings Bintaro' , 'Efraim'
	UNION SELECT 'Apotek Wellings Gandaria' , 'Efraim'
	UNION SELECT 'Apotek Wellings Citra Garden' , 'Ristia'
	UNION SELECT 'Apotek Wellings Erajaya Plaza' , 'Other'
	UNION SELECT 'Apotek Wellings Green Lake' , 'Ristia'
	UNION SELECT 'Apotek Wellings Greenville' , 'Ristia'
	UNION SELECT 'Apotek Wellings Harapan Indah' , 'Efraim'
	UNION SELECT 'Apotek Wellings Jatiwaringin' , 'Other'
	UNION SELECT 'Apotek Wellings Kelapa Gading' , 'Ristia'
	UNION SELECT 'Apotek Wellings PIK' , 'Ristia'
	UNION SELECT 'Apotek Wellings PIK 2' , 'Ristia'
	UNION SELECT 'Apotek Wellings Pondok Kelapa' , 'Efraim'
	UNION SELECT 'Apotek Wellings Rempoa' , 'Efraim'
	UNION SELECT 'Apotek Wellings Sunter' , 'Ristia'
	UNION SELECT 'Apotek Wellings Tebet' , 'Efraim'
	UNION SELECT 'Apotek Wellings Teluk Gong' , 'Ristia'
	UNION SELECT 'Apotek Wellings Veteran' , 'Efraim'
	UNION SELECT 'Apotek Wellings Greenville (Ecom)' , 'Hakim'
	UNION SELECT 'Marketplace Greenville' , 'Hakim'
	) am on am.StoreLocation COLLATE utf8mb4_general_ci = a.STORE_LOCATION 
LEFT JOIN (
	SELECT 'Nida' PIC_CM, 'Pharma' GCAT1
	UNION SELECT 'Irma' PIC_CM , 'Beauty & GMS' GCAT1
	UNION SELECT 'Winey' PIC_CM , 'Health Nutrition & Support' GCAT1
	UNION SELECT 'Putri' PIC_CM , 'Health Vit. Supplement' GCAT1
	) gcat ON a.PIC_CM = gcat.PIC_CM COLLATE utf8mb4_general_ci
WHERE a.SALES_DATE BETWEEN '20231201' AND '20231231'
GROUP BY a.STORE_LOCATION, IFNULL(am.AM,'Other'), a.SALES_NUMBER,
		a.SALES_DATE, a.SALES_TIME, a.CLIENT_ID, a.CLIENT_NAME,
		ttc.PHONE, ttc.GENDER, 
		CASE WHEN a.MEMBER_TYPE = 'member'
				 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
				 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
				 AND a.CLIENT_STORE_REGISTRATION = a.STORE_LOCATION 
			 THEN 'New Member'
			 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
			 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
		ELSE 'Existing' END,
		a.ITEM_CODE, a.ITEM_NAME, IFNULL(gcat.GCAT1, 'Unknown'),
		a.ITEM_DIVISION , a.ITEM_GROUP, a.ITEM_MODEL, a.NON_FUNCTIONAL_25, 
		a.CHANNEL_TYPE, a.PROMOTION_TYPE





