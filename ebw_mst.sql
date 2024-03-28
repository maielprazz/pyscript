CREATE TABLE `mst_group_category_cm` (
  `PIC_CM` varchar(100) NOT NULL,
  `group_category` varchar(100) NOT NULL,
  KEY `ix_gcat1` (`PIC_CM`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

insert into mst_group_category_cm
select 'Nida',	'Pharma'
UNION select 'Irma','Beauty & GMS'
UNION select 'Winey','Health Nutrition & Support'
UNION select 'Putri','Health Vit. Supplement'

select * from mst_group_category_cm

CREATE TABLE `dim_store` (
  `store_id` int auto_increment,
  `store_name` varchar(200) NULL,
  `store_location` varchar(250) NULL,
  KEY `ix_dim_store` (`store_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

CREATE TABLE `mst_store_area_manager` (
  `store_id` INT not NULL,
  `area_manager` varchar(100) NOT NULL DEFAULT '',
  KEY `ix_mststmgr` (`store_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 ;


insert into mst_store_area_manager
select a.store_id, b.areamanager 
from dim_store a
join (
Select 'Apotek Wellings Cirendeu' storelocation,'Efraim' areamanager
UNION Select 'Apotek Wellings Bintaro' storelocation,'Efraim' areamanager
UNION Select 'Apotek Wellings Gandaria' storelocation,'Efraim' areamanager
UNION Select 'Apotek Wellings Citra Garden' storelocation,'Ristia' areamanager
UNION Select 'Apotek Wellings Erajaya Plaza' storelocation,'Other' areamanager
UNION Select 'Apotek Wellings Green Lake' storelocation,'Ristia' areamanager
UNION Select 'Apotek Wellings Greenville' storelocation,'Ristia' areamanager
UNION Select 'Apotek Wellings Harapan Indah' storelocation,'Efraim' areamanager
UNION Select 'Apotek Wellings Jatiwaringin' storelocation,'Other' areamanager
UNION Select 'Apotek Wellings Kelapa Gading' storelocation,'Ristia' areamanager
UNION Select 'Apotek Wellings PIK' storelocation,'Ristia' areamanager
UNION Select 'Apotek Wellings PIK 2' storelocation,'Ristia' areamanager
UNION Select 'Apotek Wellings Pondok Kelapa' storelocation,'Efraim' areamanager
UNION Select 'Apotek Wellings Rempoa' storelocation,'Efraim' areamanager
UNION Select 'Apotek Wellings Sunter' storelocation,'Ristia' areamanager
UNION Select 'Apotek Wellings Tebet' storelocation,'Efraim' areamanager
UNION Select 'Apotek Wellings Teluk Gong' storelocation,'Ristia' areamanager
UNION Select 'Apotek Wellings Veteran' storelocation,'Efraim' areamanager
UNION Select 'Apotek Wellings Greenville (Ecom)' storelocation,'Hakim' areamanager
UNION Select 'Marketplace Greenville' storelocation,'Hakim' areamanager
) b on a.store_name = b.storelocation

insert into dim_store (store_name, store_location)
Select 'Apotek Wellings Cirendeu','Cirendeu'
UNION Select 'Apotek Wellings Bintaro','Bintaro'
UNION Select 'Apotek Wellings Gandaria','Gandaria'
UNION Select 'Apotek Wellings Citra Garden','Citra Garden'
UNION Select 'Apotek Wellings Erajaya Plaza','Erajaya Plaza'
UNION Select 'Apotek Wellings Green Lake','Green Lake'
UNION Select 'Apotek Wellings Greenville','Greenville'
UNION Select 'Apotek Wellings Harapan Indah','Harapan Indah'
UNION Select 'Apotek Wellings Jatiwaringin','Jatiwaringin'
UNION Select 'Apotek Wellings Kelapa Gading','Kelapa Gading'
UNION Select 'Apotek Wellings PIK','PIK'
UNION Select 'Apotek Wellings PIK 2','PIK 2'
UNION Select 'Apotek Wellings Pondok Kelapa','Pondok Kelapa'
UNION Select 'Apotek Wellings Rempoa','Rempoa'
UNION Select 'Apotek Wellings Sunter','Sunter'
UNION Select 'Apotek Wellings Tebet','Tebet'
UNION Select 'Apotek Wellings Teluk Gong','Teluk Gong'
UNION Select 'Apotek Wellings Veteran','Veteran'
UNION Select 'Apotek Wellings Greenville (Ecom)','Greenville (Ecom)'
UNION Select 'Marketplace Greenville','Marketplace Greenville'

select  b.store_name, a.area_manager
from mst_store_area_manager a
join dim_store b on a.store_id = b.store_id


CREATE DEFINER=`ismail`@`%` PROCEDURE `db_ip`.`ip_getSalesForPPT`(
	IN  sdate date,
	IN 	edate date
)
BEGIN
	
SELECT 	a.STORE_LOCATION, 
		IFNULL(am.AM, 'Other') AreaManager, 
		a.SALES_NUMBER,
		a.SALES_DATE, 
		a.SALES_TIME, 
		a.CLIENT_ID, 
		a.CLIENT_NAME,
		ttc.PHONE, 
		ttc.GENDER, 
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
		IFNULL(gc1.GCAT1, 'Unknown') GroupCategory1,
		a.ITEM_DIVISION, 
		a.ITEM_GROUP, 
		a.ITEM_MODEL, 
		-- a.NON_FUNCTIONAL_25, 
		-- a.CHANNEL_TYPE, 
		-- a.PROMOTION_TYPE,
		-- a.APPLIED_PROMO, 
		SUM(a.PARENT_QTY) PARENT_QTY, 
		SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
		-- SUM(a.REBATE) REBATE, SUM(a.TOTAL_GROSS_MARGIN_ACCURATE) TOTAL_GROSS_MARGIN_ACCURATE
FROM tr_tbl_sales_by_item a
LEFT JOIN tr_tbl_customer ttc ON ttc.CODE = a.CLIENT_ID 
LEFT JOIN ip_store_am am ON am.StoreLocation = a.STORE_LOCATION 
LEFT JOIN mst_group_category_cm gc1 ON gc1.PIC_CM = a.PIC_CM
WHERE a.SALES_DATE BETWEEN sdate AND edate
AND a.Store_location <> 'Apotek Wellings Erajaya Plaza'
AND a.SALES_STATUS not like 'CANCEL%'
GROUP BY a.STORE_LOCATION,
		IFNULL(am.AM, 'Other'),
		a.SALES_NUMBER,
		a.SALES_DATE, 
		a.SALES_TIME, 
		a.CLIENT_ID, 
		a.CLIENT_NAME,
		ttc.PHONE, 
		ttc.GENDER, 
		CASE WHEN a.MEMBER_TYPE = 'member'
				 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
				 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
				 AND a.CLIENT_STORE_REGISTRATION = a.STORE_LOCATION 
			 THEN 'New Member'
			 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
			 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
		ELSE 'Existing' END,
		a.ITEM_CODE, 
		a.ITEM_NAME, 
		IFNULL(gc1.GCAT1, 'Unknown'),
		a.ITEM_DIVISION , 
		a.ITEM_GROUP, 
		a.ITEM_MODEL;
		
END;

select * from stg_tr_tbl_sales_by_item limit 100

-- == table dim_product
item_id
item_code
item_name
description
fullname
[type]
model
category_id
department_id
division_id
[group_id]
brand_id
promotion_type_id
principal_id
sales_person_id
channel_type_id
member_type_id
customer_id
sales_status_code
active_sku_code
client_id
pic_cm
non_functional_25_code

select COUNT(*) from ip_tr_tbl_sales_by_item
select * from dim_store

SELECT 	[Year], [Month], a.[Store Location], a.[CHANNEL_TYPE], [Group Category 1], ITEM_DIVISION,
		ITEM_GROUP,	BRAND,	NON_FUNCTIONAL_25,	Member Classification,	Promo_Type,	
		PROMOTION_TYPE,	APPLIED_PROMO,	[Client Id], [Sales Number], 
		PARENT_QTY, [Sales Value], REBATE
FROM stg_tr_tbl_sales_by_item a

select Year(now()), MONTH(now())

SELECT 	year(a.sales_date) `YEAR`, 
		MONTH(a.Sales_date) `MONTH`, 
		ds.Store_name `Store Location`,
		a.CHANNEL_TYPE,
		IFNULL(mcm.group_category, 'Unknown') `Group Category 1`, 
		ITEM_DIVISION, 
		ITEM_GROUP,
		BRAND, 
		NON_FUNCTIONAL_25, 
		CASE WHEN a.MEMBER_TYPE = 'member'
				 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
				 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
				 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
			 THEN 'New Member'
			 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
			 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
		ELSE 'Existing' end `Member Classification`, 
		a.PROMOTION_TYPE, 
		a.APPLIED_PROMO,
		a.CLIENT_ID, 
		a.SALES_NUMBER, 
		SUM(a.PARENT_QTY) PARENT_QTY, 
		SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
from ip_tr_tbl_sales_by_item a
join dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
join dim_store ds on ds.store_id  = a.store_id
left join mst_group_category_cm mcm on mcm.PIC_CM = a.PIC_CM
-- left join mst_store_area_manager
where year(a.sales_date) = 2024 
group by year(a.sales_date), MONTH(a.Sales_date), ds.Store_name,
		a.CHANNEL_TYPE,
		IFNULL(mcm.group_category, 'Unknown'), 
		ITEM_DIVISION, ITEM_GROUP,
		BRAND, NON_FUNCTIONAL_25, 
		CASE WHEN a.MEMBER_TYPE = 'member'
				 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
				 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
				 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
			 THEN 'New Member'
			 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
			 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
		ELSE 'Existing' end, 
		a.PROMOTION_TYPE, a.APPLIED_PROMO,
		a.CLIENT_ID, a.SALES_NUMBER


select  b.store_name, a.area_manager
select * from mst_store_area_manager a
select * from mst_group_category_cm
join dim_store b on a.store_id = b.store_id




