CREATE DEFINER=`xilnex_user`@`%` PROCEDURE `xilnex_rpt`.`sp_tbl_sales_by_item`()
BEGIN
SET @tgl = DATE(NOW() - INTERVAL 1 MONTH);
drop temporary table if exists tbl_si;
drop temporary table if exists tbl_source;
drop temporary table if exists tbl_xsource;
drop temporary table if exists tbl_calc;

SELEC
FROM xilnex_rpt.tr_tbl_sales_by_item;
/*
create TEMPORARY TABLE tbl_sales_package
select si.ID,si.ITEM_CODE as ID_ITEM_CODE,ap.DOUBLE_QUANTITY from xilnex.APP_4_SALESITEM si 
join xilnex.APP_4_PACKAGE ap on si.ITEM_CODE = ap.PACKAGE_ITEM_ID and
				ap.ID = 
						(select ID from xilnex.APP_4_PACKAGE ax 
						where ax.PACKAGE_ITEM_ID = si.ITEM_CODE 
                        order by ax.ID DESC limit 1) 
where si.WBDELETED = 0 and STR_TO_DATE(si.SALES_DATE, '%d/%m/%Y') >= @tgl;
ALTER TABLE tbl_sales_package 
ADD INDEX `t_txp` (`ID` ASC);

        
CREATE TEMPORARY TABLE IF NOT EXISTS tbl_calc
select SALES_DATE,ITEM_CODE,QTY,
SUM(IF(SALES_DATE BETWEEN MOVING_DATE and SALES_DATE ,QTY,0)) as TOTAL_QTY_3_M, 
ROUND((QTY / SUM(IF(SALES_DATE BETWEEN MOVING_DATE and SALES_DATE ,QTY,0))) * 100 , 0) as percentage,
DOUBLE_SUB_TOTAL,
SUM(IF(SALES_DATE BETWEEN MOVING_DATE and SALES_DATE ,DOUBLE_SUB_TOTAL,0)) as PARETO_TOTAL
from 
(
	select
	STR_TO_DATE(apsi.SALES_DATE, '%d/%m/%Y') as SALES_DATE,
	STR_TO_DATE(apsi.SALES_DATE, '%d/%m/%Y') - INTERVAL 3 MONTH as MOVING_DATE,
	mi.ITEM_CODE as ITEM_CODE,
	apsi.DOUBLE_UOM_Quantity * ABS(apsi.`INT_QUANTITY`) as QTY,
    DOUBLE_SUB_TOTAL
	from xilnex.APP_4_SALESITEM apsi
	left join xilnex.vw_item_code_mapping_sap mi on mi.ID = apsi.ITEM_CODE
	join xilnex.APP_4_SALES aps on aps.SALES_NO = apsi.SALES_NO
	-- left join xilnex.ms_item_sap mis on mis.ITEM_CODE_XILNEX = mi.ITEM_CODE
	where apsi.WBDELETED = 0 and aps.SALES_STATUS = 'COMPLETED'

)x 
GROUP BY ITEM_CODE,SALES_DATE
;
ALTER  TABLE tbl_calc 
ADD INDEX `t_c` (`SALES_DATE` ASC, `ITEM_CODE` ASC);
*/

CREATE TEMPORARY TABLE IF NOT EXISTS ip_tbl_source
select
	STR_TO_DATE(apsi.SALES_DATE, '%d/%m/%Y') as SALES_DATE,
	STR_TO_DATE(apsi.SALES_DATE, '%d/%m/%Y') - INTERVAL 3 MONTH as MOVING_DATE,
	mi.ITEM_CODE as ITEM_CODE,
	apsi.DOUBLE_UOM_Quantity * ABS(apsi.`INT_QUANTITY`) as QTY,
    DOUBLE_SUB_TOTAL
	from xilnex.APP_4_SALESITEM apsi
	left join xilnex.vw_item_code_mapping_sap mi on mi.ID = apsi.ITEM_CODE
	join xilnex.APP_4_SALES aps on aps.SALES_NO = apsi.SALES_NO
	-- left join xilnex.ms_item_sap mis on mis.ITEM_CODE_XILNEX = mi.ITEM_CODE
	where apsi.WBDELETED = 0 and aps.SALES_STATUS = 'COMPLETED';
    
ALTER  TABLE ip_tbl_source 
ADD INDEX `t_xxs` (`SALES_DATE` ASC, `ITEM_CODE` ASC);

    
CREATE TEMPORARY TABLE IF NOT EXISTS ip_tbl_xsource
select SALES_DATE,SUM(IF(SALES_DATE BETWEEN MOVING_DATE and SALES_DATE ,QTY,0)) as TOTAL_QTY_3_M,
SUM(IF(SALES_DATE BETWEEN MOVING_DATE and SALES_DATE ,DOUBLE_SUB_TOTAL,0)) as PARETO_TOTAL
from ip_tbl_source
group by SALES_DATE;

ALTER  TABLE ip_tbl_xsource 
ADD INDEX `t_xxsx` (`SALES_DATE` ASC);

CREATE TEMPORARY TABLE IF NOT EXISTS ip_tbl_calc
select a.SALES_DATE,a.ITEM_CODE,a.QTY,b.TOTAL_QTY_3_M,b.PARETO_TOTAL,
((a.QTY / b.TOTAL_QTY_3_M) * 100) as percentage 
from ip_tbl_source a
join ip_tbl_xsource b on a.SALES_DATE = b.SALES_DATE;

ALTER  TABLE ip_tbl_calc 
ADD INDEX `t_c` (`SALES_DATE` ASC, `ITEM_CODE` ASC);

CREATE TEMPORARY TABLE IF NOT EXISTS ip_tbl_si
select
apsi.ID as SALES_ITEM_NUMBER,
apsi.SALES_NO as SALES_NUMBER,
apsi.ITEM_CODE as ITEM_ID_OLD,
LEFT(ai.ITEM_CODE,10) as ITEM_CODE_OLD,
@ITEM_CODE := mis.ITEM_CODE as ITEM_CODE,
REPLACE(CONVERT(ai.ITEM_NAME USING UTF8),'\n','') as ITEM_NAME,
REPLACE(ai.CUSTOM_FIELD_VALUE_4,'\n','') as ITEM_TYPE,
REPLACE(apsi.MODEL,'\n','') as ITEM_MODEL,
REPLACE(ai.CATEGORY,'\n','') as ITEM_CATEGORY,
REPLACE(ai.DEPARTMENT_CODE,'\n','') as ITEM_DEPARTEMENT,
REPLACE(ai.ITEM_DIVISION,'\n','') as ITEM_DIVISION,
REPLACE(ai.GROUP_NAME,'\n','') as ITEM_GROUP,
REPLACE(ai.ITEM_DESCRIPTION,'\n','') as DESCRIPTION,
REPLACE(ai.ITEM_BRAND,'\n','') as BRAND,
ROUND(apsi.DOUBLE_PRICE,4) as UNIT_PRICE_BEFORE_TAX,
ROUND(apsi.DOUBLE_Enter_Price,4) as UNIT_PRICE_AFTER_TAX,
ROUND(apsi.INT_QUANTITY,4) as QTY,
ROUND((apsi.INT_QUANTITY * apsi.DOUBLE_UOM_Quantity),4) as PARENT_QTY,
ROUND(apsi.DOUBLE_SALE_PRICE,4) as SALE_PRICE,
ROUND(apsi.DOUBLE_SUB_TOTAL,4) as TOTAL_PRICE_AFTER_TAX,
ROUND((apsi.DOUBLE_SUB_TOTAL - apsi.DOUBLE_TOTAL_BILL_LEVEL_TAX_AMOUNT),4) as TOTAL_PRICE_BEFORE_TAX,  
ROUND(apsi.DOUBLE_MGST_Tax_Percentage,4) as TAX_PERCENTAGE,
ROUND(apsi.DOUBLE_TOTAL_BILL_LEVEL_TAX_AMOUNT,4) as TAX_AMOUNT,
ROUND(apsi.DOUBLE_DISCOUNT_PERCENTAGE,4) as TOTAL_DISCOUNT_PERCENTAGE,
ROUND(apsi.DOUBLE_TOTAL_DISCOUNT_AMOUNT,4) as TOTAL_DISCOUNT,
ROUND(apsi.item_Custom_6 * (apsi.DOUBLE_COST * apsi.INT_QUANTITY),4) as DISCOUNT_COST_ON_PO,
ROUND((apsi.DOUBLE_COST * apsi.INT_QUANTITY),4) as UNIT_COST,
ROUND((apsi.DOUBLE_COST * apsi.INT_QUANTITY) -  (apsi.item_Custom_6 * (apsi.DOUBLE_COST * apsi.INT_QUANTITY)),4) as FINAL_COST,
ROUND((apsi.DOUBLE_PROFIT - apsi.DOUBLE_COST),4) as UNIT_PROFIT, -- ROUND(apsi.DOUBLE_PROFIT,4)
ROUND((apsi.INT_QUANTITY * apsi.DOUBLE_Enter_Price) - (apsi.INT_QUANTITY * apsi.DOUBLE_COST),4) as TOTAL_PROFIT, -- (ROUND(apsi.DOUBLE_PROFIT * apsi.INT_QUANTITY,4)
REPLACE(apsi.CUSTOMER_NAME,'\n','') as CUSTOMER_NAME,
CASE 
	WHEN apsi.CUSTOMER_NAME in ('WALK IN','CASH') THEN 'NON MEMBER'
    WHEN apsi.CUSTOMER_NAME in ('SHOPEE','TOKOPEDIA','Good Doctor','Sehatq','Halodoc','dr. Deddy Gouw - TEST','Ecom Client Test') THEN 'OTHER'
	ELSE 'MEMBER'
END  as MEMBER_TYPE,
CASE
	WHEN apsi.SALES_PERSON like '%Drf%' THEN 'DOKTER'
	WHEN apy.STRING_EXTEND_2 = 'MARKETPLACE' THEN 'MARKETPLACE' 
	WHEN apy.STRING_EXTEND_2 IN ('HALODOC','GOODDOCTOR','SEHATQ') THEN 'TELEMEDICINE' 
	WHEN apy.STRING_EXTEND_2 Like 'DOKU%' THEN 'WHATSAPP'
    ELSE 'IN STORE'
END as 'CHANNEL_TYPE',
STR_TO_DATE(apsi.SALES_DATE, '%d/%m/%Y') as SALES_DATE,
REPLACE(aps.NOW_TIME,'\n','') as SALES_TIME,
REPLACE(mld.LOCATIONNAME,'\n','') as STORE_LOCATION,
REPLACE(apsi.SALESITEM_REMARK,'\n','') as SALES_ITEM_REMARK,
REPLACE(apsi.SALES_PERSON,'\n','') as SALES_PERSON,
REPLACE(aps.SALES_STATUS,'\n','') as SALES_STATUS,
REPLACE(apsi.ALT_LOOKUP,'\n','') as BARCODE,
REPLACE(ai.UNIT_OF_MEASURE,'\n','') UOM,
NULL as APPLIED_PROMO,
NULL  as 'Promotion_Type',
REPLACE(apsi.Voucher_No,'\n','') as VOUCHER_NUMBER,
REPLACE(ai.CUSTOM_FIELD_VALUE_7,'\n','') as PRINSIPAL,
REPLACE(ai.CUSTOM_FIELD_VALUE_5,'\n','') as FULL_NAME_PRODUCT,
REPLACE(apsi.ITEM_ADDITIONAL_INFO_1,'\n','') as ITEM_ADDITIONAL_INFO,
REPLACE(ai.CUSTOM_FIELD_VALUE_13,'\n','') as ACTIVE_SKU,
NULL as ITEM_SMI,
NULL as PARETO_TYPE,
mt.target,
COALESCE(apc.ID,'C.00004') as 'CLIENT_ID',
COALESCE(apc.COMPANY_NAME,'WALK IN') as 'CLIENT_NAME',
apc.GENDER as 'CLIENT_GENDER',
STR_TO_DATE(apc.DOB, '%d/%m/%Y') as 'CLIENT_DOB',
STR_TO_DATE(apc.CREATION_DATE, '%d/%m/%Y') as 'CLIENT_CREATION_DATE',
REPLACE(mldc.LOCATIONNAME,'\n','') as 'CLIENT_STORE_REGISTRATION', 
NULL as TOTAL_NET_MARGIN_ACCURATE,
NULL as TOTAL_GROSS_MARGIN_ACCURATE,
NULL as MCH,
NULL as FAST_MOVING,
NULL as CATEGORY_CM,
NULL as PIC_CM,
NULL as CAMPAIGN_NAME,
NULL as NON_FUNCTIONAL_25,
NULL as REBATE,
NULL as REBATE_ERROR,
NOW() as TIMESTAMP
from xilnex.APP_4_SALESITEM apsi
join xilnex.APP_4_SALES aps on aps.SALES_NO = apsi.SALES_NO and aps.WBDELETED = 0
join xilnex.APP_4_ITEM ai on ai.ID = apsi.ITEM_CODE
left join (select distinct INVOICE_ID,
				case when STRING_EXTEND_2 like 'DOKU%' THEN 'DOKU' 
                else STRING_EXTEND_2 end as STRING_EXTEND_2 
			from xilnex.APP_4_PAYMENT 
            where WBDELETED = 0 
            and (STRING_EXTEND_2 IN ('MARKETPLACE','HALODOC','GOODDOCTOR','SEHATQ') 
				OR STRING_EXTEND_2 Like 'DOKU%')
			) apy on apy.INVOICE_ID = aps.SALES_NO 
-- left join xilnex.ms_item_sap mis on mis.ITEM_CODE_XILNEX = LEFT(ai.ITEM_CODE,10)
left join xilnex.vw_item_code_mapping_sap mis on mis.ID = apsi.ITEM_CODE
-- left join tbl_sales_package ap on apsi.ID = ap.ID
JOIN xilnex.ms_location_detail mld on mld.ID = apsi.SALES_LOCATION and mld.LOCATION_INT_ID <> 13
left join (select tanggal,
				cabang,
				target 
			from xilnex_rpt.MS_TARGET 
            GROUP BY tanggal,cabang) mt 
		on left(mt.tanggal,7) = left(STR_TO_DATE(apsi.SALES_DATE, '%d/%m/%Y'),7) and mt.cabang = mld.LOCATIONNAME 
left join xilnex.APP_4_CUSTOMER apc on apc.ID = aps.CUSTOMER_ID
left join xilnex.ms_location_detail mldc on mldc.ID = apc.CREATED_LOCATION_ID
-- left join xilnex.tr_margin tm on apsi.SALES_NO = tm.INVOICE_NO 
--	and tm.ITEM_CODE = LEFT(ai.ITEM_CODE,10) and ROUND(apsi.INT_QUANTITY,0) = ROUND(tm.QTY,0) and ROUND(apsi.DOUBLE_SUB_TOTAL,0) = ROUND(tm.TOTAL_PRICE,0)
where apsi.WBDELETED = 0 and ai.ITEM_CODE NOT LIKE 'K1%'
and STR_TO_DATE(apsi.SALES_DATE, '%d/%m/%Y') >= @tgl
;
SELECT COUNT(*) From xilnex.APP_4_SALESITEM
SELECT MIN(SALES_DATE) From xilnex.APP_4_SALESITEM limit 100
SELECT COUNT(*) FROM xilnex.APP_4_SALES
SELECT COUNT(*) FROM xilnex.APP_4_ITEM


from xilnex.APP_4_SALESITEM apsi
join xilnex.APP_4_SALES aps on aps.SALES_NO = apsi.SALES_NO and aps.WBDELETED = 0
join xilnex.APP_4_ITEM ai on ai.ID = apsi.ITEM_CODE

-- create table if not exists xilnex_rpt.tr_tbl_sales_by_item like tbl_si;
-- TRUNCATE `xilnex_rpt`.`tr_tbl_sales_by_item`;
DELETE FROM `xilnex_rpt`.`tr_tbl_sales_by_item` where SALES_DATE >= @tgl;
insert xilnex_rpt.tr_tbl_sales_by_item select * from tbl_si;
-- insert xilnex_rpt.tr_tbl_sales_by_item select * from tbl_sidel;

UPDATE xilnex_rpt.tr_tbl_sales_by_item a
left join xilnex.ms_item_sap b on b.ITEM_CODE_XILNEX = a.ITEM_CODE_OLD
set a.ITEM_CODE = COALESCE(b.ITEM_CODE_SAP,a.ITEM_CODE_OLD);

UPDATE xilnex_rpt.tr_tbl_sales_by_item a
JOIN xilnex.ms_item mi on mi.ITEM_CODE = a.ITEM_CODE and mi.UOM = 'EA'
SET
	a.ITEM_NAME = REPLACE(CONVERT(mi.ITEM_NAME USING UTF8),'\n',''),
	a.ITEM_TYPE = REPLACE(mi.ITEM_TYPE,'\n',''),
--	a.ITEM_MODEL = REPLACE(mi.ITEM_MODEL,'\n',''),
	a.ITEM_CATEGORY = REPLACE(mi.CATEGORY,'\n',''),
	a.ITEM_DEPARTEMENT = REPLACE(mi.DEPARTMENT_CODE,'\n',''),
	a.ITEM_GROUP = REPLACE(mi.GROUP_NAME,'\n',''),
	a.DESCRIPTION = REPLACE(mi.ITEM_DESCRIPTION,'\n',''),
	a.BRAND = REPLACE(mi.ITEM_BRAND,'\n',''),
	a.UOM = REPLACE(mi.CONVERT_UOM,'\n',''),
	a.PRINSIPAL = REPLACE(mi.PRINSIPAL,'\n',''),
	a.ITEM_TYPE = REPLACE(mi.TYPE,'\n',''),
	a.FULL_NAME_PRODUCT = REPLACE(mi.FULL_NAME,'\n',''),
	a.ACTIVE_SKU = REPLACE(mi.ACTIVE_SKU,'\n','')
--	where a.SALES_DATE >= @tgl
;
 -- drop temporary table if exists tbl_sidel;

UPDATE xilnex_rpt.tr_tbl_sales_by_item a
left join tbl_calc tp on tp.ITEM_CODE = a.ITEM_CODE and tp.SALES_DATE = a.SALES_DATE
set ITEM_SMI =  CASE 
					when a.TOTAL_PRICE_AFTER_TAX * 1 = tp.PARETO_TOTAL * 1 THEN 'YES' 
					ELSE 'NO' 
				END ,
    PARETO_TYPE = CASE 
					WHEN a.TOTAL_PRICE_AFTER_TAX / tp.PARETO_TOTAL > 0.7 THEN 'A'
					WHEN a.TOTAL_PRICE_AFTER_TAX / tp.PARETO_TOTAL > 0.1 THEN 'B'
				ELSE 'C' END ;
 
UPDATE xilnex_rpt.tr_tbl_sales_by_item a
left join xilnex.tr_mch b on b.ITEM_CODE = a.ITEM_CODE_OLD
set a.MCH = COALESCE(b.TOPIC,'NORMAL');

UPDATE xilnex_rpt.tr_tbl_sales_by_item a
join xilnex.tr_margin tm on a.SALES_NUMBER = tm.INVOICE_NO 
	and a.ITEM_CODE_OLD = tm.ITEM_CODE 
    and ROUND(a.QTY,0) = ROUND(tm.QTY,0) 
    and ROUND(a.TOTAL_PRICE_AFTER_TAX,0) = ROUND(tm.TOTAL_PRICE,0)
set a.TOTAL_NET_MARGIN_ACCURATE = tm.TOTAL_NET_MARGIN,
	a.TOTAL_GROSS_MARGIN_ACCURATE = tm.TOTAL_GROSS_MARGIN ;
/*
UPDATE xilnex_rpt.tr_tbl_sales_by_item a
join tbl_calc b on b.ITEM_CODE = a.ITEM_CODE and a.SALES_DATE = b.SALES_DATE
set a.FAST_MOVING = CASE 
						WHEN b.PERCENTAGE > 90 THEN 'SLOW MOVING'
                        WHEN b.PERCENTAGE >= 71 THEN 'NORMAL'
                        WHEN b.PERCENTAGE <= 70 THEN 'FAST MOVING'
                   --     else NULL
					END;*/
UPDATE xilnex_rpt.tr_tbl_sales_by_item a
join xilnex_rpt.tr_tbl_fast_moving b on b.ITEM_CODE = a.ITEM_CODE
set a.FAST_MOVING = b.FAST_MOVING_TYPE;

UPDATE xilnex_rpt.tr_tbl_sales_by_item a
join xilnex.ms_category_pic_cm cpc on cpc.ITEM_DIVISION = a.ITEM_GROUP
set a.CATEGORY_CM = cpc.CATEGORY_CM,
	a.PIC_CM = cpc.PIC_CM
where a.CATEGORY_CM IS NULL or a.PIC_CM IS NULL
;
UPDATE xilnex_rpt.tr_tbl_sales_by_item a
join xilnex.APP_4_PROMOTIONHISTORY ph on a.SALES_ITEM_NUMBER = ph.SALESITEM_ID and ph.WBDELETED = 0
set a.APPLIED_PROMO = ph.RULESNAME, 
	a.CAMPAIGN_NAME = ph.CAMPAIGNNAME
where a.APPLIED_PROMO IS NULL or a.CAMPAIGN_NAME IS NULL;


UPDATE xilnex_rpt.tr_tbl_sales_by_item a 
left join xilnex_rpt.ms_promo b on a.APPLIED_PROMO like CONCAT('%',b.promo_2,'%') and a.APPLIED_PROMO like CONCAT('%',b.promo_1,'%')
set a.PROMOTION_TYPE = CASE
						WHEN a.APPLIED_PROMO like '%CAT%' and a.APPLIED_PROMO like '%Fair%' THEN 'Brand Fair'
						WHEN a.APPLIED_PROMO IS NULL or a.APPLIED_PROMO = '' THEN  'Regular Sales'
                        WHEN b.type is NULL THEN 'Other'
                        ELSE b.type
					END
where a.PROMOTION_TYPE IS NULL;

UPDATE xilnex_rpt.tr_tbl_sales_by_item a
join APP_4_SALESITEM b on a.SALES_ITEM_NUMBER = b.ID
join APP_4_ITEMNONFUNCTIONALINFO c on c.ITEM_ID = b.ITEM_CODE and ISNULL(c.FIELD_25) = 0
set a.NON_FUNCTIONAL_25 = c.FIELD_25
where a.NON_FUNCTIONAL_25 IS NULL;


/*UPDATE xilnex_rpt.tr_tbl_sales_by_item a
join tbl_package tp on a.SALES_ITEM_NUMBER = tp.ID
set a.PARENT_QTY = CASE WHEN a.UOM <> 'EA' THEN ROUND((a.QTY * COALESCE(tp.DOUBLE_QUANTITY,1)),4) 
						WHEN a.UOM = 'EA' THEN ROUND(a.QTY,4) 
					END
WHERE a.PARENT_QTY IS NULL;
*/
drop temporary table if exists tbl_si;
drop temporary table if exists tbl_source;
drop temporary table if exists tbl_xsource;
drop temporary table if exists tbl_calc;


UPDATE `xilnex_rpt`.`tr_tbl_sales_by_item` a
join xilnex.APP_4_ITEM c on a.ITEM_ID_OLD = c.ID
join xilnex.ms_item_rafaksi b on a.CAMPAIGN_NAME = b.CAMPAIGN_NAME and c.ITEM_CODE = b.ITEM_CODE
SET 
a.REBATE = ROUND(COALESCE((b.VALUE * a.QTY),(a.SALE_PRICE * b.rafaksi * a.QTY)),4),
a.REBATE_ERROR = NULL
where a.REBATE IS NULL and a.SALES_DATE >= @tgl;

UPDATE `xilnex_rpt`.`tr_tbl_sales_by_item` a
join xilnex.APP_4_ITEM c on a.ITEM_ID_OLD = c.ID
join xilnex.ms_item_rafaksi b on a.CAMPAIGN_NAME = b.CAMPAIGN_NAME
SET 
a.REBATE_ERROR = 'REBATE NOT SET'
where a.REBATE IS NULL;


UPDATE `xilnex_rpt`.`tr_tbl_sales_by_item`
SET
`SALES_ITEM_NUMBER` = REPLACE(SALES_ITEM_NUMBER,'\n',''),
`SALES_NUMBER` = REPLACE(SALES_NUMBER,'\n',''),
`ITEM_CODE_OLD` = REPLACE(ITEM_CODE_OLD,'\n',''),
`ITEM_CODE` = REPLACE(ITEM_CODE,'\n',''),
`ITEM_NAME` = REPLACE(ITEM_NAME,'\n',''),
`ITEM_TYPE` = REPLACE(ITEM_TYPE,'\n',''),
`ITEM_MODEL` = REPLACE(ITEM_MODEL,'\n',''),
`ITEM_CATEGORY` = REPLACE(ITEM_CATEGORY,'\n',''),
`ITEM_DEPARTEMENT` = REPLACE(ITEM_DEPARTEMENT,'\n',''),
`ITEM_DIVISION` = REPLACE(ITEM_DIVISION,'\n',''),
`ITEM_GROUP` = REPLACE(ITEM_GROUP,'\n',''),
`DESCRIPTION` = REPLACE(DESCRIPTION,'\n',''),
`BRAND` = REPLACE(BRAND,'\n',''),
`UNIT_PRICE_BEFORE_TAX` = REPLACE(UNIT_PRICE_BEFORE_TAX,'\n',''),
`UNIT_PRICE_AFTER_TAX` = REPLACE(UNIT_PRICE_AFTER_TAX,'\n',''),
`QTY` = REPLACE(QTY,'\n',''),
`PARENT_QTY` = REPLACE(PARENT_QTY,'\n',''),
`TOTAL_PRICE_AFTER_TAX` = REPLACE(TOTAL_PRICE_AFTER_TAX,'\n',''),
`TOTAL_PRICE_BEFORE_TAX` = REPLACE(TOTAL_PRICE_BEFORE_TAX,'\n',''),
`TAX_PERCENTAGE` = REPLACE(TAX_PERCENTAGE,'\n',''),
`TAX_AMOUNT` = REPLACE(TAX_AMOUNT,'\n',''),
`TOTAL_DISCOUNT_PERCENTAGE` = REPLACE(TOTAL_DISCOUNT_PERCENTAGE,'\n',''),
`TOTAL_DISCOUNT` = REPLACE(TOTAL_DISCOUNT,'\n',''),
`DISCOUNT_COST_ON_PO` = REPLACE(DISCOUNT_COST_ON_PO,'\n',''),
`UNIT_COST` = REPLACE(UNIT_COST,'\n',''),
`FINAL_COST` = REPLACE(FINAL_COST,'\n',''),
`UNIT_PROFIT` = REPLACE(UNIT_PROFIT,'\n',''),
`TOTAL_PROFIT` = REPLACE(TOTAL_PROFIT,'\n',''),
`CUSTOMER_NAME` = REPLACE(CUSTOMER_NAME,'\n',''),
`MEMBER_TYPE` = REPLACE(MEMBER_TYPE,'\n',''),
`CHANNEL_TYPE` = REPLACE(CHANNEL_TYPE,'\n',''),
`SALES_DATE` = REPLACE(SALES_DATE,'\n',''),
`SALES_TIME` = REPLACE(SALES_TIME,'\n',''),
`STORE_LOCATION` = REPLACE(STORE_LOCATION,'\n',''),
`SALES_ITEM_REMARK` = REPLACE(SALES_ITEM_REMARK,'\n',''),
`SALES_PERSON` = REPLACE(SALES_PERSON,'\n',''),
`SALES_STATUS` = REPLACE(SALES_STATUS,'\n',''),
`BARCODE` = REPLACE(BARCODE,'\n',''),
`UOM` = REPLACE(UOM,'\n',''),
`APPLIED_PROMO` = REPLACE(APPLIED_PROMO,'\n',''),
`PROMOTION_TYPE` = REPLACE(PROMOTION_TYPE,'\n',''),
`VOUCHER_NUMBER` = REPLACE(VOUCHER_NUMBER,'\n',''),
`PRINSIPAL` = REPLACE(PRINSIPAL,'\n',''),
`FULL_NAME_PRODUCT` = REPLACE(FULL_NAME_PRODUCT,'\n',''),
`ITEM_ADDITIONAL_INFO` = REPLACE(ITEM_ADDITIONAL_INFO,'\n',''),
`ACTIVE_SKU` = REPLACE(ACTIVE_SKU,'\n',''),
`ITEM_SMI` = REPLACE(ITEM_SMI,'\n',''),
`PARETO_TYPE` = REPLACE(PARETO_TYPE,'\n',''),
`target` = REPLACE(target,'\n',''),
`CLIENT_ID` = REPLACE(CLIENT_ID,'\n',''),
`CLIENT_NAME` = REPLACE(CLIENT_NAME,'\n',''),
`CLIENT_CREATION_DATE` = REPLACE(CLIENT_CREATION_DATE,'\n',''),
`CLIENT_STORE_REGISTRATION` = REPLACE(CLIENT_STORE_REGISTRATION,'\n',''),
`TOTAL_NET_MARGIN_ACCURATE` = REPLACE(TOTAL_NET_MARGIN_ACCURATE,'\n',''),
`TOTAL_GROSS_MARGIN_ACCURATE` = REPLACE(TOTAL_GROSS_MARGIN_ACCURATE,'\n',''),
`MCH` = REPLACE(MCH,'\n',''),
`FAST_MOVING` = REPLACE(FAST_MOVING,'\n',''),
`CATEGORY_CM` = TRIM(REPLACE(CATEGORY_CM,'\n','')),
`PIC_CM` = REPLACE(PIC_CM,'\n',''),
`CAMPAIGN_NAME` = TRIM(REPLACE(CAMPAIGN_NAME,'\n','')),
`NON_FUNCTIONAL_25` = REPLACE(NON_FUNCTIONAL_25,'\n',''),
`TIMESTAMP` = REPLACE(TIMESTAMP,'\n','');
END;

