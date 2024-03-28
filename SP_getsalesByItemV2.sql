	declare begindate as DATE
	declare sdate_lm DATE;
	declare sdate_lm2 DATE;
	declare sdate_lm3 DATE;
	declare eodayedate_lm DATE;
	declare edate_lm DATE;
	declare edate_lm2 DATE;
	declare edate_lm3 DATE;
	
	declare edate DATE;


	set @edate = date(NOW() - INTERVAL 1 day);	
	set @sdate = LAST_DAY(@edate - INTERVAL 4 MONTH) + interval 1 day;
	set @begindate = LAST_DAY(@edate - INTERVAL 3 MONTH) + interval 1 day;
	set @sdate_lm = @begindate + interval 1 month;
	set @sdate_lm2 = @begindate + interval 2 month;
	
	set @eodayedate_lm = LAST_DAY(@edate - interval 1 month);
	SET @edate_lm = @edate - interval 1 month;
	set @edate_lm2 = @edate - interval 2 month;
	set @edate_lm3 = @edate - interval 3 month;
	

select @begindate, @sdate, @sdate_lm, @sdate_lm2,  @edate, @edate_lm, @edate_lm2, @edate_lm3


	SELECT 	year(a.sales_date) `YEAR`, 
			DATE_FORMAT(a.Sales_date, '%b') `MONTH`, 
			ds.Store_name `Store Location`,
			case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end CHANNEL_TYPE,
			IFNULL(mcm.group_category, 'Unknown') `Group Category 1`, 
			case when IFNULL(ITEM_DIVISION,'') = '' then 'Unknown' else ITEM_DIVISION end ITEM_DIVISION, 
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
			case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end Promo_type,
			a.PROMOTION_TYPE, 
			IFNULL(a.APPLIED_PROMO,'(blank)') APPLIED_PROMO,
			case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end FLAG_3M,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end FLAG_3LM,
			COUNT(distinct a.CLIENT_ID) CLIENT_ID, 
			COUNT(distinct a.SALES_NUMBER) SALES_NUMBER, 
			SUM(a.PARENT_QTY) PARENT_QTY, 
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE,
			SUM(IFNULL(a.REBATE,0)) REBATE,
			SUM(a.TOTAL_NET_MARGIN_ACCURATE + IFNULL(REBATE,0)) FINAL_MARGIN 
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	left join db_ip.mst_group_category_cm mcm on mcm.PIC_CM = a.PIC_CM
	-- left join mst_store_area_manager
	where a.sales_date between sdate and edate
	and a.SALES_STATUS not like 'CANCEL%'
	group by year(a.sales_date), DATE_FORMAT(a.Sales_date, '%b') , ds.Store_name,
			case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end,
			IFNULL(mcm.group_category, 'Unknown'), 
			case when IFNULL(ITEM_DIVISION,'') = '' then 'Unknown' else ITEM_DIVISION end, ITEM_GROUP,
			BRAND, NON_FUNCTIONAL_25, 
			CASE WHEN a.MEMBER_TYPE = 'member'
					 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
					 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
					 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
				 THEN 'New Member'
				 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
				 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
			ELSE 'Existing' end, 
			case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end,
			a.PROMOTION_TYPE, IFNULL(a.APPLIED_PROMO,'(blank)'),
			case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end ;	


call SP_RPT_getSalesByItem (4, null, null)

select year(NOW())*100 + MONTH(NOW())

select CAST(date_format(now(),'%Y%m') as UNSIGNED)


SELECT 	year(a.sales_date) `YEAR`, 
				DATE_FORMAT(a.Sales_date, '%b') `MONTH`, 
				ds.Store_name `Store Location`,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end CHANNEL_TYPE,
-- 				IFNULL(mcm.group_category, 'Unknown') `Group Category 1`, 
-- 				ITEM_DIVISION, 
-- 				ITEM_GROUP,
-- 				BRAND, 
-- 				NON_FUNCTIONAL_25, 
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end `Member Classification`, 
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end Promo_type,
-- 				a.PROMOTION_TYPE, 
-- 				IFNULL(a.APPLIED_PROMO,'(blank)') APPLIED_PROMO,
				case when year(a.SALES_DATE)*100 + MONTH(a.SALES_DATE) between CAST(date_format('20240101','%Y%m') as UNSIGNED) and CAST(date_format('20240330','%Y%m') as UNSIGNED) then 1 else 0 end FLAG_3M,
				COUNT(distinct a.CLIENT_ID) CLIENT_ID, 
				COUNT(distinct a.SALES_NUMBER) SALES_NUMBER, 
				SUM(a.PARENT_QTY) PARENT_QTY, 
				SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE,
-- 				SUM(IFNULL(a.REBATE,0)) REBATE,
				SUM(a.TOTAL_NET_MARGIN_ACCURATE + IFNULL(REBATE,0)) FINAL_MARGIN 
		from db_ip.ip_tr_tbl_sales_by_item a
		join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
		join db_ip.dim_store ds on ds.store_id  = a.store_id
		left join db_ip.mst_group_category_cm mcm on mcm.PIC_CM = a.PIC_CM
		-- left join mst_store_area_manager
		where a.sales_date between '20230101' and '20240306'
		and a.SALES_STATUS not like 'CANCEL%'
		group by year(a.sales_date), DATE_FORMAT(a.Sales_date, '%b') , 
				ds.Store_name,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end,
-- 				IFNULL(mcm.group_category, 'Unknown'), 
-- 				ITEM_DIVISION, ITEM_GROUP,
-- 				BRAND, NON_FUNCTIONAL_25, 
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end,
				case when year(a.SALES_DATE)*100 + MONTH(a.SALES_DATE) between CAST(date_format('20240101','%Y%m') as UNSIGNED) and CAST(date_format('20240330','%Y%m') as UNSIGNED) then 1 else 0 end
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end,
-- 				a.PROMOTION_TYPE, IFNULL(a.APPLIED_PROMO,'(blank)');
				-- a.CLIENT_ID, a.SALES_NUMBER;
				order by 1,2;


			
			
			
			

call SP_RPT_getSalesByItemV2 (1, null, null)
call SP_RPT_getSalesByItemV2 (10, null, null)


-- ============== DARI SINI 


DELIMITER $$

create PROCEDURE `db_ip`.`SP_RPT_getSalesByItemV2`(
	in  T INT, -- type parameter 
	IN  sdate date,
	IN 	edate date)
BEGIN

-- Raw data 	
	declare begindate DATE;
	declare sdate_lm DATE;
	declare sdate_lm2 DATE;
	declare sdate_lm3 DATE;

	declare eodayedate_lm DATE;

	declare edate_lm DATE;
	declare edate_lm2 DATE;

	declare edate_lm3 DATE;

	if IFNULL(edate,'') = '' then 
		set edate = date(NOW() - INTERVAL 1 day);	
	end if;

	if IFNULL(sdate,'') = '' then 
		set sdate = LAST_DAY(edate - INTERVAL 4 MONTH) + interval 1 day;
	end if;
	
	set begindate = LAST_DAY(edate - INTERVAL 3 MONTH) + interval 1 day;

	set sdate_lm = begindate + interval 1 month;
	set sdate_lm2 = begindate + interval 2 month;
	-- set sdate_lm3 = begindate + interval 3 month;
	set eodayedate_lm = LAST_DAY(edate - interval 1 month);

	SET edate_lm = edate - interval 1 month;
	set edate_lm2 = edate - interval 2 month;
	set edate_lm3 = edate - interval 3 month;
	
	IF T = 1 then 
		-- select 1;
	SELECT 	year(a.sales_date) `YEAR`, 
			DATE_FORMAT(a.Sales_date, '%b') `MONTH`, 
			ds.Store_name `Store Location`,
			case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end CHANNEL_TYPE,
			IFNULL(mcm.group_category, 'Unknown') `Group Category 1`, 
			case when IFNULL(ITEM_DIVISION,'') = '' then 'Unknown' else ITEM_DIVISION end ITEM_DIVISION, 
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
			case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end Promo_type,
			case when a.PROMOTION_TYPE = 'Flush Out' then 'Other' else a.PROMOTION_TYPE end PROMOTION_TYPE,
			IFNULL(a.APPLIED_PROMO,'(blank)') APPLIED_PROMO,
			case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end FLAG_3M,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end FLAG_3LM,
			COUNT(distinct a.CLIENT_ID) CLIENT_ID, 
			COUNT(distinct a.SALES_NUMBER) SALES_NUMBER, 
			SUM(a.PARENT_QTY) PARENT_QTY, 
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE,
			SUM(IFNULL(a.REBATE,0)) REBATE,
			SUM(a.TOTAL_NET_MARGIN_ACCURATE + IFNULL(REBATE,0)) FINAL_MARGIN 
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	left join db_ip.mst_group_category_cm mcm on mcm.PIC_CM = a.PIC_CM
	-- left join mst_store_area_manager
	where a.sales_date between sdate and edate
	and a.SALES_STATUS not like 'CANCEL%'
	group by year(a.sales_date), DATE_FORMAT(a.Sales_date, '%b') , ds.Store_name,
			case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end,
			IFNULL(mcm.group_category, 'Unknown'), 
			case when IFNULL(ITEM_DIVISION,'') = '' then 'Unknown' else ITEM_DIVISION end, ITEM_GROUP,
			BRAND, NON_FUNCTIONAL_25, 
			CASE WHEN a.MEMBER_TYPE = 'member'
					 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
					 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
					 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
				 THEN 'New Member'
				 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
				 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
			ELSE 'Existing' end, 
			case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end,
			case when a.PROMOTION_TYPE = 'Flush Out' then 'Other' else a.PROMOTION_TYPE end,  IFNULL(a.APPLIED_PROMO,'(blank)'),
			case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end ;	
			-- a.CLIENT_ID, a.SALES_NUMBER;
	end if;

	-- D.Channel
	IF T = 2 then 
		SELECT 	year(a.sales_date) `YEAR`, 
				DATE_FORMAT(a.Sales_date, '%b') `MONTH`, 
				-- ds.Store_name `Store Location`,
				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end CHANNEL_TYPE,
-- 				IFNULL(mcm.group_category, 'Unknown') `Group Category 1`, 
-- 				ITEM_DIVISION, 
-- 				ITEM_GROUP,
-- 				BRAND, 
-- 				NON_FUNCTIONAL_25, 
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end `Member Classification`, 
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end Promo_type,
-- 				a.PROMOTION_TYPE, 
-- 				IFNULL(a.APPLIED_PROMO,'(blank)') APPLIED_PROMO,
				case when 
					(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
					or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
					or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
					 then 1 
					 else 0 
				end FLAG_3M,
	-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
				case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				)
				then 1 
					 else 0 
				end FLAG_3LM, 
				COUNT(distinct a.CLIENT_ID) CLIENT_ID, 
				COUNT(distinct a.SALES_NUMBER) SALES_NUMBER, 
				SUM(a.PARENT_QTY) PARENT_QTY, 
				SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE,
-- 				SUM(IFNULL(a.REBATE,0)) REBATE,
				SUM(a.TOTAL_NET_MARGIN_ACCURATE + IFNULL(REBATE,0)) FINAL_MARGIN 
		from db_ip.ip_tr_tbl_sales_by_item a
		join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
		join db_ip.dim_store ds on ds.store_id  = a.store_id
		left join db_ip.mst_group_category_cm mcm on mcm.PIC_CM = a.PIC_CM
		-- left join mst_store_area_manager
		where a.sales_date between '20230101' and edate
		and a.SALES_STATUS not like 'CANCEL%'
		group by year(a.sales_date), DATE_FORMAT(a.Sales_date, '%b') , ds.Store_name,
				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end,
-- 				IFNULL(mcm.group_category, 'Unknown'), 
-- 				ITEM_DIVISION, ITEM_GROUP,
-- 				BRAND, NON_FUNCTIONAL_25, 
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end,
			case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end 
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end,
-- 				a.PROMOTION_TYPE, IFNULL(a.APPLIED_PROMO,'(blank)');
				-- a.CLIENT_ID, a.SALES_NUMBER;
				order by 1,2;
		end if;

	-- D.Cat
	IF T = 3 then 
		SELECT 	year(a.sales_date) `YEAR`, 
				DATE_FORMAT(a.Sales_date, '%b') `MONTH`, 
				-- ds.Store_name `Store Location`,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end CHANNEL_TYPE,
				IFNULL(mcm.group_category, 'Unknown') `Group Category 1`, 
-- 				ITEM_DIVISION, 
-- 				ITEM_GROUP,
-- 				BRAND, 
-- 				NON_FUNCTIONAL_25, 
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end `Member Classification`, 
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end Promo_type,
-- 				a.PROMOTION_TYPE, 
-- 				IFNULL(a.APPLIED_PROMO,'(blank)') APPLIED_PROMO,
				case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end FLAG_3M,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end FLAG_3LM,
				COUNT(distinct a.CLIENT_ID) CLIENT_ID, 
				COUNT(distinct a.SALES_NUMBER) SALES_NUMBER, 
				SUM(a.PARENT_QTY) PARENT_QTY, 
				SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE,
-- 				SUM(IFNULL(a.REBATE,0)) REBATE,
				SUM(a.TOTAL_NET_MARGIN_ACCURATE + IFNULL(REBATE,0)) FINAL_MARGIN 
		from db_ip.ip_tr_tbl_sales_by_item a
		join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
		join db_ip.dim_store ds on ds.store_id  = a.store_id
		left join db_ip.mst_group_category_cm mcm on mcm.PIC_CM = a.PIC_CM
		-- left join mst_store_area_manager
		where a.sales_date between '20230101' and edate
		and a.SALES_STATUS not like 'CANCEL%'
		group by year(a.sales_date), DATE_FORMAT(a.Sales_date, '%b'),
-- 				,ds.Store_name,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end,
				IFNULL(mcm.group_category, 'Unknown'), 
-- 				ITEM_DIVISION, ITEM_GROUP,
-- 				BRAND, NON_FUNCTIONAL_25, 
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end,
			case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end  
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end,
-- 				a.PROMOTION_TYPE, IFNULL(a.APPLIED_PROMO,'(blank)');
				-- a.CLIENT_ID, a.SALES_NUMBER;
				order by 1,2;
		end if;	
	
	-- D.Store
	IF T = 4 then 
		SELECT 	year(a.sales_date) `YEAR`, 
				DATE_FORMAT(a.Sales_date, '%b') `MONTH`, 
				ds.Store_name `Store Location`,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end CHANNEL_TYPE,
-- 				IFNULL(mcm.group_category, 'Unknown') `Group Category 1`, 
-- 				ITEM_DIVISION, 
-- 				ITEM_GROUP,
-- 				BRAND, 
-- 				NON_FUNCTIONAL_25, 
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end `Member Classification`, 
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end Promo_type,
-- 				a.PROMOTION_TYPE, 
-- 				IFNULL(a.APPLIED_PROMO,'(blank)') APPLIED_PROMO,
-- 				case when year(a.SALES_DATE)*100 + MONTH(a.SALES_DATE) between CAST(date_format(sdate,'%Y%m') as UNSIGNED) and CAST(date_format(edate,'%Y%m') as UNSIGNED) then 1 else 0 end FLAG_3M,
				case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end FLAG_3M,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end FLAG_3LM,
				COUNT(distinct a.CLIENT_ID) CLIENT_ID, 
				COUNT(distinct a.SALES_NUMBER) SALES_NUMBER, 
				SUM(a.PARENT_QTY) PARENT_QTY, 
				SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE,
-- 				SUM(IFNULL(a.REBATE,0)) REBATE,
				SUM(a.TOTAL_NET_MARGIN_ACCURATE + IFNULL(REBATE,0)) FINAL_MARGIN 
		from db_ip.ip_tr_tbl_sales_by_item a
		join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
		join db_ip.dim_store ds on ds.store_id  = a.store_id
		left join db_ip.mst_group_category_cm mcm on mcm.PIC_CM = a.PIC_CM
		-- left join mst_store_area_manager
		where a.sales_date between '20230101' and edate
		and a.SALES_STATUS not like 'CANCEL%'
		group by year(a.sales_date), DATE_FORMAT(a.Sales_date, '%b') , 
				ds.Store_name,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end,
-- 				IFNULL(mcm.group_category, 'Unknown'), 
-- 				ITEM_DIVISION, ITEM_GROUP,
-- 				BRAND, NON_FUNCTIONAL_25, 
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end,
			case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end  
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end,
-- 				a.PROMOTION_TYPE, IFNULL(a.APPLIED_PROMO,'(blank)');
				-- a.CLIENT_ID, a.SALES_NUMBER;
				order by 1,2;
		end if;
	
	-- D.Member
	IF T = 5 then 
		SELECT 	year(a.sales_date) `YEAR`, 
				DATE_FORMAT(a.Sales_date, '%b') `MONTH`, 
-- 				ds.Store_name `Store Location`,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end CHANNEL_TYPE,
-- 				IFNULL(mcm.group_category, 'Unknown') `Group Category 1`, 
-- 				ITEM_DIVISION, 
-- 				ITEM_GROUP,
-- 				BRAND, 
-- 				NON_FUNCTIONAL_25, 
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end `Member Classification`, 
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end Promo_type,
-- 				a.PROMOTION_TYPE, 
-- 				IFNULL(a.APPLIED_PROMO,'(blank)') APPLIED_PROMO,
			case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end FLAG_3M,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end FLAG_3LM,
				COUNT(distinct a.CLIENT_ID) CLIENT_ID, 
				COUNT(distinct a.SALES_NUMBER) SALES_NUMBER, 
				SUM(a.PARENT_QTY) PARENT_QTY, 
				SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE,
-- 				SUM(IFNULL(a.REBATE,0)) REBATE,
				SUM(a.TOTAL_NET_MARGIN_ACCURATE + IFNULL(REBATE,0)) FINAL_MARGIN 
		from db_ip.ip_tr_tbl_sales_by_item a
		join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
		join db_ip.dim_store ds on ds.store_id  = a.store_id
		left join db_ip.mst_group_category_cm mcm on mcm.PIC_CM = a.PIC_CM
		-- left join mst_store_area_manager
		where a.sales_date between '20230101' and edate
		and a.SALES_STATUS not like 'CANCEL%'
		group by year(a.sales_date), DATE_FORMAT(a.Sales_date, '%b') , 
-- 				ds.Store_name,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end,
-- 				IFNULL(mcm.group_category, 'Unknown'), 
-- 				ITEM_DIVISION, ITEM_GROUP,
-- 				BRAND, NON_FUNCTIONAL_25, 
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end,
			case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end, 
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end  
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end,
-- 				a.PROMOTION_TYPE, IFNULL(a.APPLIED_PROMO,'(blank)');
				-- a.CLIENT_ID, a.SALES_NUMBER;
				order by 1,2;
		end if;

	-- D.CatMov
	IF T = 6 then 
		SELECT 	year(a.sales_date) `YEAR`, 
				DATE_FORMAT(a.Sales_date, '%b') `MONTH`, 
-- 				ds.Store_name `Store Location`,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end CHANNEL_TYPE,
-- 				IFNULL(mcm.group_category, 'Unknown') `Group Category 1`, 
-- 				ITEM_DIVISION, 
-- 				ITEM_GROUP,
-- 				BRAND, 
				case when NON_FUNCTIONAL_25 = '#N/A' then 'UNKNOWN' else NON_FUNCTIONAL_25 end NON_FUNCTIONAL_25, 
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end `Member Classification`, 
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end Promo_type,
-- 				a.PROMOTION_TYPE, 
-- 				IFNULL(a.APPLIED_PROMO,'(blank)') APPLIED_PROMO,
			case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end FLAG_3M,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end FLAG_3LM,
				COUNT(distinct a.CLIENT_ID) CLIENT_ID, 
				COUNT(distinct a.SALES_NUMBER) SALES_NUMBER, 
				SUM(a.PARENT_QTY) PARENT_QTY, 
				SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE,
-- 				SUM(IFNULL(a.REBATE,0)) REBATE,
				SUM(a.TOTAL_NET_MARGIN_ACCURATE + IFNULL(REBATE,0)) FINAL_MARGIN 
		from db_ip.ip_tr_tbl_sales_by_item a
		join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
		join db_ip.dim_store ds on ds.store_id  = a.store_id
		left join db_ip.mst_group_category_cm mcm on mcm.PIC_CM = a.PIC_CM
		-- left join mst_store_area_manager
		where a.sales_date between '20230101' and edate
		and a.SALES_STATUS not like 'CANCEL%'
		group by year(a.sales_date), DATE_FORMAT(a.Sales_date, '%b') , 
-- 				ds.Store_name,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end,
-- 				IFNULL(mcm.group_category, 'Unknown'), 
-- 				ITEM_DIVISION, ITEM_GROUP,
-- 				BRAND, 
				case when NON_FUNCTIONAL_25 = '#N/A' then 'UNKNOWN' else NON_FUNCTIONAL_25 end, 
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end,
				case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end ,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end  
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end,
-- 				a.PROMOTION_TYPE, IFNULL(a.APPLIED_PROMO,'(blank)');
				-- a.CLIENT_ID, a.SALES_NUMBER;
				order by 1,2;
		end if;
	
	-- D.SubCat1
	IF T = 7 then 
		SELECT 	year(a.sales_date) `YEAR`, 
				DATE_FORMAT(a.Sales_date, '%b') `MONTH`, 
-- 				ds.Store_name `Store Location`,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end CHANNEL_TYPE,
-- 				IFNULL(mcm.group_category, 'Unknown') `Group Category 1`, 
				ITEM_DIVISION, 
-- 				ITEM_GROUP,
-- 				BRAND, 
-- 				NON_FUNCTIONAL_25, 
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end `Member Classification`, 
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end Promo_type,
-- 				a.PROMOTION_TYPE, 
-- 				IFNULL(a.APPLIED_PROMO,'(blank)') APPLIED_PROMO,
				case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end FLAG_3M,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end FLAG_3LM,
				COUNT(distinct a.CLIENT_ID) CLIENT_ID, 
				COUNT(distinct a.SALES_NUMBER) SALES_NUMBER, 
				SUM(a.PARENT_QTY) PARENT_QTY, 
				SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE,
-- 				SUM(IFNULL(a.REBATE,0)) REBATE,
				SUM(a.TOTAL_NET_MARGIN_ACCURATE + IFNULL(REBATE,0)) FINAL_MARGIN 
		from db_ip.ip_tr_tbl_sales_by_item a
		join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
		join db_ip.dim_store ds on ds.store_id  = a.store_id
		left join db_ip.mst_group_category_cm mcm on mcm.PIC_CM = a.PIC_CM
		-- left join mst_store_area_manager
		where a.sales_date between '20230101' and edate
		and a.SALES_STATUS not like 'CANCEL%'
		group by year(a.sales_date), DATE_FORMAT(a.Sales_date, '%b') , 
-- 				ds.Store_name,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end,
-- 				IFNULL(mcm.group_category, 'Unknown'), 
				ITEM_DIVISION, 
-- 				ITEM_GROUP,
-- 				BRAND, 
-- 				NON_FUNCTIONAL_25, 
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end,
			case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end 
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end,
-- 				a.PROMOTION_TYPE, IFNULL(a.APPLIED_PROMO,'(blank)');
				-- a.CLIENT_ID, a.SALES_NUMBER;
				order by 1,2;
		end if;
	
	-- D.Promo
	IF T = 8 then 
		SELECT 	year(a.sales_date) `YEAR`, 
				DATE_FORMAT(a.Sales_date, '%b') `MONTH`, 
-- 				ds.Store_name `Store Location`,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end CHANNEL_TYPE,
-- 				IFNULL(mcm.group_category, 'Unknown') `Group Category 1`, 
-- 				ITEM_DIVISION, 
-- 				ITEM_GROUP,
-- 				BRAND, 
-- 				NON_FUNCTIONAL_25, 
-- 				a.PROMOTION_TYPE,
				case when a.PROMOTION_TYPE = 'Flush Out' then 'Other' else a.PROMOTION_TYPE end PROMOTION_TYPE,
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end `Member Classification`, 
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end Promo_type,
-- 				IFNULL(a.APPLIED_PROMO,'(blank)') APPLIED_PROMO,
			case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end FLAG_3M,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end FLAG_3LM,
				COUNT(distinct a.CLIENT_ID) CLIENT_ID, 
				COUNT(distinct a.SALES_NUMBER) SALES_NUMBER, 
				SUM(a.PARENT_QTY) PARENT_QTY, 
				SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE,
-- 				SUM(IFNULL(a.REBATE,0)) REBATE,
				SUM(a.TOTAL_NET_MARGIN_ACCURATE + IFNULL(REBATE,0)) FINAL_MARGIN 
		from db_ip.ip_tr_tbl_sales_by_item a
		join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
		join db_ip.dim_store ds on ds.store_id  = a.store_id
		left join db_ip.mst_group_category_cm mcm on mcm.PIC_CM = a.PIC_CM
		-- left join mst_store_area_manager
		where a.sales_date between '20230101' and edate
		and a.SALES_STATUS not like 'CANCEL%'
		and a.applied_promo not like '%voucher%'
		and a.APPLIED_PROMO like 'CAT%'
		group by year(a.sales_date), DATE_FORMAT(a.Sales_date, '%b') , 
-- 				ds.Store_name,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end,
-- 				IFNULL(mcm.group_category, 'Unknown'), 
-- 				ITEM_DIVISION, 
-- 				ITEM_GROUP,
-- 				BRAND, 
-- 				NON_FUNCTIONAL_25, 
				case when a.PROMOTION_TYPE = 'Flush Out' then 'Other' else a.PROMOTION_TYPE end,
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end,
			case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end ,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end  
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end,
-- 				IFNULL(a.APPLIED_PROMO,'(blank)');
				-- a.CLIENT_ID, a.SALES_NUMBER;
				order by 1,2;
		end if;
	
	-- D.Brand
	IF T = 9 then 
		SELECT 	year(a.sales_date) `YEAR`, 
				DATE_FORMAT(a.Sales_date, '%b') `MONTH`, 
-- 				ds.Store_name `Store Location`,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end CHANNEL_TYPE,
-- 				IFNULL(mcm.group_category, 'Unknown') `Group Category 1`, 
-- 				ITEM_DIVISION, 
-- 				ITEM_GROUP,
				BRAND, 
-- 				NON_FUNCTIONAL_25, 
-- 				a.PROMOTION_TYPE,
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end `Member Classification`, 
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end Promo_type,
-- 				IFNULL(a.APPLIED_PROMO,'(blank)') APPLIED_PROMO,
			case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end FLAG_3M,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end FLAG_3LM,
				COUNT(distinct a.CLIENT_ID) CLIENT_ID, 
				COUNT(distinct a.SALES_NUMBER) SALES_NUMBER, 
				SUM(a.PARENT_QTY) PARENT_QTY, 
				SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE,
-- 				SUM(IFNULL(a.REBATE,0)) REBATE,
				SUM(a.TOTAL_NET_MARGIN_ACCURATE + IFNULL(REBATE,0)) FINAL_MARGIN 
		from db_ip.ip_tr_tbl_sales_by_item a
		join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
		join db_ip.dim_store ds on ds.store_id  = a.store_id
		left join db_ip.mst_group_category_cm mcm on mcm.PIC_CM = a.PIC_CM
		-- left join mst_store_area_manager
		where a.sales_date between '20230101' and edate
		and a.SALES_STATUS not like 'CANCEL%'
		and a.applied_promo not like '%voucher%'
		and a.APPLIED_PROMO like 'CAT%'
		group by year(a.sales_date), DATE_FORMAT(a.Sales_date, '%b') , 
-- 				ds.Store_name,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end,
-- 				IFNULL(mcm.group_category, 'Unknown'), 
-- 				ITEM_DIVISION, 
-- 				ITEM_GROUP,
				BRAND, 
-- 				NON_FUNCTIONAL_25, 
-- 				a.PROMOTION_TYPE,
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end,
				case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end 
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end,
-- 				IFNULL(a.APPLIED_PROMO,'(blank)');
				-- a.CLIENT_ID, a.SALES_NUMBER;
				order by 1,2;
		end if;	
	
-- D.SubCat2
	IF T = 10 then 
		SELECT 	year(a.sales_date) `YEAR`, 
				DATE_FORMAT(a.Sales_date, '%b') `MONTH`, 
-- 				ds.Store_name `Store Location`,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end CHANNEL_TYPE,
-- 				IFNULL(mcm.group_category, 'Unknown') `Group Category 1`, 
-- 				ITEM_DIVISION, 
				ITEM_GROUP,
-- 				BRAND, 
-- 				NON_FUNCTIONAL_25, 
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end `Member Classification`, 
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end Promo_type,
-- 				a.PROMOTION_TYPE, 
-- 				IFNULL(a.APPLIED_PROMO,'(blank)') APPLIED_PROMO,
				case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end FLAG_3M,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end FLAG_3LM,
				COUNT(distinct a.CLIENT_ID) CLIENT_ID, 
				COUNT(distinct a.SALES_NUMBER) SALES_NUMBER, 
				SUM(a.PARENT_QTY) PARENT_QTY, 
				SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE,
-- 				SUM(IFNULL(a.REBATE,0)) REBATE,
				SUM(a.TOTAL_NET_MARGIN_ACCURATE + IFNULL(REBATE,0)) FINAL_MARGIN 
		from db_ip.ip_tr_tbl_sales_by_item a
		join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
		join db_ip.dim_store ds on ds.store_id  = a.store_id
		left join db_ip.mst_group_category_cm mcm on mcm.PIC_CM = a.PIC_CM
		-- left join mst_store_area_manager
		where a.sales_date between '20230101' and edate
		and a.SALES_STATUS not like 'CANCEL%'
		group by year(a.sales_date), DATE_FORMAT(a.Sales_date, '%b') , 
-- 				ds.Store_name,
-- 				case when a.CHANNEL_TYPE like 'IN STORE%' then 'STORE' else a.CHANNEL_TYPE end,
-- 				IFNULL(mcm.group_category, 'Unknown'), 
-- 				ITEM_DIVISION, 
				ITEM_GROUP,
-- 				BRAND, 
-- 				NON_FUNCTIONAL_25, 
				CASE WHEN a.MEMBER_TYPE = 'member'
						 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
						 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
						 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
					 THEN 'New Member'
					 WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
					 WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
				ELSE 'Existing' end,
			case when 
				(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
				or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm2,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate,'%Y%m%d') as UNSIGNED))
				 then 1 
				 else 0 
			end,
-- 			case when YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(LAST_DAY(edate_lm),'%Y%m%d') as UNSIGNED)
			case when 
			(YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm3,'%Y%m%d') as UNSIGNED)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(begindate,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm2,'%Y%m%d') as unsigned)
			or YEAR(a.sales_date)*10000+MONTH(a.sales_date)*100+DAY(a.sales_date) between CAST(date_format(sdate_lm,'%Y%m%d') as UNSIGNED) and CAST(date_format(edate_lm,'%Y%m%d') as UNSIGNED)
			)
			then 1 
				 else 0 
			end 
-- 				case when a.APPLIED_PROMO like 'CAT%' then 'Promo' else '' end,
-- 				a.PROMOTION_TYPE, IFNULL(a.APPLIED_PROMO,'(blank)');
				-- a.CLIENT_ID, a.SALES_NUMBER;
				order by 1,2;
		end if;	
	
end$$

DELIMITER ;



select a.*, b.ITEM_CODE, c.store_name  
from ip_tr_tbl_stock_history a 
join dim_product_stockh b on a.IP_ITEM_ID = b.IP_ITEM_ID 
join dim_store c on a.STORE_ID = c.store_id 
where a.STOCK_DATE = '20240312'
and b.ITEM_CODE = '1116240088'

CREATE TABLE `dim_store` (
  `store_id` int NOT NULL AUTO_INCREMENT,
  `store_name` varchar(200) DEFAULT NULL,
  `store_location` varchar(250) DEFAULT NULL,
  KEY `ix_dim_store` (`store_id`)
) ENGINE=InnoDB AUTO_INCREMENT=21 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;

alter table dim_store 
add column `store_code` VARCHAR(5) null after `store_id`


select * from dim_store

update dim_store 
set store_code = 'VTR'
where store_id = 18
and store_code is null



