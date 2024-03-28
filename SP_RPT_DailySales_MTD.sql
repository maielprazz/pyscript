
call SP_RPT_DailySalesMTD(1, null, 'CRD') -- total sales -->OK
call SP_RPT_DailySalesMTD(2, null, 'CRD') -- halodoc -->OK
call SP_RPT_DailySalesMTD(3, null, 'CRD') -- CEKKES -->OK
call SP_RPT_DailySalesMTD(4, null, 'CRD') -- PWP  -->OK
call SP_RPT_DailySalesMTD(5, null, 'CRD') -- THEMATIC -->OK
call SP_RPT_DailySalesMTD(6, null, 'CRD') -- BC -->OK
call SP_RPT_DailySalesMTD(7, null, 'CRD') -- PHARMA -->OK
call SP_RPT_DailySalesMTD(8, null, 'BT5') -- no Existing -->OK
call SP_RPT_DailySalesMTD(9, null, 'BT5') -- TRX Existing -->OK
call SP_RPT_DailySalesMTD(10, null, 'BT5') -- no of new member -->OK
call SP_RPT_DailySalesMTD(11, null, 'BT5') -- TRX Non & new member -->OK
call SP_RPT_DailySalesMTD(12, '20240321', 'CRD') -- TOP 30 SKU
call SP_RPT_DailySalesMTD(13, NULL, 'BT5') -- cek kes by store by sales person
call SP_RPT_DailySalesMTD(14, NULL, null) 
call SP_RPT_DailySalesMTD(15, NULL, null) 
select * from ip_tr_tbl_sales_by_item ittsbi 
select * from dim_product_tsbi dpt 
SELECT DISTINCT STORE_CODE, STORE_NAME FROM DIM_STORE WHERE STORE_CODE IS NOT NULL order by 1
select dpt.ITEM_NAME, ittsbi.* 
from ip_tr_tbl_sales_by_item ittsbi  
join dim_product_tsbi dpt on dpt.IP_ITEM_ID = ittsbi.IP_ITEM_ID 
join dim_store ds on ds.store_id = ittsbi.STORE_ID 
where ittsbi.SALES_DATE between '20240301' and '20240304'
and ds.store_code = 'CRD'
and dpt.ITEM_NAME like 'CEK%'


create index x on CEKKES_PERSON (store_name, SALES_PERSON)

create table CEKKES_PERSON

drop procedure if exists `db_ip`.`SP_RPT_DailySalesMTD`;

DELIMITER $$

create PROCEDURE `db_ip`.`SP_RPT_DailySalesMTD`(
	in  T INT, -- type parameter 
	IN  asof date,
	IN 	store varchar(5))
BEGIN

	declare sdate DATE;
	declare edate DATE;
	declare sdate_lm DATE;
	declare edate_lm DATE;

	if IFNULL(asof, '19700101') = '19700101' then 
	set edate = CURDATE() - interval 1 day;
	else 
		set edate = asof;
	end if;

	set sdate = (edate) - interval day(edate) day + interval 1 day;
	set sdate_lm = sdate - interval 1 month;
	set edate_lm = edate - interval 1 month;

-- T = 1 Total Sales By Store
IF T = 1 then 
	
	-- storename, actual, target, achive, trx, lastmonthtrx, vs LM trx, ABV
	DROP temporary TABLE IF EXISTS slsgab; 
	create temporary table slsgab 
	-- Current Month
	SELECT 	d.`date`, 
			ds.store_code,
			'CM' Fg,
			COUNT(distinct a.SALES_NUMBER) SALES_NUMBER,
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate and edate
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	group by d.`date`, ds.store_code
	union 
	-- Last Month
	SELECT 	d.`date`, 
			ds.store_code,
			'LM' Fg,
			COUNT(distinct a.SALES_NUMBER) SALES_NUMBER,
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate_lm and edate_lm
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	group by d.`date`, ds.store_code;

	DROP temporary TABLE IF EXISTS slsgab_b; 
	create temporary table slsgab_b
	select * from slsgab;
	
	select 	date_format(d.`date`, '%Y%m%d') `DATE`, x.sales_value Sales, '' D4, t.Target, '' F,'' G, '' H, 
			x.trx, x.trx_lm,'' K, '' L, 
			case when IFNULL(x.trx_lm,0) = 0 then 0 ELSE (x.sales_value_lm/x.trx_lm) end LM_ABV, '' N
	from dim_date d
	left join (
	select 	a.`date`, 
			-- s.store_name, 
			a.store_code,
			SUM(a.sales_number) trx, SUM(b.sales_number) trx_lm, 
			SUM(a.sales_value) sales_value, SUM(b.sales_value) sales_value_lm  
	from slsgab a
	join (select distinct store_code, store_name from dim_store) s on s.store_code = a.store_code
	left join slsgab_b b on b.fg = 'LM' and b.`date` =  (a.`date` - interval 1 month)
	where a.fg = 'CM'
	group by a.`date`, 
			-- s.store_name, 
			a.store_code, 
			a.fg
	) x	on x.`date` = d.`date`
	left join db_ip.StoreDailyTarget t 
			on t.typ = '0' and t.sitecode = store and d.`date` = t.asofdate 
	where d.`date` between sdate and last_day(edate)
	order by 1 asc;
	
	
end if;

-- T = 2 Total Sales Halodoc By Store
IF T = 2 then 
	
	DROP temporary TABLE IF EXISTS slsgab; 
	create temporary table slsgab 
	-- Current Month
	SELECT 	d.date, 
			ds.store_code,
			'CM' Fg,
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate and edate
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and a.CLIENT_NAME = 'halodoc'
	and a.CHANNEL_TYPE = 'TELEMEDICINE'
	group by d.date, ds.store_code
	union 
	-- Last Month
	SELECT 	d.date, 
			ds.store_code,
			'LM' Fg,
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate_lm and edate_lm
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and a.CLIENT_NAME = 'halodoc'
	and a.CHANNEL_TYPE = 'TELEMEDICINE'
	group by d.date, ds.store_code;

	DROP temporary TABLE IF EXISTS slsgab_b; 
	create temporary table slsgab_b
	select * from slsgab;
	
	select date_format(d.`date`, '%Y%m%d') `DATE`, x.sales_value,'' R, t.target, '' T, '' U, '' V
	from dim_date d
	left join (
	select 	a.`date`, 
			-- s.store_name, 
			a.store_code,
			SUM(a.sales_value) sales_value, SUM(b.sales_value) sales_value_lm  
	from slsgab a
	join (select distinct store_code, store_name from dim_store) s on s.store_code = a.store_code
	left join slsgab_b b on b.fg = 'LM' and b.`date` =  (a.`date` - interval 1 month)
	where a.fg = 'CM'
	group by a.`date`, 
			-- s.store_name, 
			a.store_code, 
			a.fg
	) x	on x.`date` = d.`date`
	left join db_ip.StoreDailyTarget t 
			on t.typ = '1' and t.sitecode = store and d.`date` = t.asofdate 
	where d.`date` between sdate and last_day(edate)
	order by 1 asc;
	
end if;


-- T = 3 Total Sales CEK-KES By Store
IF T = 3 then 
	
	DROP temporary TABLE IF EXISTS slsgab; 
	create temporary table slsgab 
	-- Current Month
	SELECT 	d.date, 
			ds.store_code,
			'CM' Fg,
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate and edate
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and dpt.ITEM_NAME like 'CEK%'
	group by d.date, ds.store_code
	union 
	-- Last Month
	SELECT 	d.date, 
			ds.store_code,
			'LM' Fg,
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate_lm and edate_lm
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and dpt.ITEM_NAME like 'CEK%'
	group by d.date, ds.store_code;

	DROP temporary TABLE IF EXISTS slsgab_b; 
	create temporary table slsgab_b
	select * from slsgab;
	
-- 	select d.`date`, x.sales_value, x.sales_value_lm, t.target
	select date_format(d.`date`, '%Y%m%d') `DATE`, x.sales_value,'' Z, t.target, '' AB, '' AC, '' AD
	from dim_date d
	left join (
	select 	a.`date`, 
			-- s.store_name, 
			a.store_code,
			SUM(a.sales_value) sales_value, SUM(b.sales_value) sales_value_lm  
	from slsgab a
	join (select distinct store_code, store_name from dim_store) s on s.store_code = a.store_code
	left join slsgab_b b on b.fg = 'LM' and b.`date` =  (a.`date` - interval 1 month)
	where a.fg = 'CM'
	group by a.`date`, 
			-- s.store_name, 
			a.store_code, 
			a.fg
	) x	on x.`date` = d.`date`
	left join db_ip.StoreDailyTarget t 
			on t.typ = '2' and t.sitecode = store and d.`date` = t.asofdate 
	where d.`date` between sdate and last_day(edate)
	order by 1 asc;
	
end if;

-- T = 4 Total Sales PWP By Store
IF T = 4 then 
	
	DROP temporary TABLE IF EXISTS slsgab; 
	create temporary table slsgab 
	-- Current Month
	SELECT 	d.date, 
			ds.store_code,
			'CM' Fg,
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate and edate
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and a.PROMOTION_TYPE like 'PWP%'
	group by d.date, ds.store_code
	union 
	-- Last Month
	SELECT 	d.date, 
			ds.store_code,
			'LM' Fg,
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate_lm and edate_lm
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and a.PROMOTION_TYPE like 'PWP%'
	group by d.date, ds.store_code;

	DROP temporary TABLE IF EXISTS slsgab_b; 
	create temporary table slsgab_b
	select * from slsgab;
	
-- 	select d.`date`, x.sales_value, x.sales_value_lm, t.target
	select date_format(d.`date`, '%Y%m%d') `DATE`, x.sales_value,'' AH, t.target, '' AJ, '' AK, '' AL
	from dim_date d
	left join (
	select 	a.`date`, 
			-- s.store_name, 
			a.store_code,
			SUM(a.sales_value) sales_value, SUM(b.sales_value) sales_value_lm  
	from slsgab a
	join (select distinct store_code, store_name from dim_store) s on s.store_code = a.store_code
	left join slsgab_b b on b.fg = 'LM' and b.`date` =  (a.`date` - interval 1 month)
	where a.fg = 'CM'
	group by a.`date`, 
			-- s.store_name, 
			a.store_code, 
			a.fg
	) x	on x.`date` = d.`date`
	left join db_ip.StoreDailyTarget t 
			on t.typ = '3' and t.sitecode = store and d.`date` = t.asofdate 
	where d.`date` between sdate and last_day(edate)
	order by 1 asc;
	
end if;

-- T = 5 Total Sales Thematic By Store
IF T = 5 then 
	
	DROP temporary TABLE IF EXISTS slsgab; 
	create temporary table slsgab 
	-- Current Month
	SELECT 	d.date, 
			ds.store_code,
			'CM' Fg,
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate and edate
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and a.PROMOTION_TYPE not like 'PWP%'
	and a.APPLIED_PROMO not like 'VOUCHER%'
	group by d.date, ds.store_code
	union 
	-- Last Month
	SELECT 	d.date, 
			ds.store_code,
			'LM' Fg,
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate_lm and edate_lm
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and a.PROMOTION_TYPE not like 'PWP%'
	and a.APPLIED_PROMO not like 'VOUCHER%'
	group by d.date, ds.store_code;

	DROP temporary TABLE IF EXISTS slsgab_b; 
	create temporary table slsgab_b
	select * from slsgab;
	
-- 	select d.`date`, x.sales_value, x.sales_value_lm, t.target
	select date_format(d.`date`, '%Y%m%d') `DATE`, x.sales_value,'' AP, t.target, '' AR, '' `AS`, '' `AT`
	from dim_date d
	left join (
	select 	a.`date`, 
			a.store_code,
			SUM(a.sales_value) sales_value, SUM(b.sales_value) sales_value_lm  
	from slsgab a
	join (select distinct store_code, store_name from dim_store) s on s.store_code = a.store_code
	left join slsgab_b b on b.fg = 'LM' and b.`date` =  (a.`date` - interval 1 month)
	where a.fg = 'CM'
	group by a.`date`, 
			-- s.store_name, 
			a.store_code, 
			a.fg
	) x	on x.`date` = d.`date`
	left join db_ip.StoreDailyTarget t 
			on t.typ = '4' and t.sitecode = store and d.`date` = t.asofdate 
	where d.`date` between sdate and last_day(edate)
	order by 1 asc;
	
end if;

-- T = 6 Total Sales BC By Store
IF T = 6 then 
	
	DROP temporary TABLE IF EXISTS slsgab; 
	create temporary table slsgab 
	-- Current Month
	SELECT 	d.date, 
			ds.store_code,
			'CM' Fg,
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate and edate
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and dpt.Item_model like 'Best Choice%'
	group by d.date, ds.store_code
	union 
	-- Last Month
	SELECT 	d.date, 
			ds.store_code,
			'LM' Fg,
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate_lm and edate_lm
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and dpt.Item_model like 'Best Choice%'
	group by d.date, ds.store_code;

	DROP temporary TABLE IF EXISTS slsgab_b; 
	create temporary table slsgab_b
	select * from slsgab;
	
-- 	select d.`date`, x.sales_value, x.sales_value_lm, t.target
	select date_format(d.`date`, '%Y%m%d') `DATE`, x.sales_value,'' AX, t.target, '' AZ, '' BA, '' BB
	from dim_date d
	left join (
	select 	a.`date`, 
			a.store_code,
			SUM(a.sales_value) sales_value, SUM(b.sales_value) sales_value_lm  
	from slsgab a
	join (select distinct store_code, store_name from dim_store) s on s.store_code = a.store_code
	left join slsgab_b b on b.fg = 'LM' and b.`date` =  (a.`date` - interval 1 month)
	where a.fg = 'CM'
	group by a.`date`, 
			-- s.store_name, 
			a.store_code, 
			a.fg
	) x	on x.`date` = d.`date`
	left join db_ip.StoreDailyTarget t 
			on t.typ = '5' and t.sitecode = store and d.`date` = t.asofdate 
	where d.`date` between sdate and last_day(edate)
	order by 1 asc;
	
end if;

-- T = 7 Total Sales Pharma By Store
IF T = 7 then 
	
	DROP temporary TABLE IF EXISTS slsgab; 
	create temporary table slsgab 
	-- Current Month
	SELECT 	d.date, 
			ds.store_code,
			'CM' Fg,
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	left join db_ip.mst_group_category_cm mcm on mcm.PIC_CM = a.PIC_CM
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate and edate
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and IFNULL(mcm.group_category, 'Unknown') = 'Pharma'
	group by d.date, ds.store_code
	union 
	-- Last Month
	SELECT 	d.date, 
			ds.store_code,
			'LM' Fg,
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	left join db_ip.mst_group_category_cm mcm on mcm.PIC_CM = a.PIC_CM
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate_lm and edate_lm
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and IFNULL(mcm.group_category, 'Unknown') = 'Pharma'
	group by d.date, ds.store_code;

	DROP temporary TABLE IF EXISTS slsgab_b; 
	create temporary table slsgab_b
	select * from slsgab;
	
-- 	select d.`date`, x.sales_value, x.sales_value_lm, t.target
	select date_format(d.`date`, '%Y%m%d') `DATE`, x.sales_value,'' BF, t.target, '' BH, '' BI, '' BJ
	from dim_date d
	left join (
	select 	a.`date`, 
			a.store_code,
			SUM(a.sales_value) sales_value, SUM(b.sales_value) sales_value_lm  
	from slsgab a
	join (select distinct store_code, store_name from dim_store) s on s.store_code = a.store_code
	left join slsgab_b b on b.fg = 'LM' and b.`date` =  (a.`date` - interval 1 month)
	where a.fg = 'CM'
	group by a.`date`, 
			-- s.store_name, 
			a.store_code, 
			a.fg
	) x	on x.`date` = d.`date`
	left join db_ip.StoreDailyTarget t 
			on t.typ = '6' and t.sitecode = store and d.`date` = t.asofdate 
	where d.`date` between sdate and last_day(edate)
	order by 1 asc;
	
end if;

-- T = 8 No of Existing Member
IF T = 8 then 
	
	DROP temporary TABLE IF EXISTS slsgab; 
	create temporary table slsgab 
	-- Current Month
	SELECT 	d.date, 
			ds.store_code,
			'CM' Fg, 
			COUNT(distinct a.CLIENT_ID) CLIENTID
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate and edate
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and CASE WHEN a.MEMBER_TYPE = 'member'
			 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
			 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
			 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
		 	THEN 'New Member'
		 	WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
		 	WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
			ELSE 'Existing' end = 'Existing'
	group by d.date, ds.store_code
	union 
	-- Last Month
	SELECT 	d.date, 
			ds.store_code,
			'LM' Fg,
			COUNT(distinct a.CLIENT_ID) CLIENTID
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate_lm and edate_lm
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and CASE WHEN a.MEMBER_TYPE = 'member'
			 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
			 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
			 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
		 	THEN 'New Member'
		 	WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
		 	WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
			ELSE 'Existing' end = 'Existing'
	group by d.date, ds.store_code;

	DROP temporary TABLE IF EXISTS slsgab_b; 
	create temporary table slsgab_b
	select * from slsgab;
	
-- 	select d.`date`, x.CLIENTID, x.CLIENTID_LM, t.target
	select date_format(d.`date`, '%Y%m%d') `DATE`, x.CLIENTID,'' D, t.target, '' F, '' G, '' H
	from dim_date d
	left join (
	select 	a.`date`, 
			-- s.store_name, 
			a.store_code,
			SUM(a.CLIENTID) CLIENTID, SUM(b.clientid) CLIENTID_LM
	from slsgab a
	join (select distinct store_code, store_name from dim_store) s on s.store_code = a.store_code
	left join slsgab_b b on b.fg = 'LM' and b.`date` =  (a.`date` - interval 1 month)
	where a.fg = 'CM'
	group by a.`date`, 
			-- s.store_name, 
			a.store_code, 
			a.fg
	) x	on x.`date` = d.`date`
	left join db_ip.StoreDailyTarget t 
			on t.typ = '7' and t.sitecode = store and d.`date` = t.asofdate 
	where d.`date` between sdate and last_day(edate)
	order by 1 asc;

end if;

-- T = 9 TRX of Existing Member
IF T = 9 then 
	
	DROP temporary TABLE IF EXISTS slsgab; 
	create temporary table slsgab 
	-- Current Month
	SELECT 	d.date, 
			ds.store_code,
			'CM' Fg, 
			COUNT(distinct a.SALES_NUMBER) SALES_NUMBER
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate and edate
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and CASE WHEN a.MEMBER_TYPE = 'member'
			 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
			 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
			 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
		 	THEN 'New Member'
		 	WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
		 	WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
			ELSE 'Existing' end = 'Existing'
	group by d.date, ds.store_code
	union 
	-- Last Month
	SELECT 	d.date, 
			ds.store_code,
			'LM' Fg,
			COUNT(distinct a.SALES_NUMBER) SALES_NUMBER
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate_lm and edate_lm
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and CASE WHEN a.MEMBER_TYPE = 'member'
			 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
			 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
			 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
		 	THEN 'New Member'
		 	WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
		 	WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
			ELSE 'Existing' end = 'Existing'
	group by d.date, ds.store_code;

	DROP temporary TABLE IF EXISTS slsgab_b; 
	create temporary table slsgab_b
	select * from slsgab;
	
-- 	select d.`date`, x.TRX, x.TRX_LM, t.target
	select date_format(d.`date`, '%Y%m%d') `DATE`, x.TRX,'' L, t.target, '' N, '' O, '' P
	from dim_date d
	left join (
	select 	a.`date`, 
			-- s.store_name, 
			a.store_code,
			SUM(a.SALES_NUMBER) TRX, SUM(b.SALES_NUMBER) TRX_LM
	from slsgab a
	join (select distinct store_code, store_name from dim_store) s on s.store_code = a.store_code
	left join slsgab_b b on b.fg = 'LM' and b.`date` =  (a.`date` - interval 1 month)
	where a.fg = 'CM'
	group by a.`date`, 
			-- s.store_name, 
			a.store_code, 
			a.fg
	) x	on x.`date` = d.`date`
	left join db_ip.StoreDailyTarget t 
			on t.typ = '8' and t.sitecode = store and d.`date` = t.asofdate 
	where d.`date` between sdate and last_day(edate)
	order by 1 asc;
	
	
end if;

-- T = 10 No of New Member
IF T = 10 then 
	
	DROP temporary TABLE IF EXISTS slsgab; 
	create temporary table slsgab 
	-- Current Month
	SELECT 	d.date, 
			ds.store_code,
			'CM' Fg, 
			COUNT(distinct a.CLIENT_ID) CLIENTID
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate and edate
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and CASE WHEN a.MEMBER_TYPE = 'member'
			 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
			 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
			 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
		 	THEN 'New Member'
		 	WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
		 	WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
			ELSE 'Existing' end = 'New Member'
	group by d.date, ds.store_code
	union 
	-- Last Month
	SELECT 	d.date, 
			ds.store_code,
			'LM' Fg,
			COUNT(distinct a.CLIENT_ID) CLIENTID
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate_lm and edate_lm
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and CASE WHEN a.MEMBER_TYPE = 'member'
			 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
			 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
			 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
		 	THEN 'New Member'
		 	WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
		 	WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
			ELSE 'Existing' end = 'New Member'
	group by d.date, ds.store_code;

	DROP temporary TABLE IF EXISTS slsgab_b; 
	create temporary table slsgab_b
	select * from slsgab;
	
-- 	select d.`date`, x.CLIENTID, x.CLIENTID_LM, t.target
	select date_format(d.`date`, '%Y%m%d') `DATE`, x.CLIENTID,'' T, t.target, '' V, '' W, '' X
	from dim_date d
	left join (
	select 	a.`date`, 
			-- s.store_name, 
			a.store_code,
			SUM(a.CLIENTID) CLIENTID, SUM(b.clientid) CLIENTID_LM
	from slsgab a
	join (select distinct store_code, store_name from dim_store) s on s.store_code = a.store_code
	left join slsgab_b b on b.fg = 'LM' and b.`date` =  (a.`date` - interval 1 month)
	where a.fg = 'CM'
	group by a.`date`, 
			-- s.store_name, 
			a.store_code, 
			a.fg
	) x	on x.`date` = d.`date`
	left join db_ip.StoreDailyTarget t 
			on t.typ = '9' and t.sitecode = store and d.`date` = t.asofdate 
	where d.`date` between sdate and last_day(edate)
	order by 1 asc;

end if;

-- T = 11 TRX of Existing Member + New
IF T = 11 then 
	
	DROP temporary TABLE IF EXISTS slsgab; 
	create temporary table slsgab 
	-- Current Month
	SELECT 	d.date, 
			ds.store_code,
			'CM' Fg, 
			COUNT(distinct a.SALES_NUMBER) SALES_NUMBER
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate and edate
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and CASE WHEN a.MEMBER_TYPE = 'member'
			 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
			 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
			 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
		 	THEN 'New Member'
		 	WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
		 	WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
			ELSE 'Existing' end IN ('New Member', 'Non Member')
	group by d.date, ds.store_code
	union 
	-- Last Month
	SELECT 	d.date, 
			ds.store_code,
			'LM' Fg,
			COUNT(distinct a.SALES_NUMBER) SALES_NUMBER
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate_lm and edate_lm
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and CASE WHEN a.MEMBER_TYPE = 'member'
			 AND MONTH(a.CLIENT_CREATION_DATE) = MONTH(a.SALES_DATE) 
			 AND YEAR(a.CLIENT_CREATION_DATE) = YEAR(a.SALES_DATE)
			 AND a.CLIENT_STORE_REGISTRATION = ds.store_name  
		 	THEN 'New Member'
		 	WHEN a.MEMBER_TYPE = 'Other' THEN 'Other'
		 	WHEN a.MEMBER_TYPE = 'Non Member' THEN 'Non Member'
			ELSE 'Existing' end IN ('New Member', 'Non Member')
	group by d.date, ds.store_code;

	DROP temporary TABLE IF EXISTS slsgab_b; 
	create temporary table slsgab_b
	select * from slsgab;
	
-- 	select d.`date`, x.TRX, x.TRX_LM, t.target
select date_format(d.`date`, '%Y%m%d') `DATE`, x.TRX,'' AB, t.target, '' AD, '' AE, '' AF
	from dim_date d
	left join (
	select 	a.`date`, 
			-- s.store_name, 
			a.store_code,
			SUM(a.SALES_NUMBER) TRX, SUM(b.SALES_NUMBER) TRX_LM
	from slsgab a
	join (select distinct store_code, store_name from dim_store) s on s.store_code = a.store_code
	left join slsgab_b b on b.fg = 'LM' and b.`date` =  (a.`date` - interval 1 month)
	where a.fg = 'CM'
	group by a.`date`, 
			-- s.store_name, 
			a.store_code, 
			a.fg
	) x	on x.`date` = d.`date`
	left join db_ip.StoreDailyTarget t 
			on t.typ = '10' and t.sitecode = store and d.`date` = t.asofdate 
	where d.`date` between sdate and last_day(edate)
	order by 1 asc;
	
	
end if;


-- T = 12 TOP 30 SKU by Store
IF T = 12 then 
	
	SELECT 	dpt.ITEM_NAME,  '' A, '' B, '' C,  
			SUM(a.TOTAL_PRICE_AFTER_TAX) Sales, '' E, 
			COUNT(distinct a.SALES_NUMBER) Trx
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	where a.sales_date between sdate and edate
	and a.SALES_STATUS not like 'CANCEL%'
	and case when store is null then 1 else ds.store_code end = case when store is null then 1 else store end
	group by dpt.ITEM_NAME
	order by SALES desc
	limit 30;
	
	
end if;

-- T = 13 CEK-KES By Store by Sales Person
IF T = 13 then 
	
	DROP TABLE IF EXISTS slsgab; 
	create  table slsgab 
	SELECT 	d.date, 
			ds.store_code,
			a.Sales_person,
			case when dpt.ITEM_GROUP like 'URIC ACID%' then 'Asam Urat'
				 when dpt.ITEM_GROUP like 'CHOLESTEROL%' then 'Kolesterol'
				 when dpt.ITEM_GROUP like 'GLUCOSE%' then 'Gula Darah'
				 else 'OTHER'
			end CEKKES,	 
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE,
			count(distinct a.SALES_NUMBER) trx 
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate and edate
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and dpt.ITEM_NAME like 'CEK%'
	group by d.date, ds.store_code, a.Sales_person, case when dpt.ITEM_GROUP like 'URIC ACID%' then 'Asam Urat'
				 when dpt.ITEM_GROUP like 'CHOLESTEROL%' then 'Kolesterol'
				 when dpt.ITEM_GROUP like 'GLUCOSE%' then 'Gula Darah'
				 else 'OTHER'
			end;
			
	DROP TABLE IF EXISTS slsgabpvt; 
	create  table slsgabpvt		
	select 	date, store_code, sales_person, 
			SUM(if (CEKKES = 'Gula Darah', SALES_VALUE, 0)) as SALES_GD,
			SUM(if (CEKKES = 'Gula Darah', trx, 0)) as TRX_GD,
			SUM(if (CEKKES = 'Kolesterol', SALES_VALUE, 0)) as SALES_KO,
			SUM(if (CEKKES = 'Kolesterol', trx, 0)) as TRX_KO,
			SUM(if (CEKKES = 'Asam Urat', SALES_VALUE, 0)) as SALES_AU,
			SUM(if (CEKKES = 'Asam Urat', trx, 0)) as TRX_AU
	from slsgab	
	group by date, store_code, sales_person;

	insert into slsgabpvt
	select 	sdate, store, a.sales_person, 
			cast(0 as decimal) SALES_GD, cast(0 as decimal) TRX_GD, 
			cast(0 as decimal) SALES_KO, cast(0 as decimal) TRX_KO,
			cast(0 as decimal) SALES_AU, cast(0 as decimal) TRX_AU
	from (select distinct a.sales_person 
			from ip_tr_tbl_sales_by_item a 
			join dim_store ds on ds.store_id  = a.store_id
			where ds.store_code = store 
			and a.sales_person not in (select distinct sales_person from slsgabpvt where store_code = store)
			and a.sales_date between (sdate - interval 1 month) and (edate - interval 1 month)
			) a ;
	
	
-- 	total
	DROP TABLE IF EXISTS slsgabtotal; 
	create  table slsgabtotal
	SELECT 	d.date, 
			ds.store_code,
			a.Sales_person,
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE,
			count(distinct a.SALES_NUMBER) trx 
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	right join db_ip.dim_date d on d.date = a.sales_date
	where a.sales_date between sdate and edate
	and a.SALES_STATUS not like 'CANCEL%'
	and ds.store_code = store
	and dpt.ITEM_NAME like 'CEK%'
	group by d.date, ds.store_code, a.Sales_person;
	
	insert into slsgabtotal
	select 	sdate, store, a.sales_person, 
			cast(0 as decimal) SALES_VALUE, cast(0 as decimal) TRX			
	from (select distinct a.sales_person 
			from ip_tr_tbl_sales_by_item a 
			join dim_store ds on ds.store_id  = a.store_id
			where ds.store_code = store 
			and a.sales_person not in (select distinct sales_person from slsgabtotal where store_code = store)
			and a.sales_date between (sdate - interval 1 month) and (edate - interval 1 month)
			) a ;
	

	DROP TABLE IF EXISTS slsgabpvt2; 
	create  table slsgabpvt2
	select a.date, a.Store_code, a.sales_person, a.SALES_VALUE, a.TRX, 
		   b.SALES_GD, b.TRX_GD, b.SALES_KO, b.TRX_KO, b.SALES_AU, b.TRX_AU 
	from slsgabtotal a
	left join slsgabpvt b 
		on a.date = b.DATE and a.sales_person = b.sales_person and a.store_code = b.Store_code;
	
	
	SET SESSION group_concat_max_len=65535;

	SET @sql_dinamis = (
		SELECT
			 GROUP_CONCAT(DISTINCT
				CONCAT('SUM( IF(Sales_person = \''
					, Sales_person 
					, '\',SALES_VALUE,0) ) AS SALES_VALUE_'
					, replace(Sales_person,' ', '_')
					, ', SUM( IF(Sales_person = \''
					, Sales_person
					, '\',TRX,0) ) AS TRX_'
					, replace(Sales_person,' ', '_')
					,', SUM( IF(Sales_person = \''
					, Sales_person
					, '\',SALES_GD,0) ) AS SALES_GD_'
					, replace(Sales_person,' ', '_')
					,', SUM( IF(Sales_person = \''
					, Sales_person
					, '\',TRX_GD,0) ) AS TRX_GD_'
					, replace(Sales_person,' ', '_')
					,', SUM( IF(Sales_person = \''
					, Sales_person
					, '\',SALES_KO,0) ) AS SALES_KO_'
					, replace(Sales_person,' ', '_')
					,', SUM( IF(Sales_person = \''
					, Sales_person
					, '\',TRX_KO,0) ) AS TRX_KO_'
					, replace(Sales_person,' ', '_')
					,', SUM( IF(Sales_person = \''
					, Sales_person
					, '\',SALES_AU,0) ) AS SALES_AU_'
					, replace(Sales_person,' ', '_')
					,', SUM( IF(Sales_person = \''
					, Sales_person
					, '\',TRX_AU,0) ) AS TRX_AU_'
					, replace(Sales_person,' ', '_')
				)
				)
		FROM slsgabpvt2
	);

	drop table if exists fintable;
	SET @SQL = CONCAT('create table fintable ','SELECT d.date, ', 
			  @sql_dinamis, ' FROM slsgabpvt2 a RIGHT JOIN dim_date d 
		on d.date = a.date where d.date between \'', sdate, '\' and \'', edate, '\' GROUP BY d.date'
	   );

	PREPARE stmt FROM @SQL;
	EXECUTE stmt;
	DEALLOCATE PREPARE stmt;
	
	
	drop table if exists fintable2;
	create table fintable2
	select date_format(d.`date`, '%Y%m%d') tgl, a.* 
	from dim_date d 
		left join fintable a on a.date = d.date 
	where d.date between sdate and last_day(sdate) 
	order by 1 asc;

	alter table fintable2
	drop column `date`;
	
	DROP TABLE IF EXISTS fintable;
	DROP TABLE IF EXISTS slsgabpvt; 
	DROP TABLE IF EXISTS slsgabpvt2; 
	DROP TABLE IF EXISTS slsgabtotal;
	DROP TABLE IF EXISTS slsgab; 

	select a.* from fintable2 a;

end if;


-- T=14 list store
IF T = 14 then 
	select distinct store_code, store_name
	from dim_store
	where store_code is not null
	order by store_name; 

end if;

-- T = 15 CEK-KES By Store by Sales Person SUMMARY
IF T = 15 then 
	
-- 	total
	DROP TABLE IF EXISTS slsgabtotal; 
	create  table slsgabtotal
	SELECT  ds.store_code,
			a.Sales_person,
			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE,
			count(distinct a.SALES_NUMBER) trx 
	from db_ip.ip_tr_tbl_sales_by_item a
	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
	join db_ip.dim_store ds on ds.store_id  = a.store_id
	where a.sales_date between sdate and edate
	and a.SALES_STATUS not like 'CANCEL%'
	and dpt.ITEM_NAME like 'CEK%'
	group by ds.store_code, a.Sales_person;
	

	select a.store_name, a.sales_person, IFNULL(b.SALES_VALUE,0) SALES_VALUE, IFNULL(b.TRX,0) TRX 
    from CEKKES_PERSON a
    join dim_store ds on a.store_name = ds.store_name 
    left join slsgabtotal b on b.store_code = ds.store_code and b.sales_person = a.sales_person;
    
    
-- 
-- 
-- 	insert into slsgabtotal
-- 	select x.store_code, x.SALES_PERSON, cast(0 as decimal) SALES_VALUE, CAST(0 as DECIMAL) TRX
-- 	from (
-- 	SELECT  ds.store_code,
-- 			a.Sales_person,
-- 			SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE,
-- 			count(distinct a.SALES_NUMBER) trx 
-- 	from db_ip.ip_tr_tbl_sales_by_item a
-- 	join db_ip.dim_product_tsbi dpt on a.ip_item_id = dpt.ip_item_id
-- 	join db_ip.dim_store ds on ds.store_id  = a.store_id
-- 	where a.SALES_STATUS not like 'CANCEL%'
-- 	and dpt.ITEM_NAME like 'CEK%'
-- 	group by ds.store_code, a.Sales_person) x
-- 	where not exists (select 1 from slsgabtotal y where y.store_code = x.store_code and y.sales_person = x.sales_person);
-- 
-- 	select ds.store_name, j.sales_person, SUM(j.SALES_VALUE) SALES_VALUE, SUM(j.TRX) TRX
-- 	from slsgabtotal j
-- 	join dim_store ds on j.store_code = ds.store_code
-- 	group by ds.store_name, j.sales_person
-- 	order by ds.store_name, j.sales_person;
	
-- 	select ds.store_name, a.sales_person, a.sales_value, a.Trx
-- 	from slsgabtotal a
-- 	join dim_store ds on a.store_code = ds.store_code 
-- 	where a.sales_person not like '%Whatsapp%'
-- 	order by ds.store_name, a.sales_person;	


end if;

end$$

DELIMITER ;



