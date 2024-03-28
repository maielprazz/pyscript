DELIMITER $$

CREATE  PROCEDURE `db_ip`.`SP_getStockHistoryV2`(
	in  T INT, -- type parameter 
	IN  stockdate date)
begin
	if IFNULL(stockdate,'19700101') = '19700101' then
		set stockdate = (select max(stock_date) mxdat from ip_tr_tbl_stock_history);
	end if;
-- Raw data 1	
	IF T = 1 then 
		-- select 1;
	select a.STOCK_DATE `Stock Date`, ds.STORE_NAME `Store`, 
		   b.ITEM_CODE `Item Code`, b.ITEM_NAME `Item Name`,
		   IFNULL(mcm.group_category, 'Unknown') `Group Category`,
		   -- b.Category `Group Category`, 
		   b.ITEM_DIVISION `Item Division`,
		   b.Item_Group `Item Group`, 
		   a.NON_FUNCTIONAL_25, 
		   b.Brand, b.`Type`,
		   SUM(a.QTY_OH) QTY_OH, SUM(a.Stock_Value) Stock_Value
		   ,CASE WHEN ds.STORE_NAME like '%Ecom%' THEN 'Marketplace' else 'Store' end CHANNEL
	from ip_tr_tbl_stock_history a
	join dim_product_stockH b on a.ip_item_id = b.ip_item_id
	join dim_store ds on ds.store_id = a.store_id
	left join db_ip.mst_group_category_cm mcm on mcm.PIC_CM = a.PIC_CM
	where STOCK_DATE = stockdate
	group by a.STOCK_DATE, ds.STORE_NAME, 
		   b.ITEM_CODE, b.ITEM_NAME,
		   -- b.Category, 
		   IFNULL(mcm.group_category, 'Unknown'),
		   b.ITEM_DIVISION ,
		   b.Item_Group, 
		   a.NON_FUNCTIONAL_25, 
		   b.Brand, b.`Type`,
		   CASE WHEN ds.STORE_NAME like '%Ecom%' THEN 'Marketplace' else 'Store' end ;


	end if;

end$$

DELIMITER ;