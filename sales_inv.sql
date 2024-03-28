CREATE TABLE ip_gcat1 AS
SELECT 'Nida' PIC_CM, 'Pharma' GCAT1
UNION SELECT 'Irma' PIC_CM, 'Beauty & GMS' GCAT1
UNION SELECT 'Winey' PIC_CM, 'Health Nutrition & Support' GCAT1
UNION SELECT 'Putri' PIC_CM, 'Health Vit. Supplement' GCAT1

SELECT * FROM ip_gcat1

CREATE INDEX idx_ip_gcat1 ON ip_gcat1 (PIC_CM);
CREATE INDEX idx_ip_sbi ON ip_gcat1 (PIC_CM);

CREATE TABLE ip_store_am AS
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

CREATE INDEX idx_ip_store_am ON ip_store_am (StoreLocation);
CREATE INDEX idx_ip_sbi ON tr_tbl_sales_by_item (CLIENT_ID, STORE_LOCATION, PIC_CM, SALES_STATUS, CLIENT_CREATION_DATE, CLIENT_STORE_REGISTRATION, MEMBER_TYPE)

-- Query from here
SELECT 	a.STORE_LOCATION, 
		IFNULL(am.AM, 'Other') AreaManager, 
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
		IFNULL(gc1.GCAT1, 'Unknown') GroupCategory1,
		a.ITEM_DIVISION , a.ITEM_GROUP, a.ITEM_MODEL, a.NON_FUNCTIONAL_25, a.CHANNEL_TYPE, a.PROMOTION_TYPE,
		a.APPLIED_PROMO, 
		SUM(a.PARENT_QTY) PARENT_QTY, SUM(a.TOTAL_PRICE_AFTER_TAX) SALES_VALUE,
		SUM(a.REBATE) REBATE, SUM(a.TOTAL_GROSS_MARGIN_ACCURATE) TOTAL_GROSS_MARGIN_ACCURATE
FROM tr_tbl_sales_by_item  a
LEFT JOIN tr_tbl_customer ttc ON ttc.CODE = a.CLIENT_ID 
LEFT JOIN ip_store_am am ON am.StoreLocation = a.STORE_LOCATION 
LEFT JOIN ip_gcat1 gc1 ON gc1.PIC_CM = a.PIC_CM
WHERE a.SALES_DATE BETWEEN '20230101' AND '20231231'
AND a.Store_location <> 'Apotek Wellings Erajaya Plaza'
AND a.SALES_STATUS not like 'CANCEL%'
GROUP BY a.STORE_LOCATION,
		IFNULL(am.AM, 'Other'),
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
		IFNULL(gc1.GCAT1, 'Unknown'),
		a.ITEM_DIVISION , a.ITEM_GROUP, a.ITEM_MODEL, a.NON_FUNCTIONAL_25, 
		a.CHANNEL_TYPE, a.PROMOTION_TYPE, a.APPLIED_PROMO


SELECT promo_1, promo_2, `type`
FROM xilnex_rpt.ms_promo;

SELECT ID_BATCH, CHANNEL_ID, LOG_DATE, TRANSNAME, STEPNAME, LINES_READ, LINES_WRITTEN, LINES_INPUT, LINES_UPDATED, LINES_OUTPUT, LINES_REJECTED, ERRORS, `RESULT`, NR_RESULT_ROWS, NR_RESULT_FILES, LOG_FIELD, COPY_NR
FROM xilnex.tr_log_pentaho;


select * 
from tr_tbl_sales_by_item
limit 100;


SELECT COUNT(*) FROM tr_tbl_sales_by_item



CREATE TABLE ip_sales_inv AS
select sal.*, IFNULL(tts.AVAILABLE_STOCK,0) StockQty, IFNULL(tts.VALUE_AVAILABLE_STOCK,0) StockAmt
FROM (SELECT 	a.STORE_LOCATION, 
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
FROM tr_tbl_sales_by_item  a
LEFT JOIN tr_tbl_customer ttc ON ttc.CODE = a.CLIENT_ID 
WHERE a.SALES_DATE BETWEEN '20240101' AND '20240131'
AND a.Store_location <> 'Apotek Wellings Erajaya Plaza'
AND a.SALES_STATUS not like 'CANCEL%'
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

	
select * from MS_TARGET mt 
select * from MS_TARGET_DAILY mt 


-- select ITEM_CODE, STORE, AVAILABLE_STOCK , VALUE_AVAILABLE_STOCK  
-- from tr_tbl_stock
	
SELECT PT, `Kode Pelanggan`, `Nama Pelanggan`, Tanggal, `KODE MATA UANG`, KURS, SALES_NO, `Diskon Faktur %`, `Diskon Faktur Rp`, `Kode Barang`, `Nama Barang`, BRAND, `Item Category`, `Item Category 2`, Kuantitas, Satuan, Harga, `Diskon Barang %`, `Diskon Barang Rp`, `Total Harga`, Profit_Margin, Gudang, `Pejual/Salesman`, Catatan, Pajak, `Total Termasuk Pajak`, `Kode Dokumen`, `Kode Transaksi`, Cabang, `Supplier Type`, `TIMESTAMP`
FROM xilnex_rpt.tr_tbl_sales;


SELECT promo_1, promo_2, `type`
FROM xilnex_rpt.ms_promo;


SELECT ID, PROMOGROUPNAME, DATESTART, DATEEND, LOCATION_ID, WBDELETED, UPDATE_TIMESTAMP
FROM xilnex.APP_4_PROMOGROUP;

SELECT SALES_NUMBER, SALES_DATE, SALES_TIME, COMPLETED_DATE, RECEPIENT, CUSTOMER_ID, CUSTOMER_NAME, CLIENT_TYPE, TOTAL_SALES_GROSS_AMOUNT, TOTAL_SALES_NET_AMOUNT, `TOTAL_TAX_%`, TOTAL_TAX, TOTAL_DISCOUNT, DISCOUNT_PO, TOTAL_COST, FINAL_COST, TOTAL_PROFIT_MARGIN, TOTAL_PAID_AMOUNT, TOTAL_BALANCE_AMOUNT, CASHIER_NAME, SALES_STATUS, PAYMENT_STATUS, SALES_PERSON, STORE_LOCATION, TOTAL_BASKET, INCLUDE_TAX, CHANNEL, PAYMENT_REMARK, `TIMESTAMP`
FROM xilnex_rpt.tr_tbl_sales_summary;


SELECT COMPANY, `Kode Pelanggan`, `Nama Pelanggan`, Tanggal, `KODE MATA UANG`, KURS, SALES_NO, `Diskon Faktur %`, `Diskon Faktur Rp`, `Kode Barang`, `Nama Barang`, Kuantitas, `Serial Number Prefix`, `Serial Number Range Start`, `Serial Number Kuantitas`, Satuan, Harga, `Diskon Barang %`, `Diskon Barang Rp`, `Total Harga`, Departemen, Projek, Gudang, `Pejual/Salesman`, `NOMOR PENAWARAN PENJUALAN`, `NOMOR PESANAN PENJUALAN`, `NOMOR PENGIRIMAN PESANAN`, Catatan, `Alamat Kirim`, Pajak, `Total Termasuk Pajak`, `Kode Dokumen`, `Kode Transaksi`, `Tgl Faktur Pajak`, `No. Faktur Pajak`, `Tgl Pengiriman`, Pengiriman, `Cabang*`, `Syarat Pembayaran`, FOB, Keterangan, `Kode Akun`, `Nama Akun`, Jumlah, Departemen2, Projek2, `Catatan Akun`, `Nomor Faktur Uang Muka`, `Jumlah Uang Muka`, `Karakter 1`, `Karakter 2`, `Karakter 3`, `Karakter 4`, `Karakter 5`, `Karakter 6`, `Karakter 7`, `Karakter 8`, `Karakter 9`, `Karakter 10`, `Angka 1`, `Angka 2`, `Angka 3`, `Tanggal 1`, `Tanggal 2`, `TIMESTAMP`
FROM xilnex.tr_sales;


SELECT VBELN, `DATE`, PO_NO
FROM xilnex.tr_sap_vbak;

SELECT VBELV, VBELN, VBTYP_N
FROM xilnex.tr_sap_vbfa;
