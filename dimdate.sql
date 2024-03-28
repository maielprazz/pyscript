create table dim_date (
 day_key int not null auto_increment,   -- primay key.
 date datetime ,     -- date: 2008-08-18 00:00:00
 date_num int(8),    -- numeric value, in YYYYMMDD, 20080818
 day_num int (2),    -- numeric value, 18
 day_of_year int(4), -- the day of the year 
 day_of_week int(2), -- the day of the week
 day_of_week_name varchar(20), -- day of week name (Monday, Tuesday,etc)
 week_num int (2), --  week of the year 
 week_begin_date datetime,  -- week begin date
 week_end_date datetime, -- week end date
 last_week_begin_date datetime,  -- priore week begin date
 last_week_end_date datetime,   -- priore week end date
 last_2_week_begin_date  datetime,   -- priore two week begin date
 last_2_week_end_date datetime,  -- priore two ween end date
 month_num int (2) ,  -- month in number, ie. 12
 month_name varchar(20),  -- month in name, ie. December
 YEARMONTH_NUM int(6),  -- year and month in number, ie. 201212
 last_month_num int (2), -- priore month in number, ie. 11
 last_month_name varchar(20), -- priore month in name, November
 last_month_year int(4),  -- priore month in year, 2012
 last_yearmonth_num int(6), -- priore year and month in  number, ie, 2o1211
 quarter_num int (2),  -- quarter in number, ie 4
 year_num int (4), -- year in number, ie, 2012
 created_date timestamp not null  ,  -- date record was created
 updated_date timestamp not null , -- date record was updated
 primary key (day_key)
 );
 

use db_ip;

drop procedure if exists sp_dim_date;

truncate table dim_date;

select * from dim_date

call sp_dim_date('2022-01-01 00:00:00.000', '2024-12-31 00:00:00.000')

delimiter //

CREATE PROCEDURE sp_dim_date (in p_start_date datetime, p_end_date datetime)
BEGIN

  Declare StartDate datetime;
  Declare EndDate datetime;
  Declare RunDate datetime;

-- Set date variables

  Set StartDate = p_start_date; -- update this value to reflect the earliest date that you will use.
  Set EndDate = p_end_date; -- update this value to reflect the latest date that you will use.
  Set RunDate = StartDate;

-- Loop through each date and insert into DimTime table

WHILE RunDate <= EndDate DO

INSERT Into dim_date(
 date ,
 date_num,
 day_num ,
 Day_of_Year,
 Day_of_Week,
 Day_of_week_name,
 Week_num,
 week_begin_date,
 week_end_date,
 last_week_begin_date,
 last_week_end_date,
 last_2_week_begin_date,
 last_2_week_end_date,
 Month_num ,
 Month_Name,
 yearmonth_num,
 last_month_num,
 last_month_name,
 last_month_year,
 last_yearmonth_num,
 Quarter_num ,
 Year_num  ,
 created_date, 
 updated_date 
)
select 
RunDate date
,CONCAT(year(RunDate), lpad(MONTH(RunDate),2,'0'),lpad(day(RunDate),2,'0')) date_num
,day(RunDate) day_num
,DAYOFYEAR(RunDate) day_of_year
,DAYOFWEEK(RunDate) day_of_week
,DAYNAME(RunDate) day_of_week_name
,WEEK(RunDate) week_num
,DATE_ADD(RunDate, INTERVAL(1-DAYOFWEEK(RunDate)) DAY) week_begin_date
,ADDTIME(DATE_ADD(RunDate, INTERVAL(7-DAYOFWEEK(RunDate)) DAY),'23:59:59') week_end_date
,DATE_ADD(RunDate, INTERVAL ((1-DAYOFWEEK(RunDate))-7) DAY) last_week_begin_date
,ADDTIME(DATE_ADD(RunDate, INTERVAL ((7-DAYOFWEEK(RunDate))-7) DAY),'23:59:59')last_week_end_date
,DATE_ADD(RunDate, INTERVAL ((1-DAYOFWEEK(RunDate))-14) DAY) last_2_week_begin_date
,ADDTIME(DATE_ADD(RunDate, INTERVAL ((7-DAYOFWEEK(RunDate))-7) DAY),'23:59:59')last_2_week_end_date
,MONTH(RunDate) month_num
,MONTHNAME(RunDate) month_name
,CONCAT(year(RunDate), lpad(MONTH(RunDate),2,'0')) YEARMONTH_NUM
,MONTH(date_add(RunDate,interval -1 month)) last_month_num
,MONTHNAME(date_add(RunDate,interval -1 month)) last_month_name
,year(date_add(RunDate,interval -1 month)) last_month_year
,CONCAT(year(date_add(RunDate,interval -1 month)),lpad(MONTH(date_add(RunDate,interval -1 month)),2,'0')) Last_YEARMONTH_NUM
,QUARTER(RunDate) quarter_num
,YEAR(RunDate) year_num
,now() created_date
,now() update_date
;
-- commit;
-- increase the value of the @date variable by 1 day

Set RunDate = ADDDATE(RunDate,1);

END WHILE;
commit;
END;
//