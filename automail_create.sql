select * from test t

show engines

CREATE TABLE REF_AUTOMAIL (
    report_id INT NULL,
    report_name VARCHAR(50) NULL,
    subject VARCHAR(100) null,
    email_to VARCHAR(50) NULL,
    email_cc VARCHAR(50) NULL,
    email_bc VARCHAR(50) NULL,
    status TINYINT(1) default 0
) ENGINE=InnoDB;

CREATE TABLE REF_AUTOMAIL_BODY (
    report_id INT not NULL,
    body TEXT NULL
) ENGINE=InnoDB;


CREATE TABLE PARAM_AUTOMAIL (
    account_no INT not null,
	sender VARCHAR(50) NULL,
    sender_email VARCHAR(50) NULL,
    password VARCHAR(50) null,
    smtp VARCHAR(50) null,
    smtp_port VARCHAR(10) null
) ENGINE=InnoDB;

select * from REF_AUTOMAIL where report_id = 0;
select * from PARAM_AUTOMAIL where account_no = 1;
select * from REF_AUTOMAIL_BODY where report_id = 0


insert into REF_AUTOMAIL_BODY 
select 0, "<h1>Dear All,<br>Berikut report daily stock</h1><br><p>Regards, Team</p>"

SELECT 	REPORT_NAME, Subject, GROUP_CONCAT(email_to order by email_to SEPARATOR ', ') AS email_to,
		GROUP_CONCAT(email_cc order by email_cc SEPARATOR ', ') AS email_cc,
		GROUP_CONCAT(email_bc order by email_bc SEPARATOR ', ') AS email_bc
FROM REF_AUTOMAIL
where REPORT_NAME = 'Daily Stock Report' 
and status = 1
group by REPORT_NAME, Subject
;




update REF_AUTOMAIL
set email_to = 'ismail.prasetyo@erajaya.com'
where report_id = 1

insert into PARAM_AUTOMAIL
select 1, 'Automail', 'ismail.prasetyo@erajaya.com','Instrument@24', 'smtp.erajaya.com', '587'

insert into REF_AUTOMAIL
select 0, 'Daily Stock Report', 'Daily Stock Report', null, 'ismail.prasetyo@map.co.id', null, 1



SELECT body
FROM db_ip.ref_automail_body
where report_id = 0;

update ref_automail_body
set body = 'Dear All,<br>Berikut report daily stock<br><p>Regards, Team</p>'
where report_id = 0;

SELECT account_no, sender, sender_email, password, smtp, smtp_port
FROM db_ip.param_automail;

SELECT report_id, report_name, subject, email_to, email_cc, email_bc, status
FROM db_ip.ref_automail;

