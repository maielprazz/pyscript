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








