SELECT
    `dynamic_value`.`dynamic_value_name` AS 'DPD',
    `leads`.`leads_acctno` AS 'ACCOUNT NUM',
    `leads`.`leads_endo_date` AS 'ENDO DATE',
    `leads`.`leads_chcode` AS 'CHCODE',
    LEADS_USER.`users_username` AS 'TAGGING',
    `leads`.`leads_ob` AS 'OB',
    `leads`.`leads_chname` AS 'CHNAME',
    `leads_status`.`leads_status_name` AS 'STATUS',
    `leads_substatus`.`leads_substatus_name` AS 'SUBSTATUS',
    DATE_FORMAT(`leads_result`.`leads_result_barcode_date`, "%m/%d/%Y") AS 'DATE',
    `users`.`users_username` AS 'LAST TOUCH'
  FROM `bcrm`.`leads`
  LEFT JOIN
    (SELECT
      leads_result_lead,
      MAX(leads_result_id) AS leads_result_id
    FROM
      leads_result
      INNER JOIN leads ON leads.`leads_id` = leads_result.`leads_result_lead`
    WHERE leads.`leads_client_id` = 156
    GROUP BY leads_result.`leads_result_lead`) latest
    ON latest.leads_result_lead = leads.leads_id
  LEFT JOIN `bcrm`.`leads_result` ON (`latest`.`leads_result_id` = `leads_result`.`leads_result_id`)
  LEFT JOIN `bcrm`.`leads_status` ON (`leads_result`.`leads_result_status_id` = `leads_status`.`leads_status_id`)
  LEFT JOIN `bcrm`.`leads_substatus` ON (`leads_result`.`leads_result_substatus_id` = `leads_substatus`.`leads_substatus_id`)
  LEFT JOIN `bcrm`.`users` ON (`leads_result`.`leads_result_users` = `users`.`users_id`)
  LEFT JOIN `bcrm`.`users` AS LEADS_USER ON (LEADS_USER.`users_id` = `leads`.`leads_users_id`)
  LEFT JOIN `bcrm`.`dynamic_value` ON (`leads`.`leads_id` = `dynamic_value`.`dynamic_value_lead_id` AND `dynamic_value`.`dynamic_value_dynamic_id` = 4361)
  WHERE `leads`.`leads_client_id` =156 
  ORDER BY `leads_result`.`leads_result_barcode_date` DESC