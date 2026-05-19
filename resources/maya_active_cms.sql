SELECT
 debtor.`id`,
 debtor.`account`,
 debtor.`name`,
 debtor.`total_debt`,
 debtor.`cycle`
 FROM debtor
 WHERE debtor.`client_id` = 7
 AND debtor.`is_aborted` = 0
 AND debtor.`is_locked` = 0