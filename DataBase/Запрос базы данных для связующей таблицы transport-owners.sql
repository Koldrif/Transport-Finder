SELECT
	owners.INN as ИНН,
	owners.Title as Название,
	transport.transport_id as 'ID Тс',
	transport.VIN as ВИН,
	`transport`.`brand` as Марка,
	`transport`.`State_Registr_Mark` as Номер,
	`transport`.`Region` as Регион,
	`transport`.`Ownership` as `Тип владения`
FROM
	transport
	join transport_owners on transport.transport_id = transport_owners.transport_id
	join owners on owners.Owner_id = transport_owners.Owner_id
WHERE
	`transport`.`State_Registr_Mark` = 'B855CT178';

UPDATE
	`transportfinder`.`owners`
SET
	`INN` = 'INN',
	`OGRN` = 'OGRN',
	`Title` = 'Title',
	`Registered_at_date` = 'Registered_at_date',
	`License_number` = 'License_number',
	`Reg_address` = 'Reg_address',
	`Implement_address` = 'Implement_address',
	`Risk_category` = 'Risk_category',
	`Starts_at` = 'Starts_at',
	`Duration_hours` = 'Duration_hours',
	`Last inspec` = 'Last_inspec',
	`Purpose` = 'Purpose',
	`other_reason` = 'Other_reason',
	`form_of_holding` = 'Form_of_holding',
	`Performs_with` = 'Performs_with',
	`Punishment` = 'Punishment',
	`Description` = 'Description'
WHERE
	(`Owner_id` = '1');