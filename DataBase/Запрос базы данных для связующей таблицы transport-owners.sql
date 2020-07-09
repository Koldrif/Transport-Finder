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
	owners.INN = '7802519724';

UPDATE
	`transportfinder`.`transport`
SET
	`VIN` = 'ВИН',
	`State_Registr_Mark` = 'Регистрационный знак',
	`Region` = 'Регион',
	`Date_of_issue` = 'Дата выпуска',
	`pass_ser` = 'Серия паспорта',
	`Ownership` = 'Тип владения',
	`brand` = 'Брэнд',
	`type` = 'Вид',
	`Registred_at` = 'Зарегестрировано',
	`License number` = 'Номери лицензии'
WHERE
	(`transport_id` = 'ID трансопрта');