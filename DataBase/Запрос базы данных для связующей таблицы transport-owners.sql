SELECT owners.INN as ИНН, owners.Title as Название, transport.VIN as ВИН, `transport`.`brand` as Марка, `transport`.`State_Registr_Mark` as Номер, `transport`.`Region` as Регион, `transport`.`Ownership` as `Тип владения` FROM transport
	join transport_owners
		on transport.transport_id = transport_owners.transport_id
	join owners 
		on owners.Owner_id = transport_owners.transport_id;