SET
    SQL_SAFE_UPDATES = 0;

DELETE FROM
    owners;

ALTER TABLE
    owners AUTO_INCREMENT = 0;

SET
    SQL_SAFE_UPDATES = 0;

SET
    SQL_SAFE_UPDATES = 0;

DELETE FROM
    transport;

ALTER TABLE
    transport AUTO_INCREMENT = 0;

SET
    SQL_SAFE_UPDATES = 0;

INSERT INTO
    `transportfinder`.`owners` (
        `INN`,
        `OGRN`,
        `Title`,
        `Registered_at_date`,
        `License_number`,
        `Reg_address`,
        `Implement_address`,
        `Risk_category`,
        `Starts_at`,
        `Duration_hours`,
        `Last inspec`,
        `Purpose`,
        `other_reason`,
        `form_of_holding`,
        `Performs_with`,
        `Punishment`,
        `Description`
    )
VALUES
    (
        'INN',
        'OGRN',
        'Название',
        'Registered_at',
        'License_number',
        'Reg_address',
        'Implement_address',
        'Risk_cat',
        'Starts_at',
        'Duration',
        'Last_Inspect',
        'Purpose',
        'Other_reason',
        'Form_of_holding',
        'Performs_with',
        'Punishment',
        'Description'
    );