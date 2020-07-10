SET
    SQL_SAFE_UPDATES = 0;

DELETE FROM
    transport;

ALTER TABLE
    transport AUTO_INCREMENT = 0;

SET
    SQL_SAFE_UPDATES = 0;

INSERT INTO
    `transportfinder`.`transport` (
        `VIN`,
        `State_Registr_Mark`,
        `Region`,
        `Date_of_issue`,
        `pass_ser`,
        `Ownership`,
        `End_date_of_ownership`,
        `brand`,
        `model`,
        `type`,
        `Registred_at`,
        `License number`,
        `Status`,
        `Action_with_vehicle`,
        `Categorized`,
        `Number_of_cat_reg`,
        `Data_in_cat_reg`,
        `ATP`,
        `Model_from_cat_reg`,
        `Owner_from_cat_reg`,
        `Purpose_into_cat_reg`,
        `Category`,
        `Date_of_cat_reg`
    )
VALUES
    (
        'VIN',
        'SRM',
        'Region',
        'Date_of_issue',
        'Pass_ser',
        'Ownership',
        'End_date_of_ownership',
        'Brand',
        'Model',
        'Type',
        'Registred_at',
        'License_number',
        'Status',
        'Action_with_vehicle',
        'Categorized',
        'Number_of_cat_reg',
        'Data_in_cat_reg',
        'ATP',
        'Model_from_cat_reg',
        'Ower_from_cat_reg',
        'Purpose_into_cat_reg',
        'Category',
        'Date_of_cat_reg'
    );