class Owner:
    def __init__(self, **arguments):
        self.inn = self.arguments['inn']
        self.ogrn = self.arguments['ogrn']
        self.company = self.arguments['company']
        self.registered_at = self.arguments['registered_at']
        self.license_number = self.arguments['license_number']
        self.reg_address = self.arguments['reg_address']
        self.implement_address = self.arguments['implement_address']
        self.risk_category = self.arguments['risk_category']
        self.inspect_start = self.arguments['inspect_start']
        self.inspect_duration = self.arguments['inspect_duration']
        self.last_inspect = self.arguments['last_inspect']
        self.purpose_of_inspect = self.arguments['purpose_of_inspect']
        self.other_reason_of_inspect = self.arguments['other_reason_of_inspect']
        self.form_of_holding_inspect = self.arguments['form_of_holding_inspect']
        self.inspect_perform = self.arguments['inspect_perform']
        self.punishment = self.arguments['punishment']
        self.description = self.arguments['description']