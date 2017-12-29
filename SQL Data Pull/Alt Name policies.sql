select policies.policyString from sms.policies WHERE
policies.business_id=@bizid
and (policies.policyType_id=1 or policies.policyType_id=7);