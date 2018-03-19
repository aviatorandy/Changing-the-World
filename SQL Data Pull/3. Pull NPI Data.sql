select concat("'","",l.id) as "Location ID", proto_get("profile.Location.StringField", data, "value") as "Location NPI"
from alpha.locations l
join profiles.location_profiles lp on lp.location_id=l.id
join profiles.profile_field_data_latest pfdl on pfdl.profile_id=lp.profile_id and pfdl.field_id=110
where l.business_id=@bizid;