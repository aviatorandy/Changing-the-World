select l.ID as "Location ID", replace(group_concat(proto_get('profile.Location.StringField',pfd.data,'value')),","," ") as "Provider Name"
from alpha.locations l 
join profiles.location_primary_profiles lpp on lpp.location_id=l.id 
left join profiles.profile_field_data_latest pfd on pfd.profile_id=lpp.profile_id and  (pfd.field_id=119 or pfd.field_id=121)
where  L.id in (@locationIDs)
group by l.id;
