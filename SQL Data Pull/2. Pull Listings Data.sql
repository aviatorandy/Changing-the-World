select  concat("'","",wl.id) as "Listing ID", p.id as "Publisher ID", p.name as "Publisher", concat("'","",wl.externalId) as "External ID", wl.name as "Listing Name", wl.address as "Listing Address", lhf.npi as "Listing NPI", wl.address2 as "Listing Address 2",
wl.city as "Listing City", wl.state as "Listing State", wl.zip as "Listing Zip", wl.phone as "Listing Phone", wl.latitude as "Listing Latitude", wl.longitude as "Listing Longitude",
wl.listingUrl as "Listing URL", IF(wl.description IS NOT NULL,'yes','no') as "Description/ExternalID", IF(wl.websiteurl IS NOT NULL, wl.websiteurl,'no') as "Website/URL",
IF((wl.reviewCount IS NULL OR wl.reviewCount=-1), 0, wl.reviewCount) AS "Review Count", wl.status as "Status", 
IF(wl.id IN (SELECT tl.listingid FROM tags_listings tl WHERE tl.tagsListingStatus_id=6), 1, 0) as "Live Sync",
IF(wl.id IN (SELECT pl.Id FROM plc_suppressed_listings plc join warehouse.listings pl ON pl.externalId=plc.externalId AND pl.partner_id=plc.partner_id), 1, 0) AS "Live Suppress",
IF(wl.isexistingadvertiser=1, "Advertiser", IF(wl.isClaimed=1, "Claimed", 0)) AS "Advertiser/Claimed",
IF(wl.partner_id=559, wl.likeCount, IF(wl.partner_id=440, wl.checkinCount, 0)) AS "Likes/Check-ins",
IF(wl.partner_id=559, STR_TO_DATE(wl.lastPostDate, '%Y%m%d%H%i%s'), 0) AS "Last Post Date"
from warehouse.listings wl 
left join alpha.partners p on p.id=wl.partner_id
left join warehouse.listings_healthcare_fields lhf on lhf.listing_id = wl.id
--left join alpha.tags_listings tl on l.id = tl.location_id

--INCLUDE FOR SUPPRESSIONS
--left join warehouse.listings_additional_google_fields lagf on lagf.listing_id=wl.id
where wl.id in (@ListingIDs)
--and tl.tagslistingstatus_id=6
--INCLUDE FOR SUPPRESSIONS
--AND (lagf.googlePlaceId not in (select tl.externalid from tags_listings tl join tags_listings_unavailable_reasons tlur on tlur.location_id = tl.location_id join alpha.tags_unavailable_reasons tur on tur.id=tlur.tagsunavailablereason_id where tlur.partner_id=715 and tur.showasWarning is false) or lagf.googlePlaceId is null)
;
