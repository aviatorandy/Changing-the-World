select  m.listing_id as "Listing ID", 1 as "User Match"
from matches.matches m
where m.source = 'Yext' 
and m.listing_id in (@listingid) 
and (m.powerlisting = "Sync" 
and m.matchType = "Match");