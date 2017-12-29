Select concat(l.id, "-", m.listing_id) as "Link ID", l.corporateCode as "Store ID", l.id as "Location ID",  "" as "Folder", 
Case l.country 
when "GB" then "GB"
when "DE" then "DE"
else "US" end
as "Country",
l.name as "Location Name", l.address as "Location Address", l.address2 as "Location Address 2", l.city as "Location City", l.state as "Location State",
l.zip as "Location Zip", l.phone as "Location Phone", l.localPhone as "Location Local Phone", l.yextDisplayLatitude as "Location Latitude", l.yextDisplayLongitude as "Location Longitude",
concat("'", "", m.listing_id) as "Listing ID", m.matchType as "SQL Override", m.powerListing as "SQL PL Status"
from alpha.locations l
join matches.match_sets_latest msl on msl.location_id=l.id
join matches.matches m on m.matchSet_id=msl.matchSet_id
--left join alpha.location_tree_nodes ltn on ltn.id=l.treeNode_id
--left join alpha.location_labels ll on ll.location_id=l.id
where l.business_id=@bizid
and l.dashboardDeleteDate is null
--and l.treeNode_id=@folderid
--and ll.label_id=@labelid
and msl.partner_id not in (715)
--and msl.partner_id in (559,...)

--INCLUDE FOR SUPPRESSION BACKUP
--and not (m.powerListing="sync" and m.matchType="match")
--and msl.partner_id in (
-- 713, -- Google Reviews 
-- 559, -- Facebook
-- 583, -- 2findlocal
-- 547, -- 8coupons
-- 579, -- ABLocal
-- 716, -- AroundMe
-- 512, -- Avantar
-- 697, -- Bizwiki.com
-- 692, -- Brownbook.net
-- 545, -- ChamberofCommerce
-- 402, -- Citysearch (Manual)
-- 524, -- Citysquares
-- 521, -- CoPilot
-- 656, -- Credibility.com
-- 661, -- Credibility Review
-- 638, -- Cylex
-- 551, -- eLocal
-- 733, -- EZLocalV2
-- 392, -- GetFave
-- 563, -- GoLocal 24/7
-- 720, -- HotFrog
-- 644, -- iBegin
-- 645, -- iGlobal
-- 662, -- Insider Pages
-- 732, -- Kudzu
-- 395, -- Local.com
-- 761, -- LocalDatabase
-- 516, -- LocalPages
-- 420, -- Mapquest
-- 398, -- MerchantCircle
-- 665, -- Mojopages
-- 590, -- My Local Services 
-- 595, -- n49
-- 580, -- Opendi
-- 560, -- Pennysaver
-- 578, -- Pointcom
-- 469, -- ShowMeLocal
-- 433, -- Superpages
-- 434, -- Topix
-- 502, -- Tupalo
-- 543, -- USCity
-- 581, -- VoteForTheBest
-- 523, -- WhereTo?
-- 393, -- Whitepages
-- 690, -- Yalwa 
-- 562, -- YaSabe
-- 499, -- Yellowise
-- 505, -- Yellowmoxie
-- 544, -- YPC
-- 576, -- YPGG
-- 553, -- YP
-- 693, -- 192.com
-- 698, -- Bizwiki.co.uk
-- 692, -- Brownbook.net
-- 598, -- BundesTelefonbuch	
-- 632, -- Busqueda-local (Marktplatz)	
-- 638, -- Cylex	
-- 597, -- Dialo
-- 724, -- eirphonebook
-- 720,	-- HotFrog
-- 645, -- iGlobal
-- 644, -- iBegin
-- 646, -- Infobel	
-- 596, -- Marktplatz-Mittelstand	
-- 605, -- Meinestadt	
-- 631, -- MyTown.ie		
-- 595, -- n49	
-- 688, -- Najisto.cz
-- 667, -- Oeffnungszeiten.com  
-- 669, -- Opendi Argentina (AR)
-- 670, -- Opendi Australia (AU)
-- 611, -- Opendi Belgium (BE)	
-- 671, -- Opendi Brazil (BR)
-- 626, -- Opendi Canada (CA)	
-- 672, -- Opendi Chile (CL)
-- 673, -- Opendi China (CN)
-- 674, -- Opendi Colombia (CO)
-- 612, -- Opendi Denmark (DK)	
-- 621, -- Opendi Finland (FI)	
-- 613, -- Opendi France (FR)	
-- 675, -- Opendi HK
-- 620, -- Opendi Hungary (HU)	
-- 677, -- Opendi India (IN)
-- 676, -- Opendi Indonesia (ID)
-- 614, -- Opendi Italy (IT)
-- 678, -- Opendi Japan (JP)
-- 622, -- Opendi Luxembourg (LU)	
-- 615, -- Opendi Netherlands (NL)	
-- 616, -- Opendi Norway (NO)	
-- 680, -- Opendi New Zealand (NZ)
-- 681, -- Opendi Panama (PA)
-- 682, -- Opendi Peru (PE)
-- 617, -- Opendi Poland (PL)	
-- 683, -- Opendi Puerto Rico (PR)
-- 623, -- Opendi Portugal (PT)	
-- 624, -- Opendi Romania (RO)	
-- 625, -- Opendi Slovenia (SI)	
-- 618, -- Opendi Spain (ES)	
-- 619, -- Opendi Sweden (SE)	
-- 684, -- Opendi Singapore (SG)
-- 627, -- Opendi Turkey (TR)	
-- 600, -- Opendi United Kingdom (GB)	
-- 580, -- Opendi United States (US)
-- 685, -- Opendi Venezuala (VE)
-- 686, -- Opendi South Africa (ZA)
-- 699, -- Ourbis.CA
-- 635, -- Pages24 (NL)
-- 700, -- Profile Canada
-- 636, -- Ricercare Imprese
-- 502, -- Tupalo
-- 719, -- Torget
-- 668, -- vebidooBIZ
-- 690, -- Yalwa 
-- 594, -- Yellowmap
-- 707, -- YourLocal.IE
-- 652, -- ZlatéStránky.cz
-- 659) -- ZlatéStránky.sk