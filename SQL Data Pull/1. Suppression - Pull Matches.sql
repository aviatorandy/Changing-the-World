Select concat(l.id, "-", m.listing_id) as "Link ID", l.corporateCode as "Store ID", l.id as "Location ID",  "" as "Folder", l.country as "Country",
l.name as "Location Name", l.address as "Location Address", l.address2 as "Location Address 2", l.city as "Location City", l.state as "Location State",
l.zip as "Location Zip", l.phone as "Location Phone", l.localPhone as "Location Local Phone", l.yextDisplayLatitude as "Location Latitude", l.yextDisplayLongitude as "Location Longitude",
concat("'", "", m.listing_id) as "Listing ID", m.matchType as "SQL Override", m.powerListing as "SQL PL Status"
from alpha.locations l
join matches.match_sets_latest msl on msl.location_id=l.id
join matches.matches m on m.matchSet_id=msl.matchSet_id
--left join alpha.location_tree_nodes ltn on ltn.id=l.treeNode_id
--left join alpha.location_labels ll on ll.location_id=l.id
WHERE l.dashboardDeleteDate is null
and (m.powerListing="NoPowerListing" and m.matchType="match")
and l.business_id=@bizid
--and l.treeNode_id=@folderid
--and ll.label_id=@labelid
and msl.partner_id not in (715)
and msl.partner_id
IN (
579, 711, 550, 716, 512, 698, 697, 657, 692, 598, 632, 592, 545, 402, 524, 656, 661, 638, 597, 735, 547, 724, 551, 736, 733, 559, 768, 523, 392, 647, 773, 563, 713, 764, 720, 644, 645, 662, 732, 395, 761, 516, 665, 420, 596, 398, 590, 591, 631, 595, 688, 693, 669, 670, 611, 671, 626, 672, 673, 674, 599, 612, 618, 621, 613, 675, 620, 580, 676, 677, 614, 678, 622, 679, 615, 616, 680, 681, 682, 617, 683, 623, 624, 619, 684, 625, 627, 600, 685, 686, 667, 655, 699, 634, 633, 635, 758, 721, 560, 578, 700, 767, 636, 717, 469, 629, 628, 433, 434, 719, 502, 583, 543, 668, 660, 581, 510, 690, 562, 499, 594, 505, 544, 553, 601, 576, 707
)
--
--713, --add if suppressing on google
--559, --add if suppressing on fb
-- 583, -- 2findlocal
-- 547, -- 8coupons
-- 579, -- ABLocal
-- 512, -- Avantar
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
-- 720, -- Hotfrog
-- 644, -- iBegin
-- 645, -- iGlobal
-- 662, -- Insider Pages
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
-- 544, -- YellowPageCity.com
-- 576, -- YellowPagesGoesGreen
-- 553, -- YP.com
-- 544, -- YPC
-- 576, -- YPGG
-- 661, -- Credibility Review
-- 644, -- iBegin
-- 656, -- Credibility.com
-- 690, -- Yalwa 
-- 692, -- Brownbook.net
-- 662, -- Insider Pages
-- 696, -- 192.com
-- 716, -- AroundMe
-- 697, -- Bizwiki.com
-- 732, --) Kudzu
-- 
--
----International Publishers
-- 698, -- Bizwiki.co.uk
-- 598, -- BundesTelefonbuch	
-- 632, -- Busqueda-local (Marktplatz)
-- 692, -- Brownbook.net	
-- 638, -- Cylex	
-- 597, -- Dialo
-- 647, -- GoldenPages.ie
-- 645, -- iGlobal
-- 644, -- iBegin	
-- 596, -- Marktplatz-Mittelstand	
-- 605, -- Meinestadt	
-- 631, -- MyTown.ie	
-- 688, -- Najisto.cz	
-- 595, -- n49	
-- 667, -- Oeffnungszeiten.com
-- 611, -- Opendi Belgium (BE)	
-- 626, -- Opendi Canada (CA)	
-- 612, -- Opendi Denmark (DK)	
-- 621, -- Opendi Finland (FI)	
-- 613, -- Opendi France (FR)	
-- 620, -- Opendi Hungary (HU)	
-- 614, -- Opendi Italy (IT)	
-- 622, -- Opendi Luxembourg (LU)	
-- 615, -- Opendi Netherlands (NL)	
-- 616, -- Opendi Norway (NO)	
-- 617, -- Opendi Poland (PL)	
-- 623, -- Opendi Portugal (PT)	
-- 624, -- Opendi Romania (RO)	
-- 625, -- Opendi Slovenia (SI)	
-- 618, -- Opendi Spain (ES)	
-- 619, -- Opendi Sweden (SE)	
-- 627, -- Opendi Turkey)	
-- 600, -- Opendi United Kingdom (GB)	
-- 580, -- Opendi United States (US)
-- 634, -- Pages24.ch (Marktplatz)
-- 633, -- Pages24.fr (Marktplatz)
-- 635, -- Pages24 (NL)	
-- 594, -- Yellowmap	
-- 502, -- Tupalo
-- 636, -- Ricercare Imprese
-- 652, -- Zlat�Str�nky.cz (Manual)
-- 659, -- Zlat�Str�nky.sk
----667,-- Oeffnungszeiten.com, not yet but coming soon after launch
-- 668, -- vebidooBIZ
-- 660, -- Voradius
-- 606, -- Yellow Pages T�rkiye
-- 669, -- Opendi AR
-- 670, -- Opendi AU
-- 671, -- Opendi BR
-- 672, -- Opendi CL
-- 673, -- Opendi CN
-- 674, -- Opendi CO
-- 675, -- Opendi HK
-- 676, -- Opendi ID
-- 677, -- Opendi IN
-- 678, -- Opendi JP
-- 680, -- Opendi NZ
-- 681, -- Opendi PA
-- 682, -- Opendi PE
-- 683, -- Opendi PR
-- 684, -- Opendi SG
-- 685, -- Opendi VE
-- 686, -- Opendi ZA
-- 688, -- Najisto.cz
-- 690, -- Yalwa 
-- 692, -- Brownbook.net
-- 700, -- Profile Canada
-- 699, -- Ourbis.CA
-- 724, -- eirphonebook
-- 719) -- Torget
-- 
;