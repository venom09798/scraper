"""
Leads Scraper — Flask version
Same logic as original. Setup form in browser, scrape runs in background.

Sources : Google Maps + Google / DuckDuckGo fallback
Output  : UK_Leads.xlsx + optional CRM

Run:
  pip install flask pandas openpyxl requests playwright
  playwright install chromium
  python leads_scraper_flask.py
  → http://127.0.0.1:5000
"""
from playwright.sync_api import sync_playwright
from openpyxl.worksheet.datavalidation import DataValidation
from flask import Flask, request, redirect, url_for, flash, send_file, render_template_string, jsonify
from threading import Thread, Lock
from datetime import datetime
from urllib.parse import quote, urlparse, parse_qs, unquote
import pandas as pd
import requests
import shutil
import time
import re
import os
import sys

SERVICES = [
    "freelance web designer",
    "freelance web developer",
    "freelance SEO consultant",
    "freelance social media manager",
    "SEO consultant",
    "social media manager",
    "digital marketing consultant",
    "small web design business",
    "web designer",
    "website design services",
    "online marketing services",
    "SEO services",
    "SEO agency",
    "SEO company",
    "web design company",
    "web development agency",
    "website development company",
    "social media marketing services",
    "social media marketing agency",
    "social media marketing company",
    "digital marketing agency",
]


def _expand(names, suffix):
    """Append a country suffix to every place name that does not already have it."""
    return [n if n.endswith(suffix) else f"{n}{suffix}" for n in names]


UK_LOCATIONS = _expand([
    "London", "Manchester", "Birmingham", "Leeds", "Liverpool", "Bristol", "Sheffield",
    "Newcastle upon Tyne", "Nottingham", "Southampton", "Brighton", "Leicester", "Coventry",
    "Bradford", "Kingston upon Hull", "Stoke-on-Trent", "Wolverhampton", "Derby", "Plymouth",
    "Portsmouth", "Sunderland", "Salford", "Preston", "Peterborough", "Norwich", "Oxford",
    "Cambridge", "York", "Exeter", "Bath", "Lincoln", "Chester", "Canterbury", "Carlisle",
    "Chelmsford", "Chichester", "Colchester", "Doncaster", "Durham", "Ely", "Gloucester",
    "Hereford", "Lancaster", "Lichfield", "Milton Keynes", "Ripon", "Salisbury",
    "Southend-on-Sea", "St Albans", "Truro", "Wakefield", "Wells", "Winchester", "Worcester",
    "Westminster, London",
    "Cardiff", "Swansea", "Newport, Wales", "Wrexham", "Bangor, Wales", "St Asaph, Wales",
    "St Davids, Wales", "Neath", "Port Talbot", "Barry, Wales", "Bridgend", "Merthyr Tydfil",
    "Llanelli", "Caerphilly", "Cwmbran", "Pontypridd", "Rhyl", "Llandudno", "Aberystwyth",
    "Carmarthen", "Haverfordwest", "Colwyn Bay", "Pembroke", "Tenby", "Aberdare", "Ebbw Vale",
    "Pontypool", "Newtown, Wales", "Holyhead", "Abergavenny", "Brecon", "Monmouth",
    "Glasgow", "Edinburgh", "Aberdeen", "Dundee", "Inverness", "Stirling", "Perth, Scotland",
    "Dunfermline", "Paisley", "East Kilbride", "Livingston", "Falkirk", "Ayr", "Kilmarnock",
    "Greenock", "Cumbernauld", "Hamilton, Scotland", "Motherwell", "Kirkcaldy", "Dumfries",
    "Elgin", "Fort William", "Oban", "Stranraer", "Arbroath", "Montrose", "Peterhead",
    "Fraserburgh", "Irvine, Scotland", "Alloa", "Bathgate", "Musselburgh", "Dunbar",
    "Galashiels", "Hawick", "Airdrie", "Coatbridge", "Clydebank", "Cambuslang", "Glenrothes",
    "Belfast", "Londonderry", "Lisburn", "Newry", "Armagh", "Bangor, Northern Ireland",
    "Ballymena", "Craigavon", "Coleraine", "Omagh", "Enniskillen", "Newtownabbey",
    "Portadown", "Lurgan", "Larne", "Carrickfergus", "Downpatrick", "Dungannon", "Cookstown",
    "Strabane", "Limavady", "Ballymoney", "Banbridge", "Antrim", "Holywood", "Newtownards",
    "Reading", "Luton", "Northampton", "Swindon", "Bournemouth", "Blackpool", "Middlesbrough",
    "Huddersfield", "Bolton", "Slough", "Watford", "Warrington", "Telford", "Ipswich",
    "Cheltenham", "Maidstone", "Crawley", "Basingstoke", "Guildford", "Gateshead", "Stockport",
    "Oldham", "Blackburn", "Rotherham", "Barnsley", "Harrogate", "Solihull", "Croydon",
    "Chesterfield", "Basildon", "Worthing", "Eastbourne", "Hastings", "Bedford", "Stevenage",
    "Hemel Hempstead", "High Wycombe", "Woking", "Maidenhead", "Bracknell", "Gillingham",
    "Rochdale", "Burnley", "Wigan", "Halifax", "Dudley", "Walsall", "West Bromwich",
    "Darlington", "Hartlepool", "Stockton-on-Tees", "South Shields", "Grimsby", "Scunthorpe",
    "Mansfield", "Kettering", "Rugby", "Nuneaton", "Redditch", "Crewe", "Macclesfield",
    "Southport", "St Helens", "Birkenhead", "Bury", "Shrewsbury", "Torquay", "Poole",
    "Weymouth", "Taunton", "Yeovil", "Aldershot", "Farnborough", "Epsom",
    "Kingston upon Thames", "Bromley", "Romford", "Ilford", "Harrow", "Ealing", "Hounslow",
    "Richmond upon Thames", "Sutton, London", "Barnet", "Enfield", "Hackney, London",
    "Camden, London", "Islington, London", "Shoreditch, London", "Canary Wharf, London",
    "Stratford, London", "Wembley", "Uxbridge", "Wimbledon, London",
    "Accrington", "Altrincham", "Andover", "Ashford", "Aylesbury", "Banbury", "Barking",
    "Barrow-in-Furness", "Bishop's Stortford", "Bognor Regis", "Boston, Lincolnshire",
    "Braintree", "Brentwood", "Bridgwater", "Burton upon Trent", "Bury St Edmunds",
    "Camberley", "Cannock", "Chatham", "Chorley", "Clacton-on-Sea", "Dartford", "Dover",
    "Eastleigh", "Evesham", "Fareham", "Folkestone", "Gosport", "Gravesend", "Hinckley",
    "Horsham", "Huntingdon", "Kidderminster", "King's Lynn", "Leamington Spa", "Leatherhead",
    "Lowestoft", "Loughborough", "Margate", "Newbury", "Newcastle-under-Lyme", "Newquay",
    "Northwich", "Rayleigh", "Redhill", "Royal Tunbridge Wells", "Runcorn", "Scarborough",
    "Sevenoaks", "Sittingbourne", "Skegness", "Stafford", "Stourbridge", "Stroud",
    "Sutton Coldfield", "Tamworth", "Tonbridge", "Trowbridge", "Wellingborough", "Widnes",
    "Worksop", "Yate", "Barnstaple", "Bideford", "Falmouth", "Penzance", "St Ives, Cornwall",
    "Newton Abbot", "Paignton", "Kendal", "Keswick", "Whitby", "Beverley", "Bridlington",
    "Hexham", "Morpeth", "Alnwick", "Berwick-upon-Tweed", "Cleethorpes", "Grantham",
    "Newark-on-Trent", "Stamford", "Oakham", "Melton Mowbray", "Banbury", "Witney",
    "Abingdon", "Didcot", "Henley-on-Thames", "Marlow", "Windsor", "Staines", "Egham",
    "Dorking", "Reigate", "Redditch", "Bromsgrove", "Malvern", "Ludlow", "Oswestry",
], ", UK")

AUSTRALIA_LOCATIONS = _expand([
    "Sydney", "Melbourne", "Brisbane", "Perth", "Adelaide", "Canberra", "Gold Coast",
    "Newcastle, NSW", "Wollongong", "Hobart", "Darwin", "Geelong", "Townsville", "Cairns",
    "Toowoomba", "Ballarat", "Bendigo", "Sunshine Coast", "Launceston", "Mackay",
    "Rockhampton", "Bunbury", "Wagga Wagga", "Albury", "Bundaberg", "Hervey Bay",
    "Gladstone, Queensland", "Mildura", "Shepparton", "Warrnambool", "Dubbo", "Orange, NSW",
    "Tamworth", "Port Macquarie", "Coffs Harbour", "Lismore", "Armidale", "Bathurst", "Nowra",
    "Queanbeyan", "Gympie", "Maryborough, Queensland", "Mount Gambier", "Whyalla",
    "Port Augusta", "Geraldton", "Kalgoorlie", "Albany, Western Australia", "Busselton",
    "Mandurah", "Alice Springs", "Palmerston, Northern Territory", "Devonport, Tasmania",
    "Burnie", "Wodonga", "Traralgon", "Frankston", "Dandenong", "Parramatta", "Penrith",
    "Liverpool, NSW", "Campbelltown, NSW", "Blacktown", "North Sydney", "Bondi, Sydney",
    "Logan, Queensland", "Ipswich, Queensland", "Redcliffe", "Wangaratta", "Horsham",
    "Griffith, NSW", "Goulburn", "Cessnock", "Maitland, NSW", "Gosford", "Wyong",
    "Southport, Queensland", "Surfers Paradise", "St Kilda, Melbourne", "Richmond, Melbourne",
    "Fremantle", "Bankstown", "Hornsby", "Chatswood", "Manly", "Sutherland", "Cronulla",
    "Camden, NSW", "Windsor, NSW", "Katoomba", "Bowral", "Wollondilly", "Kiama", "Ulladulla",
    "Batemans Bay", "Bega", "Cooma", "Tumut", "Young, NSW", "Cowra", "Parkes", "Forbes",
    "Mudgee", "Muswellbrook", "Singleton", "Taree", "Forster", "Kempsey", "Grafton",
    "Ballina", "Byron Bay", "Tweed Heads", "Murwillumbah", "Casino, NSW", "Inverell",
    "Moree", "Narrabri", "Gunnedah", "Broken Hill", "Bairnsdale", "Sale, Victoria",
    "Warragul", "Wonthaggi", "Sunbury", "Melton", "Werribee", "Pakenham", "Cranbourne",
    "Ringwood", "Box Hill", "Doncaster, Melbourne", "Essendon", "Footscray", "Brunswick",
    "Carlton, Melbourne", "Fitzroy, Melbourne", "Prahran", "Brighton, Melbourne", "Hawthorn",
    "Camberwell", "Portland, Victoria", "Hamilton, Victoria", "Colac", "Swan Hill", "Echuca",
    "Benalla", "Wallan", "Kilmore", "Maryborough, Victoria", "Castlemaine", "Daylesford",
    "Caboolture", "Redland Bay", "Cleveland, Queensland", "Beenleigh", "Nambour", "Maroochydore",
    "Caloundra", "Noosa", "Emerald, Queensland", "Charters Towers", "Mount Isa", "Bowen",
    "Proserpine", "Airlie Beach", "Warwick, Queensland", "Kingaroy", "Roma, Queensland",
    "Dalby", "Goondiwindi", "Yeppoon", "Biloela", "Mandurah", "Rockingham", "Joondalup",
    "Armadale, Western Australia", "Midland, Western Australia", "Karratha", "Port Hedland",
    "Broome", "Esperance", "Northam", "Collie", "Margaret River", "Murray Bridge",
    "Victor Harbor", "Port Lincoln", "Elizabeth, South Australia", "Salisbury, South Australia",
    "Glenelg", "Mount Barker, South Australia", "Gawler", "Clare, South Australia",
    "Ulverstone", "Glenorchy, Tasmania", "Kingston, Tasmania", "Katherine, Northern Territory",
    "Nhulunbuy", "Tennant Creek", "Gungahlin", "Tuggeranong", "Belconnen", "Woden Valley",
], ", Australia")

CANADA_LOCATIONS = _expand([
    "Toronto", "Montreal", "Vancouver", "Calgary", "Edmonton", "Ottawa", "Winnipeg",
    "Quebec City", "Hamilton, Ontario", "Kitchener", "London, Ontario",
    "Victoria, British Columbia", "Halifax", "Oshawa", "Windsor, Ontario", "Saskatoon", "Regina",
    "St. John's, Newfoundland", "Mississauga", "Brampton", "Surrey, British Columbia", "Laval",
    "Markham", "Barrie", "Kelowna", "Abbotsford", "Sherbrooke", "Fredericton", "Moncton",
    "Charlottetown", "Saint John, New Brunswick", "Thunder Bay", "Sudbury",
    "Kingston, Ontario", "Guelph", "Cambridge, Ontario", "Waterloo, Ontario", "Niagara Falls",
    "St. Catharines", "Burlington, Ontario", "Oakville", "Richmond Hill", "Vaughan",
    "Pickering", "Ajax, Ontario", "Whitby", "Peterborough, Ontario", "Brantford", "Sarnia",
    "Belleville", "North Bay", "Sault Ste. Marie", "Timmins", "Red Deer", "Lethbridge",
    "Medicine Hat", "Grande Prairie", "Fort McMurray", "Airdrie", "St. Albert", "Nanaimo",
    "Kamloops", "Prince George", "Chilliwack", "Vernon, British Columbia", "Burnaby",
    "Richmond, British Columbia", "Coquitlam", "Langley, British Columbia",
    "Delta, British Columbia", "Maple Ridge", "Trois-Rivieres", "Gatineau", "Longueuil",
    "Saguenay", "Levis", "Drummondville", "Saint-Jerome", "Brandon", "Moose Jaw",
    "Prince Albert", "Yellowknife", "Whitehorse", "Iqaluit", "Sydney, Nova Scotia", "Dartmouth",
    "Scarborough, Toronto", "North York, Toronto", "Etobicoke, Toronto",
    "Newmarket, Ontario", "Aurora, Ontario", "Milton, Ontario", "Georgetown, Ontario",
    "Orangeville", "Stratford, Ontario", "Woodstock, Ontario", "Chatham-Kent", "Welland",
    "Cobourg", "Orillia", "Midland, Ontario", "Collingwood", "Owen Sound", "Cornwall, Ontario",
    "Brockville", "Pembroke, Ontario", "Kenora", "Fort Frances", "Kapuskasing", "Cochrane, Alberta",
    "Okotoks", "Spruce Grove", "Leduc", "Camrose", "Lloydminster", "Cold Lake", "Banff",
    "Canmore", "Squamish", "Whistler", "Penticton", "Salmon Arm", "Cranbrook", "Nelson, British Columbia",
    "Courtenay", "Campbell River", "Duncan, British Columbia", "Port Alberni", "Terrace, British Columbia",
    "Fort St. John", "Dawson Creek", "Quesnel", "Williams Lake", "Mission, British Columbia",
    "White Rock", "Port Coquitlam", "New Westminster", "North Vancouver", "West Vancouver",
    "Steinbach", "Portage la Prairie", "Thompson, Manitoba", "Swift Current", "Yorkton",
    "North Battleford", "Estevan", "Weyburn", "Granby, Quebec", "Saint-Hyacinthe",
    "Shawinigan", "Rimouski", "Victoriaville", "Rouyn-Noranda", "Val-d'Or", "Blainville",
    "Terrebonne", "Repentigny", "Brossard", "Saint-Jean-sur-Richelieu", "Chicoutimi",
    "Truro, Nova Scotia", "Bridgewater, Nova Scotia", "Kentville", "Amherst, Nova Scotia",
    "New Glasgow", "Miramichi", "Bathurst, New Brunswick", "Edmundston", "Summerside",
    "Corner Brook", "Gander", "Grand Falls-Windsor", "Mount Pearl",
], ", Canada")

USA_LOCATIONS = _expand([
    "New York, NY", "Los Angeles, CA", "Chicago, IL", "Houston, TX", "Phoenix, AZ",
    "Philadelphia, PA", "San Antonio, TX", "San Diego, CA", "Dallas, TX", "San Jose, CA",
    "Austin, TX", "Jacksonville, FL", "Fort Worth, TX", "Columbus, OH", "Charlotte, NC",
    "San Francisco, CA", "Indianapolis, IN", "Seattle, WA", "Denver, CO", "Washington, DC",
    "Boston, MA", "El Paso, TX", "Nashville, TN", "Detroit, MI", "Oklahoma City, OK",
    "Portland, OR", "Las Vegas, NV", "Memphis, TN", "Louisville, KY", "Baltimore, MD",
    "Milwaukee, WI", "Albuquerque, NM", "Tucson, AZ", "Fresno, CA", "Sacramento, CA",
    "Mesa, AZ", "Kansas City, MO", "Atlanta, GA", "Omaha, NE", "Colorado Springs, CO",
    "Raleigh, NC", "Miami, FL", "Virginia Beach, VA", "Oakland, CA", "Minneapolis, MN",
    "Tulsa, OK", "Tampa, FL", "Arlington, TX", "New Orleans, LA", "Cleveland, OH",
    "Honolulu, HI", "Pittsburgh, PA", "Cincinnati, OH", "St. Louis, MO", "Orlando, FL",
    "Salt Lake City, UT", "Boise, ID", "Richmond, VA", "Birmingham, AL", "Buffalo, NY",
    "Rochester, NY", "Anchorage, AK", "Charleston, SC", "Providence, RI",
    "Long Beach, CA", "Bakersfield, CA", "Anaheim, CA", "Santa Ana, CA", "Riverside, CA",
    "Stockton, CA", "Irvine, CA", "Chula Vista, CA", "Fremont, CA", "San Bernardino, CA",
    "Modesto, CA", "Fontana, CA", "Oxnard, CA", "Moreno Valley, CA", "Huntington Beach, CA",
    "Glendale, CA", "Santa Clarita, CA", "Garden Grove, CA", "Santa Rosa, CA", "Oceanside, CA",
    "Rancho Cucamonga, CA", "Ontario, CA", "Elk Grove, CA", "Corona, CA", "Lancaster, CA",
    "Palmdale, CA", "Salinas, CA", "Hayward, CA", "Pomona, CA", "Sunnyvale, CA",
    "Escondido, CA", "Torrance, CA", "Pasadena, CA", "Orange, CA", "Fullerton, CA",
    "Roseville, CA", "Thousand Oaks, CA", "Visalia, CA", "Simi Valley, CA", "Concord, CA",
    "Santa Clara, CA", "Vallejo, CA", "Victorville, CA", "Berkeley, CA", "Santa Barbara, CA",
    "San Luis Obispo, CA", "Santa Monica, CA", "Beverly Hills, CA", "Burbank, CA",
    "Carlsbad, CA", "Temecula, CA", "Murrieta, CA", "Menifee, CA", "Costa Mesa, CA",
    "Newport Beach, CA", "Mountain View, CA", "Palo Alto, CA", "San Mateo, CA", "Redding, CA",
    "Chico, CA", "Napa, CA", "Santa Cruz, CA", "Monterey, CA", "Davis, CA", "Folsom, CA",
    "Corpus Christi, TX", "Plano, TX", "Laredo, TX", "Lubbock, TX", "Garland, TX",
    "Irving, TX", "Amarillo, TX", "Grand Prairie, TX", "Brownsville, TX", "McKinney, TX",
    "Frisco, TX", "Pasadena, TX", "Killeen, TX", "McAllen, TX", "Mesquite, TX", "Waco, TX",
    "Carrollton, TX", "Denton, TX", "Midland, TX", "Abilene, TX", "Beaumont, TX",
    "Round Rock, TX", "Odessa, TX", "Wichita Falls, TX", "Richardson, TX", "Lewisville, TX",
    "Tyler, TX", "College Station, TX", "Pearland, TX", "San Angelo, TX", "Allen, TX",
    "League City, TX", "Sugar Land, TX", "The Woodlands, TX", "Katy, TX", "Conroe, TX",
    "Cypress, TX", "Flower Mound, TX", "Georgetown, TX", "San Marcos, TX", "New Braunfels, TX",
    "Longview, TX", "Edinburg, TX", "Harlingen, TX", "Temple, TX", "Bryan, TX",
    "Hialeah, FL", "St. Petersburg, FL", "Port St. Lucie, FL", "Cape Coral, FL",
    "Tallahassee, FL", "Fort Lauderdale, FL", "Pembroke Pines, FL", "Hollywood, FL",
    "Gainesville, FL", "Miramar, FL", "Coral Springs, FL", "Clearwater, FL", "Palm Bay, FL",
    "Pompano Beach, FL", "West Palm Beach, FL", "Lakeland, FL", "Davie, FL",
    "Miami Gardens, FL", "Boca Raton, FL", "Sunrise, FL", "Deltona, FL", "Sarasota, FL",
    "Naples, FL", "Fort Myers, FL", "Daytona Beach, FL", "Melbourne, FL", "Ocala, FL",
    "Pensacola, FL", "Kissimmee, FL", "Jupiter, FL", "Delray Beach, FL", "Homestead, FL",
    "Largo, FL", "Brandon, FL", "Panama City, FL", "Key West, FL",
    "Lexington, KY", "Lincoln, NE", "Greensboro, NC", "Durham, NC", "Winston-Salem, NC",
    "Fayetteville, NC", "Cary, NC", "Wilmington, NC", "Asheville, NC", "Concord, NC",
    "Newark, NJ", "Jersey City, NJ", "Paterson, NJ", "Elizabeth, NJ", "Trenton, NJ",
    "Toms River, NJ", "Princeton, NJ", "Edison, NJ", "Toledo, OH", "Akron, OH", "Dayton, OH",
    "Henderson, NV", "Reno, NV", "North Las Vegas, NV", "Spokane, WA", "Tacoma, WA",
    "Vancouver, WA", "Bellevue, WA", "Kent, WA", "Everett, WA", "Olympia, WA", "Redmond, WA",
    "Aurora, CO", "Fort Collins, CO", "Lakewood, CO", "Thornton, CO", "Arvada, CO",
    "Boulder, CO", "Wichita, KS", "Overland Park, KS", "Kansas City, KS", "Topeka, KS",
    "Des Moines, IA", "Cedar Rapids, IA", "Davenport, IA", "Madison, WI", "Green Bay, WI",
    "Aurora, IL", "Naperville, IL", "Joliet, IL", "Rockford, IL", "Springfield, IL",
    "Fort Wayne, IN", "Evansville, IN", "South Bend, IN", "Grand Rapids, MI", "Warren, MI",
    "Sterling Heights, MI", "Ann Arbor, MI", "Lansing, MI", "Chattanooga, TN", "Knoxville, TN",
    "Clarksville, TN", "Murfreesboro, TN", "Franklin, TN", "Little Rock, AR", "Fayetteville, AR",
    "Jackson, MS", "Mobile, AL", "Montgomery, AL", "Huntsville, AL", "Baton Rouge, LA",
    "Shreveport, LA", "Lafayette, LA", "Savannah, GA", "Augusta, GA", "Columbus, GA",
    "Macon, GA", "Alpharetta, GA", "Marietta, GA", "Columbia, SC", "Greenville, SC",
    "Myrtle Beach, SC", "Norfolk, VA", "Chesapeake, VA", "Newport News, VA", "Hampton, VA",
    "Alexandria, VA", "Arlington, VA", "Roanoke, VA", "Hartford, CT", "New Haven, CT",
    "Stamford, CT", "Bridgeport, CT", "Worcester, MA", "Springfield, MA", "Cambridge, MA",
    "Manchester, NH", "Nashua, NH", "Portland, ME", "Burlington, VT", "Wilmington, DE",
    "Syracuse, NY", "Albany, NY", "Yonkers, NY", "Allentown, PA", "Erie, PA", "Scranton, PA",
    "Harrisburg, PA", "Lancaster, PA", "Tempe, AZ", "Chandler, AZ", "Gilbert, AZ",
    "Glendale, AZ", "Scottsdale, AZ", "Peoria, AZ", "Flagstaff, AZ", "Las Cruces, NM",
    "Santa Fe, NM", "West Valley City, UT", "Provo, UT", "Ogden, UT", "St. George, UT",
    "Billings, MT", "Missoula, MT", "Cheyenne, WY", "Fargo, ND", "Bismarck, ND",
    "Sioux Falls, SD", "Rapid City, SD", "Brooklyn, NY", "Queens, NY", "Bronx, NY",
    "Staten Island, NY", "Manhattan, NY", "Eugene, OR", "Salem, OR", "Bend, OR",
    "Columbia, MO", "Springfield, MO", "Charleston, WV", "Duluth, MN", "Rochester, MN",
    "St. Paul, MN", "Bloomington, MN", "Lexington-Fayette, KY", "Bowling Green, KY",
    "Idaho Falls, ID", "Meridian, ID", "Nampa, ID", "Juneau, AK", "Fairbanks, AK",
    "Hilo, HI", "Kailua, HI", "Burlington, NC", "Baltimore County, MD", "Frederick, MD",
    "Rockville, MD", "Annapolis, MD", "Dover, DE", "Warwick, RI", "Cranston, RI",
], ", USA")

PAKISTAN_LOCATIONS = _expand([
    "Karachi", "Lahore", "Islamabad", "Rawalpindi", "Faisalabad", "Multan", "Peshawar",
    "Quetta", "Gujranwala", "Sialkot", "Hyderabad, Sindh", "Bahawalpur", "Sargodha",
    "Sukkur", "Larkana", "Sheikhupura", "Jhang", "Rahim Yar Khan", "Gujrat", "Mardan",
    "Kasur", "Dera Ghazi Khan", "Sahiwal", "Okara", "Abbottabad", "Mirpur, Azad Kashmir",
    "Muzaffarabad", "Gilgit", "Wah Cantonment", "Mingora, Swat", "Nawabshah", "Chiniot",
    "Hafizabad", "Sadiqabad", "Jhelum", "Attock", "Murree", "Burewala", "Mandi Bahauddin",
    "Kohat", "Mirpur Khas", "Kotri", "Jamshoro", "Thatta", "Dadu", "Khairpur", "Shikarpur",
    "Ghotki", "Sanghar", "Badin", "Tando Allahyar", "Tando Adam", "Umerkot", "Gwadar",
    "Turbat", "Khuzdar", "Sibi", "Zhob", "Chaman", "Loralai", "Nowshera", "Charsadda",
    "Swabi", "Bannu", "Dera Ismail Khan", "Haripur", "Mansehra", "Timergara", "Chitral",
    "Lakki Marwat", "Hangu", "Karak", "Batkhela", "Gojra", "Toba Tek Singh", "Khanewal",
    "Vehari", "Lodhran", "Muzaffargarh", "Layyah", "Bhakkar", "Mianwali", "Khushab",
    "Narowal", "Daska", "Wazirabad", "Lalamusa", "Kharian", "Pakpattan", "Bahawalnagar",
    "Haroonabad", "Khanpur", "Ahmedpur East", "Mailsi", "Kamalia", "Samundri", "Jaranwala",
    "Tandlianwala", "Chishtian", "Kot Addu", "Taxila", "Sambrial", "Pasrur", "Shakargarh",
    "Jauharabad", "Chakwal", "Talagang", "Fateh Jang", "Hassan Abdal", "Kamra", "Kamoke",
    "Muridke", "Skardu", "Hunza", "Kotli", "Rawalakot", "Bagh, Azad Kashmir", "Bhimber",
    "Hub, Balochistan", "Kharan", "Dera Murad Jamali", "Jacobabad", "Kashmore", "Matiari",
    "Naushahro Feroze", "Tando Muhammad Khan", "Sujawal", "Jati", "Mithi", "Dir", "Buner",
    "Shangla", "Kohistan", "Tank", "Malakand", "Mohmand", "Bajaur", "Khyber", "Kurram",
    "Parachinar", "Miranshah", "Wana", "Sheikhupura City", "Nankana Sahib", "Rajanpur",
    "Narang Mandi", "Hasilpur", "Liaquatpur", "Alipur", "Jatoi", "Shorkot", "Pir Mahal",
    "Renala Khurd", "Depalpur", "Chunian", "Pattoki", "Raiwind", "Ferozewala", "Sharaqpur",
    "Gakhar Mandi", "Rabwah", "Pind Dadan Khan", "Dina", "Sohawa", "Domeli", "Gujar Khan",
    "Kallar Syedan", "Rawat", "Chakri", "Bhera", "Sillanwali", "Shahpur", "Quaidabad",
    "Clifton, Karachi", "Gulshan-e-Iqbal, Karachi", "DHA, Karachi", "North Nazimabad, Karachi",
    "Saddar, Karachi", "PECHS, Karachi", "Bahria Town, Karachi", "Gulistan-e-Johar, Karachi",
    "Korangi, Karachi", "Malir, Karachi", "Landhi, Karachi", "Orangi Town, Karachi",
    "Nazimabad, Karachi", "Federal B Area, Karachi", "Tariq Road, Karachi", "Bahadurabad, Karachi",
    "Gulberg, Lahore", "DHA, Lahore", "Johar Town, Lahore", "Model Town, Lahore",
    "Bahria Town, Lahore", "Iqbal Town, Lahore", "Shadman, Lahore", "Samanabad, Lahore",
    "Wapda Town, Lahore", "Allama Iqbal Town, Lahore", "Valencia Town, Lahore", "Ravi Road, Lahore",
    "Blue Area, Islamabad", "F-7, Islamabad", "F-10, Islamabad", "G-11, Islamabad",
    "G-9, Islamabad", "I-8, Islamabad", "E-11, Islamabad", "DHA, Islamabad",
    "Bahria Town, Rawalpindi", "Saddar, Rawalpindi", "Satellite Town, Rawalpindi",
    "University Town, Peshawar", "Hayatabad, Peshawar", "Saddar, Peshawar",
    "Satiana Road, Faisalabad", "Madina Town, Faisalabad", "Cantt, Multan", "Gulgasht, Multan",
], ", Pakistan")

QATAR_LOCATIONS = _expand([
    "Doha", "West Bay, Doha", "Al Rayyan", "Al Wakrah", "Al Khor", "Umm Salal", "Lusail",
    "The Pearl, Doha", "Al Daayen", "Mesaieed", "Dukhan", "Al Shamal", "Al Sadd, Doha",
    "Msheireb, Doha", "Old Airport, Doha", "Al Waab, Doha", "Bin Mahmoud, Doha",
    "Najma, Doha", "Al Mansoura, Doha", "Fereej Abdel Aziz, Doha", "Madinat Khalifa, Doha",
    "Al Gharrafa, Doha", "Muaither", "Abu Hamour, Doha", "Industrial Area, Doha",
    "Ain Khaled, Doha", "Umm Salal Mohammed", "Umm Salal Ali", "Al Thumama, Doha",
    "Al Sailiya", "Al Kharaitiyat", "Ras Laffan", "Simaisma", "Al Ruwais", "Al Zubarah",
    "Al Shahaniya", "Abu Samra", "Al Kheesa", "Al Aziziyah, Doha", "Al Hilal, Doha",
    "Al Nasr, Doha", "Corniche, Doha", "Katara, Doha", "Education City, Doha",
    "Al Dafna, Doha", "Al Muntazah, Doha", "Al Markhiya, Doha", "Umm Ghuwailina, Doha",
    "Al Duhail, Doha", "Al Maamoura, Doha", "Al Souq, Doha", "Rawdat Al Khail, Doha",
    "Al Mirqab, Doha", "Fereej Bin Omran, Doha", "Fereej Kulaib, Doha", "Wakra Souq",
    "Al Wukair", "Al Egla", "Barwa Village, Doha", "Gharaffat Al Rayyan", "Izghawa",
    "Al Luqta", "Al Mamoura", "Onaiza, Doha", "Musheireb Downtown, Doha",
], ", Qatar")

REGIONS = {
    "UK": UK_LOCATIONS,
    "Australia": AUSTRALIA_LOCATIONS,
    "Canada": CANADA_LOCATIONS,
    "USA": USA_LOCATIONS,
    "Pakistan": PAKISTAN_LOCATIONS,
    "Qatar": QATAR_LOCATIONS,
}

ACTIVE_REGIONS = ["UK", "Australia", "Canada", "USA", "Pakistan", "Qatar"]

EXTRA_LOCATIONS_FILE = "extra_locations.txt"


def build_locations():
    cities = [c for region in ACTIVE_REGIONS for c in REGIONS[region]]
    if os.path.exists(EXTRA_LOCATIONS_FILE):
        with open(EXTRA_LOCATIONS_FILE, encoding="utf-8") as f:
            cities += [ln.strip() for ln in f if ln.strip() and not ln.startswith("#")]
    return list(dict.fromkeys(cities))


LOCATIONS = build_locations()

LEAD_TYPE_OPTIONS = {
    "1": "Businesses WITHOUT a website",
    "2": "Businesses WITH a website",
}

INDUSTRY_OPTIONS = {
    "1": ("Web design / SEO / marketing businesses", SERVICES),
    "2": ("Restaurants, cafes & food", ["restaurant", "cafe", "takeaway", "fast food", "bakery"]),
    "3": ("Salons, spas & barbers", ["hair salon", "barber shop", "beauty salon", "spa", "nail salon"]),
    "4": ("Gyms & fitness", ["gym", "fitness center", "yoga studio", "personal trainer"]),
    "5": ("Dentists & clinics", ["dentist", "dental clinic", "medical clinic", "physiotherapist"]),
    "6": ("Real estate & property", ["real estate agent", "property dealer", "estate agency", "property management"]),
    "7": ("Home services (plumber, electrician, cleaning)", ["plumber", "electrician", "cleaning service",
                                                              "painter and decorator", "handyman", "pest control"]),
    "8": ("Auto & car services", ["car repair", "auto mechanic", "car wash", "car dealer", "tyre shop"]),
    "9": ("Retail shops", ["clothing store", "boutique", "furniture store", "electronics shop", "gift shop"]),
    "10": ("Education & coaching", ["tutor", "coaching center", "driving school", "language school", "music school"]),
    "11": ("Construction & design", ["builder", "construction company", "contractor", "interior designer", "architect"]),
    "12": ("Events & photography", ["wedding planner", "photographer", "event planner", "catering service", "florist"]),
    "13": ("Legal, finance & travel", ["lawyer", "accountant", "tax consultant", "insurance agent", "travel agent"]),
}
CUSTOM_KEY = "14" 

SERVICE_OPTIONS = {
    "1": "SEO",
    "2": "Web Design & Development",
    "3": "Social Media Marketing",
}

LEAD_FILTER = "no_website"  
SEARCH_TERMS = [(s, "Web / SEO / Marketing businesses") for s in SERVICES] 
OFFER_SERVICES = list(SERVICE_OPTIONS.values())



def configure_from_form(lead_choice, region_picks, industry_picks, custom_terms, offer_picks):
    """Apply web-form choices (same effect as original configure_run)."""
    global LEAD_FILTER, SEARCH_TERMS, OFFER_SERVICES, ACTIVE_REGIONS, LOCATIONS
    LEAD_FILTER = "no_website" if lead_choice == "1" else "with_website"

    region_names = list(REGIONS)
    # region_picks are region names (UK, Australia, ...)
    ACTIVE_REGIONS = [r for r in region_picks if r in REGIONS] or region_names[:1]
    LOCATIONS = build_locations()

    terms = []
    for k in industry_picks:
        if k == CUSTOM_KEY:
            terms += [(t.strip(), "Custom") for t in (custom_terms or "").split(",") if t.strip()]
        elif k in INDUSTRY_OPTIONS:
            label, words = INDUSTRY_OPTIONS[k]
            terms += [(w, label) for w in words]
    SEARCH_TERMS = list(dict.fromkeys(terms)) or SEARCH_TERMS

    OFFER_SERVICES = [SERVICE_OPTIONS[k] for k in offer_picks if k in SERVICE_OPTIONS] or list(SERVICE_OPTIONS.values())

    print("\\n✅ Setup done")
    print(f"   Lead type : {LEAD_TYPE_OPTIONS.get(lead_choice, lead_choice)}")
    print(f"   Regions   : {', '.join(ACTIVE_REGIONS)} ({len(LOCATIONS)} locations)")
    print(f"   Industries: {', '.join(dict.fromkeys(lbl for _, lbl in SEARCH_TERMS))}")
    print(f"   Services  : {', '.join(OFFER_SERVICES)}")


def notes_text(industry):
    return f"Industry: {industry} | Offer: {', '.join(OFFER_SERVICES)}"


MAX_SCROLLS = 15                   
MAX_PLACES_PER_QUERY = 80
HEADLESS = True
OUTPUT_EXCEL = "UK_Leads.xlsx"      
DELAY = 5                           
SAVE_EVERY = 5                      


ONLY_WITH_CONTACT = True

USE_SEARCH_FALLBACK = True

WEB_FALLBACK = "always"
SEARCH_ENGINES = ["google", "duckduckgo"]
FALLBACK_MAX_RESULTS = 15

API_ENABLED = True
API_URL = "https://crm.rhoflow.net/api/external/leads"
API_KEY = os.getenv("CRM_API_KEY") or "f33f43afb069e63f3ab24b0ee3d121275736ccae9531ef54"
API_KEY_HEADER = "X-API-Key"     
PAYLOAD_MODE = "single"

CRM_FIELD_MAP = {
    "location": "location",
    "lead_type": "lead_type",
    "name": "company_name",
    "phone": "phone_number",
    "email": "email",
    "website": "website_url",
    "address": "address",
    "instagram": "instagram",
    "facebook": "facebook",
    "linkedin": "linkedin",
    "twitter": "twitter",
    "youtube": "youtube",
    "status": "status",
    "notes": "notes",
}

CRM_STATUSES = ["New", "Contacted", "Replied", "Meeting Booked",
                "Proposal Sent", "Won", "Lost", "Not Interested"]

FIELD_TO_HEADER = {
    "location": "Location",
    "lead_type": "Lead Type",
    "name": "Company Name",
    "phone": "Phone Number",
    "email": "Email",
    "website": "Website URL",
    "address": "Address",
    "instagram": "Instagram",
    "facebook": "Facebook",
    "linkedin": "LinkedIn",
    "twitter": "Twitter",
    "youtube": "YouTube",
    "status": "Status",
    "notes": "Notes",
}
HEADER_TO_FIELD = {v: k for k, v in FIELD_TO_HEADER.items()}
FIELDS = list(FIELD_TO_HEADER.keys())

STATE_SHEET = "_state"
STATE_COLS = ["key", "place_id", "crm_synced"]
QUERY_PREFIX = "QUERY|"
BACKUP_FILE = "UK_Leads_old_backup.xlsx"

EMAIL_RE = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
BAD_EMAIL_WORDS = ["example", "sentry", "wix", "wordpress", "domain",
                   ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", "schema.org", "google"]
PHONE_RE = (r"(?:\+\d{1,3}[\s.-]?\(?\d{1,4}\)?(?:[\s.-]?\d{2,5}){2,4}"
            r"|\(?\b0\d{1,4}\)?[\s.-]?\d{3,4}[\s.-]?\d{3,4}"
            r"|\(?\b[2-9]\d{2}\)?[\s.-]\d{3}[\s.-]\d{4})")

SOCIAL_HOSTS = ("facebook.com", "instagram.com", "linkedin.com", "twitter.com", "x.com",
                "youtube.com", "tiktok.com", "pinterest.com")
IGNORE_HOSTS = SOCIAL_HOSTS + ("google.", "goo.gl", "bit.ly", "wa.me", "whatsapp.com", "linktr.ee",
                               "fb.me", "fb.com", "duckduckgo", "w3.org", "schema.org",
                               "gstatic", "fbcdn", "fbsbx", "cloudfront")


def clean(text):
    if not text:
        return ""
    return re.sub(r"\s+", " ", str(text)).strip()


STRICT_EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+-]+@(?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,24}$")


def valid_email(e):
    """True only for a clean, CRM-safe address (rejects junk such as 'a@b.comPhone' or 'a@b.c')."""
    e = (e or "").strip()
    return bool(STRICT_EMAIL_RE.match(e)) and ".." not in e and len(e) <= 100


def good_emails(text):
    found = re.findall(EMAIL_RE, text or "")
    out = []
    for e in found:
        e = e.strip(".-_")
        if any(b in e.lower() for b in BAD_EMAIL_WORDS) or not valid_email(e):
            continue
        out.append(e)
    return out


def good_phones(text):
    out = []
    for m in re.findall(PHONE_RE, text or ""):
        digits = re.sub(r"\D", "", m)
        if 8 <= len(digits) <= 15:
            out.append(clean(m))
    return out


def is_social_url(url):
    host = urlparse(url or "").netloc.lower()
    return any(h in host for h in SOCIAL_HOSTS)


def place_id(url):
    """Stable Google Maps ID (the same business is recognised even if the URL differs)."""
    m = re.search(r"0x[0-9a-f]+:0x[0-9a-f]+", url or "")
    return m.group(0) if m else (url or "").split("?")[0]


def make_key(name, phone, address):
    """Duplicate key built from name + phone + address."""
    digits = re.sub(r"\D", "", str(phone or ""))
    return f"{str(name or '').lower().strip()}|{digits}|{str(address or '').lower().strip()}"


def empty_leads():
    return pd.DataFrame(columns=FIELDS)


def empty_state():
    return pd.DataFrame(columns=STATE_COLS)


def _read_sheet(sheet):
    """Read a sheet. If Excel has the file open, wait until it is closed (never return an
    empty DataFrame, or old leads would be overwritten). Returns None if the sheet is missing."""
    while True:
        try:
            return pd.read_excel(OUTPUT_EXCEL, sheet_name=sheet, dtype=str).fillna("")
        except PermissionError:
            print(f"⚠️ {OUTPUT_EXCEL} is open in Excel. Close it, retrying in 10 seconds...")
            time.sleep(10)
        except ValueError:
            return None          
        except Exception as e:
            print(f"❌ Could not read {OUTPUT_EXCEL}: {e}")
            print("   The file may be corrupted. Make a backup and delete or repair it.")
            raise


def load_data():
    """Returns (leads_df, state_df)."""
    if not os.path.exists(OUTPUT_EXCEL):
        return empty_leads(), empty_state()

    leads = _read_sheet("Leads")
    if leads is None:
        leads = empty_leads()
    leads = leads.rename(columns=HEADER_TO_FIELD).reindex(columns=FIELDS).fillna("")

    state = _read_sheet(STATE_SHEET)
    if state is None:
        state = empty_state()
    state = state.reindex(columns=STATE_COLS).fillna("")
    return leads, state


def write_data(leads, state):
    """One file: 'Leads' sheet (14 columns) + hidden '_state' sheet."""
    out = leads.reindex(columns=FIELDS).fillna("").rename(columns=FIELD_TO_HEADER)
    state = state.reindex(columns=STATE_COLS).fillna("")

    while True:
        try:
            with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
                out.to_excel(writer, sheet_name="Leads", index=False)
                ws = writer.sheets["Leads"]
                ws.freeze_panes = "A2"
                ws.auto_filter.ref = ws.dimensions
                for col in ws.columns:
                    width = max(len(str(c.value or "")) for c in list(col)[:50]) + 2
                    ws.column_dimensions[col[0].column_letter].width = min(width, 45)

                headers = [c.value for c in ws[1]]
                letter = ws.cell(row=1, column=headers.index(FIELD_TO_HEADER["status"]) + 1).column_letter
                last = max(ws.max_row, 3000)
                dv = DataValidation(type="list",
                                    formula1='"' + ",".join(CRM_STATUSES) + '"',
                                    allow_blank=True)
                ws.add_data_validation(dv)
                dv.add(f"{letter}2:{letter}{last}")

                state.to_excel(writer, sheet_name=STATE_SHEET, index=False)
                writer.sheets[STATE_SHEET].sheet_state = "hidden"
            return
        except PermissionError:
            print(f"⚠️ {OUTPUT_EXCEL} is open in Excel. Close it, retrying in 10 seconds...")
            time.sleep(10)


def create_empty_excel():
    if not os.path.exists(OUTPUT_EXCEL):
        write_data(empty_leads(), empty_state())
        print(f"✅ Empty Excel created → {OUTPUT_EXCEL}")


def migrate_old_file():
    """Convert an old file (lead_id, query, rating, maps_url ... columns) into the new
    14-column format. A backup is created first."""
    if not os.path.exists(OUTPUT_EXCEL):
        return
    df = _read_sheet("Leads")
    if df is None or "name" not in df.columns or "Company Name" in df.columns:
        return

    if not os.path.exists(BACKUP_FILE):
        shutil.copy2(OUTPUT_EXCEL, BACKUP_FILE)
        print(f"🛟 Backup of the old file created: {BACKUP_FILE} (delete it whenever you like)")

    leads = pd.DataFrame(index=df.index)
    for f in FIELDS:
        leads[f] = df[f] if f in df.columns else ""
    if "query" in df.columns:
        leads["location"] = df["query"].apply(lambda q: q.split(" in ", 1)[-1] if q else "")
    leads["status"] = leads["status"].replace("", "New")

    keys = [make_key(n, p, a) for n, p, a in zip(leads["name"], leads["phone"], leads["address"])]
    pids = [place_id(u) for u in df["maps_url"]] if "maps_url" in df.columns else [""] * len(df)
    synced = df["crm_synced"] if "crm_synced" in df.columns else ["No"] * len(df)
    state = pd.DataFrame({"key": keys, "place_id": pids,
                          "crm_synced": ["Yes" if s == "Yes" else "No" for s in synced]})
    write_data(leads.reset_index(drop=True), state)
    print(f"🔁 Old file converted to the new format ({len(leads)} leads)")


def load_seen():
    leads, state = load_data()
    ids = {p for p in state["place_id"] if p}
    keys = {make_key(n, p, a) for n, p, a in zip(leads["name"], leads["phone"], leads["address"])}
    done = {k[len(QUERY_PREFIX):] for k in state["key"] if k.startswith(QUERY_PREFIX)}
    return ids, keys, done


def mark_queries_done(queries):
    """Record finished queries in the hidden sheet so they are skipped on the next run."""
    if not queries:
        return
    leads_df, state_df = load_data()
    new = pd.DataFrame([{"key": QUERY_PREFIX + q, "place_id": "", "crm_synced": "Done"} for q in queries])
    state_df = new if state_df.empty else pd.concat([state_df, new], ignore_index=True)
    write_data(leads_df, state_df)


def save_leads(leads, quiet=False):
    if not leads:
        return []

    leads_df, state_df = load_data()
    old_ids = {p for p in state_df["place_id"] if p}
    old_keys = {make_key(n, p, a) for n, p, a in
                zip(leads_df["name"], leads_df["phone"], leads_df["address"])}

    fresh, batch_keys = [], set()
    for l in leads:
        k = make_key(l["name"], l["phone"], l["address"])
        if place_id(l["maps_url"]) in old_ids or k in old_keys or k in batch_keys:
            continue
        batch_keys.add(k)
        fresh.append(l)

    if not fresh:
        if not quiet:
            print("   (all these leads were already saved, nothing new)")
        return []

    df_new = pd.DataFrame(fresh).reindex(columns=FIELDS).astype(str).fillna("")
    df_new["status"] = "New"
    leads_df = df_new if leads_df.empty else pd.concat([leads_df, df_new], ignore_index=True)

    new_state = pd.DataFrame([{"key": make_key(l["name"], l["phone"], l["address"]),
                               "place_id": place_id(l["maps_url"]),
                               "crm_synced": "No"} for l in fresh])
    state_df = new_state if state_df.empty else pd.concat([state_df, new_state], ignore_index=True)

    write_data(leads_df, state_df)
    print(f"💾 Excel saved | +{len(fresh)} new | Total: {len(leads_df)}")
    sync_to_crm()          # send the new leads to the CRM right away
    return fresh


OK_CODES = (200, 201, 202)    
DUPLICATE_CODE = 409 


def _crm_headers(style=None):
    style = style or API_KEY_HEADER
    h = {"Content-Type": "application/json", "Accept": "application/json"}
    if style.lower() == "authorization":
        h["Authorization"] = f"Bearer {API_KEY}"
    else:
        h[style] = API_KEY
    return h


def _crm_payload(row):
    """Only the 14 fields listed in CRM_FIELD_MAP."""
    payload = {crm_field: row.get(src, "") for src, crm_field in CRM_FIELD_MAP.items()}
    payload["status"] = (payload.get("status") or "New").strip().lower()
    if not valid_email(payload.get("email")):
        payload["email"] = "" 
    return payload


def _crm_outcome(r):
    """Classify a CRM reply: 'ok' | 'duplicate' | 'auth' | 'error'."""
    if r.status_code == DUPLICATE_CODE:
        return "duplicate"
    if r.status_code in (401, 403):
        return "auth"
    if r.status_code in OK_CODES:
        try:
            d = r.json()
            if isinstance(d, dict) and (d.get("success") is False or d.get("error")
                                        or str(d.get("status", "")).lower() in ("error", "failed", "fail")):
                return "error"
        except ValueError:
            pass
        return "ok"
    return "error"


def sync_to_crm():
    """Send leads that are not yet in the CRM. Failed leads are retried next time,
    sent leads are never sent twice. Every reply is printed / logged so problems are visible."""
    if not API_ENABLED:
        return
    if "PASTE_" in API_URL or "PASTE_" in API_KEY or not API_KEY:
        print("⚠️ CRM URL / API key is not set, CRM sync skipped (data is saved in Excel).")
        return

    leads_df, state_df = load_data()
    if leads_df.empty:
        return

    smap = {k: [p, s] for k, p, s in zip(state_df["key"], state_df["place_id"], state_df["crm_synced"])}
    keys = [make_key(n, p, a) for n, p, a in zip(leads_df["name"], leads_df["phone"], leads_df["address"])]
    pending = [(i, k) for i, k in zip(leads_df.index, keys) if smap.get(k, ["", "No"])[1] != "Yes"]
    if not pending:
        return

    def mark(k):
        smap.setdefault(k, ["", "No"])[1] = "Yes"

    def state_now():
        return pd.DataFrame([{"key": k, "place_id": v[0], "crm_synced": v[1]}
                             for k, v in smap.items()], columns=STATE_COLS)

    headers = _crm_headers()
    created = duplicates = failed = 0
    try:
        if PAYLOAD_MODE == "bulk":
            body = {"leads": [_crm_payload(leads_df.loc[i].to_dict()) for i, _ in pending]}
            r = requests.post(API_URL, json=body, headers=headers, timeout=60)
            outcome = _crm_outcome(r)
            if outcome in ("ok", "duplicate"):
                for _, k in pending:
                    mark(k)
                created = len(pending)
            else:
                failed = len(pending)
        else:
            for i, k in pending:
                row = leads_df.loc[i].to_dict()
                r = requests.post(API_URL, json=_crm_payload(row), headers=headers, timeout=30)
                if r.status_code == 422 and "email" in r.text.lower() and row.get("email"):
                    retry = _crm_payload(row)
                    retry["email"] = ""
                    r = requests.post(API_URL, json=retry, headers=headers, timeout=30)
                    if r.status_code in OK_CODES:
                        print(f"   ✉️  {row['name'][:35]}: CRM rejected the email, lead sent without it")
                outcome = _crm_outcome(r)
                if outcome == "ok":
                    mark(k)
                    created += 1
                elif outcome == "duplicate":
                    mark(k)
                    duplicates += 1
                    print(f"   ♻️  {row['name'][:35]}: CRM says it already exists (409)")
                elif outcome == "auth":
                    print(f"❌ CRM rejected the API key ({r.status_code}). Check API_KEY / API_KEY_HEADER.")
                    failed += 1
                    break
                else:
                    failed += 1
                    print(f"❌ {row['name'][:35]}: CRM error {r.status_code} {clean(r.text)[:150]}")
                if (created + duplicates) and (created + duplicates) % 25 == 0:
                    write_data(leads_df, state_now()) 
                time.sleep(0.2)
    except requests.exceptions.RequestException as e:
        print(f"❌ Could not connect to the CRM: {str(e)[:150]}")
        print("   Check the internet / VPN and API_URL. Run  python leads_scraper.py --crm-test  to diagnose.")
    except Exception as e:
        print(f"❌ CRM sync error: {str(e)[:150]}")
    finally:
        if created or duplicates:
            write_data(leads_df, state_now())
        total_in_crm = sum(1 for k in keys if smap.get(k, ["", "No"])[1] == "Yes")
        if created or duplicates:
            extra = f" | {duplicates} already existed" if duplicates else ""
            print(f"📡 CRM saved | +{created} new | Total: {total_in_crm}{extra}")
        left = len(pending) - created - duplicates
        if left:
            print(f"⚠️ {left} leads not sent to the CRM yet (they are saved in Excel and will be retried)")


def crm_test():
    """Send ONE clearly labelled test lead and print the raw CRM reply.
    Run:  python leads_scraper.py --crm-test     (delete the test lead in the CRM afterwards)"""
    test = {f: "" for f in FIELDS}
    test.update({"location": "Test City", "lead_type": "No Website", "name": "TEST LEAD - delete me",
                 "phone": "+10000000000", "email": "test@example.com", "status": "New",
                 "notes": "Connection test from leads_scraper.py"})
    attempts = [
        ("X-API-Key header, one lead per request", "X-API-Key", "single"),
        ("Authorization: Bearer header, one lead per request", "Authorization", "single"),
        ("X-API-Key header, bulk body", "X-API-Key", "bulk"),
    ]
    print(f"Testing the CRM at {API_URL} ...")
    for label, style, mode in attempts:
        payload = _crm_payload(test)
        body = {"leads": [payload]} if mode == "bulk" else payload
        try:
            r = requests.post(API_URL, json=body, headers=_crm_headers(style), timeout=30)
        except requests.exceptions.RequestException as e:
            print(f"❌ {label}: cannot connect: {str(e)[:150]}")
            print("   → The server cannot be reached from this computer (internet / VPN / wrong URL).")
            return
        outcome = _crm_outcome(r)
        print(f"- {label}: HTTP {r.status_code} → {outcome} | {clean(r.text)[:200]}")
        if outcome in ("ok", "duplicate"):
            print(f'\n✅ This works. Set  API_KEY_HEADER = "{style}"  and  PAYLOAD_MODE = "{mode}"  in the script.')
            print("   A lead named 'TEST LEAD - delete me' was created in the CRM, please delete it there.")
            return
    print("\n❌ None of the combinations worked. If the replies above are 400/422, the CRM wants different")
    print("   field names (edit the right side of CRM_FIELD_MAP). If 404, API_URL is wrong.")


def empty_socials():
    return {"instagram": "", "facebook": "", "linkedin": "", "twitter": "",
            "youtube": "", "email": "", "phone": ""}


def classify_link(href, data):
    """If the link is a social profile, store it in data."""
    href = (href or "").lower().split("?")[0]
    if not href.startswith("http") or "google" in href:
        return
    if "instagram.com" in href and not data["instagram"]:
        data["instagram"] = href
    elif "facebook.com" in href and not data["facebook"]:
        data["facebook"] = href
    elif "linkedin.com" in href and not data["linkedin"]:
        data["linkedin"] = href
    elif ("twitter.com" in href or "x.com/" in href) and not data["twitter"]:
        data["twitter"] = href
    elif "youtube.com" in href and not data["youtube"]:
        data["youtube"] = href


def get_phone(page):
    """Phone from a Maps page: 1) phone button, 2) tel: link, 3) number in the page text."""
    try:
        btn = page.locator('button[data-item-id^="phone"], button[aria-label*="Phone"]')
        if btn.count():
            raw = btn.first.get_attribute("aria-label") or btn.first.inner_text()
            p = clean(re.sub(r"[^\d+\-\s()]", "", raw))
            if len(re.sub(r"\D", "", p)) >= 7:
                return p
    except Exception:
        pass
    try:
        tel = page.locator('a[href^="tel:"]')
        if tel.count():
            p = clean(unquote(tel.first.get_attribute("href") or "").replace("tel:", ""))
            if len(re.sub(r"\D", "", p)) >= 7:
                return p
    except Exception:
        pass
    try:
        txt = page.locator('div[role="main"]').first.inner_text(timeout=3000)[:2500]
        phones = good_phones(txt)
        if phones:
            return phones[0]
    except Exception:
        pass
    return ""


def get_socials(page, website):
    """Business with a website: open it and collect email, phone and social links."""
    data = empty_socials()
    if not website or not website.startswith("http"):
        return data
    try:
        page.goto(website, timeout=15000, wait_until="domcontentloaded")
        time.sleep(2)
        for a in page.locator("a[href]").all():
            try:
                href = (a.get_attribute("href") or "")
                if href.lower().startswith("mailto:"):
                    if not data["email"]:
                        mail = unquote(href[7:].split("?")[0]).strip()
                        if valid_email(mail):
                            data["email"] = mail
                elif href.lower().startswith("tel:"):
                    if not data["phone"]:
                        data["phone"] = clean(unquote(href[4:]))
                else:
                    classify_link(href, data)
            except Exception:
                continue
        if not data["email"]:
            emails = good_emails(page.content())
            if emails:
                data["email"] = emails[0]
    except Exception:
        pass
    return data


def get_nowebsite_contacts(page):
    """Business without a website: look for email / social links on the Maps page."""
    data = empty_socials()
    try:
        emails = good_emails(page.content())
        if emails:
            data["email"] = emails[0]
        for a in page.locator('a[href*="facebook.com"], a[href*="instagram.com"], a[href*="linkedin.com"]').all():
            classify_link(a.get_attribute("href"), data)
    except Exception:
        pass
    return data


def search_fallback(page, name, address, need_phone, need_email):
    """No contact on Maps: search the business name on DuckDuckGo. Only results that
    contain the business name are used (avoids wrong numbers). Then try its Facebook page."""
    data = empty_socials()
    try:
        parts = [p.strip() for p in (address or "").split(",") if p.strip()]
        place = " ".join(parts[-2:]) if parts else ""
        q = f"{name} {place} contact email phone"
        page.goto("https://html.duckduckgo.com/html/?q=" + quote(q),
                  timeout=15000, wait_until="domcontentloaded")
        time.sleep(2)

        name_key = name.lower()[:15]
        for block in page.locator("div.result").all()[:8]:
            try:
                text = block.inner_text(timeout=2000)
                if name_key not in text.lower():
                    continue
                if need_email and not data["email"]:
                    em = good_emails(text)
                    if em:
                        data["email"] = em[0]
                if need_phone and not data["phone"]:
                    ph = good_phones(text)
                    if ph:
                        data["phone"] = ph[0]
                for a in block.locator("a[href]").all():
                    href = a.get_attribute("href") or ""
                    if "uddg=" in href:
                        href = unquote(parse_qs(urlparse(href).query).get("uddg", [""])[0])
                    classify_link(href, data)
            except Exception:
                continue

        if data["facebook"] and ((need_email and not data["email"]) or (need_phone and not data["phone"])):
            try:
                page.goto(data["facebook"], timeout=12000, wait_until="domcontentloaded")
                time.sleep(2)
                content = page.content()
                if need_email and not data["email"]:
                    fb_em = good_emails(content)
                    if fb_em:
                        data["email"] = fb_em[0]
                if need_phone and not data["phone"]:
                    fb_ph = good_phones(page.inner_text("body")[:4000])
                    if fb_ph:
                        data["phone"] = fb_ph[0]
            except Exception:
                pass
    except Exception:
        pass
    return data


FB_SKIP = ("/posts/", "/photos/", "/videos/", "/groups/", "/events/", "/watch", "/reel",
           "/story", "/share", "/hashtag", "/marketplace", "/login", "/sharer", "/permalink", "/public/")
URL_RE = r"(?:https?://|www\.)[a-z0-9][a-z0-9.-]*\.[a-z]{2,}(?:/[^\s]*)?"
GOOGLE_BLOCKED = False


def normalize_social(url):
    """Turn a Facebook / Instagram result into the page root URL. Posts/photos/groups -> None."""
    try:
        u = urlparse(url or "")
        host = re.sub(r"^(www|m|mobile|web)\.", "", u.netloc.lower())
        path = u.path.strip("/")
        if host == "facebook.com":
            if not path or any(x in "/" + path + "/" for x in FB_SKIP):
                return None
            parts = path.split("/")
            if parts[0] in ("pages", "people") and len(parts) >= 3:
                return "facebook", f"https://www.facebook.com/{parts[0]}/{parts[1]}/{parts[2]}"
            return "facebook", f"https://www.facebook.com/{parts[0]}"
        if host == "instagram.com":
            parts = path.split("/")
            if not parts[0] or parts[0] in ("p", "reel", "reels", "explore", "tv", "stories", "accounts"):
                return None
            return "instagram", f"https://www.instagram.com/{parts[0]}/"
    except Exception:
        pass
    return None


def name_from_title(title, platform):
    t = clean(title)
    if platform == "instagram":
        t = t.split(" (@")[0].split(" • ")[0]
    t = re.split(r"\s[-|–•·]\s", t)[0]
    t = re.sub(r"\b(Facebook|Instagram)\b", "", t, flags=re.I).strip(" -|•·")
    return t[:80]


def mentions_own_website(text):
    """If the page text links to any non-social website, the business has its own website."""
    for m in re.findall(URL_RE, text or "", flags=re.I):
        host = re.sub(r"^(https?://)?(www\.)?", "", m.lower()).split("/")[0]
        if not any(h in host for h in IGNORE_HOSTS):
            return True
    return False


def _accept_consent(page):
    for txt in ["Accept all", "Accept All", "I agree", "Agree"]:
        try:
            btn = page.locator(f'button:has-text("{txt}")')
            if btn.count() > 0:
                btn.first.click(timeout=2000)
                time.sleep(1)
                return
        except Exception:
            pass


def _google_results(page, q):
    global GOOGLE_BLOCKED
    page.goto("https://www.google.com/search?num=20&hl=en&q=" + quote(q),
              timeout=20000, wait_until="domcontentloaded")
    time.sleep(2)
    _accept_consent(page)

    def blocked():
        try:
            return "/sorry/" in page.url or "unusual traffic" in page.content()[:6000].lower()
        except Exception:
            return False

    if blocked():
        if not HEADLESS:
            print("   ⚠️ Google showed a captcha. Solve it in the browser (waiting up to 90 seconds)...")
            for _ in range(30):
                time.sleep(3)
                if not blocked():
                    break
        if blocked():
            GOOGLE_BLOCKED = True
            print("   ⚠️ Google is blocking us, DuckDuckGo will be used for the rest of this run.")
            return []

    out = []
    for b in page.locator("div.g").all()[:FALLBACK_MAX_RESULTS * 2]:
        try:
            title = clean(b.locator("h3").first.inner_text(timeout=1500))
            href = b.locator("a[href^='http']").first.get_attribute("href") or ""
            out.append({"title": title, "url": href, "text": clean(b.inner_text(timeout=1500))})
        except Exception:
            continue
    if not out:
        for a in page.locator("a:has(h3)").all()[:FALLBACK_MAX_RESULTS * 2]:
            try:
                out.append({"title": clean(a.locator("h3").first.inner_text(timeout=1500)),
                            "url": a.get_attribute("href") or "", "text": ""})
            except Exception:
                continue
    return out


def _ddg_results(page, q):
    page.goto("https://html.duckduckgo.com/html/?q=" + quote(q),
              timeout=20000, wait_until="domcontentloaded")
    time.sleep(2)
    out = []
    for b in page.locator("div.result").all()[:FALLBACK_MAX_RESULTS * 2]:
        try:
            a = b.locator("a.result__a").first
            href = a.get_attribute("href") or ""
            if "uddg=" in href:
                href = unquote(parse_qs(urlparse(href).query).get("uddg", [""])[0])
            out.append({"title": clean(a.inner_text(timeout=1500)), "url": href,
                        "text": clean(b.inner_text(timeout=1500))})
        except Exception:
            continue
    return out


def search_web(page, q):
    for engine in SEARCH_ENGINES:
        if engine == "google" and GOOGLE_BLOCKED:
            continue
        try:
            results = _google_results(page, q) if engine == "google" else _ddg_results(page, q)
        except Exception:
            results = []
        if results:
            return results
    return []


def web_fallback(page, service, city, seen_ids, seen_keys, industry=""):
    """Find businesses without a website through web search. Only businesses that have just
    a Facebook / Instagram page (and no website of their own) are kept."""
    leads = []
    candidates = {}
    for q in (f"{service} {city} site:facebook.com", f"{service} {city} site:instagram.com"):
        for r in search_web(page, q)[:FALLBACK_MAX_RESULTS]:
            n = normalize_social(r["url"])
            if not n:
                continue
            plat, root = n
            if root not in candidates:
                candidates[root] = (plat, name_from_title(r["title"], plat), r.get("text", ""))
    print(f"   🔎 Web search fallback: {len(candidates)} social pages found")

    city_token = city.split(",")[0].strip().lower()
    for root, (plat, name, snippet) in candidates.items():
        pid = place_id(root)
        if pid in seen_ids or not name:
            continue
        try:
            body, html = "", ""
            try:
                page.goto(root, timeout=12000, wait_until="domcontentloaded")
                time.sleep(2)
                body = page.inner_text("body")[:6000]
                html = page.content()
            except Exception:
                pass
            alltext = snippet + "\n" + body

            if city_token and city_token not in alltext.lower():
                seen_ids.add(pid)                    
                continue
            if mentions_own_website(alltext):
                seen_ids.add(pid)
                print(f"   🔎 {name[:40]} (has a website, skipped)")
                continue

            phone = (good_phones(alltext) or [""])[0]
            email = (good_emails(alltext + "\n" + html) or [""])[0]
            if ONLY_WITH_CONTACT and not phone and not email:
                seen_ids.add(pid)
                print(f"   🔎 {name[:40]} (no phone/email found, skipped)")
                continue

            key = make_key(name, phone, "")
            if key in seen_keys:
                seen_ids.add(pid)
                continue

            leads.append({
                "location": city,
                "lead_type": "No Website",
                "name": name,
                "phone": phone,
                "email": email,
                "website": "",
                "address": "",
                "instagram": root if plat == "instagram" else "",
                "facebook": root if plat == "facebook" else "",
                "linkedin": "",
                "twitter": "",
                "youtube": "",
                "notes": notes_text(industry) + " | Found via web search",
                "maps_url": root,
            })
            seen_ids.add(pid)
            seen_keys.add(key)
            print(f"   🔎 📞 NO-WEBSITE {name[:40]} | {phone or '-'} | {email or '-'}")
        except Exception as e:
            print(f"   🔎 ❌ error: {str(e)[:80]}")
            continue
    return leads


def scrape(page, query, location, seen_ids, seen_keys, industry=""):
    print(f"\n🔍 [{'WITHOUT website' if LEAD_FILTER == 'no_website' else 'WITH website'}] {query}")
    leads = []
    saved_upto = 0
    stats = {"saved": 0, "no_contact": 0, "skipped_filter": 0}
    completed = False
    try:
        url = f"https://www.google.com/maps/search/{quote(query)}"
        page.goto(url, timeout=60000, wait_until="domcontentloaded")
        time.sleep(4)
        _accept_consent(page)

        try:
            feed = page.locator('div[role="feed"]')
            if feed.count() > 0:
                for _ in range(MAX_SCROLLS):
                    feed.evaluate("el => el.scrollTop = el.scrollHeight")
                    time.sleep(1.5)
        except Exception:
            pass

        links = page.locator('a[href*="/maps/place/"]').all()
        places, seen = [], set()
        for link in links:
            href = link.get_attribute("href")
            if href and "/maps/place/" in href:
                clean_url = href.split("?")[0].split("&")[0]
                if clean_url not in seen:
                    seen.add(clean_url)
                    places.append(clean_url)

        print(f"   Found {len(places)} places")

        for i, place_url in enumerate(places[:MAX_PLACES_PER_QUERY]):
            pid = place_id(place_url)
            if pid in seen_ids:
                continue
            try:
                page.goto(place_url, timeout=20000, wait_until="domcontentloaded")

                try:
                    page.wait_for_selector("h1", timeout=8000)
                    name = clean(page.locator("h1").first.inner_text(timeout=5000))
                except Exception:
                    print(f"   [{i+1}] ⚠️ page did not load, skipped")
                    continue
                if not name:
                    print(f"   [{i+1}] ⚠️ no name found, skipped")
                    continue
                time.sleep(1.5)

                phone = get_phone(page)

                website = ""
                try:
                    web = page.locator('a[data-item-id="authority"], a[aria-label*="Website"]')
                    if web.count():
                        website = web.first.get_attribute("href") or ""
                        if not website.startswith("http"):
                            website = ""
                except Exception:
                    pass

                socials = empty_socials()
                if website and is_social_url(website):
                    classify_link(website, socials)
                    website = ""

                has_website = bool(website)
                if (LEAD_FILTER == "no_website" and has_website) or \
                   (LEAD_FILTER == "with_website" and not has_website):
                    seen_ids.add(pid)
                    stats["skipped_filter"] += 1
                    print(f"   [{i+1}] ⏭️  {name[:45]} "
                          f"({'has a website' if has_website else 'no website'}, skipped - wrong lead type)")
                    continue

                address = ""
                try:
                    addr = page.locator('button[data-item-id="address"]')
                    if addr.count():
                        address = clean(addr.first.get_attribute("aria-label") or "")
                        address = re.sub(r"^Address:\s*", "", address)
                except Exception:
                    pass

                key = make_key(name, phone, address)
                if key in seen_keys:
                    seen_ids.add(pid)
                    print(f"   [{i+1}] ♻️  {name[:45]} (already saved, skipped)")
                    continue

                found = get_socials(page, website) if has_website else get_nowebsite_contacts(page)
                for k in ("email", "instagram", "facebook", "linkedin", "twitter", "youtube"):
                    if not socials[k] and found[k]:
                        socials[k] = found[k]
                if not phone and found.get("phone"):
                    phone = found["phone"]

                if USE_SEARCH_FALLBACK and (not phone or not socials["email"]):
                    fb = search_fallback(page, name, address,
                                         need_phone=not phone,
                                         need_email=not socials["email"])
                    for k in ("email", "instagram", "facebook", "linkedin", "twitter", "youtube"):
                        if not socials[k] and fb[k]:
                            socials[k] = fb[k]
                    if not phone and fb["phone"]:
                        phone = fb["phone"]

                email = socials.get("email", "")

                if ONLY_WITH_CONTACT and not phone and not email:
                    seen_ids.add(pid)
                    stats["no_contact"] += 1
                    print(f"   [{i+1}] ⏭️  {name[:45]} (no phone/email found, skipped)")
                    continue

                key = make_key(name, phone, address)
                if key in seen_keys:
                    seen_ids.add(pid)
                    continue

                leads.append({
                    "location": location,
                    "lead_type": "Has Website" if has_website else "No Website",
                    "name": name,
                    "phone": phone,
                    "email": email,
                    "website": website,
                    "address": address,
                    "instagram": socials.get("instagram", ""),
                    "facebook": socials.get("facebook", ""),
                    "linkedin": socials.get("linkedin", ""),
                    "twitter": socials.get("twitter", ""),
                    "youtube": socials.get("youtube", ""),
                    "notes": notes_text(industry),
                    "maps_url": place_url,
                })
                seen_ids.add(pid)
                seen_keys.add(key)
                stats["saved"] += 1
                print(f"   [{i+1}] {'🌐 HAS-WEBSITE' if has_website else '📞 NO-WEBSITE'} {name[:45]} | {phone or '-'} | {email or '-'}")

                if len(leads) - saved_upto >= SAVE_EVERY:
                    save_leads(leads[saved_upto:], quiet=True)
                    saved_upto = len(leads)

            except Exception as e:
                print(f"   [{i+1}] ❌ error: {str(e)[:80]}")
                continue

        completed = bool(places)

    except Exception as e:
        print(f"❌ Failed: {str(e)[:80]}")

    if saved_upto < len(leads):
        save_leads(leads[saved_upto:])

    print(f"   📊 This query: saved {stats['saved']} | "
          f"no contact found {stats['no_contact']} | "
          f"skipped (other lead type) {stats['skipped_filter']}")
    return leads, completed



# ── job state for web UI ───────────────────────────────────────────────────
_job_lock = Lock()
_job = {
    "running": False,
    "stop": False,
    "log": [],
    "stats": {"saved": 0, "queries_done": 0, "queries_total": 0},
    "config_summary": "",
    "error": None,
}


def job_print(msg):
    print(msg)
    with _job_lock:
        _job["log"].append(msg)
        if len(_job["log"]) > 400:
            _job["log"] = _job["log"][-300:]


def run_scrape_job():
    """Same body as original main(), without terminal configure_run."""
    global GOOGLE_BLOCKED
    GOOGLE_BLOCKED = False
    with _job_lock:
        _job["running"] = True
        _job["stop"] = False
        _job["log"] = []
        _job["error"] = None
        _job["stats"] = {"saved": 0, "queries_done": 0, "queries_total": len(LOCATIONS) * len(SEARCH_TERMS)}
        _job["config_summary"] = (
            f"{'WITHOUT' if LEAD_FILTER == 'no_website' else 'WITH'} website | "
            f"{', '.join(ACTIVE_REGIONS)} ({len(LOCATIONS)} cities) | "
            f"{len(SEARCH_TERMS)} terms | Offer: {', '.join(OFFER_SERVICES)}"
        )
    try:
        job_print("=" * 60)
        job_print(f"Leads Scraper - {'businesses WITHOUT a website' if LEAD_FILTER == 'no_website' else 'businesses WITH a website'}")
        job_print(f"Regions: {', '.join(ACTIVE_REGIONS)} | Locations: {len(LOCATIONS)} | Search terms: {len(SEARCH_TERMS)}")
        job_print(f"CRM sync: {'ON' if API_ENABLED else 'OFF'}")
        job_print("=" * 60)

        create_empty_excel()
        migrate_old_file()
        seen_ids, seen_keys, done_queries = load_seen()
        job_print(f"📂 {len(seen_keys)} leads already in Excel.")
        if done_queries:
            job_print(f"⏩ {len(done_queries)} queries already done (resume).")
        sync_to_crm()

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=HEADLESS,
                                        args=["--disable-blink-features=AutomationControlled"])
            context = browser.new_context(
                viewport={"width": 1280, "height": 900},
                user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                            "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"),
            )
            page = context.new_page()
            page.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

            pending_done = []
            try:
                for city in LOCATIONS:
                    if _job.get("stop"):
                        job_print("⏹ Stopped")
                        break
                    for service, industry in SEARCH_TERMS:
                        if _job.get("stop"):
                            break
                        query = f"{service} in {city}"
                        done_key = query if LEAD_FILTER == "no_website" else query + " [with website]"
                        if done_key in done_queries:
                            with _job_lock:
                                _job["stats"]["queries_done"] += 1
                            continue

                        leads, completed = scrape(page, query, city, seen_ids, seen_keys, industry)

                        fb_ran = False
                        if LEAD_FILTER == "no_website" and (WEB_FALLBACK == "always" or (WEB_FALLBACK == "no_leads" and not leads)):
                            fb_ran = True
                            fb_leads = web_fallback(page, service, city, seen_ids, seen_keys, industry)
                            if fb_leads:
                                save_leads(fb_leads)
                                leads = leads + fb_leads

                        if completed or fb_ran:
                            done_queries.add(done_key)
                            pending_done.append(done_key)
                        with _job_lock:
                            _job["stats"]["queries_done"] += 1
                        time.sleep(DELAY)

                    mark_queries_done(pending_done)
                    pending_done = []
            finally:
                if pending_done:
                    mark_queries_done(pending_done)
                browser.close()
        job_print(f"\n✅ Done! Check file → {OUTPUT_EXCEL}")
    except Exception as e:
        job_print(f"❌ Fatal: {e}")
        with _job_lock:
            _job["error"] = str(e)
    finally:
        with _job_lock:
            _job["running"] = False


# ── Flask ──────────────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "leads-scraper-key")

PAGE = """
<!DOCTYPE html>
<html><head><meta charset="utf-8"/><meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Leads Scraper</title>
<style>
body{font-family:system-ui,sans-serif;max-width:720px;margin:1.5rem auto;padding:0 1rem;background:#f4f6fa;color:#111}
.card{background:#fff;border-radius:10px;padding:1.25rem;box-shadow:0 2px 10px rgba(0,0,0,.06);margin-bottom:1rem}
h1{font-size:1.3rem;margin:0 0 .5rem;color:#1e40af}
label{display:block;margin:.7rem 0 .25rem;font-weight:600;font-size:.9rem}
.row{display:flex;flex-wrap:wrap;gap:.5rem 1rem;font-size:.88rem}
.row label{font-weight:400;margin:0}
input[type=text]{width:100%;padding:.5rem;border:1px solid #ddd;border-radius:6px}
button,.btn{background:#1e40af;color:#fff;border:none;border-radius:8px;padding:.65rem 1.1rem;font-weight:600;cursor:pointer;text-decoration:none;display:inline-block}
button.sec,.btn.sec{background:#e2e8f0;color:#111}
button:disabled{opacity:.5}
#log{background:#0f172a;color:#e2e8f0;font:12px/1.4 ui-monospace,monospace;padding:1rem;border-radius:8px;max-height:340px;overflow:auto;white-space:pre-wrap}
.flash{padding:.6rem 1rem;border-radius:8px;margin-bottom:.8rem;background:#ecfdf5;color:#047857}
.flash.err{background:#fef2f2;color:#b91c1c}
</style></head><body>
<h1>Leads Scraper</h1>
{% with msgs = get_flashed_messages(with_categories=1) %}
  {% for c,m in msgs %}<div class="flash {% if c=='error' %}err{% endif %}">{{ m }}</div>{% endfor %}
{% endwith %}

{% if page=='form' %}
<div class="card">
<form method="post" action="{{ url_for('start') }}">
  <label>1) Lead type</label>
  <div class="row">
    <label><input type="radio" name="lead_type" value="1" checked> Without website</label>
    <label><input type="radio" name="lead_type" value="2"> With website</label>
  </div>
  <label>2) Regions</label>
  <div class="row">
    {% for r in regions %}<label><input type="checkbox" name="regions" value="{{ r }}" {% if r=='UK' %}checked{% endif %}> {{ r }}</label>{% endfor %}
  </div>
  <label>3) Industries</label>
  <div class="row">
    {% for k,v in industries.items() %}<label><input type="checkbox" name="industries" value="{{ k }}" {% if k=='1' %}checked{% endif %}> {{ v[0] }}</label>{% endfor %}
    <label><input type="checkbox" name="industries" value="14"> Custom</label>
  </div>
  <input type="text" name="custom_terms" placeholder="Custom terms: dentist, orthodontist, ...">
  <label>4) Services you offer</label>
  <div class="row">
    {% for k,v in services.items() %}<label><input type="checkbox" name="offers" value="{{ k }}" checked> {{ v }}</label>{% endfor %}
  </div>
  <p style="margin-top:1rem">
    <button type="submit" {% if running %}disabled{% endif %}>Start scrape</button>
    <a class="btn sec" href="{{ url_for('status') }}">Status</a>
    {% if excel %}<a class="btn sec" href="{{ url_for('download') }}">Download Excel</a>{% endif %}
  </p>
</form>
</div>
{% else %}
<div class="card">
  <p><b>{{ 'RUNNING' if running else 'IDLE' }}</b> — {{ summary }}</p>
  <p>Saved: {{ stats.saved }} | Queries: {{ stats.queries_done }}/{{ stats.queries_total }}</p>
  <p>
    {% if running %}<form method="post" action="{{ url_for('stop') }}" style="display:inline"><button class="sec" type="submit">Stop</button></form>{% endif %}
    <a class="btn sec" href="{{ url_for('index') }}">Setup</a>
    {% if excel %}<a class="btn" href="{{ url_for('download') }}">Download Excel</a>{% endif %}
  </p>
  <div id="log">{{ log }}</div>
</div>
<script>
function poll(){fetch("{{ url_for('api_status') }}").then(r=>r.json()).then(d=>{
  document.getElementById("log").textContent=d.log.join("\\n");
  if(d.running) setTimeout(poll,2500);
}).catch(()=>setTimeout(poll,4000));}
{% if running %}poll();{% endif %}
</script>
{% endif %}
</body></html>
"""


@app.route("/")
def index():
    return render_template_string(
        PAGE, page="form",
        regions=list(REGIONS.keys()),
        industries=INDUSTRY_OPTIONS,
        services=SERVICE_OPTIONS,
        running=_job["running"],
        excel=os.path.exists(OUTPUT_EXCEL),
    )


@app.route("/start", methods=["POST"])
def start():
    if _job["running"]:
        flash("Already running.", "error")
        return redirect(url_for("status"))
    lead = request.form.get("lead_type", "1")
    regions = request.form.getlist("regions") or ["UK"]
    industries = request.form.getlist("industries") or ["1"]
    custom = request.form.get("custom_terms", "")
    offers = request.form.getlist("offers") or list(SERVICE_OPTIONS.keys())
    configure_from_form(lead, regions, industries, custom, offers)
    Thread(target=run_scrape_job, daemon=True).start()
    flash("Scrape started.", "success")
    return redirect(url_for("status"))


@app.route("/stop", methods=["POST"])
def stop():
    with _job_lock:
        _job["stop"] = True
    flash("Stop requested.", "success")
    return redirect(url_for("status"))


@app.route("/status")
def status():
    with _job_lock:
        return render_template_string(
            PAGE, page="status",
            running=_job["running"],
            summary=_job.get("config_summary") or "—",
            stats=_job["stats"],
            log="\n".join(_job["log"]) or "(waiting…)",
            excel=os.path.exists(OUTPUT_EXCEL),
        )


@app.route("/api/status")
def api_status():
    with _job_lock:
        return jsonify(running=_job["running"], stats=_job["stats"], log=_job["log"][-200:], error=_job["error"])


@app.route("/download")
def download():
    if not os.path.exists(OUTPUT_EXCEL):
        flash("No Excel yet.", "error")
        return redirect(url_for("index"))
    return send_file(OUTPUT_EXCEL, as_attachment=True, download_name=OUTPUT_EXCEL)


if __name__ == "__main__":
    if "--crm-test" in sys.argv:
        crm_test()
    else:
        port = int(os.environ.get("PORT", 5000))
        print(f"Open http://0.0.0.0:{port}")
        app.run(host="0.0.0.0", port=port, debug=False, use_reloader=False)
