"""
Leads Scraper — Flask version (Railway / Docker ready)
Setup form in browser, scrape runs in background. Single file (design is built in).

Sources : Google Maps  +  web search fallback (Google / Bing / DuckDuckGo) that looks for
          Facebook, Instagram, Threads, LinkedIn and X pages of businesses WITHOUT a website.
Output  : UK_Leads.xlsx + optional CRM

Environment variables (Railway -> Variables, or set in your terminal when running locally):
  LOGIN_USER    login username   (default is set in the script)
  LOGIN_EMAIL   login email      (default is set in the script, username or email both work)
  LOGIN_PASS    login password   (optional: overrides the built-in password, set it on Railway)
  SECRET_KEY    long random text used to sign the login cookie (recommended on Railway)
  CRM_API_KEY   your CRM API key (required for CRM sync)
  CRM_API_URL   optional, overrides the CRM endpoint
  DATA_DIR      optional, folder for the Excel file (set to /data when a Volume is mounted there)
  PORT          set by Railway automatically

Run locally (Windows PowerShell):
  pip install flask pandas openpyxl requests playwright psutil
  playwright install chromium
  $env:CRM_API_KEY="f33f43afb069e63f3ab24b0ee3d121275736ccae9531ef54"
  python leads_scraper.py
  -> http://127.0.0.1:5000

(psutil is optional but recommended: it lets the Stop button close the browser instantly.)
"""
from playwright.sync_api import sync_playwright
from openpyxl.worksheet.datavalidation import DataValidation
from flask import Flask, request, redirect, url_for, flash, send_file, render_template_string, jsonify, session
from threading import Thread, Lock
import hmac
import secrets
from datetime import datetime
from urllib.parse import quote, urlparse, parse_qs, unquote
import pandas as pd
import requests
import base64
import shutil
import time
import re
import os
import sys

# ── job state + live logging (every print() also goes to the web page) ─────
_job_lock = Lock()
_job = {
    "running": False,
    "stop": False,
    "log": [],
    "stats": {"saved": 0, "queries_done": 0, "queries_total": 0},
    "config_summary": "",
    "error": None,
}

_real_print = print


def print(*args, **kwargs):  # noqa: A001  (intentionally shadows the builtin)
    """Print to the server log (unbuffered) AND to the web status page."""
    kwargs.pop("flush", None)
    try:
        _real_print(*args, flush=True, **kwargs)
    except Exception:
        pass
    msg = kwargs.get("sep", " ").join(str(a) for a in args)
    with _job_lock:
        _job["log"].append(msg)
        if len(_job["log"]) > 400:
            _job["log"] = _job["log"][-300:]


def job_print(msg):
    print(msg)


# ── instant stop ────────────────────────────────────────────────────────────
class StopJob(BaseException):
    """Raised when the Stop button was pressed. BaseException on purpose, so the many
    'except Exception' blocks in the scraper can never swallow it."""


def check_stop():
    if _job.get("stop"):
        raise StopJob()


def nap(seconds):
    """Sleep, but wake up immediately when Stop is pressed."""
    end = time.time() + seconds
    while True:
        check_stop()
        left = end - time.time()
        if left <= 0:
            return
        time.sleep(min(0.2, left))


def kill_browser():
    """Close the Chromium processes right away so any page load in progress fails at once.
    Needs psutil (optional). Without it, the scraper still stops at the next step."""
    try:
        import psutil
    except ImportError:
        return
    try:
        for child in psutil.Process().children(recursive=True):
            try:
                name = child.name().lower()
                if "chrom" in name or "headless" in name:
                    child.kill()
            except Exception:
                pass
    except Exception:
        pass


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

# ── data folder (use a Railway Volume: mount it at /data and set DATA_DIR=/data) ──
DATA_DIR = os.environ.get("DATA_DIR", ".")
os.makedirs(DATA_DIR, exist_ok=True)

EXTRA_LOCATIONS_FILE = os.path.join(DATA_DIR, "extra_locations.txt")


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
    """Apply web-form choices."""
    global LEAD_FILTER, SEARCH_TERMS, OFFER_SERVICES, ACTIVE_REGIONS, LOCATIONS
    LEAD_FILTER = "no_website" if lead_choice == "1" else "with_website"

    region_names = list(REGIONS)
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


def notes_text(industry):
    return f"Industry: {industry} | Offer: {', '.join(OFFER_SERVICES)}"


MAX_SCROLLS = 15
MAX_PLACES_PER_QUERY = 80
HEADLESS = True
OUTPUT_EXCEL = os.path.join(DATA_DIR, "UK_Leads.xlsx")
SEED_EXCEL = "UK_Leads.xlsx"        # copy shipped in the repo, used once to seed an empty /data volume
DEBUG_SHOT = os.path.join(DATA_DIR, "debug.png")
DELAY = 3
SAVE_EVERY = 5

# Page loading: shorter timeouts so a slow page never blocks the run for long
NAV_TIMEOUT = 30000          # ms for each page.goto attempt
RECYCLE_EVERY = 10           # open a fresh browser tab after this many searches (prevents crashes)
NAV_RETRIES = 1              # extra attempts after the first one
BLOCK_MAP_TILES = True        # do not draw the map itself (lighter + far fewer browser crashes)
MAP_TILE_PARTS = ("/maps/vt", "/vt/pb", "khms", "/kh?v=", "StaticMapService", "/maps/vt/stream")
BLOCK_HEAVY_RESOURCES = True # do not load images / media / fonts (much faster pages)


ONLY_WITH_CONTACT = True

USE_SEARCH_FALLBACK = True

# Web search for businesses WITHOUT a website (Facebook, Instagram, Threads, LinkedIn, X).
# "always"  -> run after every Maps query      "no_leads" -> only when Maps found nothing
WEB_FALLBACK = "always"
SEARCH_ENGINES = ["google", "bing", "duckduckgo"]
FALLBACK_MAX_RESULTS = 15
# Sites searched with  "<service> <city> site:<site>"  (remove one to make the run faster)
FALLBACK_SITES = ["facebook.com", "instagram.com", "threads.net", "linkedin.com/company", "linkedin.com/in", "x.com"]

API_ENABLED = True
API_URL = os.getenv("CRM_API_URL", "https://crm.rhoflow.net/api/external/leads")
API_KEY = os.getenv("CRM_API_KEY", "f33f43afb069e63f3ab24b0ee3d121275736ccae9531ef54").strip()
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
BACKUP_FILE = os.path.join(DATA_DIR, "UK_Leads_old_backup.xlsx")

EMAIL_RE = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
BAD_EMAIL_WORDS = ["example", "sentry", "wix", "wordpress", "domain",
                   ".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", "schema.org", "google"]
PHONE_RE = (r"(?:\+\d{1,3}[\s.-]?\(?\d{1,4}\)?(?:[\s.-]?\d{2,5}){2,4}"
            r"|\(?\b0\d{1,4}\)?[\s.-]?\d{3,4}[\s.-]?\d{3,4}"
            r"|\(?\b[2-9]\d{2}\)?[\s.-]\d{3}[\s.-]\d{4})")

SOCIAL_HOSTS = ("facebook.com", "instagram.com", "linkedin.com", "twitter.com", "x.com",
                "youtube.com", "tiktok.com", "pinterest.com", "threads.net", "threads.com")
IGNORE_HOSTS = SOCIAL_HOSTS + ("google.", "goo.gl", "bit.ly", "wa.me", "whatsapp.com", "linktr.ee",
                               "fb.me", "fb.com", "duckduckgo", "bing.com", "microsoft.com",
                               "w3.org", "schema.org", "gstatic", "fbcdn", "fbsbx", "cloudfront")


def clean(text):
    if not text:
        return ""
    return re.sub(r"\s+", " ", str(text)).strip()


def short_err(e):
    """Short, readable version of a Playwright error (no 'Call log', no long URLs)."""
    msg = clean(str(e).split("Call log")[0])
    if "Timeout" in msg:
        return "page timed out"
    m = re.search(r"net::\w+", msg)
    if m:
        return f"network error ({m.group(0)})"
    if "closed" in msg.lower():
        return "browser closed"
    return msg[:110]


PAGE_CRASHED = False


def mem_info():
    """Memory limit / usage of this container in MB (Railway = cgroup files). '' if unknown."""
    def rd(*paths):
        for pth in paths:
            try:
                v = open(pth).read().strip()
                if v.isdigit() and int(v) < 1 << 50:
                    return int(v) / 1048576
            except Exception:
                pass
        return None
    lim = rd("/sys/fs/cgroup/memory.max", "/sys/fs/cgroup/memory/memory.limit_in_bytes")
    use = rd("/sys/fs/cgroup/memory.current", "/sys/fs/cgroup/memory/memory.usage_in_bytes")
    if lim is None and use is None:
        return ""
    return f"RAM used {use:.0f} MB / limit {lim:.0f} MB" if lim and use else \
        (f"RAM used {use:.0f} MB" if use else f"RAM limit {lim:.0f} MB")


def note_crash(e):
    """True (and remembers it) when Chromium reported 'Page crashed' / 'Target closed'.
    The run loop then opens a fresh browser tab, because a crashed tab never recovers."""
    global PAGE_CRASHED
    m = str(e).lower()
    if "crash" in m or "target closed" in m or "has been closed" in m:
        PAGE_CRASHED = True
        return True
    return False


def safe_goto(page, url, timeout=None, retries=None, wait_until="domcontentloaded"):
    """page.goto with a timeout and retries. Raises only after all attempts fail.
    Stops immediately when the Stop button was pressed."""
    timeout = timeout or NAV_TIMEOUT
    retries = NAV_RETRIES if retries is None else retries
    for attempt in range(retries + 1):
        check_stop()
        try:
            page.goto(url, timeout=timeout, wait_until=wait_until)
            return
        except Exception:
            check_stop()
            if attempt >= retries:
                raise
            nap(1 + attempt)


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
            nap(10)
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
        # First run on a fresh volume: start from the Excel shipped in the repo (keeps old leads / resume)
        if os.path.exists(SEED_EXCEL) and os.path.abspath(SEED_EXCEL) != os.path.abspath(OUTPUT_EXCEL):
            shutil.copy2(SEED_EXCEL, OUTPUT_EXCEL)
            print(f"📋 Copied {SEED_EXCEL} -> {OUTPUT_EXCEL} (first run on this data folder)")
            return
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
    with _job_lock:
        _job["stats"]["saved"] += len(fresh)
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


def _post_lead(row, headers):
    """Send one lead. If the CRM answers with a server error (500) or rejects the data, retry with
    fewer fields (no email / social links, then only the basics). Returns (reply, what_was_removed)."""
    full = _crm_payload(row)
    lite = dict(full)
    for k in ("email", "instagram", "facebook", "linkedin", "twitter", "youtube"):
        lite[k] = ""
    lite["notes"] = (lite.get("notes") or "")[:200]
    basic = {k: "" for k in full}
    basic.update({
        "location": full.get("location", ""),
        "lead_type": full.get("lead_type", ""),
        "company_name": re.sub(r"[^\x20-\x7E]", "", full.get("company_name", "")).strip() or full.get("company_name", ""),
        "phone_number": re.sub(r"[^\d+]", "", full.get("phone_number", "")),
        "status": full.get("status", "new"),
    })
    r = None
    for label, payload in (("", full), ("email, social links", lite), ("everything except name/phone", basic)):
        r = requests.post(API_URL, json=payload, headers=headers, timeout=15)
        if r.status_code < 500 and r.status_code not in (400, 422):
            return r, label
    return r, "failed"


def sync_to_crm():
    """Send leads that are not yet in the CRM. Failed leads are retried next time,
    sent leads are never sent twice. Every reply is printed / logged so problems are visible."""
    if not API_ENABLED or _job.get("stop"):
        return
    if not API_KEY or "PASTE_" in API_URL:
        print("⚠️ CRM_API_KEY / CRM URL is not set, CRM sync skipped (data is saved in Excel).")
        return

    leads_df, state_df = load_data()
    if leads_df.empty:
        return

    smap = {k: [p, s] for k, p, s in zip(state_df["key"], state_df["place_id"], state_df["crm_synced"])}
    keys = [make_key(n, p, a) for n, p, a in zip(leads_df["name"], leads_df["phone"], leads_df["address"])]
    pending = [(i, k) for i, k in zip(leads_df.index, keys) if smap.get(k, ["", "No"])[1] not in ("Yes", "Error")]
    if not pending:
        return

    print(f"📡 CRM: sending {len(pending)} pending leads...")

    def mark(k):
        smap.setdefault(k, ["", "No"])[1] = "Yes"

    def mark_err(k):
        smap.setdefault(k, ["", "No"])[1] = "Error"

    def state_now():
        return pd.DataFrame([{"key": k, "place_id": v[0], "crm_synced": v[1]}
                             for k, v in smap.items()], columns=STATE_COLS)

    headers = _crm_headers()
    created = duplicates = failed = errored = 0
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
                if _job.get("stop"):
                    break
                row = leads_df.loc[i].to_dict()
                r, used = _post_lead(row, headers)
                outcome = _crm_outcome(r)
                if outcome == "ok":
                    mark(k)
                    created += 1
                    if used:
                        print(f"   ✉️  {row['name'][:35]}: CRM had trouble with some fields, lead sent without: {used}")
                elif outcome == "duplicate":
                    mark(k)
                    duplicates += 1
                    print(f"   ♻️  {row['name'][:35]}: CRM says it already exists (409)")
                elif outcome == "auth":
                    print(f"❌ CRM rejected the API key ({r.status_code}). Check CRM_API_KEY / API_KEY_HEADER.")
                    failed += 1
                    break
                else:
                    failed += 1
                    errored += 1
                    mark_err(k)      # never retried again, so the log does not repeat this error
                    print(f"❌ {row['name'][:35]}: CRM refuses this lead ({r.status_code}), kept in Excel only")
                if (created + duplicates) and (created + duplicates) % 25 == 0:
                    write_data(leads_df, state_now())
                time.sleep(0.2)
    except requests.exceptions.RequestException as e:
        print(f"❌ Could not connect to the CRM: {str(e)[:150]}")
        print(f"   Check CRM_API_URL ({API_URL}). Run  python leads_scraper.py --crm-test  to diagnose.")
    except Exception as e:
        print(f"❌ CRM sync error: {str(e)[:150]}")
    finally:
        if created or duplicates or errored:
            write_data(leads_df, state_now())
        total_in_crm = sum(1 for k in keys if smap.get(k, ["", "No"])[1] == "Yes")
        if created or duplicates:
            extra = f" | {duplicates} already existed" if duplicates else ""
            print(f"📡 CRM saved | +{created} new | Total: {total_in_crm}{extra}")
        left = len(pending) - created - duplicates - errored
        if left and not _job.get("stop"):
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
        safe_goto(page, website, timeout=15000, retries=0)
        nap(1.5)
        for a in page.locator("a[href]").all():
            check_stop()
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
        check_stop()
        parts = [p.strip() for p in (address or "").split(",") if p.strip()]
        place = " ".join(parts[-2:]) if parts else ""
        q = f"{name} {place} contact email phone"
        page.goto("https://html.duckduckgo.com/html/?q=" + quote(q),
                  timeout=15000, wait_until="domcontentloaded")
        nap(1.5)

        name_key = name.lower()[:15]
        for block in page.locator("div.result").all()[:8]:
            check_stop()
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
                check_stop()
                page.goto(data["facebook"], timeout=12000, wait_until="domcontentloaded")
                nap(1.5)
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
LI_SKIP = ("/jobs", "/posts", "/pulse", "/feed", "/learning", "/school", "/showcase", "/events",
           "/groups", "/directory", "/salary", "/login", "/signup", "/legal", "/help")
X_SKIP = ("home", "search", "explore", "i", "intent", "share", "hashtag", "login", "signup", "settings", "tos", "privacy")
URL_RE = r"(?:https?://|www\.)[a-z0-9][a-z0-9.-]*\.[a-z]{2,}(?:/[^\s]*)?"
GOOGLE_BLOCKED = False


def _host_is(host, domain):
    return host == domain or host.endswith("." + domain)


def normalize_social(url):
    """Turn a Facebook / Instagram / Threads / LinkedIn / X result into the profile root URL.
    Posts, photos, groups, job pages ... -> None.  Returns (platform, root_url)."""
    try:
        u = urlparse(url or "")
        host = u.netloc.lower()
        path = u.path.strip("/")
        parts = path.split("/") if path else []

        if _host_is(host, "facebook.com"):
            if not path or any(x in "/" + path + "/" for x in FB_SKIP):
                return None
            if parts[0] in ("pages", "people") and len(parts) >= 3:
                return "facebook", f"https://www.facebook.com/{parts[0]}/{parts[1]}/{parts[2]}"
            return "facebook", f"https://www.facebook.com/{parts[0]}"

        if _host_is(host, "instagram.com"):
            if not parts or parts[0] in ("p", "reel", "reels", "explore", "tv", "stories", "accounts"):
                return None
            return "instagram", f"https://www.instagram.com/{parts[0]}/"

        if _host_is(host, "threads.net") or _host_is(host, "threads.com"):
            if not parts or not parts[0].startswith("@") or "post" in parts:
                return None
            return "threads", f"https://www.threads.net/{parts[0]}"

        if _host_is(host, "linkedin.com"):
            if len(parts) >= 2 and parts[0] in ("company", "in") and not any(
                    x in "/" + path for x in LI_SKIP):
                return "linkedin", f"https://www.linkedin.com/{parts[0]}/{parts[1]}"
            return None

        if _host_is(host, "x.com") or _host_is(host, "twitter.com"):
            if not parts or parts[0].lower() in X_SKIP or "status" in parts:
                return None
            return "twitter", f"https://x.com/{parts[0]}"
    except Exception:
        pass
    return None


def name_from_title(title, platform):
    t = clean(title)
    if platform == "instagram":
        t = t.split(" (@")[0].split(" • ")[0]
    if platform == "threads":
        t = t.split(" (@")[0]
    if platform == "linkedin":
        t = re.split(r"\s[-|–•·]\s", t)[0]
    t = re.split(r"\s[-|–•·]\s", t)[0]
    t = re.sub(r"\b(Facebook|Instagram|Threads|LinkedIn|Twitter)\b", "", t, flags=re.I).strip(" -|•·")
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
                nap(1)
                return
        except Exception:
            pass


def _google_results(page, q):
    global GOOGLE_BLOCKED
    check_stop()
    page.goto("https://www.google.com/search?num=20&hl=en&q=" + quote(q),
              timeout=20000, wait_until="domcontentloaded")
    nap(1.5)
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
                nap(3)
                if not blocked():
                    break
        if blocked():
            GOOGLE_BLOCKED = True
            print("   ⚠️ Google is blocking us, Bing / DuckDuckGo will be used for the rest of this run.")
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
    check_stop()
    page.goto("https://html.duckduckgo.com/html/?q=" + quote(q),
              timeout=20000, wait_until="domcontentloaded")
    nap(1.5)
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


def _decode_bing(href):
    """Bing sometimes wraps result links (bing.com/ck/a?...&u=a1<base64>)."""
    try:
        if "bing.com/ck/a" in href:
            u = parse_qs(urlparse(href).query).get("u", [""])[0]
            if u.startswith("a1"):
                u = u[2:]
                u += "=" * (-len(u) % 4)
                return base64.urlsafe_b64decode(u).decode("utf-8", "ignore")
    except Exception:
        pass
    return href


def _bing_results(page, q):
    check_stop()
    page.goto("https://www.bing.com/search?count=30&setlang=en&q=" + quote(q),
              timeout=20000, wait_until="domcontentloaded")
    nap(1.5)
    _accept_consent(page)
    out = []
    for b in page.locator("li.b_algo").all()[:FALLBACK_MAX_RESULTS * 2]:
        try:
            a = b.locator("h2 a").first
            out.append({"title": clean(a.inner_text(timeout=1500)),
                        "url": _decode_bing(a.get_attribute("href") or ""),
                        "text": clean(b.inner_text(timeout=1500))})
        except Exception:
            continue
    return out


_ENGINE_FUNCS = {"google": _google_results, "bing": _bing_results, "duckduckgo": _ddg_results}


def search_web(page, q, merge=False):
    """Run a web search. With merge=True results of several engines are combined until
    enough are found (more leads); otherwise the first engine that answers is used."""
    collected, seen = [], set()
    for engine in SEARCH_ENGINES:
        check_stop()
        if engine == "google" and GOOGLE_BLOCKED:
            continue
        func = _ENGINE_FUNCS.get(engine)
        if not func:
            continue
        try:
            results = func(page, q)
        except Exception:
            results = []
        for r in results:
            if r["url"] and r["url"] not in seen:
                seen.add(r["url"])
                collected.append(r)
        if collected and (not merge or len(collected) >= FALLBACK_MAX_RESULTS):
            break
    return collected


def web_fallback(page, service, city, seen_ids, seen_keys, industry="", out=None):
    """Find businesses WITHOUT a website from anywhere on the web: Facebook, Instagram, Threads,
    LinkedIn and X pages found through Google / Bing / DuckDuckGo. A business is kept only if its
    page does not point to an own website. New leads are appended to `out` as soon as found."""
    leads = out if out is not None else []
    candidates = {}

    queries = [f"{service} {city} site:{site}" for site in FALLBACK_SITES]
    queries.append(f"{service} {city} facebook instagram linkedin contact")
    for q in queries:
        check_stop()
        for r in search_web(page, q, merge=True)[:FALLBACK_MAX_RESULTS]:
            n = normalize_social(r["url"])
            if not n:
                continue
            plat, root = n
            if root not in candidates:
                candidates[root] = (plat, name_from_title(r["title"], plat), r.get("text", ""))
    print(f"   🔎 Web search: {len(candidates)} social pages found")

    city_token = city.split(",")[0].strip().lower()
    for root, (plat, name, snippet) in candidates.items():
        check_stop()
        pid = place_id(root)
        if pid in seen_ids or not name:
            continue
        try:
            body, html = "", ""
            try:
                page.goto(root, timeout=12000, wait_until="domcontentloaded")
                nap(1.5)
                body = page.inner_text("body")[:6000]
                html = page.content()
            except Exception:
                check_stop()
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

            note = notes_text(industry) + f" | Found via web search ({plat})"
            if plat == "threads":
                note += f" | Threads: {root}"
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
                "linkedin": root if plat == "linkedin" else "",
                "twitter": root if plat == "twitter" else "",
                "youtube": "",
                "notes": note,
                "maps_url": root,
            })
            seen_ids.add(pid)
            seen_keys.add(key)
            print(f"   🔎 📞 NO-WEBSITE ({plat}) {name[:40]} | {phone or '-'} | {email or '-'}")
        except Exception as e:
            check_stop()
            if note_crash(e):
                print("   🔎 ⚠️ browser tab crashed, opening a fresh one...")
                break
            print(f"   🔎 ⚠️ {short_err(e)} (skipped)")
            continue
    return leads


def scrape(page, query, location, seen_ids, seen_keys, industry=""):
    print(f"\n🔍 [{'WITHOUT website' if LEAD_FILTER == 'no_website' else 'WITH website'}] {query}")
    leads = []
    saved_upto = 0
    stats = {"saved": 0, "no_contact": 0, "skipped_filter": 0}
    completed = False
    stopped = False
    try:
        url = f"https://www.google.com/maps/search/{quote(query)}"
        safe_goto(page, url, timeout=60000)
        _accept_consent(page)
        # wait until the result list (or a single place page) is really there
        try:
            page.wait_for_selector('div[role="feed"], a[href*="/maps/place/"], h1', timeout=20000)
        except Exception:
            check_stop()
            _accept_consent(page)
        nap(1.5)

        try:
            feed = page.locator('div[role="feed"]')
            if feed.count() > 0:
                for _ in range(MAX_SCROLLS):
                    check_stop()
                    feed.evaluate("el => el.scrollTop = el.scrollHeight")
                    nap(1.2)
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
        if not places:
            # Diagnostic: what did Google actually show? (open the debug screenshot on the Status page)
            try:
                print(f"   ⚠️ No places. URL: {page.url[:90]} | Title: {page.title()[:60]}")
                page.screenshot(path=DEBUG_SHOT)
            except Exception:
                pass

        for i, place_url in enumerate(places[:MAX_PLACES_PER_QUERY]):
            check_stop()
            pid = place_id(place_url)
            if pid in seen_ids:
                continue
            try:
                # 'commit' returns as soon as Google answers; we then wait for the name ourselves.
                # This avoids the "Page.goto: Timeout exceeded" errors on slow Maps pages.
                safe_goto(page, place_url, wait_until="commit")

                try:
                    page.wait_for_selector("h1", timeout=15000)
                    name = clean(page.locator("h1").first.inner_text(timeout=5000))
                except Exception:
                    check_stop()
                    print(f"   [{i+1}] ⚠️ page did not load, skipped")
                    continue
                if not name:
                    print(f"   [{i+1}] ⚠️ no name found, skipped")
                    continue
                nap(1.2)

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
                check_stop()
                if note_crash(e):
                    print(f"   [{i+1}] ⚠️ browser tab crashed, opening a fresh one and moving on...")
                    break
                print(f"   [{i+1}] ⚠️ {short_err(e)} (skipped)")
                continue

        completed = bool(places) and not PAGE_CRASHED

    except StopJob:
        stopped = True
    except Exception as e:
        check_stop()
        if note_crash(e):
            print("   ⚠️ browser tab crashed, opening a fresh one...")
        else:
            print(f"❌ Failed: {short_err(e)}")

    # always save what was found so far (also when Stop was pressed)
    if saved_upto < len(leads):
        save_leads(leads[saved_upto:])

    if stopped:
        raise StopJob()

    print(f"   📊 This query: saved {stats['saved']} | "
          f"no contact found {stats['no_contact']} | "
          f"skipped (other lead type) {stats['skipped_filter']}")
    return leads, completed


def run_scrape_job():
    """Background job started from the web form."""
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
        job_print(f"CRM sync: {'ON' if (API_ENABLED and API_KEY) else 'OFF (CRM_API_KEY not set)'}")
        job_print(f"Data file: {OUTPUT_EXCEL}")
        job_print("=" * 60)

        create_empty_excel()
        migrate_old_file()
        seen_ids, seen_keys, done_queries = load_seen()
        job_print(f"📂 {len(seen_keys)} leads already in Excel.")
        if done_queries:
            job_print(f"⏩ {len(done_queries)} queries already done (resume).")
        sync_to_crm()

        job_print("🌐 Starting browser...")
        pending_done = []
        try:
            with sync_playwright() as p:
                BROWSER_ARGS = ["--disable-blink-features=AutomationControlled",
                                "--no-sandbox",
                                "--disable-setuid-sandbox",
                                "--disable-dev-shm-usage",
                                "--disable-gpu",
                                "--disable-extensions",
                                "--no-first-run",
                                "--mute-audio"]
                bholder = {"b": None}

                def launch_browser():
                    """(Re)start Chromium completely. Frees all memory used by the old one."""
                    if bholder["b"] is not None:
                        try:
                            bholder["b"].close()
                        except BaseException:
                            pass
                    bholder["b"] = p.chromium.launch(headless=HEADLESS, args=BROWSER_ARGS)
                    return bholder["b"]

                browser = launch_browser()
                mi = mem_info()
                if mi:
                    job_print(f"🧠 {mi}")

                def _route(route):
                    try:
                        if _job.get("stop"):
                            route.abort()
                        elif BLOCK_HEAVY_RESOURCES and route.request.resource_type in ("image", "media", "font"):
                            route.abort()
                        elif BLOCK_MAP_TILES and any(x in route.request.url for x in MAP_TILE_PARTS):
                            route.abort()       # map drawing is what makes Chromium crash; the contact data is not in the tiles
                        else:
                            route.continue_()
                    except Exception:
                        pass

                holder = {"ctx": None, "state": None}

                def fresh_page(restart=False):
                    """New browser context + tab. Used at start, after a crash, and every
                    RECYCLE_EVERY searches to free memory. restart=True also relaunches Chromium."""
                    global PAGE_CRASHED
                    if holder["ctx"] is not None:
                        try:
                            holder["state"] = holder["ctx"].storage_state()   # keeps Google consent cookies
                        except BaseException:
                            pass
                        try:
                            holder["ctx"].close()
                        except BaseException:
                            pass
                        holder["ctx"] = None
                    if restart:
                        launch_browser()
                        nap(2)
                    ctx = bholder["b"].new_context(
                        storage_state=holder.get("state"),
                        viewport={"width": 1024, "height": 720},
                        locale="en-US",
                        user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                                    "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"),
                    )
                    ctx.set_default_navigation_timeout(60000)
                    ctx.set_default_timeout(30000)
                    ctx.route("**/*", _route)
                    pg = ctx.new_page()
                    pg.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
                    holder["ctx"] = ctx
                    PAGE_CRASHED = False
                    return pg

                page = fresh_page()
                searches_since_fresh = 0
                job_print("✅ Browser ready, starting searches...")

                try:
                    for city in LOCATIONS:
                        check_stop()
                        for service, industry in SEARCH_TERMS:
                            check_stop()
                            query = f"{service} in {city}"
                            done_key = query if LEAD_FILTER == "no_website" else query + " [with website]"
                            if done_key in done_queries:
                                with _job_lock:
                                    _job["stats"]["queries_done"] += 1
                                continue

                            if PAGE_CRASHED or searches_since_fresh >= RECYCLE_EVERY:
                                if PAGE_CRASHED:
                                    job_print("🔄 Browser tab crashed, restarting Chromium...")
                                page = fresh_page(restart=PAGE_CRASHED)
                                searches_since_fresh = 0
                            searches_since_fresh += 1

                            crash_tries = 0
                            while True:
                                leads, completed = scrape(page, query, city, seen_ids, seen_keys, industry)
                                if PAGE_CRASHED and crash_tries < 2:
                                    crash_tries += 1
                                    mi = mem_info()
                                    job_print(f"🔄 Browser crashed{(' (' + mi + ')') if mi else ''}. "
                                              f"Restarting Chromium and retrying this search ({crash_tries}/2)...")
                                    page = fresh_page(restart=True)
                                    searches_since_fresh = 0
                                    continue
                                break

                            fb_ran = False
                            if LEAD_FILTER == "no_website" and not PAGE_CRASHED and \
                               (WEB_FALLBACK == "always" or (WEB_FALLBACK == "no_leads" and not leads)):
                                fb_ran = True
                                fb_leads = []
                                try:
                                    web_fallback(page, service, city, seen_ids, seen_keys, industry, out=fb_leads)
                                finally:
                                    # saved even when Stop is pressed in the middle
                                    if fb_leads:
                                        save_leads(fb_leads)

                            if (completed or fb_ran) and not PAGE_CRASHED:
                                done_queries.add(done_key)
                                pending_done.append(done_key)
                            with _job_lock:
                                _job["stats"]["queries_done"] += 1
                            nap(DELAY)

                        mark_queries_done(pending_done)
                        pending_done = []
                finally:
                    try:
                        if bholder["b"] is not None:
                            bholder["b"].close()
                    except BaseException:
                        pass
        finally:
            if pending_done:
                try:
                    mark_queries_done(pending_done)
                except Exception:
                    pass
        job_print(f"\n✅ Done! Check file → {OUTPUT_EXCEL}")
    except StopJob:
        job_print("\n⏹ Stopped. Everything found so far is saved.")
    except Exception as e:
        if _job.get("stop"):
            job_print("\n⏹ Stopped. Everything found so far is saved.")
        else:
            job_print(f"❌ Fatal: {short_err(e)}")
            with _job_lock:
                _job["error"] = short_err(e)
    finally:
        with _job_lock:
            _job["running"] = False


# ── Flask ──────────────────────────────────────────────────────────────────
app = Flask(__name__)

# ── login ───────────────────────────────────────────────────────────────────
# Login works with the username OR the email below. The password is stored only as a hash.
# To change it on Railway set LOGIN_USER / LOGIN_EMAIL / LOGIN_PASS in Variables.
from datetime import timedelta
from functools import wraps
from werkzeug.security import check_password_hash, generate_password_hash

AUTH_USER = os.environ.get("LOGIN_USER", "venom09798").strip()
AUTH_EMAIL = os.environ.get("LOGIN_EMAIL", "venom09798@gmail.com").strip()
_env_pass = os.environ.get("LOGIN_PASS")
AUTH_HASH = (generate_password_hash(_env_pass, method="pbkdf2:sha256") if _env_pass else
             "pbkdf2:sha256:1000000$vTayL2p2gyb1SYR2$a0da9531052df0ba668894da191ee26f273a41b97a6611618c781ecbb7d33605")


def _load_secret_key():
    """Cookie signing key: SECRET_KEY variable, else a random key kept in the data folder
    (never a fixed text from the source code, so nobody can forge a login cookie)."""
    key = os.environ.get("SECRET_KEY")
    if key:
        return key
    path = os.path.join(DATA_DIR, ".secret_key")
    try:
        if os.path.exists(path):
            with open(path, encoding="utf-8") as f:
                saved = f.read().strip()
            if saved:
                return saved
        key = secrets.token_hex(32)
        with open(path, "w", encoding="utf-8") as f:
            f.write(key)
        return key
    except Exception:
        return secrets.token_hex(32)


app.secret_key = _load_secret_key()
app.config.update(SESSION_COOKIE_HTTPONLY=True, SESSION_COOKIE_SAMESITE="Lax",
                  PERMANENT_SESSION_LIFETIME=timedelta(days=30))

_login_fails = {}            # ip -> [failed attempts, locked until]
MAX_FAILS = 5
LOCK_SECONDS = 300


def _client_ip():
    fwd = request.headers.get("X-Forwarded-For", "")
    return (fwd.split(",")[0].strip() if fwd else request.remote_addr) or "?"


def _check_login(identifier, password):
    ident = (identifier or "").strip().lower()
    user_ok = hmac.compare_digest(ident, AUTH_USER.lower()) or hmac.compare_digest(ident, AUTH_EMAIL.lower())
    pass_ok = check_password_hash(AUTH_HASH, password or "")      # always checked (same timing)
    return user_ok and pass_ok


def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("auth"):
            if request.path.startswith("/api/") or request.headers.get("X-Requested-With") == "fetch":
                return jsonify(error="login required"), 401
            return redirect(url_for("login", next=request.path))
        return f(*args, **kwargs)
    return wrapper


LOGIN_PAGE = r"""
<!DOCTYPE html>
<html lang="en"><head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"/>
<meta name="theme-color" content="#084F3A"/>
<title>Sign in · Leads Scraper</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<style>
:root{--brand:#0B6B4E;--brand-2:#128A66;--dark:#062F23;--ink:#13201A;--muted:#62716A;--line:#DCE4DB;--bad:#B42318;--bad-bg:#FDECEA}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%}
body{margin:0;min-height:100vh;min-height:100dvh;display:grid;place-items:center;padding:20px;color:var(--ink);overflow-x:hidden;
  font:15px/1.5 "Plus Jakarta Sans",system-ui,-apple-system,"Segoe UI",sans-serif;
  background:radial-gradient(1100px 600px at 12% -10%,#14996F 0%,transparent 60%),
             radial-gradient(900px 600px at 100% 110%,#0E7A58 0%,transparent 55%),
             linear-gradient(160deg,#062F23,#0A4A38 55%,#0B6B4E)}
.orb{position:fixed;border-radius:50%;filter:blur(60px);opacity:.35;pointer-events:none;animation:float 14s ease-in-out infinite}
.orb.a{width:280px;height:280px;background:#35D399;top:-80px;right:-60px}
.orb.b{width:240px;height:240px;background:#0E7A58;bottom:-80px;left:-60px;animation-delay:-6s}
@keyframes float{0%,100%{transform:translateY(0)}50%{transform:translateY(24px)}}
@media (prefers-reduced-motion:reduce){.orb{animation:none}}
.box{position:relative;width:100%;max-width:410px;background:#fff;border-radius:26px;padding:34px 30px 28px;
  box-shadow:0 30px 80px rgba(0,0,0,.35),0 0 0 1px rgba(255,255,255,.06)}
.logo{width:56px;height:56px;border-radius:17px;background:linear-gradient(135deg,var(--brand),var(--brand-2));display:grid;place-items:center;
  margin:0 auto 16px;box-shadow:0 10px 24px rgba(11,107,78,.35)}
.logo svg{width:28px;height:28px}
h1{margin:0;text-align:center;font-size:1.55rem;font-weight:800;letter-spacing:-.025em}
.sub{margin:6px 0 24px;text-align:center;color:var(--muted);font-size:.9rem}
.err{background:var(--bad-bg);color:var(--bad);border-radius:12px;padding:11px 14px;margin-bottom:16px;font-weight:600;font-size:.88rem}
label.lbl{display:block;font-size:.8rem;font-weight:700;margin:0 0 6px;color:#2B3B33}
.inp{position:relative;margin-bottom:16px}
.inp svg.ic{position:absolute;left:14px;top:50%;transform:translateY(-50%);width:18px;height:18px;stroke:#8A9A92;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;pointer-events:none}
.inp input{width:100%;padding:13px 44px 13px 42px;border:1.5px solid var(--line);border-radius:13px;font:inherit;font-size:16px;background:#F8FAF8;color:var(--ink);transition:border-color .15s,background .15s}
.inp input:focus{outline:none;border-color:var(--brand);background:#fff;box-shadow:0 0 0 4px rgba(18,138,102,.15)}
.eye{position:absolute;right:6px;top:50%;transform:translateY(-50%);background:none;border:0;padding:8px;cursor:pointer;border-radius:10px;color:#8A9A92;display:grid;place-items:center}
.eye:hover{background:#EDF2EC;color:var(--brand)}
.eye svg{width:19px;height:19px;stroke:currentColor;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round}
.row{display:flex;align-items:center;justify-content:space-between;margin:2px 0 20px;font-size:.85rem;color:var(--muted);font-weight:600}
.row label{display:flex;align-items:center;gap:8px;cursor:pointer}
.row input{width:17px;height:17px;accent-color:var(--brand)}
button.go{width:100%;padding:14px;border:0;border-radius:14px;font:inherit;font-weight:800;font-size:1rem;color:#fff;cursor:pointer;
  background:linear-gradient(135deg,var(--brand),var(--brand-2));box-shadow:0 10px 24px rgba(11,107,78,.35);transition:transform .1s,box-shadow .15s}
button.go:hover{box-shadow:0 14px 30px rgba(11,107,78,.42)}
button.go:active{transform:translateY(1px)}
button.go:focus-visible,.eye:focus-visible{outline:3px solid #86D4B6;outline-offset:2px}
.foot{margin:20px 0 0;text-align:center;color:#8A9A92;font-size:.78rem}
@media (max-width:420px){.box{padding:28px 20px 22px;border-radius:22px}h1{font-size:1.4rem}}
</style></head><body>
<div class="orb a"></div><div class="orb b"></div>
<main class="box">
  <div class="logo"><svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/><path d="M11 8v6M8 11h6"/></svg></div>
  <h1>Welcome back</h1>
  <p class="sub">Sign in to open Leads Scraper</p>

  {% if error %}<div class="err" role="alert">{{ error }}</div>{% endif %}

  <form method="post" action="{{ url_for('login') }}" autocomplete="on">
    <input type="hidden" name="next" value="{{ nxt }}">
    <label class="lbl" for="u">Username or email</label>
    <div class="inp">
      <svg class="ic" viewBox="0 0 24 24"><circle cx="12" cy="8" r="4"/><path d="M4 21c0-4 4-6 8-6s8 2 8 6"/></svg>
      <input id="u" name="identifier" type="text" value="{{ ident }}" placeholder="username or email" autocomplete="username" autocapitalize="none" spellcheck="false" required autofocus>
    </div>
    <label class="lbl" for="p">Password</label>
    <div class="inp">
      <svg class="ic" viewBox="0 0 24 24"><rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/></svg>
      <input id="p" name="password" type="password" placeholder="••••••••" autocomplete="current-password" required>
      <button class="eye" type="button" id="eye" aria-label="Show password">
        <svg viewBox="0 0 24 24"><path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12z"/><circle cx="12" cy="12" r="3"/></svg>
      </button>
    </div>
    <div class="row"><label><input type="checkbox" name="remember" value="1" checked> Keep me signed in</label></div>
    <button class="go" type="submit">Sign in</button>
  </form>
  <p class="foot">Private access only</p>
</main>
<script>
(function(){
  var p=document.getElementById("p"), b=document.getElementById("eye");
  b.addEventListener("click", function(){
    var show = p.type === "password";
    p.type = show ? "text" : "password";
    b.setAttribute("aria-label", show ? "Hide password" : "Show password");
  });
})();
</script>
</body></html>
"""


def _safe_next(nxt):
    nxt = nxt or ""
    return nxt if nxt.startswith("/") and not nxt.startswith("//") and "\\" not in nxt else ""


@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("auth"):
        return redirect(url_for("index"))
    error, ident = "", ""
    nxt = _safe_next(request.values.get("next", ""))
    if request.method == "POST":
        ip = _client_ip()
        fails, locked_until = _login_fails.get(ip, [0, 0])
        now = time.time()
        ident = request.form.get("identifier", "")
        if locked_until > now:
            wait = int(locked_until - now) // 60 + 1
            error = f"Too many failed attempts. Try again in {wait} min."
        elif _check_login(ident, request.form.get("password", "")):
            _login_fails.pop(ip, None)
            session.clear()
            session["auth"] = True
            session.permanent = bool(request.form.get("remember"))
            return redirect(nxt or url_for("index"))
        else:
            fails += 1
            if fails >= MAX_FAILS:
                _login_fails[ip] = [0, now + LOCK_SECONDS]
                error = "Too many failed attempts. Try again in 5 min."
            else:
                _login_fails[ip] = [fails, 0]
                error = "Wrong username/email or password."
    return render_template_string(LOGIN_PAGE, error=error, ident=ident, nxt=nxt)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))

PAGE = r"""
<!DOCTYPE html>
<html lang="en"><head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"/>
<meta name="theme-color" content="#0B5C45"/>
<title>Leads Scraper</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
:root{
  --bg:#F3F5F2; --card:#FFFFFF; --ink:#13201A; --muted:#62716A; --line:#E0E6DF;
  --brand:#0B6B4E; --brand-2:#128A66; --brand-dark:#084F3A; --tint:#E3F3EB;
  --bad:#B42318; --bad-bg:#FDECEA; --warn:#B7791F;
  --console:#0E1915; --console-line:#21322B;
  --shadow:0 1px 2px rgba(19,32,26,.06),0 6px 20px rgba(19,32,26,.05);
}
*{box-sizing:border-box}
html{-webkit-text-size-adjust:100%;scroll-behavior:smooth}
body{margin:0;background:var(--bg);color:var(--ink);overflow-x:hidden;
  font:15px/1.5 "Plus Jakarta Sans",system-ui,-apple-system,"Segoe UI",sans-serif}
a{color:inherit}
.wrap{max-width:860px;margin:0 auto;padding:0 18px 48px}

/* ---- top bar ---- */
.hero{background:linear-gradient(135deg,var(--brand-dark),var(--brand) 55%,var(--brand-2));color:#fff;
  padding:22px 0;margin-bottom:22px}
.hero .wrap{padding-bottom:0}
.hero-row{display:flex;align-items:center;justify-content:space-between;gap:14px;flex-wrap:wrap}
.brand{display:flex;align-items:center;gap:12px}
.logo{width:42px;height:42px;border-radius:12px;background:rgba(255,255,255,.16);display:grid;place-items:center;flex:none}
.logo svg{width:22px;height:22px}
.hero h1{margin:0;font-size:1.4rem;font-weight:800;letter-spacing:-.02em;line-height:1.2}
.hero p{margin:2px 0 0;font-size:.85rem;opacity:.82}
nav.tabs{display:flex;gap:4px;background:rgba(0,0,0,.2);padding:4px;border-radius:12px}
nav.tabs a{padding:8px 18px;border-radius:9px;color:rgba(255,255,255,.85);text-decoration:none;font-weight:600;font-size:.88rem;text-align:center}
nav.tabs a.on{background:#fff;color:var(--brand-dark)}
a:focus-visible,button:focus-visible,.opt input:focus-visible+span{outline:3px solid #86D4B6;outline-offset:2px}

.flash{padding:12px 16px;border-radius:12px;margin-bottom:14px;background:var(--tint);color:var(--brand-dark);font-weight:600;box-shadow:var(--shadow)}
.flash.err{background:var(--bad-bg);color:var(--bad)}

/* ---- cards ---- */
.card{background:var(--card);border:1px solid var(--line);border-radius:20px;padding:22px;margin-bottom:14px;box-shadow:var(--shadow)}
.card-head{display:flex;align-items:flex-start;gap:12px;margin-bottom:14px}
.num{width:28px;height:28px;border-radius:9px;background:var(--tint);color:var(--brand-dark);font-weight:800;font-size:.85rem;display:grid;place-items:center;flex:none;margin-top:1px}
.card-head h2{margin:0;font-size:1.02rem;font-weight:700;line-height:1.3}
.card-head p{margin:2px 0 0;color:var(--muted);font-size:.84rem}
.card-head .grow{flex:1;min-width:0}
.mini{display:flex;gap:6px;flex:none}
.mini button{background:none;border:0;color:var(--brand);font:inherit;font-size:.8rem;font-weight:700;padding:4px 6px;cursor:pointer;border-radius:8px}
.mini button:hover{background:var(--tint)}

/* ---- option chips ---- */
.opts{display:flex;flex-wrap:wrap;gap:8px}
.opt{position:relative;cursor:pointer;max-width:100%}
.opt input{position:absolute;opacity:0;inset:0;width:100%;height:100%;margin:0;cursor:pointer}
.opt span{display:inline-flex;align-items:center;gap:7px;max-width:100%;padding:9px 15px;border:1.5px solid var(--line);border-radius:999px;background:#FAFBF9;
  font-size:.88rem;font-weight:600;transition:background .12s,border-color .12s,color .12s}
.opt:hover span{border-color:#B3C4B8}
.opt input:checked+span{background:var(--brand);border-color:var(--brand);color:#fff}
.opt input:checked+span::before{content:"✓";font-size:.8rem}

/* big radio cards */
.choice{display:grid;grid-template-columns:1fr 1fr;gap:10px}
.choice .opt span{display:block;width:100%;border-radius:16px;padding:14px 16px;border-width:2px}
.choice .opt b{display:block;font-size:.98rem;margin-bottom:2px}
.choice .opt small{display:block;font-weight:500;font-size:.8rem;color:var(--muted);line-height:1.35}
.choice .opt input:checked+span{background:var(--tint);border-color:var(--brand);color:var(--ink)}
.choice .opt input:checked+span::before{content:none}
.choice .opt input:checked+span small{color:var(--brand-dark)}

.field{width:100%;margin-top:12px;padding:12px 14px;border:1.5px solid var(--line);border-radius:12px;
  font:inherit;font-size:16px;background:#FAFBF9;color:var(--ink)}
.field:focus{outline:3px solid #86D4B6;outline-offset:1px;border-color:var(--brand)}
#customBox{display:none}
#customBox.show{display:block}

/* ---- buttons ---- */
button,.btn{font:inherit;font-weight:700;border:1.5px solid transparent;border-radius:12px;padding:12px 22px;
  cursor:pointer;text-decoration:none;display:inline-flex;align-items:center;justify-content:center;gap:8px;line-height:1.2;text-align:center}
.btn.primary,button.primary{background:var(--brand);color:#fff}
.btn.primary:hover,button.primary:hover{background:var(--brand-dark)}
.btn.ghost,button.ghost{background:var(--card);color:var(--ink);border-color:var(--line)}
.btn.ghost:hover,button.ghost:hover{border-color:#B3C4B8}
button.danger{background:var(--bad);color:#fff}
button.danger:hover{background:#8F1B12}
button:disabled{opacity:.5;cursor:not-allowed}

.actions{display:flex;flex-wrap:wrap;gap:10px;align-items:center}
.dock{position:sticky;bottom:0;z-index:20;margin:0 -18px;padding:12px 18px calc(12px + env(safe-area-inset-bottom));
  background:linear-gradient(to top,var(--bg) 70%,rgba(243,245,242,0));}
.dock .actions{background:var(--card);border:1px solid var(--line);border-radius:18px;padding:10px;box-shadow:var(--shadow)}
.dock .primary{flex:1 1 auto;min-width:160px}

/* ---- status ---- */
.state{display:flex;align-items:center;gap:10px;font-weight:800;font-size:1.2rem;letter-spacing:-.01em}
.dot{width:12px;height:12px;border-radius:50%;background:#9AA7A0;flex:none}
.dot.run{background:var(--brand-2);animation:pulse 1.6s ease-in-out infinite}
.dot.stopping{background:var(--warn);animation:pulse 0.8s ease-in-out infinite}
@keyframes pulse{0%,100%{box-shadow:0 0 0 0 rgba(18,138,102,.45)}50%{box-shadow:0 0 0 8px rgba(18,138,102,0)}}
@media (prefers-reduced-motion:reduce){.dot.run,.dot.stopping{animation:none}.bar i{transition:none}}
.chips{display:flex;flex-wrap:wrap;gap:6px;margin:10px 0 18px}
.chips span{background:#F1F5F0;border:1px solid var(--line);border-radius:999px;padding:4px 11px;font-size:.78rem;font-weight:600;color:var(--muted)}
.bar{height:12px;background:#E6EBE4;border-radius:999px;overflow:hidden}
.bar i{display:block;height:100%;width:0;background:linear-gradient(90deg,var(--brand),#2FB587);border-radius:999px;transition:width .6s ease}
.pct{display:flex;justify-content:space-between;color:var(--muted);font-size:.8rem;font-weight:600;margin:7px 0 0}
.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin:18px 0}
.stat{background:#F6F8F5;border:1px solid var(--line);border-radius:16px;padding:14px 16px;min-width:0}
.stat b{display:block;font-size:1.8rem;line-height:1.1;letter-spacing:-.03em;font-weight:800}
.stat small{color:var(--muted);font-size:.8rem;font-weight:600}
.stat.hl{background:var(--tint);border-color:#C5E3D4}
.stat.hl b{color:var(--brand-dark)}
.err-box{background:var(--bad-bg);color:var(--bad);border-radius:12px;padding:10px 14px;margin-bottom:14px;display:none;overflow-wrap:anywhere;font-weight:600}

.console{background:var(--console);border-radius:20px;overflow:hidden;margin-top:14px;box-shadow:var(--shadow)}
.console .bar2{display:flex;justify-content:space-between;align-items:center;padding:11px 18px;
  border-bottom:1px solid var(--console-line);color:#8FA79C;font-size:.8rem;font-weight:600}
.console .bar2 .live{display:flex;align-items:center;gap:7px}
.console .bar2 .live i{width:8px;height:8px;border-radius:50%;background:#35D399}
#log{font:12.5px/1.65 "JetBrains Mono",ui-monospace,Consolas,monospace;color:#D5E3DC;padding:14px 18px;
  height:460px;overflow:auto;white-space:pre-wrap;overflow-wrap:anywhere;-webkit-overflow-scrolling:touch}
#log div{padding:0 0 0 10px;border-left:2px solid transparent}
#log .ok{color:#7FE0B0}
#log .warn{color:#F2C26B}
#log .bad{color:#FF8D82;border-left-color:#FF8D82}
#log .q{color:#fff;font-weight:500;margin-top:8px}

/* ---- mobile ---- */
@media (max-width:640px){
  .wrap{padding:0 14px 40px}
  .hero{padding:16px 0;margin-bottom:16px}
  .hero-row{flex-direction:column;align-items:stretch;gap:14px}
  .hero h1{font-size:1.25rem}
  nav.tabs{width:100%}
  nav.tabs a{flex:1;padding:10px 8px}
  .card{padding:16px 14px;border-radius:16px;margin-bottom:12px}
  .choice{grid-template-columns:1fr}
  .opts{gap:7px}
  .opt span{padding:9px 13px;font-size:.86rem}
  .dock{margin:0 -14px;padding:10px 14px calc(10px + env(safe-area-inset-bottom))}
  .dock .actions .btn,.dock .actions button{flex:1 1 calc(50% - 6px)}
  .dock .actions .primary{flex-basis:100%}
  .stats{gap:8px}
  .stat{padding:11px 11px}
  .stat b{font-size:1.35rem}
  .stat small{font-size:.74rem}
  .status-actions .btn,.status-actions button{flex:1 1 calc(50% - 6px)}
  .console{border-radius:16px}
  .console .bar2{padding:10px 14px}
  #log{height:55vh;min-height:300px;padding:12px;font-size:11.5px}
}
@media (max-width:360px){.stats{grid-template-columns:1fr}}
</style></head><body>

<div class="hero"><div class="wrap">
  <div class="hero-row">
    <div class="brand">
      <div class="logo"><svg viewBox="0 0 24 24" fill="none" stroke="#fff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/><path d="M11 8v6M8 11h6"/></svg></div>
      <div>
        <h1>Leads Scraper</h1>
        <p>Find businesses and save their contact details</p>
      </div>
    </div>
    <nav class="tabs">
      <a href="{{ url_for('index') }}" class="{% if page=='form' %}on{% endif %}">Setup</a>
      <a href="{{ url_for('status') }}" class="{% if page!='form' %}on{% endif %}">Status{% if running %} ●{% endif %}</a>
      <a href="{{ url_for('logout') }}">Logout</a>
    </nav>
  </div>
</div></div>

<div class="wrap">

{% with msgs = get_flashed_messages(with_categories=1) %}
  {% for c,m in msgs %}<div class="flash {% if c=='error' %}err{% endif %}">{{ m }}</div>{% endfor %}
{% endwith %}

{% if page=='form' %}
<form method="post" action="{{ url_for('start') }}">

  <section class="card">
    <div class="card-head"><div class="num">1</div><div class="grow"><h2>Lead type</h2><p>Which businesses do you want to collect?</p></div></div>
    <div class="choice">
      <label class="opt"><input type="radio" name="lead_type" value="1" checked><span><b>Without a website</b><small>Maps + Facebook, Instagram, Threads, LinkedIn, X</small></span></label>
      <label class="opt"><input type="radio" name="lead_type" value="2"><span><b>With a website</b><small>Google Maps businesses that have a site</small></span></label>
    </div>
  </section>

  <section class="card">
    <div class="card-head"><div class="num">2</div><div class="grow"><h2>Regions</h2><p>Pick one or more countries to search.</p></div>
      <div class="mini"><button type="button" data-all="regions">All</button><button type="button" data-none="regions">None</button></div></div>
    <div class="opts">
      {% for r in regions %}
      <label class="opt"><input type="checkbox" name="regions" value="{{ r }}" {% if r=='UK' %}checked{% endif %}><span>{{ r }}</span></label>
      {% endfor %}
    </div>
  </section>

  <section class="card">
    <div class="card-head"><div class="num">3</div><div class="grow"><h2>Industries</h2><p>Each industry searches several related business types.</p></div>
      <div class="mini"><button type="button" data-all="industries">All</button><button type="button" data-none="industries">None</button></div></div>
    <div class="opts">
      {% for k,v in industries.items() %}
      <label class="opt"><input type="checkbox" name="industries" value="{{ k }}" {% if k=='1' %}checked{% endif %}><span>{{ v[0] }}</span></label>
      {% endfor %}
      <label class="opt"><input type="checkbox" name="industries" value="14" id="customToggle"><span>Custom</span></label>
    </div>
    <div id="customBox">
      <input class="field" type="text" name="custom_terms" placeholder="Separate with commas: dentist, orthodontist, ...">
    </div>
  </section>

  <section class="card">
    <div class="card-head"><div class="num">4</div><div class="grow"><h2>Services you offer</h2><p>Saved in each lead's notes.</p></div></div>
    <div class="opts">
      {% for k,v in services.items() %}
      <label class="opt"><input type="checkbox" name="offers" value="{{ k }}" checked><span>{{ v }}</span></label>
      {% endfor %}
    </div>
  </section>

  <div class="dock"><div class="actions">
    <button class="primary" type="submit" {% if running %}disabled{% endif %}>▶ Start scrape</button>
    {% if running %}<a class="btn ghost" href="{{ url_for('status') }}">View running scrape</a>{% endif %}
    {% if excel %}<a class="btn ghost" href="{{ url_for('download') }}">⬇ Download Excel</a>{% endif %}
  </div></div>
</form>
<script>
(function(){
  var t=document.getElementById("customToggle"), b=document.getElementById("customBox");
  function sync(){ b.classList.toggle("show", t.checked); }
  t.addEventListener("change", sync); sync();
  document.querySelectorAll("[data-all],[data-none]").forEach(function(btn){
    btn.addEventListener("click", function(){
      var name = btn.getAttribute("data-all") || btn.getAttribute("data-none");
      var on = btn.hasAttribute("data-all");
      document.querySelectorAll('input[name="'+name+'"]').forEach(function(i){ i.checked = on; });
      sync();
    });
  });
})();
</script>

{% else %}
<section class="card">
  <div class="state"><span class="dot {% if running %}run{% endif %}" id="dot"></span><span id="st">{{ 'Running' if running else 'Idle' }}</span></div>
  <div class="chips" id="chips"></div>

  <div class="err-box" id="errbox"></div>

  <div class="bar" role="progressbar" aria-valuemin="0" aria-valuemax="100"
       aria-valuenow="{{ ((stats.queries_done / stats.queries_total * 100) if stats.queries_total else 0)|round|int }}">
    <i id="bar" style="width:{{ ((stats.queries_done / stats.queries_total * 100) if stats.queries_total else 0) }}%"></i>
  </div>
  <div class="pct"><span id="pct"></span><span id="pct2"></span></div>

  <div class="stats">
    <div class="stat hl"><b id="sv">{{ stats.saved }}</b><small>Leads saved</small></div>
    <div class="stat"><b id="qd">{{ stats.queries_done }}</b><small>Searches done</small></div>
    <div class="stat"><b id="qt">{{ stats.queries_total }}</b><small>Searches total</small></div>
  </div>

  <div class="actions status-actions">
    <button class="danger" type="button" id="stopBtn" {% if not running %}style="display:none"{% endif %}>■ Stop</button>
    <a class="btn ghost" href="{{ url_for('index') }}">Change setup</a>
    {% if excel %}<a class="btn primary" href="{{ url_for('download') }}">⬇ Download Excel</a>{% endif %}
    {% if debug_shot %}<a class="btn ghost" href="{{ url_for('debug_png') }}" target="_blank">Debug screenshot</a>{% endif %}
  </div>
</section>

<div class="console">
  <div class="bar2"><span class="live"><i></i>Live log</span><span>updates every 2s</span></div>
  <div id="log"></div>
</div>

<script>
var INITIAL = {{ log|tojson }};
var SUMMARY = {{ summary|tojson }};

(function(){
  var box = document.getElementById("chips");
  (SUMMARY || "").split("|").forEach(function(s){
    s = s.trim(); if(!s || s === "—") return;
    var e = document.createElement("span"); e.textContent = s; box.appendChild(e);
  });
})();

function lineClass(l){
  if (/❌|Fatal/.test(l)) return "bad";
  if (/⚠️/.test(l)) return "warn";
  if (/💾|📡|✅|⏹|NO-WEBSITE|HAS-WEBSITE/.test(l)) return "ok";
  if (/^\s*🔍 \[/.test(l)) return "q";
  return "";
}

function renderLog(lines){
  var el = document.getElementById("log");
  var atEnd = el.scrollTop + el.clientHeight >= el.scrollHeight - 40;
  var frag = document.createDocumentFragment();
  lines.forEach(function(l){
    var d = document.createElement("div");
    d.textContent = l === "" ? "\u00a0" : l;
    var c = lineClass(l); if (c) d.className = c;
    frag.appendChild(d);
  });
  el.replaceChildren(frag);
  if (atEnd) el.scrollTop = el.scrollHeight;
}

var stopBtn = document.getElementById("stopBtn");

function setStats(d){
  var done = d.stats.queries_done, total = d.stats.queries_total;
  var pct = total ? Math.min(100, done / total * 100) : 0;
  document.getElementById("sv").textContent = d.stats.saved;
  document.getElementById("qd").textContent = done;
  document.getElementById("qt").textContent = total;
  document.getElementById("bar").style.width = pct + "%";
  document.getElementById("pct").textContent = total ? pct.toFixed(1) + "% of searches" : "";
  document.getElementById("pct2").textContent = total ? done + " / " + total : "";
  var label = d.stopping ? "Stopping…" : (d.running ? "Running" : "Idle");
  document.getElementById("st").textContent = label;
  document.getElementById("dot").className = "dot" + (d.stopping ? " stopping" : (d.running ? " run" : ""));
  stopBtn.style.display = d.running ? "" : "none";
  if (!d.stopping) { stopBtn.disabled = false; stopBtn.textContent = "■ Stop"; }
  var eb = document.getElementById("errbox");
  if (d.error){ eb.textContent = d.error; eb.style.display = "block"; } else { eb.style.display = "none"; }
}

stopBtn.addEventListener("click", function(){
  stopBtn.disabled = true; stopBtn.textContent = "Stopping…";
  document.getElementById("st").textContent = "Stopping…";
  document.getElementById("dot").className = "dot stopping";
  fetch("{{ url_for('stop') }}", {method:"POST", headers:{"X-Requested-With":"fetch"}})
    .then(function(){ poll(true); });
});

renderLog(INITIAL.split("\n"));
document.getElementById("log").scrollTop = 1e9;

var timer = null;
function poll(once){
  clearTimeout(timer);
  fetch("{{ url_for('api_status') }}").then(function(r){
    if (r.status === 401) { location.href = "{{ url_for('login') }}"; throw new Error("login"); }
    return r.json();
  }).then(function(d){
    setStats(d); renderLog(d.log);
    if (!once) timer = setTimeout(poll, d.stopping ? 700 : 2000);
    else timer = setTimeout(poll, 700);
  }).catch(function(){ timer = setTimeout(poll, 4000); });
}
poll();
</script>
{% endif %}

</div>
</body></html>
"""


@app.route("/")
@login_required
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
@login_required
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
@login_required
def stop():
    with _job_lock:
        _job["stop"] = True
    kill_browser()          # closes Chromium at once so the scraper stops immediately
    if request.headers.get("X-Requested-With") == "fetch":
        return jsonify(ok=True)
    flash("Stopping...", "success")
    return redirect(url_for("status"))


@app.route("/status")
@login_required
def status():
    with _job_lock:
        return render_template_string(
            PAGE, page="status",
            running=_job["running"],
            summary=_job.get("config_summary") or "—",
            stats=_job["stats"],
            log="\n".join(_job["log"]) or "(waiting…)",
            excel=os.path.exists(OUTPUT_EXCEL),
            debug_shot=os.path.exists(DEBUG_SHOT),
        )


@app.route("/api/status")
@login_required
def api_status():
    with _job_lock:
        return jsonify(running=_job["running"],
                       stopping=bool(_job["running"] and _job["stop"]),
                       stats=_job["stats"], log=_job["log"][-200:], error=_job["error"])


@app.route("/download")
@login_required
def download():
    if not os.path.exists(OUTPUT_EXCEL):
        flash("No Excel yet.", "error")
        return redirect(url_for("index"))
    return send_file(OUTPUT_EXCEL, as_attachment=True, download_name=os.path.basename(OUTPUT_EXCEL))


@app.route("/debug.png")
@login_required
def debug_png():
    if not os.path.exists(DEBUG_SHOT):
        flash("No debug screenshot yet.", "error")
        return redirect(url_for("status"))
    return send_file(DEBUG_SHOT, mimetype="image/png")


if __name__ == "__main__":
    if "--crm-test" in sys.argv:
        crm_test()
    else:
        port = int(os.environ.get("PORT", 5000))
        print(f"Open http://0.0.0.0:{port}")
        app.run(host="0.0.0.0", port=port, debug=False, use_reloader=False, threaded=True)
