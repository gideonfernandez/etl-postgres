import time
from datetime import datetime
from pytz import timezone

fmt = "%Y-%m-%d %H:%M:%S %Z%z"
NOW_TIME = datetime.now(timezone('US/Eastern'))
TODAY = NOW_TIME.strftime(fmt)  #Eastern Time, ie: 2022-05-19 09:36:02 EDT-0400
TIMESTR = time.strftime("%Y%m%d-%H%M%S")
TODAYSTR = time.strftime("%Y%m%d")
TODAY_SLASH = time.strftime("%m/%d/%Y")

BITLY_TOKEN = '87e91da4d9efdc2513496785c84e92d20da4d992'

SHAREPOINT_MMG_URL = 'https://montagemarketing.sharepoint.com'
SHAREPOINT_BASE_URL = 'https://montagemarketing.sharepoint.com/sites/MMGDataTeam/'
SHAREPOINT_SUBFOLDER = '/sites/MMGDataTeam/Shared%20Documents/General/Database/Daily Data Sources/NetworkNinja/'
MMG_USER = 'analytics@montagemarketinggroup.com'
MMG_PASSWORD = 'xJfDB94zqWS4I4M2%$xf'

RECIPIENT = 'analytics@montagemarketinggroup.com, gfernandez@montagemarketinggroup.com'

# API RATE LIMITER
ONE_MINUTE = 60
MAX_CALLS_PER_MINUTE = 30

BITLY_LIST = [
	'bit.ly/LA-GoldenChange',
	'bit.ly/togetherDMV',
	'bit.ly/PMT_DREF',
	'bit.ly/togetherBOS',
	'bit.ly/togetherMIA',
	'bit.ly/togetherATL',
	'bit.ly/togetherCHI',
	'bit.ly/togetherDTX',
	'bit.ly/togetherDTM',
	'bit.ly/togetherNY',
	'bit.ly/togetherWI',
	'bit.ly/juntosATL',
	'bit.ly/juntosBOS',
	'bit.ly/juntosCHI',
	'bit.ly/juntosDTX',
	'bit.ly/juntosDTM',
	'bit.ly/juntosMIA',
	'bit.ly/juntosNY',
	'bit.ly/juntosWI',
	'bit.ly/togetherDEN',
	'bit.ly/togetherEP',
	'bit.ly/togetherHTX',
	'bit.ly/togetherMS',
	'bit.ly/togetherNC',
	'bit.ly/juntosDEN',
	'bit.ly/juntosDMV',
	'bit.ly/juntosEP',
	'bit.ly/juntosHTX',
	'bit.ly/juntosMS',
	'bit.ly/juntosNC',
	'bit.ly/togetherBR',
	'bit.ly/togetherNOLA',
	'bit.ly/togetherSHV',
	'bit.ly/juntosBR',
	'bit.ly/juntosNOLA',
	'bit.ly/juntosSHV',
	'bit.ly/exploreAoU',
	'bit.ly/researchercommunity',
	'bit.ly/AllofUs_MSRS',
	'bit.ly/togetherWA',
	'bit.ly/PMT_NBCUSA',
	'bit.ly/NAHH_PMT',
	'bit.ly/PrideNet_PMT',
	'bit.ly/NNLM_PMT',
	'bit.ly/AHC_PMT',
	'bit.ly/Unidos_PMT',
	'bit.ly/NHMA_PMT',
	'bit.ly/PMT_NAHN',
	'bit.ly/LGBTQI_PMT',
	'bit.ly/NHCOA_PMT',
	'bit.ly/NBNA_PMT',
	'bit.ly/ACCESSNY_PMT',
	'bit.ly/LULAC_PMT',
	'bit.ly/AACN_PMT',
	'bit.ly/CenterLink_PMT',
	'bit.ly/BGLC_PMT',
	'bit.ly/NAHN_PMT',
	'bit.ly/NBCUSA_PMT',
	'bit.ly/LGBTDetroit_PMT',
	'bit.ly/BWHI_PMT',
	'bit.ly/AFP-PMT',
	'bit.ly/Cobb_PMT',
	'bit.ly/DREF_PMT',
	'bit.ly/ATLProviders',
	'bit.ly/MSProviders',
	'bit.ly/BOS_Provider',
	'bit.ly/DTX_Providers',
	'bit.ly/EP_Providers',
	'bit.ly/NOLA_Providers',
	'bit.ly/CHIProviders',
	'bit.ly/DTMProviders',
	'bit.ly/DMVProviders',
	'bit.ly/HTXProviders',
	'bit.ly/MiamiProviders',
	'bit.ly/NYProviders',
	'bit.ly/NCProviders',
	'bit.ly/SHVProvider',
	'bit.ly/WIProviders',
	'bit.ly/cpgiadvocates',
	'bit.ly/houstonadvocates',
	'bit.ly/dmvadvocates',
	'bit.ly/chicagoadvocates',
	'bit.ly/LA-GoldenChange-es',
	'bit.ly/LA-DREF',
	'bit.ly/LA-SevenStar',
	'bit.ly/LA-NOCOA',
	'bit.ly/LA-Dillard',
	'bit.ly/LA-Delgado',
	'bit.ly/LA-Xavier',
	'bit.ly/LA-SUNO',
	'bit.ly/NNLM_MI',
	'bit.ly/togetherAoU',
	'bit.ly/AoUTodosjuntos',
	'bit.ly/AoUEats',
	'bit.ly/AllofUsEats',
	'bit.ly/AoU_Eats',
	'bit.ly/Providers_ATL',
	'bit.ly/togetherchelsea',
	'bit.ly/togetherKEAHU',
	'bit.ly/togetherLisa',
	'bit.ly/togetherMARSHALL',
	'bit.ly/togetherMuneera',
	'bit.ly/togetherPaola',
	'bit.ly/togetherSWIFT',
	'bit.ly/LA-NORBCC',
	'bit.ly/LA-NORBCC-es',
]