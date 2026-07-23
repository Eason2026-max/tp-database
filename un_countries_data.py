"""
UN Member States Data for ITP Comparability Database
=====================================================
All 194 UN member states + 2 observers (Holy See, Palestine)
with realistic transfer pricing data distribution.

Data distribution logic:
- High-income developed economies: 40-120 comparables, many industries
- Upper-middle-income: 15-50 comparables
- Lower-middle-income: 5-25 comparables
- Low-income developing: 1-12 comparables
- Tax havens/financial centers: 0-5 comparables (limited TP data)
"""

import numpy as np
import pandas as pd
from datetime import datetime, timedelta

# ===== Full UN Member States (194) + 2 Observers =====
# Format: (country_name, iso3, region, income_group, developing)
# Income groups: HIC=High Income, UMC=Upper Middle, LMC=Lower Middle, LIC=Low Income

UN_ECONOMIES = [
    # === Africa (54) ===
    ("Algeria", "DZA", "Africa", "UMC", True),
    ("Angola", "AGO", "Africa", "LMC", True),
    ("Benin", "BEN", "Africa", "LIC", True),
    ("Botswana", "BWA", "Africa", "UMC", True),
    ("Burkina Faso", "BFA", "Africa", "LIC", True),
    ("Burundi", "BDI", "Africa", "LIC", True),
    ("Cabo Verde", "CPV", "Africa", "LMC", True),
    ("Cameroon", "CMR", "Africa", "LMC", True),
    ("Central African Republic", "CAF", "Africa", "LIC", True),
    ("Chad", "TCD", "Africa", "LIC", True),
    ("Comoros", "COM", "Africa", "LIC", True),
    ("Congo, Republic of the", "COG", "Africa", "LMC", True),
    ("Congo, Democratic Republic of the", "COD", "Africa", "LIC", True),
    ("Cote d'Ivoire", "CIV", "Africa", "LMC", True),
    ("Djibouti", "DJI", "Africa", "LMC", True),
    ("Egypt", "EGY", "Africa", "LMC", True),
    ("Equatorial Guinea", "GNQ", "Africa", "UMC", True),
    ("Eritrea", "ERI", "Africa", "LIC", True),
    ("Eswatini", "SWZ", "Africa", "LMC", True),
    ("Ethiopia", "ETH", "Africa", "LIC", True),
    ("Gabon", "GAB", "Africa", "UMC", True),
    ("Gambia", "GMB", "Africa", "LIC", True),
    ("Ghana", "GHA", "Africa", "LMC", True),
    ("Guinea", "GIN", "Africa", "LIC", True),
    ("Guinea-Bissau", "GNB", "Africa", "LIC", True),
    ("Kenya", "KEN", "Africa", "LMC", True),
    ("Lesotho", "LSO", "Africa", "LMC", True),
    ("Liberia", "LBR", "Africa", "LIC", True),
    ("Libya", "LBY", "Africa", "UMC", True),
    ("Madagascar", "MDG", "Africa", "LIC", True),
    ("Malawi", "MWI", "Africa", "LIC", True),
    ("Mali", "MLI", "Africa", "LIC", True),
    ("Mauritania", "MRT", "Africa", "LMC", True),
    ("Mauritius", "MUS", "Africa", "UMC", True),
    ("Morocco", "MAR", "Africa", "LMC", True),
    ("Mozambique", "MOZ", "Africa", "LIC", True),
    ("Namibia", "NAM", "Africa", "UMC", True),
    ("Niger", "NER", "Africa", "LIC", True),
    ("Nigeria", "NGA", "Africa", "LMC", True),
    ("Rwanda", "RWA", "Africa", "LIC", True),
    ("Sao Tome and Principe", "STP", "Africa", "LMC", True),
    ("Senegal", "SEN", "Africa", "LMC", True),
    ("Seychelles", "SYC", "Africa", "HIC", True),
    ("Sierra Leone", "SLE", "Africa", "LIC", True),
    ("Somalia", "SOM", "Africa", "LIC", True),
    ("South Africa", "ZAF", "Africa", "UMC", True),
    ("South Sudan", "SSD", "Africa", "LIC", True),
    ("Sudan", "SDN", "Africa", "LIC", True),
    ("Tanzania", "TZA", "Africa", "LMC", True),
    ("Togo", "TGO", "Africa", "LIC", True),
    ("Tunisia", "TUN", "Africa", "LMC", True),
    ("Uganda", "UGA", "Africa", "LIC", True),
    ("Zambia", "ZMB", "Africa", "LMC", True),
    ("Zimbabwe", "ZWE", "Africa", "LMC", True),

    # === Asia (48) ===
    ("Afghanistan", "AFG", "Asia", "LIC", True),
    ("Armenia", "ARM", "Asia", "UMC", True),
    ("Azerbaijan", "AZE", "Asia", "UMC", True),
    ("Bahrain", "BHR", "Asia", "HIC", True),
    ("Bangladesh", "BGD", "Asia", "LMC", True),
    ("Bhutan", "BTN", "Asia", "LMC", True),
    ("Brunei Darussalam", "BRN", "Asia", "HIC", True),
    ("Cambodia", "KHM", "Asia", "LMC", True),
    ("China", "CHN", "Asia", "UMC", True),
    ("Cyprus", "CYP", "Asia", "HIC", False),
    ("Georgia", "GEO", "Asia", "UMC", True),
    ("India", "IND", "Asia", "LMC", True),
    ("Indonesia", "IDN", "Asia", "UMC", True),
    ("Iran, Islamic Republic of", "IRN", "Asia", "LMC", True),
    ("Iraq", "IRQ", "Asia", "UMC", True),
    ("Israel", "ISR", "Asia", "HIC", False),
    ("Japan", "JPN", "Asia", "HIC", False),
    ("Jordan", "JOR", "Asia", "LMC", True),
    ("Kazakhstan", "KAZ", "Asia", "UMC", True),
    ("Kuwait", "KWT", "Asia", "HIC", True),
    ("Kyrgyzstan", "KGZ", "Asia", "LMC", True),
    ("Laos", "LAO", "Asia", "LMC", True),
    ("Lebanon", "LBN", "Asia", "LMC", True),
    ("Malaysia", "MYS", "Asia", "UMC", True),
    ("Maldives", "MDV", "Asia", "UMC", True),
    ("Mongolia", "MNG", "Asia", "LMC", True),
    ("Myanmar", "MMR", "Asia", "LMC", True),
    ("Nepal", "NPL", "Asia", "LMC", True),
    ("North Korea", "PRK", "Asia", "LIC", True),
    ("Oman", "OMN", "Asia", "HIC", True),
    ("Pakistan", "PAK", "Asia", "LMC", True),
    ("Philippines", "PHL", "Asia", "LMC", True),
    ("Qatar", "QAT", "Asia", "HIC", True),
    ("Saudi Arabia", "SAU", "Asia", "HIC", True),
    ("Singapore", "SGP", "Asia", "HIC", False),
    ("South Korea", "KOR", "Asia", "HIC", False),
    ("Sri Lanka", "LKA", "Asia", "LMC", True),
    ("Syria", "SYR", "Asia", "LIC", True),
    ("Tajikistan", "TJK", "Asia", "LMC", True),
    ("Thailand", "THA", "Asia", "UMC", True),
    ("Timor-Leste", "TLS", "Asia", "LMC", True),
    ("Turkey", "TUR", "Asia", "UMC", True),
    ("Turkmenistan", "TKM", "Asia", "UMC", True),
    ("United Arab Emirates", "ARE", "Asia", "HIC", True),
    ("Uzbekistan", "UZB", "Asia", "LMC", True),
    ("Vietnam", "VNM", "Asia", "LMC", True),
    ("Yemen", "YEM", "Asia", "LIC", True),

    # === Europe (44) ===
    ("Albania", "ALB", "Europe", "UMC", True),
    ("Andorra", "AND", "Europe", "HIC", False),
    ("Austria", "AUT", "Europe", "HIC", False),
    ("Belarus", "BLR", "Europe", "UMC", True),
    ("Belgium", "BEL", "Europe", "HIC", False),
    ("Bosnia and Herzegovina", "BIH", "Europe", "UMC", True),
    ("Bulgaria", "BGR", "Europe", "UMC", True),
    ("Croatia", "HRV", "Europe", "HIC", False),
    ("Czech Republic", "CZE", "Europe", "HIC", False),
    ("Denmark", "DNK", "Europe", "HIC", False),
    ("Estonia", "EST", "Europe", "HIC", False),
    ("Finland", "FIN", "Europe", "HIC", False),
    ("France", "FRA", "Europe", "HIC", False),
    ("Germany", "DEU", "Europe", "HIC", False),
    ("Greece", "GRC", "Europe", "HIC", False),
    ("Hungary", "HUN", "Europe", "HIC", False),
    ("Iceland", "ISL", "Europe", "HIC", False),
    ("Ireland", "IRL", "Europe", "HIC", False),
    ("Italy", "ITA", "Europe", "HIC", False),
    ("Latvia", "LVA", "Europe", "HIC", False),
    ("Liechtenstein", "LIE", "Europe", "HIC", False),
    ("Lithuania", "LTU", "Europe", "HIC", False),
    ("Luxembourg", "LUX", "Europe", "HIC", False),
    ("Malta", "MLT", "Europe", "HIC", False),
    ("Moldova", "MDA", "Europe", "LMC", True),
    ("Monaco", "MCO", "Europe", "HIC", False),
    ("Montenegro", "MNE", "Europe", "UMC", True),
    ("Netherlands", "NLD", "Europe", "HIC", False),
    ("North Macedonia", "MKD", "Europe", "UMC", True),
    ("Norway", "NOR", "Europe", "HIC", False),
    ("Poland", "POL", "Europe", "HIC", False),
    ("Portugal", "PRT", "Europe", "HIC", False),
    ("Romania", "ROU", "Europe", "UMC", True),
    ("Russia", "RUS", "Europe", "UMC", True),
    ("San Marino", "SMR", "Europe", "HIC", False),
    ("Serbia", "SRB", "Europe", "UMC", True),
    ("Slovakia", "SVK", "Europe", "HIC", False),
    ("Slovenia", "SVN", "Europe", "HIC", False),
    ("Spain", "ESP", "Europe", "HIC", False),
    ("Sweden", "SWE", "Europe", "HIC", False),
    ("Switzerland", "CHE", "Europe", "HIC", False),
    ("Ukraine", "UKR", "Europe", "LMC", True),
    ("United Kingdom", "GBR", "Europe", "HIC", False),
    ("Vatican City (Holy See)", "VAT", "Europe", "HIC", False),

    # === Latin America & Caribbean (33) ===
    ("Antigua and Barbuda", "ATG", "Latin America", "HIC", True),
    ("Argentina", "ARG", "Latin America", "UMC", True),
    ("Bahamas", "BHS", "Latin America", "HIC", True),
    ("Barbados", "BRB", "Latin America", "HIC", True),
    ("Belize", "BLZ", "Latin America", "UMC", True),
    ("Bolivia", "BOL", "Latin America", "LMC", True),
    ("Brazil", "BRA", "Latin America", "UMC", True),
    ("Chile", "CHL", "Latin America", "HIC", True),
    ("Colombia", "COL", "Latin America", "UMC", True),
    ("Costa Rica", "CRI", "Latin America", "UMC", True),
    ("Cuba", "CUB", "Latin America", "UMC", True),
    ("Dominica", "DMA", "Latin America", "UMC", True),
    ("Dominican Republic", "DOM", "Latin America", "UMC", True),
    ("Ecuador", "ECU", "Latin America", "UMC", True),
    ("El Salvador", "SLV", "Latin America", "UMC", True),
    ("Grenada", "GRD", "Latin America", "UMC", True),
    ("Guatemala", "GTM", "Latin America", "UMC", True),
    ("Guyana", "GUY", "Latin America", "LMC", True),
    ("Haiti", "HTI", "Latin America", "LIC", True),
    ("Honduras", "HND", "Latin America", "LMC", True),
    ("Jamaica", "JAM", "Latin America", "UMC", True),
    ("Mexico", "MEX", "Latin America", "UMC", True),
    ("Nicaragua", "NIC", "Latin America", "LMC", True),
    ("Panama", "PAN", "Latin America", "HIC", True),
    ("Paraguay", "PRY", "Latin America", "UMC", True),
    ("Peru", "PER", "Latin America", "UMC", True),
    ("Saint Kitts and Nevis", "KNA", "Latin America", "HIC", True),
    ("Saint Lucia", "LCA", "Latin America", "UMC", True),
    ("Saint Vincent and the Grenadines", "VCT", "Latin America", "UMC", True),
    ("Suriname", "SUR", "Latin America", "UMC", True),
    ("Trinidad and Tobago", "TTO", "Latin America", "HIC", True),
    ("Uruguay", "URY", "Latin America", "HIC", True),
    ("Venezuela", "VEN", "Latin America", "LMC", True),

    # === North America (2) ===
    ("Canada", "CAN", "North America", "HIC", False),
    ("United States", "USA", "North America", "HIC", False),

    # === Oceania (14) ===
    ("Australia", "AUS", "Oceania", "HIC", False),
    ("Cook Islands", "COK", "Oceania", "UMC", True),
    ("Fiji", "FJI", "Oceania", "UMC", True),
    ("Kiribati", "KIR", "Oceania", "LMC", True),
    ("Marshall Islands", "MHL", "Oceania", "UMC", True),
    ("Micronesia", "FSM", "Oceania", "LMC", True),
    ("Nauru", "NRU", "Oceania", "HIC", True),
    ("New Zealand", "NZL", "Oceania", "HIC", False),
    ("Palau", "PLW", "Oceania", "HIC", True),
    ("Papua New Guinea", "PNG", "Oceania", "LMC", True),
    ("Samoa", "WSM", "Oceania", "LMC", True),
    ("Solomon Islands", "SLB", "Oceania", "LMC", True),
    ("Tonga", "TON", "Oceania", "UMC", True),
    ("Tuvalu", "TUV", "Oceania", "UMC", True),

    # === Observers / Special ===
    ("Palestine", "PSE", "Asia", "LMC", True),
    ("Taiwan (China)", "TWN", "Asia", "HIC", False),
    ("Hong Kong SAR (China)", "HKG", "Asia", "HIC", False),
    ("Macao SAR (China)", "MAC", "Asia", "HIC", False),
    ("Puerto Rico", "PRI", "Latin America", "HIC", False),
    ("Greenland", "GRL", "Europe", "HIC", False),
    ("New Caledonia", "NCL", "Oceania", "HIC", False),
    ("French Polynesia", "PYF", "Oceania", "HIC", False),
    ("Bermuda", "BMU", "North America", "HIC", False),
    ("Cayman Islands", "CYM", "Latin America", "HIC", False),
    ("British Virgin Islands", "VGB", "Latin America", "HIC", False),
    ("Isle of Man", "IMN", "Europe", "HIC", False),
    ("Jersey", "JEY", "Europe", "HIC", False),
    ("Guernsey", "GGY", "Europe", "HIC", False),
    ("Gibraltar", "GIB", "Europe", "HIC", False),
    ("Curacao", "CUW", "Latin America", "HIC", False),
    ("Aruba", "ABW", "Latin America", "HIC", False),
    ("Faroe Islands", "FRO", "Europe", "HIC", False),
]


def generate_countries_data() -> pd.DataFrame:
    """Generate country-level data for all UN member states with realistic distribution."""
    np.random.seed(42)
    data = []
    today = datetime.now()

    # Known tax havens / financial centers — fewer TP comparables
    tax_havens = {
        "Bermuda", "Cayman Islands", "British Virgin Islands", "Isle of Man",
        "Jersey", "Guernsey", "Gibraltar", "British Virgin Islands",
        "San Marino", "Monaco", "Liechtenstein", "Andorra", "Nauru",
        "Marshall Islands", "Vanuatu", "Panama", "Bahamas", "Seychelles",
    }

    # Countries with known strong TP regimes get more data
    high_tp_coverage = {
        "United States", "Germany", "United Kingdom", "France", "Japan",
        "China", "India", "Brazil", "South Africa", "Australia", "Canada",
        "Netherlands", "South Korea", "Italy", "Spain", "Mexico", "Russia",
        "Singapore", "Switzerland", "Belgium", "Sweden", "Norway", "Finland",
        "Denmark", "Ireland", "Poland", "Portugal", "Czech Republic", "Austria",
    }

    # Major economies get specific high comparables
    specific_comparables = {
        "United States": 125, "China": 110, "Germany": 92, "Japan": 88,
        "United Kingdom": 85, "France": 76, "India": 95, "South Korea": 72,
        "Italy": 68, "Spain": 65, "Canada": 78, "Australia": 70,
        "Netherlands": 62, "Brazil": 78, "Mexico": 55, "Switzerland": 58,
        "Singapore": 52, "Sweden": 48, "Belgium": 45, "Norway": 42,
        "South Africa": 67, "Ireland": 40, "Poland": 38, "Austria": 40,
        "Denmark": 35, "Finland": 33, "Portugal": 30, "Russia": 48,
        "Turkey": 35, "Indonesia": 45, "Malaysia": 42, "Thailand": 38,
        "Vietnam": 33, "Saudi Arabia": 35, "United Arab Emirates": 30,
        "Argentina": 28, "Chile": 35, "Colombia": 30, "Egypt": 22,
        "Nigeria": 38, "Kenya": 42, "Ghana": 18, "Morocco": 25,
        "Pakistan": 20, "Bangladesh": 15, "Philippines": 28,
        "Czech Republic": 32, "Hungary": 28, "Romania": 22,
        "Greece": 25, "Israel": 35, "New Zealand": 30,
    }

    for country, iso, region, income, developing in UN_ECONOMIES:
        # Determine comparables count
        if country in tax_havens:
            comp_count = np.random.randint(0, 3)
            ind_count = np.random.randint(0, 2)
            map_count = 0
        elif country in specific_comparables:
            comp_count = specific_comparables[country]
            ind_count = min(comp_count // 8 + 2, 12)
            map_count = np.random.randint(comp_count // 15, comp_count // 8 + 1)
        elif income == "HIC":
            comp_count = np.random.randint(15, 50)
            ind_count = np.random.randint(4, 8)
            map_count = np.random.randint(1, 5)
        elif income == "UMC":
            comp_count = np.random.randint(8, 30)
            ind_count = np.random.randint(3, 6)
            map_count = np.random.randint(0, 4)
        elif income == "LMC":
            comp_count = np.random.randint(2, 15)
            ind_count = np.random.randint(1, 4)
            map_count = np.random.randint(0, 2)
        else:  # LIC
            comp_count = np.random.randint(0, 6)
            ind_count = np.random.randint(0, 3)
            map_count = np.random.randint(0, 2)

        # Last updated date — more recent for active data contributors
        if comp_count > 30:
            days_ago = np.random.randint(1, 30)
        elif comp_count > 10:
            days_ago = np.random.randint(7, 60)
        elif comp_count > 0:
            days_ago = np.random.randint(14, 120)
        else:
            days_ago = np.random.randint(60, 180)

        last_updated = (today - timedelta(days=days_ago)).strftime("%Y-%m-%d")

        data.append({
            "country": country,
            "iso_code": iso,
            "comparables_count": comp_count,
            "last_updated": last_updated,
            "industries_count": ind_count,
            "region": region,
            "developing": developing,
            "map_cases": map_count,
        })

    df = pd.DataFrame(data)
    # Mirror Taiwan with China's data for map display
    chn = df[df["iso_code"] == "CHN"]
    if not chn.empty:
        twn = chn.iloc[0].copy()
        twn["iso_code"] = "TWN"
        twn["country"] = "Taiwan (China)"
        df = pd.concat([df, pd.DataFrame([twn])], ignore_index=True)
    return df


def generate_country_policies() -> pd.DataFrame:
    """Generate TP policy data for all economies."""
    np.random.seed(88)
    data = []

    # Known specific policies for major economies
    specific_policies = {
        "United States": (True, "Three-tier (482 + CbCR)", False, 65, 18, True, "EOIR + CbC + spontaneous"),
        "China": (True, "Three-tier + Special File", True, 55, 15, True, "EOIR + CbC"),
        "Germany": (True, "Three-tier (GAufzV)", False, 60, 11, True, "EOIR + CbC + spontaneous"),
        "United Kingdom": (True, "Three-tier (DPTL)", False, 58, 10, True, "EOIR + CbC"),
        "France": (True, "Three-tier (Art. 13 VBA)", False, 62, 9, True, "EOIR + CbC + spontaneous"),
        "Japan": (True, "Three-tier", False, 55, 12, True, "EOIR + CbC"),
        "India": (True, "Three-tier + Form 3CEB", True, 45, 12, True, "EOIR + CbC"),
        "Brazil": (True, "Local File + CbCR (new TP rules 2024)", True, 35, 9, True, "EOIR"),
        "South Africa": (True, "Three-tier", True, 30, 8, True, "EOIR + CbC + spontaneous"),
        "Mexico": (True, "Three-tier", True, 28, 5, True, "EOIR + CbC"),
        "Australia": (True, "Three-tier (PCG 2019/1)", True, 42, 7, True, "EOIR + CbC"),
        "Canada": (True, "Three-tier", False, 48, 6, True, "EOIR + CbC"),
        "Netherlands": (True, "Three-tier (DOC)", False, 52, 8, True, "EOIR + CbC + spontaneous"),
        "South Korea": (True, "Three-tier", True, 38, 6, True, "EOIR + CbC"),
        "Switzerland": (True, "Three-tier (KFG)", False, 40, 5, True, "EOIR + CbC"),
        "Singapore": (True, "Three-tier (TPG-D04/2024)", False, 35, 4, True, "EOIR + CbC"),
        "Ireland": (True, "Three-tier", False, 32, 5, True, "EOIR + CbC"),
        "Italy": (True, "Three-tier (Art. 276 DPR 917)", True, 45, 8, True, "EOIR + CbC"),
        "Spain": (True, "Three-tier", True, 42, 7, True, "EOIR + CbC"),
        "Nigeria": (True, "Local File + CbCR", True, 18, 5, True, "EOIR"),
        "Kenya": (True, "Three-tier (Master, Local, CbCR)", False, 15, 3, True, "EOIR + CbC"),
        "Indonesia": (True, "Three-tier (PMK-213)", True, 40, 4, True, "EOIR + CbC"),
        "Vietnam": (True, "Local File + CbCR (Decree 132)", True, 22, 3, True, "EOIR"),
        "Chile": (True, "Local File + CbCR", True, 20, 2, True, "EOIR"),
        "Russia": (True, "Three-tier (Art. 25.6-25.9 NKK)", False, 48, 7, True, "EOIR"),
        "Turkey": (True, "Three-tier (Gen. Teblig 2017)", True, 35, 4, True, "EOIR + CbC"),
        "Saudi Arabia": (True, "Three-tier (DZIT)", False, 28, 3, True, "EOIR + CbC"),
        "Thailand": (True, "Three-tier (RD. Por. 113/2554)", True, 22, 3, True, "EOIR"),
        "Malaysia": (True, "Three-tier (TPG 2024)", True, 25, 4, True, "EOIR + CbC"),
        "Argentina": (True, "Local File + CbCR", True, 24, 4, True, "EOIR"),
    }

    for country, iso, region, income, developing in UN_ECONOMIES:
        if country in specific_policies:
            tp_reg, doc_req, sh, tc, mc, apa, em = specific_policies[country]
        else:
            # Auto-generate based on income level
            if income == "HIC":
                tp_reg = True
                doc_req = np.random.choice([
                    "Three-tier (Master + Local + CbCR)",
                    "Local File + CbCR",
                ])
                sh = bool(np.random.choice([True, False], p=[0.3, 0.7]))
                tc = np.random.randint(20, 65)
                mc = np.random.randint(1, 8)
                apa = bool(np.random.choice([True, False], p=[0.85, 0.15]))
                em = np.random.choice([
                    "EOIR + CbC",
                    "EOIR + CbC + spontaneous",
                    "EOIR",
                ])
            elif income == "UMC":
                tp_reg = bool(np.random.choice([True, True, False], p=[0.85, 0.10, 0.05]))
                doc_req = np.random.choice([
                    "Three-tier",
                    "Local File + CbCR",
                    "Local File only",
                ])
                sh = bool(np.random.choice([True, False], p=[0.55, 0.45]))
                tc = np.random.randint(10, 35)
                mc = np.random.randint(0, 5)
                apa = bool(np.random.choice([True, False], p=[0.70, 0.30]))
                em = "EOIR" if np.random.rand() > 0.5 else "EOIR + CbC"
            elif income == "LMC":
                tp_reg = bool(np.random.choice([True, False], p=[0.75, 0.25]))
                doc_req = np.random.choice([
                    "Local File + CbCR",
                    "Local File only",
                    "Not yet required",
                ])
                sh = bool(np.random.choice([True, False], p=[0.45, 0.55]))
                tc = np.random.randint(5, 22)
                mc = np.random.randint(0, 3)
                apa = bool(np.random.choice([True, False], p=[0.40, 0.60]))
                em = "EOIR"
            else:  # LIC
                tp_reg = bool(np.random.choice([True, False], p=[0.40, 0.60]))
                doc_req = np.random.choice([
                    "Not yet required",
                    "Local File only",
                ])
                sh = bool(np.random.choice([True, False], p=[0.20, 0.80]))
                tc = np.random.randint(2, 15)
                mc = np.random.randint(0, 2)
                apa = bool(np.random.choice([True, False], p=[0.15, 0.85]))
                em = "EOIR (limited)"

        data.append({
            "country": country,
            "has_tp_regulations": tp_reg,
            "documentation_req": doc_req,
            "safe_harbor": sh,
            "treaty_count": tc,
            "map_cases_2024": mc,
            "apa_available": apa,
            "exchange_mechanism": em,
        })

    return pd.DataFrame(data)


def generate_companies_data(n_companies: int = 800) -> pd.DataFrame:
    """Generate comparable company data spread across many economies."""
    np.random.seed(42)

    # Industries with ISIC codes
    industries = [
        ("Manufacturing", "C"),
        ("Software & IT Services", "J"),
        ("Pharmaceuticals", "C"),
        ("Financial Services", "K"),
        ("Retail & Distribution", "G"),
        ("Mining & Extractive", "B"),
        ("Energy & Utilities", "D"),
        ("Telecommunications", "J"),
        ("Construction", "F"),
        ("Transportation & Logistics", "H"),
        ("Agriculture & Food Processing", "A"),
        ("Automotive", "C"),
        ("Chemicals", "C"),
        ("Electronics", "C"),
        ("Textiles & Apparel", "C"),
        ("Professional Services", "M"),
        ("Real Estate", "L"),
        ("Healthcare", "Q"),
        ("Education", "P"),
        ("Media & Entertainment", "J"),
    ]

    # Weight countries by data availability
    country_weights = {}
    for country, iso, region, income, developing in UN_ECONOMIES:
        if income == "HIC":
            country_weights[country] = (iso, 15)
        elif income == "UMC":
            country_weights[country] = (iso, 8)
        elif income == "LMC":
            country_weights[country] = (iso, 4)
        else:  # LIC
            country_weights[country] = (iso, 1)

    countries_list = list(country_weights.keys())
    weights = np.array([country_weights[c][1] for c in countries_list], dtype=float)
    weights /= weights.sum()
    iso_map = {c: country_weights[c][0] for c in countries_list}

    companies = []
    for i in range(n_companies):
        country = np.random.choice(countries_list, p=weights)
        iso_code = iso_map[country]
        ind_name, isic_code = industries[np.random.randint(0, len(industries))]

        revenue = np.random.uniform(5, 800)
        cogs_ratio = np.random.uniform(0.45, 0.85)
        cogs = revenue * cogs_ratio
        gross_profit = revenue - cogs
        gross_margin = gross_profit / revenue
        op_expense = revenue * np.random.uniform(0.05, 0.25)
        op_profit = gross_profit - op_expense
        op_margin = op_profit / revenue

        # Industry-specific margin adjustments
        if ind_name in ["Pharmaceuticals", "Software & IT Services"]:
            op_margin = max(op_margin, np.random.uniform(0.10, 0.35))
        elif ind_name in ["Mining & Extractive", "Energy & Utilities"]:
            op_margin = np.random.uniform(0.02, 0.20)

        cost_plus = op_profit / cogs if cogs > 0 else 0.15
        net_assets = revenue * np.random.uniform(0.3, 1.5)

        companies.append({
            "company_id": f"CMP{i+1:04d}",
            "company_name": f"Company {i+1:04d} Ltd.",
            "country": country,
            "country_code": iso_code,
            "industry_code_isic": isic_code,
            "industry": ind_name,
            "isic_section": isic_code,
            "revenue_musd": round(revenue, 2),
            "net_assets_musd": round(net_assets, 2),
            "operating_margin": round(op_margin, 4),
            "gross_margin": round(gross_margin, 4),
            "cost_plus_markup": round(cost_plus, 4),
            "roe": round(np.random.uniform(0.02, 0.30), 4),
            "roa": round(np.random.uniform(0.01, 0.20), 4),
            "berry_ratio": round(gross_profit / op_expense if op_expense > 0 else 1.05, 4),
            "roce": round(np.random.uniform(0.05, 0.25), 4),
            "net_profit_margin": round(np.random.uniform(0.02, 0.20), 4),
            "related_party_pct": round(np.random.uniform(0, 0.60), 4),
            "fiscal_year": int(np.random.choice([2022, 2023, 2024])),
            "functions": ", ".join(np.random.choice(
                ["R&D", "Manufacturing", "Distribution", "Marketing", "Digital Platform", "Logistics"],
                size=np.random.randint(1, 4), replace=False)),
            "risks": np.random.choice(["Market Risk", "Market Risk; Inventory Risk", "Market Risk; Credit Risk", "Full Risk"]),
            "data_tier": np.random.choice(["Tier1", "Tier2", "Tier3"], p=[0.35, 0.45, 0.20]),
            "data_source": "Company Annual Reports",
            "has_inventory_risk": bool(np.random.choice([True, False], p=[0.6, 0.4])),
            "has_market_risk": bool(np.random.choice([True, False], p=[0.7, 0.3])),
            "functional_complexity_score": np.random.randint(3, 10),
        })

    return pd.DataFrame(companies)


def generate_map_cases(n_cases: int = 120) -> pd.DataFrame:
    """Generate MAP/APA cases across many economies."""
    np.random.seed(55)

    # Use a wide range of countries
    developing_economies = [e[0] for e in UN_ECONOMIES if e[4]]
    developed_economies = [e[0] for e in UN_ECONOMIES if not e[4]]

    # Ensure we have at least some countries
    if len(developing_economies) < 10:
        developing_economies.extend(["Kenya", "Nigeria", "South Africa", "India", "China"])
    if len(developed_economies) < 10:
        developed_economies.extend(["Germany", "USA", "UK", "Netherlands", "France"])

    statuses = ["Active", "Resolved", "Pending", "Closed - No agreement"]
    status_probs = [0.35, 0.30, 0.25, 0.10]
    case_types = ["MAP", "Bilateral APA", "Unilateral APA", "Multilateral APA"]
    issues = ["Royalty pricing", "Intra-group services", "Goods transfer pricing",
              "Financial transactions", "Business restructuring", "Digital services",
              "Cost contribution arrangements", "Intangibles migration"]

    cases = []
    for i in range(n_cases):
        start_date = datetime(2022, 1, 1) + timedelta(days=int(np.random.randint(0, 900)))
        duration_days = int(np.random.randint(60, 730))
        status = str(np.random.choice(statuses, p=status_probs))

        cases.append({
            "case_id": f"MAP{i+1:04d}",
            "case_type": str(np.random.choice(case_types)),
            "taxpayer": f"Taxpayer {i+1:03d} Ltd.",
            "country_a": str(np.random.choice(developing_economies)),
            "country_b": str(np.random.choice(developed_economies)),
            "issue": str(np.random.choice(issues)),
            "status": status,
            "start_date": start_date.strftime("%Y-%m-%d"),
            "duration_days": duration_days,
            "amount_disputed_musd": round(float(np.random.uniform(0.5, 50)), 2),
            "resolution_date": (start_date + timedelta(days=duration_days)).strftime("%Y-%m-%d")
                if status in ["Resolved", "Closed - No agreement"] else None,
        })

    return pd.DataFrame(cases)


def generate_digital_comparables(n_companies: int = 150) -> pd.DataFrame:
    """Generate digital economy comparable companies."""
    np.random.seed(77)

    service_types = [
        "Cloud Computing", "SaaS", "Digital Advertising", "E-commerce Platform",
        "Data Analytics", "IT Outsourcing", "Digital Content", "Fintech Services",
        "Social Media Platform", "Streaming Services",
    ]

    # Digital economy hubs — countries with significant digital sector
    digital_hubs = [
        ("United States", 20), ("China", 18), ("India", 15), ("South Korea", 8),
        ("Singapore", 8), ("Ireland", 6), ("Israel", 6), ("Germany", 8),
        ("United Kingdom", 7), ("Japan", 8), ("Brazil", 6), ("Nigeria", 3),
        ("Estonia", 4), ("United Arab Emirates", 5), ("Canada", 6),
        ("Australia", 5), ("Netherlands", 5), ("Sweden", 4),
        ("France", 5), ("Spain", 4), ("Indonesia", 4),
        ("Malaysia", 3), ("Vietnam", 3), ("Thailand", 3),
        ("Mexico", 4), ("Poland", 3), ("Czech Republic", 3),
        ("Hungary", 2), ("Portugal", 2), ("Finland", 3),
        ("South Africa", 3), ("Kenya", 2), ("Egypt", 2),
        ("Philippines", 3), ("Turkey", 3), ("Argentina", 2),
        ("Chile", 2), ("Colombia", 2), ("Pakistan", 2),
        ("Bangladesh", 1), ("Rwanda", 1), ("Nigeria", 3),
    ]

    countries = [c[0] for c in digital_hubs]
    weights = np.array([c[1] for c in digital_hubs], dtype=float)
    weights /= weights.sum()

    data = []
    for i in range(n_companies):
        country = np.random.choice(countries, p=weights)
        revenue = np.random.uniform(10, 500)
        op_margin = round(float(np.random.uniform(0.05, 0.45)), 4)

        data.append({
            "company_id": f"DIG{i+1:04d}",
            "company_name": f"Digital Co {i+1:04d}",
            "country": country,
            "service_type": str(np.random.choice(service_types)),
            "revenue_musd": round(revenue, 2),
            "operating_margin": op_margin,
            "cost_plus_markup": round(op_margin / (1 - op_margin), 4) if op_margin < 1 else 0.5,
            "tnmm_profit_margin": op_margin,
            "digital_intensity": round(float(np.random.uniform(0.3, 1.0)), 4),
            "automation_level": np.random.choice(["Automated", "Semi-Automated", "Manual"]),
            "users_millions": int(np.random.uniform(1, 500)),
            "fiscal_year": int(np.random.choice([2022, 2023, 2024])),
        })

    return pd.DataFrame(data)


def generate_customs_data(n_records: int = 500) -> pd.DataFrame:
    """Generate customs valuation records across many economies."""
    np.random.seed(99)

    # Commodities with HS codes
    commodities = [
        ("Copper Concentrate", "2603.00", "Metals & Minerals"),
        ("Iron Ore (Fe 62%)", "2601.11", "Metals & Minerals"),
        ("Gold Dore (Au 85%)", "7108.12", "Metals & Minerals"),
        ("Crude Petroleum", "2709.00", "Energy & Petroleum"),
        ("Refined Copper", "7401.00", "Metals & Minerals"),
        ("Aluminum Ingots", "7601.10", "Metals & Minerals"),
        ("Cotton", "5201.00", "Agriculture & Food"),
        ("Coffee Beans", "0901.00", "Agriculture & Food"),
        ("Electronic Components", "8542.00", "Electronics"),
        ("Pharmaceuticals", "3004.90", "Chemicals & Pharma"),
        ("Automotive Parts", "8708.00", "Automotive"),
        ("Plastics", "3901.10", "Chemicals & Pharma"),
    ]

    # Trading economies — broader list
    reporter_economies = [e[0] for e in UN_ECONOMIES if e[3] in ("HIC", "UMC")]
    if len(reporter_economies) < 20:
        reporter_economies = [e[0] for e in UN_ECONOMIES]

    partner_economies = [e[0] for e in UN_ECONOMIES]

    directions = ["Import", "Export"]
    val_methods = ["Transaction Value (Art. 1)", "Transaction Value (Art. 2)",
                   "Deductive Value (Art. 5)", "Computed Value (Art. 6)"]

    records = []
    for i in range(n_records):
        comm, hs, cat = commodities[np.random.randint(0, len(commodities))]
        direction = np.random.choice(directions)
        reporter = np.random.choice(reporter_economies)
        partner = np.random.choice(partner_economies)
        while partner == reporter:
            partner = np.random.choice(partner_economies)

        related = bool(np.random.choice([True, False], p=[0.35, 0.65]))
        qty = np.random.uniform(10, 50000)
        unit_price = np.random.uniform(50, 9000)
        declared = qty * unit_price
        # Related-party transactions may have pricing deviations
        deviation = np.random.uniform(-0.25, 0.25) if related else np.random.uniform(-0.05, 0.05)

        records.append({
            "customs_id": f"CUS{i+1:05d}",
            "trade_direction": direction,
            "reporting_country_name": reporter,
            "partner_country_name": partner,
            "commodity_name": comm,
            "hs_code": hs,
            "commodity_category": cat,
            "quantity": round(qty, 2),
            "unit_price_usd": round(unit_price, 2),
            "declared_value_usd": round(declared, 2),
            "related_party_flag": related,
            "valuation_method": np.random.choice(val_methods),
            "deviation_pct": round(deviation * 100, 2),
            "date": (datetime.now() - timedelta(days=np.random.randint(1, 365))).strftime("%Y-%m-%d"),
        })

    return pd.DataFrame(records)
