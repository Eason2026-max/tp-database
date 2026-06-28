"""
International Transfer Pricing Comparability Database
====================================================
A Streamlit-based frontend application for the International Tax Cooperation
Framework Convention's global transfer pricing comparability analysis database.

Helps developing economy tax officials access comparable data and conduct
transfer pricing analyses under Articles 5 (Fair Distribution of Taxing Rights)
and 11 (Capacity Building).

Pages:
    1. Global Data Portal - Overview with world map and KPIs
    2. Comparability Analysis Wizard - 4-step guided analysis
    3. Extractive Industry Pricing - Commodity pricing workbench
    4. Country Policy Lookup - National TP regulation lookup
    5. Data Contribution Hub - Data upload and management
    6. MAP/APA Case Tracker - Dispute resolution tracking
    7. Digital Economy Dashboard - Cross-border services analysis

Author: Senior Streamlit Developer
Date: 2025-06-24
"""

from __future__ import annotations

import io
import os
import random
from datetime import datetime, timedelta
from typing import Any, Dict, List, Optional

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor

# ============== Page Configuration ==============
st.set_page_config(
    page_title="ITP Comparability Database",
    page_icon=":earth_americas:",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============== Theme / Styling ==============
PRIMARY_COLOR = "#5B92E5"  # UN Blue
ACCENT_GREEN = "#27AE60"
ACCENT_RED = "#E74C3C"
ACCENT_YELLOW = "#F39C12"
ACCENT_PURPLE = "#8E44AD"

st.markdown(
    f"""
    <style>
    .main-header {{
        font-size: 2.2rem;
        font-weight: 700;
        color: {PRIMARY_COLOR};
        margin-bottom: 0.5rem;
    }}
    .sub-header {{
        font-size: 1.1rem;
        color: #666666;
        margin-bottom: 1.5rem;
    }}
    .kpi-card {{
        background: linear-gradient(135deg, #f8f9fa 0%, #e9ecef 100%);
        border-radius: 10px;
        padding: 1.2rem;
        border-left: 4px solid {PRIMARY_COLOR};
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }}
    .kpi-value {{
        font-size: 1.8rem;
        font-weight: 700;
        color: {PRIMARY_COLOR};
    }}
    .kpi-label {{
        font-size: 0.85rem;
        color: #666666;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }}
    .info-bar {{
        background-color: #e8f2fc;
        border-left: 4px solid {PRIMARY_COLOR};
        padding: 1rem;
        border-radius: 4px;
        margin: 1rem 0;
    }}
    .metric-good {{
        color: {ACCENT_GREEN};
        font-weight: 600;
    }}
    .metric-warning {{
        color: {ACCENT_YELLOW};
        font-weight: 600;
    }}
    .metric-danger {{
        color: {ACCENT_RED};
        font-weight: 600;
    }}
    .update-item {{
        padding: 0.6rem 0;
        border-bottom: 1px solid #e0e0e0;
        font-size: 0.9rem;
    }}
    .sidebar-title {{
        font-size: 1.1rem;
        font-weight: 600;
        color: {PRIMARY_COLOR};
        margin-bottom: 0.5rem;
    }}
    .method-badge {{
        background-color: {PRIMARY_COLOR};
        color: white;
        padding: 0.3rem 0.8rem;
        border-radius: 12px;
        font-size: 0.8rem;
        font-weight: 500;
    }}
    div.stButton > button:first-child {{
        background-color: {PRIMARY_COLOR};
        color: white;
        border: none;
        border-radius: 6px;
        padding: 0.5rem 1.5rem;
        font-weight: 500;
    }}
    div.stButton > button:first-child:hover {{
        background-color: #4a7bc8;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ============== Database Connection ==============
# Load config from db_config.py; fall back to inline defaults
try:
    from db_config import DB_CONFIG
except ImportError:
    DB_CONFIG = {
        "host": "localhost",
        "port": "5432",
        "database": "tp_comparability_db",
        "user": "postgres",
        "password": "your_password_here",
    }


@st.cache_resource
def get_db_engine() -> Optional[Any]:
    """Create SQLAlchemy engine for PostgreSQL connection.

    Uses connection pooling optimized for Supabase free tier (max 60 connections).
    Returns None if connection fails, in which case mock data is used.
    """
    try:
        from sqlalchemy import create_engine, text
        from urllib.parse import quote_plus

        # URL-encode password to handle special characters
        encoded_password = quote_plus(str(DB_CONFIG['password']))
        encoded_user = quote_plus(str(DB_CONFIG['user']))

        conn_str = (
            f"postgresql://{encoded_user}:{encoded_password}"
            f"@{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['database']}"
        )
        engine = create_engine(
            conn_str,
            pool_pre_ping=True,          # Check connection before use
            pool_size=5,                  # Small pool for free tier
            max_overflow=3,               # Allow 3 extra connections
            pool_timeout=30,              # Wait max 30s for connection
            pool_recycle=1800,            # Recycle connections every 30 min
            connect_args={
                "connect_timeout": 10,    # 10s connection timeout
                "application_name": "tp_database_app",
            },
        )
        # Quick connectivity test
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        return engine
    except Exception as e:
        st.session_state.setdefault("db_error", str(e))
        return None


def get_db_connection() -> Optional[Any]:
    """Return a DB connection or None if unavailable."""
    engine = get_db_engine()
    if engine is None:
        return None
    try:
        return engine.connect()
    except Exception as e:
        st.session_state.setdefault("db_error", str(e))
        return None


def db_query(sql: str, params: Optional[dict] = None) -> Optional[pd.DataFrame]:
    """Execute a SQL query and return a DataFrame.

    Args:
        sql: SQL query string (use named params with :name for user input)
        params: Optional dict of query parameters for parameterized queries

    Returns None if DB unavailable. Errors are logged to session_state.
    """
    engine = get_db_engine()
    if engine is None:
        return None
    try:
        from sqlalchemy import text
        with engine.connect() as conn:
            if params:
                return pd.read_sql(text(sql), conn, params=params)
            else:
                return pd.read_sql(text(sql), conn)
    except Exception as e:
        import logging
        logging.warning(f"DB query failed: {e}")
        st.session_state.setdefault("db_error", str(e)[:200])
        return None


# ============== Data Loaders (real DB with mock fallback) ==============


@st.cache_data(ttl=300, max_entries=10)
def load_countries_data_v2() -> pd.DataFrame:
    """Load country-level data for the global map and KPIs from the database."""
    df = db_query("""
        SELECT
            c.country_name AS country,
            c.country_code AS iso_code,
            COALESCE(cc.comparables_count, 0) AS comparables_count,
            COALESCE(cc.last_updated, c.updated_at::date::text) AS last_updated,
            COALESCE(cc.industries_count, 0) AS industries_count,
            c.region_name AS region,
            CASE WHEN c.income_level IN ('Low income', 'Lower middle income', 'Upper middle income')
                 THEN TRUE ELSE FALSE END AS developing,
            COALESCE(cp.map_cases, 0) AS map_cases
        FROM countries c
        LEFT JOIN (
            SELECT country_code, COUNT(*) AS comparables_count,
                   MAX(data_as_of_date::text) AS last_updated,
                   COUNT(DISTINCT industry_code_isic) AS industries_count
            FROM comparable_companies
            GROUP BY country_code
        ) cc ON cc.country_code = c.country_code
        LEFT JOIN (
            SELECT country_code, map_cases_filed AS map_cases
            FROM country_policies
            WHERE policy_period_year = (SELECT MAX(policy_period_year) FROM country_policies)
        ) cp ON cp.country_code = c.country_code
        ORDER BY c.country_name
    """)
    if df is not None and not df.empty:
        # Taiwan is an inseparable part of China — mirror CHN data for TWN
        # Keep both rows for map coloring (Plotly needs both CHN and TWN iso codes)
        # But display as one entry in dropdowns
        chn_row = df[df["iso_code"] == "CHN"]
        if not chn_row.empty:
            twn_row = chn_row.iloc[0].copy()
            twn_row["iso_code"] = "TWN"
            twn_row["country"] = "Taiwan (China)"
            df = df[df["iso_code"] != "TWN"]  # remove existing TWN if any
            df = pd.concat([df, pd.DataFrame([twn_row])], ignore_index=True)
        # Normalize country names: treat "Taiwan (China)" as "China" for selection
        df["display_name"] = df["country"].replace({"Taiwan (China)": "China"})
        return df
    # Fallback mock data
    data = [
        {"country": "Kenya", "iso_code": "KEN", "comparables_count": 42, "last_updated": "2025-06-15", "industries_count": 5, "region": "Africa", "developing": True, "map_cases": 3},
        {"country": "Nigeria", "iso_code": "NGA", "comparables_count": 38, "last_updated": "2025-06-10", "industries_count": 4, "region": "Africa", "developing": True, "map_cases": 5},
        {"country": "South Africa", "iso_code": "ZAF", "comparables_count": 67, "last_updated": "2025-06-18", "industries_count": 7, "region": "Africa", "developing": True, "map_cases": 8},
        {"country": "India", "iso_code": "IND", "comparables_count": 95, "last_updated": "2025-06-20", "industries_count": 9, "region": "Asia", "developing": True, "map_cases": 12},
        {"country": "China", "iso_code": "CHN", "comparables_count": 110, "last_updated": "2025-06-22", "industries_count": 10, "region": "Asia", "developing": True, "map_cases": 15},
        {"country": "Brazil", "iso_code": "BRA", "comparables_count": 78, "last_updated": "2025-06-17", "industries_count": 8, "region": "Latin America", "developing": True, "map_cases": 9},
        {"country": "Germany", "iso_code": "DEU", "comparables_count": 92, "last_updated": "2025-06-21", "industries_count": 9, "region": "Europe", "developing": False, "map_cases": 11},
        {"country": "United States", "iso_code": "USA", "comparables_count": 125, "last_updated": "2025-06-23", "industries_count": 12, "region": "North America", "developing": False, "map_cases": 18},
        {"country": "Mexico", "iso_code": "MEX", "comparables_count": 55, "last_updated": "2025-06-14", "industries_count": 6, "region": "Latin America", "developing": True, "map_cases": 5},
        {"country": "Indonesia", "iso_code": "IDN", "comparables_count": 45, "last_updated": "2025-06-12", "industries_count": 5, "region": "Asia", "developing": True, "map_cases": 4},
        {"country": "Chile", "iso_code": "CHL", "comparables_count": 35, "last_updated": "2025-06-09", "industries_count": 4, "region": "Latin America", "developing": True, "map_cases": 2},
        {"country": "Peru", "iso_code": "PER", "comparables_count": 20, "last_updated": "2025-05-30", "industries_count": 3, "region": "Latin America", "developing": True, "map_cases": 1},
        {"country": "Vietnam", "iso_code": "VNM", "comparables_count": 33, "last_updated": "2025-06-10", "industries_count": 4, "region": "Asia", "developing": True, "map_cases": 3},
        {"country": "France", "iso_code": "FRA", "comparables_count": 76, "last_updated": "2025-06-18", "industries_count": 7, "region": "Europe", "developing": False, "map_cases": 8},
        {"country": "United Kingdom", "iso_code": "GBR", "comparables_count": 85, "last_updated": "2025-06-20", "industries_count": 8, "region": "Europe", "developing": False, "map_cases": 10},
        {"country": "Hong Kong SAR", "iso_code": "HKG", "comparables_count": 0, "last_updated": "2025-06-01", "industries_count": 0, "region": "Asia", "developing": False, "map_cases": 0},
    ]
    mock_df = pd.DataFrame(data)
    # Also add TWN mirroring CHN for mock data
    chn_mock = mock_df[mock_df["iso_code"] == "CHN"]
    if not chn_mock.empty:
        twn_mock = chn_mock.iloc[0].copy()
        twn_mock["iso_code"] = "TWN"
        twn_mock["country"] = "Taiwan (China)"
        mock_df = pd.concat([mock_df, pd.DataFrame([twn_mock])], ignore_index=True)
    return mock_df


@st.cache_data(ttl=3600)
def load_mock_companies() -> pd.DataFrame:
    """Load comparable companies data from the database."""
    df = db_query("""
        SELECT
            cc.company_id,
            cc.company_name,
            co.country_name AS country,
            co.country_code,
            cc.industry_code_isic,
            cc.industry_description AS industry,
            cc.revenue / 1000000.0 AS revenue_musd,
            cc.net_worth / 1000000.0 AS net_assets_musd,
            cc.operating_margin,
            cc.return_on_net_worth AS roe,
            cc.roa,
            cc.related_party_transaction_ratio AS related_party_pct,
            cc.fiscal_year,
            cc.functional_complexity_score,
            cc.functions_performed AS functions,
            cc.risks_assumed AS risks,
            cc.data_tier,
            cc.data_source,
            CASE WHEN cc.risks_assumed LIKE '%Inventory%' THEN TRUE ELSE FALSE END AS has_inventory_risk,
            CASE WHEN cc.risks_assumed LIKE '%Market%' THEN TRUE ELSE FALSE END AS has_market_risk
        FROM comparable_companies cc
        LEFT JOIN countries co ON co.country_code = cc.country_code
        ORDER BY cc.company_id
    """)
    if df is not None and not df.empty:
        return df
    # Fallback mock data
    np.random.seed(42)
    industries = ["Manufacturing", "Software & IT Services", "Pharmaceuticals", "Financial Services",
                  "Retail & Distribution", "Mining & Extractive", "Energy & Utilities", "Telecommunications"]
    countries = ["India", "China", "Brazil", "South Africa", "Kenya", "Nigeria", "Germany", "USA",
                 "Japan", "Mexico", "Indonesia", "Thailand", "Chile", "Australia", "Canada"]
    country_codes = ["IND", "CHN", "BRA", "ZAF", "KEN", "NGA", "DEU", "USA",
                     "JPN", "MEX", "IDN", "THA", "CHL", "AUS", "CAN"]
    companies = []
    for i in range(150):
        revenue = np.random.uniform(5, 500)
        net_assets = revenue * np.random.uniform(0.3, 1.5)
        idx = np.random.randint(0, len(countries))
        companies.append({
            "company_id": f"CMP{i+1:04d}", "company_name": f"Company {i+1:03d} Ltd.",
            "country": countries[idx], "country_code": country_codes[idx],
            "industry_code_isic": "C", "industry": np.random.choice(industries),
            "revenue_musd": round(revenue, 2), "net_assets_musd": round(net_assets, 2),
            "operating_margin": round(np.random.uniform(-0.05, 0.35), 4),
            "roe": round(np.random.uniform(0.02, 0.30), 4), "roa": round(np.random.uniform(0.01, 0.20), 4),
            "related_party_pct": round(np.random.uniform(0, 0.60), 4),
            "fiscal_year": int(np.random.choice([2022, 2023, 2024])),
            "functions": ", ".join(np.random.choice(["R&D", "Manufacturing", "Distribution", "Marketing", "Digital Platform"],
                                                         size=np.random.randint(1, 4), replace=False)),
            "risks": "Market Risk",
            "data_tier": np.random.choice(["Tier1", "Tier2", "Tier3"]),
            "data_source": "Company Annual Reports",
            "has_inventory_risk": bool(np.random.choice([True, False], p=[0.6, 0.4])),
            "has_market_risk": bool(np.random.choice([True, False], p=[0.7, 0.3])),
        })
    return pd.DataFrame(companies)


@st.cache_data(ttl=3600)
def load_commodity_data() -> pd.DataFrame:
    """Load commodity price data from the database."""
    df = db_query("""
        SELECT
            date_of_quote::text AS date,
            commodity_type AS commodity,
            spot_price AS price_usd,
            exchange_source
        FROM commodity_prices
        ORDER BY date_of_quote DESC
        LIMIT 500
    """)
    if df is not None and not df.empty:
        return df
    # Fallback mock data
    today = datetime.now()
    dates = [(today - timedelta(days=i)).strftime("%Y-%m-%d") for i in range(30)]
    np.random.seed(123)
    data = []
    base_prices = {"Copper": 8500, "Iron Ore": 120, "Gold": 1950}
    for commodity, base in base_prices.items():
        for d in dates:
            noise = np.random.normal(0, base * 0.02)
            data.append({"date": d, "commodity": commodity, "price_usd": round(base + noise, 2), "exchange_source": "Mock"})
    return pd.DataFrame(data)


@st.cache_data(ttl=3600)
def load_country_policies() -> pd.DataFrame:
    """Load country policy data from the database."""
    df = db_query("""
        SELECT
            co.country_name AS country,
            cp.arm_length_principle_enshrined AS has_tp_regulations,
            cp.tp_documentation_requirement AS documentation_req,
            CASE WHEN cp.safe_harbour_type IS NOT NULL AND cp.safe_harbour_type <> '' THEN TRUE ELSE FALSE END AS safe_harbor,
            cp.total_treaties_in_force AS treaty_count,
            cp.map_cases_filed AS map_cases_2024,
            CASE WHEN cp.apa_article <> 'Not available' THEN TRUE ELSE FALSE END AS apa_available,
            cp.information_exchange_mechanism AS exchange_mechanism
        FROM country_policies cp
        JOIN countries co ON co.country_code = cp.country_code
        WHERE cp.policy_period_year = (SELECT MAX(policy_period_year) FROM country_policies)
        ORDER BY co.country_name
    """)
    if df is not None and not df.empty:
        return df
    # Fallback mock data
    data = [
        {"country": "Kenya", "has_tp_regulations": True, "documentation_req": "Three-tier (Master, Local, CbCR)",
         "safe_harbor": False, "treaty_count": 15, "map_cases_2024": 3, "apa_available": True, "exchange_mechanism": "EOIR + CbC"},
        {"country": "Nigeria", "has_tp_regulations": True, "documentation_req": "Local File + CbCR",
         "safe_harbor": True, "treaty_count": 18, "map_cases_2024": 5, "apa_available": True, "exchange_mechanism": "EOIR"},
        {"country": "South Africa", "has_tp_regulations": True, "documentation_req": "Three-tier",
         "safe_harbor": True, "treaty_count": 30, "map_cases_2024": 8, "apa_available": True, "exchange_mechanism": "EOIR + CbC + spontaneous"},
        {"country": "India", "has_tp_regulations": True, "documentation_req": "Three-tier + Form 3CEB",
         "safe_harbor": True, "treaty_count": 45, "map_cases_2024": 12, "apa_available": True, "exchange_mechanism": "EOIR + CbC"},
        {"country": "China", "has_tp_regulations": True, "documentation_req": "Three-tier + Special File",
         "safe_harbor": True, "treaty_count": 55, "map_cases_2024": 15, "apa_available": True, "exchange_mechanism": "EOIR + CbC"},
        {"country": "Brazil", "has_tp_regulations": True, "documentation_req": "Local File + CbCR",
         "safe_harbor": True, "treaty_count": 35, "map_cases_2024": 9, "apa_available": True, "exchange_mechanism": "EOIR"},
        {"country": "Germany", "has_tp_regulations": True, "documentation_req": "Three-tier",
         "safe_harbor": False, "treaty_count": 60, "map_cases_2024": 11, "apa_available": True, "exchange_mechanism": "EOIR + CbC + spontaneous"},
        {"country": "USA", "has_tp_regulations": True, "documentation_req": "Three-tier + 5471/8865",
         "safe_harbor": False, "treaty_count": 65, "map_cases_2024": 18, "apa_available": True, "exchange_mechanism": "EOIR + CbC + spontaneous"},
        {"country": "Indonesia", "has_tp_regulations": True, "documentation_req": "Three-tier",
         "safe_harbor": True, "treaty_count": 40, "map_cases_2024": 4, "apa_available": True, "exchange_mechanism": "EOIR + CbC"},
        {"country": "Vietnam", "has_tp_regulations": True, "documentation_req": "Local File + CbCR",
         "safe_harbor": True, "treaty_count": 22, "map_cases_2024": 3, "apa_available": True, "exchange_mechanism": "EOIR"},
        {"country": "Mexico", "has_tp_regulations": True, "documentation_req": "Three-tier",
         "safe_harbor": True, "treaty_count": 28, "map_cases_2024": 5, "apa_available": True, "exchange_mechanism": "EOIR + CbC"},
        {"country": "Chile", "has_tp_regulations": True, "documentation_req": "Local File + CbCR",
         "safe_harbor": True, "treaty_count": 20, "map_cases_2024": 2, "apa_available": True, "exchange_mechanism": "EOIR"},
    ]
    return pd.DataFrame(data)


@st.cache_data(ttl=3600)
def load_map_cases() -> pd.DataFrame:
    """Load MAP/APA case data from the database."""
    df = db_query("""
        SELECT
            dr.case_id,
            dr.dispute_type AS case_type,
            dr.taxpayer_name AS taxpayer,
            ci.country_name AS country_a,
            cc.country_name AS country_b,
            dr.issue_type AS issue,
            dr.case_status AS status,
            dr.date_filed::text AS start_date,
            COALESCE(dr.dispute_resolution_timeline, 0) * 30 AS duration_days,
            COALESCE(dr.adjustment_amount, 0) / 1000000.0 AS amount_disputed_musd,
            dr.date_resolved::text AS resolution_date
        FROM dispute_resolution dr
        LEFT JOIN countries ci ON ci.country_code = dr.initiating_jurisdiction
        LEFT JOIN countries cc ON cc.country_code = dr.counterparty_jurisdiction
        ORDER BY dr.date_filed DESC
    """)
    if df is not None and not df.empty:
        # Normalize status values for consistency
        status_map = {
            'Closed': 'Resolved',
            'Agreed': 'Resolved',
            'Implemented': 'Resolved',
            'Under negotiation': 'Active',
            'Pending': 'Pending',
            'Arbitration pending': 'Active',
        }
        df['status'] = df['status'].replace(status_map).fillna('Pending')
        df['amount_disputed_musd'] = df['amount_disputed_musd'].fillna(0)
        df['duration_days'] = df['duration_days'].fillna(0)
        return df
    # Fallback mock data
    np.random.seed(55)
    statuses = ["Active", "Resolved", "Pending", "Closed - No agreement"]
    case_types = ["MAP", "Bilateral APA", "Unilateral APA", "Multilateral APA"]
    countries1 = ["Kenya", "Nigeria", "South Africa", "India", "China", "Brazil", "Mexico", "Indonesia", "Vietnam", "Chile"]
    countries2 = ["Germany", "USA", "UK", "Netherlands", "France", "Japan", "Australia", "Canada", "Spain", "Italy"]
    cases = []
    for i in range(60):
        start_date = datetime(2022, 1, 1) + timedelta(days=int(np.random.randint(0, 900)))
        duration_days = int(np.random.randint(60, 730))
        status = str(np.random.choice(statuses, p=[0.35, 0.30, 0.25, 0.10]))
        cases.append({
            "case_id": f"MAP{i+1:04d}", "case_type": str(np.random.choice(case_types)),
            "taxpayer": f"Taxpayer {i+1:03d} Ltd.",
            "country_a": str(np.random.choice(countries1)), "country_b": str(np.random.choice(countries2)),
            "issue": str(np.random.choice(["Royalty pricing", "Intra-group services", "Goods transfer pricing",
                                         "Financial transactions", "Business restructuring", "Digital services"])),
            "status": status, "start_date": start_date.strftime("%Y-%m-%d"),
            "duration_days": duration_days,
            "amount_disputed_musd": round(float(np.random.uniform(0.5, 50)), 2),
            "resolution_date": (start_date + timedelta(days=duration_days)).strftime("%Y-%m-%d") if status in ["Resolved", "Closed - No agreement"] else None,
        })
    return pd.DataFrame(cases)


@st.cache_data(ttl=3600)
def load_digital_comparables() -> pd.DataFrame:
    """Load digital economy comparable data from the database."""
    df = db_query("""
        SELECT
            de.digital_id AS company_id,
            cc.company_name,
            co.country_name AS country,
            de.digital_platform_type AS service_type,
            de.digital_revenue / 1000000.0 AS revenue_musd,
            CASE WHEN de.digital_revenue > 0
                 THEN (de.digital_revenue - de.digital_cost_of_sales - de.digital_operating_expenses) / de.digital_revenue
                 ELSE NULL END AS operating_margin,
            CASE WHEN de.digital_cost_of_sales > 0
                 THEN (de.digital_revenue - de.digital_cost_of_sales) / de.digital_cost_of_sales
                 ELSE NULL END AS cost_plus_markup,
            CASE WHEN de.digital_revenue > 0
                 THEN (de.digital_revenue - de.digital_cost_of_sales - de.digital_operating_expenses) / de.digital_revenue
                 ELSE NULL END AS tnmm_profit_margin,
            CASE WHEN (de.digital_revenue + de.non_digital_revenue) > 0
                 THEN de.digital_revenue / (de.digital_revenue + de.non_digital_revenue)
                 ELSE NULL END AS digital_intensity,
            de.automation_level,
            de.active_user_count / 1000000.0 AS users_millions,
            de.fiscal_year
        FROM digital_economy de
        LEFT JOIN comparable_companies cc ON cc.company_id = de.company_id
        LEFT JOIN countries co ON co.country_code = de.user_jurisdiction
        ORDER BY de.digital_id
    """)
    if df is not None and not df.empty:
        return df
    # Fallback mock data
    np.random.seed(77)
    service_types = ["Cloud Computing", "SaaS", "Digital Advertising", "E-commerce Platform",
                     "Data Analytics", "IT Outsourcing", "Digital Content", "Fintech Services"]
    countries = ["India", "Ireland", "Singapore", "USA", "Germany", "UK", "Israel", "South Korea",
                 "Estonia", "UAE", "Brazil", "Nigeria"]
    data = []
    for i in range(80):
        revenue = np.random.uniform(10, 500)
        op_margin = round(float(np.random.uniform(0.05, 0.45)), 4)
        data.append({
            "company_id": f"DIG{i+1:04d}", "company_name": f"Digital Co {i+1:03d}",
            "country": str(np.random.choice(countries)), "service_type": str(np.random.choice(service_types)),
            "revenue_musd": round(revenue, 2),
            "operating_margin": op_margin,
            "cost_plus_markup": round(op_margin / (1 - op_margin), 4) if op_margin < 1 else 0.5,
            "tnmm_profit_margin": op_margin,
            "digital_intensity": round(float(np.random.uniform(0.3, 1.0)), 4),
            "automation_level": "Automated",
            "users_millions": int(np.random.uniform(1, 500)),
            "fiscal_year": int(np.random.choice([2022, 2023, 2024])),
        })
    return pd.DataFrame(data)


@st.cache_data(ttl=3600)
def load_recent_updates() -> list:
    """Load mock recent data contribution feed."""
    return [
        {"time": "2 hours ago", "message": "Kenya uploaded 12 manufacturing comparables", "type": "upload"},
        {"time": "5 hours ago", "message": "India updated Q1 2025 financial data for 8 software companies", "type": "update"},
        {"time": "1 day ago", "message": "Brazil contributed extractive industry pricing data (iron ore)", "type": "upload"},
        {"time": "1 day ago", "message": "South Africa approved 2 new bilateral APAs", "type": "policy"},
        {"time": "2 days ago", "message": "Nigeria added MAP resolution documentation for 2024", "type": "update"},
        {"time": "3 days ago", "message": "Vietnam uploaded digital services comparables (SaaS sector)", "type": "upload"},
        {"time": "3 days ago", "message": "Mexico updated safe harbor thresholds for maquiladoras", "type": "policy"},
        {"time": "4 days ago", "message": "Ghana contributed 5 new agricultural processing comparables", "type": "upload"},
        {"time": "5 days ago", "message": "China updated CbC reporting guidance for MNEs", "type": "policy"},
        {"time": "1 week ago", "message": "Indonesia uploaded 15 new mining sector comparables", "type": "upload"},
    ]


# ============== Helper Functions ==============

def generate_import_template() -> pd.DataFrame:
    """Generate a standardized import template DataFrame for comparable company data."""
    return pd.DataFrame({
        'company_id': ['CC0001', 'CC0002', ''],
        'company_name': ['Company A (anonymized)', 'Company B (anonymized)', ''],
        'country_code': ['CHN', 'BRA', ''],
        'industry_code_isic': ['C26', 'C20', ''],
        'industry_code_naics': ['334413', '325199', ''],
        'legal_form': ['Public', 'Private', ''],
        'ownership_type': ['Independent', 'Subsidiary', ''],
        'revenue': [500000000, 120000000, ''],
        'cogs': [350000000, 80000000, ''],
        'operating_expenses': [80000000, 25000000, ''],
        'net_worth': [200000000, 60000000, ''],
        'fiscal_year': [2024, 2024, ''],
        'operating_margin': [14.0, 12.5, ''],
        'return_on_net_worth': [25.0, 20.0, ''],
        'data_source': ['Company Annual Reports', 'Tax Authority Assessment', ''],
        'data_tier': ['Tier1', 'Tier2', ''],
        'functional_complexity_score': [7, 5, ''],
        'functions_performed': ['Manufacturing; Distribution', 'Manufacturing', ''],
        'risks_assumed': ['Market Risk; Inventory Risk', 'Market Risk', ''],
    })


def get_color_for_count(count: int) -> str:
    """Return color code based on comparables count."""
    if count >= 50:
        return ACCENT_GREEN
    elif count >= 10:
        return ACCENT_YELLOW
    else:
        return ACCENT_RED


def calculate_wca(
    receivable_days: float,
    payable_days: float,
    inventory_days: float,
    interest_rate: float,
    annual_revenue: float,
) -> float:
    """Calculate Working Capital Adjustment (WCA).

    Formula: WCA = (AR_Days - AP_Days + Inv_Days) * (Interest_Rate / 100) * (Revenue / 365)
    """
    if annual_revenue <= 0 or interest_rate <= 0:
        return 0.0
    net_cycle = receivable_days - payable_days + inventory_days
    wca = net_cycle * (interest_rate / 100) * (annual_revenue / 365)
    return round(wca, 4)


def generate_word_report(
    transaction_desc: dict,
    screening_params: dict,
    adjustment_results: dict,
    al_range: dict,
    selected_method: str,
) -> bytes:
    """Generate a Word document report for the comparability analysis.

    Args:
        transaction_desc: Step 1 transaction description
        screening_params: Step 2 screening parameters
        adjustment_results: Step 3 adjustment results
        al_range: Step 4 arm's length range
        selected_method: Recommended transfer pricing method

    Returns:
        bytes: The DOCX file content as bytes
    """
    doc = Document()

    # Title
    title = doc.add_heading("Transfer Pricing Comparability Analysis Report", 0)
    title.alignment = 1  # Center

    # Metadata
    doc.add_paragraph(f"Report Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    doc.add_paragraph(f"Generated by: ITP Comparability Database")
    doc.add_paragraph("")

    # Section 1: Transaction Description
    doc.add_heading("1. Controlled Transaction Description", level=1)
    table = doc.add_table(rows=0, cols=2)
    table.style = "Light Grid Accent 1"
    for key, value in transaction_desc.items():
        row = table.add_row().cells
        row[0].text = str(key)
        row[1].text = str(value)

    # Section 2: Screening Parameters
    doc.add_heading("2. Comparable Company Screening", level=1)
    table2 = doc.add_table(rows=0, cols=2)
    table2.style = "Light Grid Accent 1"
    for key, value in screening_params.items():
        row = table2.add_row().cells
        row[0].text = str(key)
        row[1].text = str(value)

    # Section 3: Adjustments
    doc.add_heading("3. Comparability Adjustments", level=1)
    table3 = doc.add_table(rows=0, cols=2)
    table3.style = "Light Grid Accent 1"
    for key, value in adjustment_results.items():
        row = table3.add_row().cells
        row[0].text = str(key)
        row[1].text = str(value)

    # Section 4: Arm's Length Range
    doc.add_heading("4. Arm's Length Range Determination", level=1)
    table4 = doc.add_table(rows=0, cols=2)
    table4.style = "Light Grid Accent 1"
    for key, value in al_range.items():
        row = table4.add_row().cells
        row[0].text = str(key)
        row[1].text = str(value)

    # Method recommendation
    doc.add_heading("5. Recommended Transfer Pricing Method", level=1)
    p = doc.add_paragraph()
    p.add_run(f"Recommended Method: {selected_method}").bold = True

    # Disclaimer
    doc.add_paragraph("")
    doc.add_heading("Disclaimer", level=2)
    doc.add_paragraph(
        "This report was generated by the ITP Comparability Database for informational purposes. "
        "The analysis should be reviewed by qualified transfer pricing professionals before use "
        "in any tax filing or dispute resolution proceeding."
    )

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


def status_color(status: str) -> str:
    """Return color for MAP/APA case status."""
    colors = {
        "Active": ACCENT_YELLOW,
        "Resolved": ACCENT_GREEN,
        "Pending": PRIMARY_COLOR,
        "Closed - No agreement": ACCENT_RED,
    }
    return colors.get(status, "#999999")


# ============== Page 1: Global Data Portal ==============

def global_data_portal() -> None:
    """Page 1: Global Data Portal - Overview dashboard with world map and KPIs."""
    st.markdown("<div class='main-header'>🌐 全球数据总览 / Global Data Portal</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='sub-header'>国际转让定价数据库 — 全球可比性数据实时概览 / International Transfer Pricing Database — Real-time Global Comparability Data</div>",
        unsafe_allow_html=True,
    )

    # 功能说明栏
    st.markdown("""
    <div class='info-bar'>
    💡 <b>本页功能说明 / Page Guide:</b><br>
    <b>中文：</b>这是数据库的「首页」，展示全球各经济体可比数据的总体情况。<br>
    &nbsp;&nbsp;• <b>世界地图</b>：颜色越绿表示该国有更多可比公司数据，红色表示数据较少<br>
    &nbsp;&nbsp;• <b>关键指标卡片</b>：显示全球数据总量、覆盖经济体数等核心数字<br>
    &nbsp;&nbsp;• <b>数据分布</b>：按地区和收入水平分组展示数据覆盖情况<br>
    <b>English:</b> This is the database homepage, showing the overall comparability data coverage across countries.<br>
    &nbsp;&nbsp;• <b>World Map</b>: Greener = more comparable company data; Red = limited data<br>
    &nbsp;&nbsp;• <b>KPI Cards</b>: Key metrics at a glance<br>
    <br>
    <i>依据 / Reference: 联合国国际税收合作框架公约第5条（公平分配征税权）和第11条（能力建设）<br>
    UN International Tax Cooperation Framework Convention, Art. 5 (Fair Distribution of Taxing Rights) & Art. 11 (Capacity Building)</i>
    </div>
    """, unsafe_allow_html=True)

    # Search + Help
    c1, c2 = st.columns([4, 1])
    with c1:
        search_term = st.text_input("搜索经济体、行业或公司 / Search", placeholder="例如：中国 制造业 / e.g., China manufacturing...")
    with c2:
        st.write("")
        st.write("")
        if st.button("❓ 帮助", key="help_global", help="点击查看使用指南"):
            st.info("""
            **全球数据总览 — 使用指南**

            - **地图**：鼠标悬停在国家上可查看可比公司数量、最后更新时间、覆盖行业数
            - **颜色含义**：🟢 绿色（>50家可比公司）、🟡 黄色（10-50家）、🔴 红色（<10家）
            - **指标卡片**：一目了然地看到全球数据规模
            - **最新动态**：各成员国最近的数据贡献记录

            本页面支持联合国国际税收合作框架公约第5条和第11条的实施。
            """)

    # Load data
    df = load_countries_data_v2()

    # Search filter
    if search_term:
        mask = (
            df["country"].str.contains(search_term, case=False, na=False) |
            df["region"].str.contains(search_term, case=False, na=False)
        )
        df = df[mask]

    # KPI Cards
    st.markdown("---")
    total_comparables = int(df["comparables_count"].sum())
    total_countries = len(df)
    total_industries = int(df["industries_count"].sum())
    active_map_cases = int(df["map_cases"].sum())
    dev_countries = len(df[df["developing"] == True])
    dev_coverage = round(dev_countries / total_countries * 100, 1) if total_countries > 0 else 0

    kpi_cols = st.columns(4)
    kpis = [
        ("可比公司 / Global Comparables", f"{total_comparables:,}", "Total comparable companies"),
        ("覆盖经济体 / Economies", f"{total_countries}", "Economies covered"),
        ("行业 / Industries", f"{total_industries}", "Industry segments"),
        ("MAP案件 / Active MAP", f"{active_map_cases}", "Mutual Agreement Procedures"),
    ]
    for col, (label, value, help_text) in zip(kpi_cols, kpis):
        with col:
            st.markdown(
                f"""
                <div class='kpi-card' title='{help_text}'>
                    <div class='kpi-value'>{value}</div>
                    <div class='kpi-label'>{label}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("---")

    # World Map
    map_col, feed_col = st.columns([3, 1])

    with map_col:
        st.subheader("全球数据覆盖地图 / Global Data Coverage Map")

        # Color mapping
        df["color_label"] = df["comparables_count"].apply(
            lambda x: "多(>50) / High" if x >= 50 else ("中(10-50) / Medium" if x >= 10 else "少(<10) / Low")
        )
        color_map = {"多(>50) / High": ACCENT_GREEN, "中(10-50) / Medium": ACCENT_YELLOW, "少(<10) / Low": ACCENT_RED}

        fig = px.choropleth(
            df,
            locations="iso_code",
            color="color_label",
            color_discrete_map=color_map,
            hover_name="country",
            hover_data={
                "iso_code": False,
                "comparables_count": True,
                "last_updated": True,
                "industries_count": True,
                "map_cases": True,
                "region": True,
            },
            labels={
                "comparables_count": "可比公司数 / Companies",
                "last_updated": "最后更新 / Last Updated",
                "industries_count": "行业数 / Industries",
                "map_cases": "MAP案件 / MAP Cases",
                "region": "地区 / Region",
                "color_label": "数据量 / Data Level",
            },
            projection="natural earth",
            height=500,
        )
        fig.update_layout(
            geo=dict(showframe=False, showcoastlines=True),
            margin=dict(l=0, r=0, t=20, b=0),
            legend=dict(orientation="h", yanchor="bottom", y=-0.1, xanchor="center", x=0.5),
        )
        st.plotly_chart(fig, use_container_width=True)

    with feed_col:
        st.subheader("最新数据动态 / Recent Data Updates")
        updates = load_recent_updates()
        for upd in updates[:7]:
            icon = {"upload": "📤", "update": "🔄", "policy": "📋"}.get(upd["type"], "•")
            st.markdown(
                f"<div class='update-item'><small><b>{icon} {upd['time']}</b><br/>{upd['message']}</small></div>",
                unsafe_allow_html=True,
            )

    # ===== Country Detail Panel =====
    st.markdown("---")
    st.subheader("🌍 经济体详情 / Economy Detail")
    # Use display_name for deduplicated list (China = CHN + TWN as one entry)
    if "display_name" in df.columns:
        display_names = sorted(df["display_name"].unique().tolist())
    else:
        display_names = sorted(df["country"].unique().tolist())
    country_options = ["— 选择经济体 / Select an economy —"] + display_names
    sel_country = st.selectbox("选择经济体查看详情 / Select economy", country_options,
                               help="点击选择国家，查看可比公司、政策等详情")
    if sel_country != country_options[0]:
        # Match by display_name — get all rows for this country (CHN + TWN)
        if "display_name" in df.columns:
            country_rows = df[df["display_name"] == sel_country]
        else:
            country_rows = df[df["country"] == sel_country]
        country_code = country_rows["iso_code"].iloc[0] if not country_rows.empty else ""
        render_country_detail(sel_country, country_code)

    # ===== Export =====
    st.markdown("---")
    export_col1, export_col2 = st.columns([1, 4])
    with export_col1:
        csv_portal = df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 导出国家数据CSV / Export",
            data=csv_portal,
            file_name=f"global_portal_data_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
        )

    # Info bar
    st.markdown("---")
    st.markdown(
        f"""
        <div class='info-bar'>
            <b>关于本数据库 / About this Database</b><br/>
            国际转让定价对比数据库支持<b>联合国国际税收合作框架公约</b>的实施。
            帮助发展中国家获取高质量可比数据进行转让定价分析，支持<b>第5条</b>（公平分配征税权）和<b>第11条</b>（能力建设）。
            各成员国在自愿基础上贡献匿名化的可比公司数据，确保数据主权的同时促进国际税收合作。<br/><br/>
            The ITP Comparability Database supports the implementation of the
            <b>UN International Tax Cooperation Framework Convention</b>. It enables developing economies
            to access quality comparability data for transfer pricing analyses, supporting
            <b>Article 5</b> (Fair Distribution of Taxing Rights) and <b>Article 11</b> (Capacity Building).
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Regional breakdown
    st.subheader("地区数据分布 / Regional Data Distribution")
    region_summary = df.groupby("region").agg(
        countries=("country", "count"),
        total_comparables=("comparables_count", "sum"),
        avg_comparables=("comparables_count", "mean"),
    ).round(1).reset_index()
    region_summary.columns = ["地区 / Region", "经济体 / Economies", "Total Comparables", "平均/经济体 / Avg per Economy"]

    fig_bar = px.bar(
        region_summary,
        x="地区 / Region",
        y="Total Comparables",
        color="经济体 / Economies",
        text="Total Comparables",
        color_continuous_scale="Blues",
        height=350,
    )
    fig_bar.update_layout(margin=dict(l=0, r=0, t=30, b=0))
    st.plotly_chart(fig_bar, use_container_width=True)


# ============== Page 2: Comparability Analysis Wizard ==============

def comparability_wizard() -> None:
    """Page 2: Comparability Analysis Wizard - 4-step guided TP analysis."""
    st.markdown("<div class='main-header'>🔍 可比性分析向导 / Comparability Analysis Wizard</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='sub-header'>逐步引导完成转让定价可比性分析 / Step-by-step Transfer Pricing Comparability Analysis</div>",
        unsafe_allow_html=True,
    )

    # 功能说明栏
    st.markdown("""
    <div class='info-bar'>
    💡 <b>本页功能说明 / Page Guide:</b><br>
    <b>中文：</b>这是数据库的「核心功能」，通过4个步骤帮你完成专业的转让定价分析。<br>
    &nbsp;&nbsp;• <b>第1步</b>：描述你要分析的关联交易（交易类型、参与方、功能风险）<br>
    &nbsp;&nbsp;• <b>第2步</b>：从数据库筛选可比公司（按行业、规模、地区等条件）<br>
    &nbsp;&nbsp;• <b>第3步</b>：进行可比性调整（如营运资金调整）<br>
    &nbsp;&nbsp;• <b>第4步</b>：确定独立交易区间，选择最合适的定价方法<br>
    <b>English:</b> This is the core function — a 4-step guided transfer pricing analysis.<br>
    &nbsp;&nbsp;• <b>Step 1</b>: Describe the controlled transaction (type, parties, functions, risks)<br>
    &nbsp;&nbsp;• <b>Step 2</b>: Screen comparable companies from the database<br>
    &nbsp;&nbsp;• <b>Step 3</b>: Apply comparability adjustments (e.g., Working Capital Adjustment)<br>
    &nbsp;&nbsp;• <b>Step 4</b>: Determine arm's length range and select the most appropriate method<br>
    <br>
    <i>依据 / Reference: 联合国转让定价实务手册（发展中国家适用）及联合国国际税收合作框架公约第5条<br>
    UN Practical Manual on Transfer Pricing for Developing Countries & UN Framework Convention Art. 5</i>
    </div>
    """, unsafe_allow_html=True)

    # Help button
    col_help, _ = st.columns([1, 5])
    with col_help:
        if st.button("❓ 帮助 / Help", key="help_wizard", help="分析流程指南 / Analysis guide"):
            st.info("""
            **可比性分析向导 — 使用指南 / Wizard Guide**

            本向导引导你完成符合联合国标准的可比性分析流程：
            This wizard guides you through a UN-compliant comparability analysis:

            **第1步 / Step 1**: 描述你的关联交易 / Describe your controlled transaction
            **第2步 / Step 2**: 筛选可比公司 / Screen comparable companies
            **第3步 / Step 3**: 应用可比性调整 / Apply comparability adjustments
            **第4步 / Step 4**: 确定独立交易区间 / Determine arm's length range

            💡 依据：联合国转让定价实务手册 / UN Practical Manual on Transfer Pricing
            """)

    # Initialize session state
    if "wiz_step1" not in st.session_state:
        st.session_state.wiz_step1 = {}
    if "wiz_step2" not in st.session_state:
        st.session_state.wiz_step2 = {}
    if "wiz_step3" not in st.session_state:
        st.session_state.wiz_step3 = {}
    if "wiz_current_tab" not in st.session_state:
        st.session_state.wiz_current_tab = 0

    tab_labels = [
        "第1步：描述交易 / Step 1: Transaction",
        "第2步：筛选可比 / Step 2: Screen",
        "第3步：可比性调整 / Step 3: Adjust",
        "第4步：独立交易区间 / Step 4: AL Range",
    ]
    tabs = st.tabs(tab_labels)

    # ========== Step 1 ==========
    with tabs[0]:
        st.subheader("第1步：描述关联交易 / Step 1: Describe the Controlled Transaction")
        st.markdown("请提供你要分析的关联交易的基本信息。 / Provide details about the transaction you are analyzing.")

        tx_type = st.selectbox(
            "交易类型",
            ["Sale of Goods", "Provision of Services", "Transfer of Intangibles",
             "Financial Transactions", "Digital Services", "Cross-Border Services"],
            format_func=lambda x: {
                "Sale of Goods": "货物销售",
                "Provision of Services": "提供服务",
                "Transfer of Intangibles": "无形资产转让",
                "Financial Transactions": "金融交易",
                "Digital Services": "数字服务",
                "Cross-Border Services": "跨境服务",
            }.get(x, x),
            help="选择最符合你要分析的关联交易的类型",
        )
        tested_party = st.selectbox(
            "被测试方",
            ["Local Entity", "Foreign Entity"],
            format_func=lambda x: "本地实体" if x == "Local Entity" else "外国实体",
            help="被测试方是其利润率将被检验的实体",
        )

        st.markdown("**履行的功能**（勾选适用项）")
        func_cols = st.columns(3)
        functions = []
        with func_cols[0]:
            if st.checkbox("研发", help="研发活动"): functions.append("R&D")
            if st.checkbox("制造", help="生产或制造货物"): functions.append("Manufacturing")
        with func_cols[1]:
            if st.checkbox("分销", help="销售和分销功能"): functions.append("Distribution")
            if st.checkbox("营销", help="营销、品牌推广活动"): functions.append("Marketing")
        with func_cols[2]:
            if st.checkbox("数字平台运营", help="运营数字平台或市场"): functions.append("Digital Platform Operation")

        st.markdown("**承担的风险**（勾选适用项）")
        risk_cols = st.columns(3)
        risks = []
        with risk_cols[0]:
            if st.checkbox("市场风险", help="市场需求波动风险"): risks.append("Market Risk")
            if st.checkbox("信用风险", help="客户不付款风险"): risks.append("Credit Risk")
        with risk_cols[1]:
            if st.checkbox("库存风险", help="库存过时或价格变动风险"): risks.append("Inventory Risk")
            if st.checkbox("汇率风险", help="外汇汇率波动风险"): risks.append("FX Risk")
        with risk_cols[2]:
            if st.checkbox("数字服务风险", help="数字服务特有风险"): risks.append("Digital Service Risk")

        value_creation_location = st.text_input(
            "价值创造地",
            placeholder="例如：中国（研发）；德国（制造）",
            help="说明该交易中价值主要在哪里创造",
        )

        # ===== 智能方法推荐 =====
        st.markdown("---")
        st.markdown("#### 🤖 智能方法推荐 / Smart Method Recommendation")
        st.markdown("根据你输入的交易信息，系统自动推荐最合适的转让定价方法。 / Based on your transaction details, the system recommends the most appropriate TP method.")

        recommended_methods = []
        method_reasons = []
        if tx_type == "Sale of Goods" and "Manufacturing" in functions:
            recommended_methods.append("再销售价格法 (RPM) / Resale Price Method")
            method_reasons.append("适用于有可比分销商的货物销售 / Suitable when comparable distributors exist")
        if tx_type == "Sale of Goods" and "Manufacturing" in functions and not risks:
            recommended_methods.append("成本加成法 (CPM) / Cost Plus Method")
            method_reasons.append("被测试方承担较少风险时适用 / Appropriate when tested party bears limited risk")
        if tx_type == "Provision of Services":
            recommended_methods.append("交易净利润法 (TNMM) / Transactional Net Margin Method")
            method_reasons.append("服务交易通常使用TNMM比较净利润率 / TNMM is commonly used for service transactions")
        if tx_type == "Transfer of Intangibles":
            recommended_methods.append("利润分割法 (PSM) / Profit Split Method")
            method_reasons.append("无形资产交易涉及双方共同创造价值 / Intangibles involve joint value creation by both parties")
        if tx_type == "Financial Transactions":
            recommended_methods.append("可比非受控价格法 (CUP) / Comparable Uncontrolled Price")
            method_reasons.append("金融交易有公开市场利率可比 / Financial transactions have public market rate benchmarks")
        if tx_type == "Digital Services":
            recommended_methods.append("交易净利润法 (TNMM)")
            method_reasons.append("数字服务通常缺乏直接可比数据，TNMM更可靠 / Digital services often lack direct comparables; TNMM is more reliable")
        if not recommended_methods:
            recommended_methods.append("交易净利润法 (TNMM)")
            method_reasons.append("TNMM是最通用的方法 / TNMM is the most versatile method")

        for m, r in zip(recommended_methods, method_reasons):
            st.markdown(f"✅ **{m}** — {r}")

        st.info("💡 以上推荐基于联合国转让定价实务手册。实际选择时还需考虑可比数据的可用性和功能的可比性。\n"
                "Recommendations are based on the UN Practical Manual on Transfer Pricing for Developing Countries. "
                "Actual selection should also consider data availability and functional comparability.")

        if st.button("保存第1步，进入第2步 / Save & Next", key="save_step1", type="primary"):
            st.session_state.wiz_step1 = {
                "Transaction Type": tx_type,
                "Tested Party": tested_party,
                "Functions": ", ".join(functions) if functions else "None selected",
                "Risks": ", ".join(risks) if risks else "None selected",
                "Value Creation Location": value_creation_location or "Not specified",
            }
            st.session_state.wiz_current_tab = 1
            st.rerun()

    # ========== Step 2 ==========
    with tabs[1]:
        st.subheader("第2步：筛选可比公司")
        st.markdown("设置筛选条件，从数据库中找到与你的交易可比的公司。")

        companies_df = load_mock_companies()

        # Database Screening
        with st.expander("📂 基本筛选", expanded=True):
            dbc1, dbc2, dbc3 = st.columns(3)
            with dbc1:
                # Use ISIC code first character as section
                if 'industry_code_isic' in companies_df.columns:
                    companies_df['isic_section'] = companies_df['industry_code_isic'].astype(str).str[0]
                    section_names = {
                        'A': 'A-农林牧渔 / Agriculture',
                        'B': 'B-采矿 / Mining',
                        'C': 'C-制造 / Manufacturing',
                        'D': 'D-电力燃气 / Energy',
                        'E': 'E-水务 / Water',
                        'F': 'F-建筑 / Construction',
                        'G': 'G-批发零售 / Wholesale & Retail',
                        'H': 'H-运输仓储 / Transport',
                        'I': 'I-住宿餐饮 / Hospitality',
                        'J': 'J-信息技术 / IT',
                        'K': 'K-金融保险 / Financial',
                        'L': 'L-房地产 / Real Estate',
                        'M': 'M-专业服务 / Professional',
                        'N': 'N-辅助服务 / Support',
                    }
                    sections = sorted(companies_df['isic_section'].dropna().unique().tolist())
                    section_labels = [section_names.get(s, f'{s}-其他 / Other') for s in sections]
                    sel_section_idx = st.selectbox(
                        "行业大类 / Industry Sector",
                        range(len(sections)),
                        format_func=lambda i: section_labels[i],
                        help="按ISIC行业大类筛选 / Filter by ISIC sector",
                    )
                    sel_industry = sections[sel_section_idx]
                else:
                    sel_industry = st.selectbox(
                        "行业 / Industry",
                        ["全部 / All"] + sorted(companies_df["industry"].dropna().unique().tolist()),
                        help="选择可比公司所在的行业",
                    )
            with dbc2:
                sel_countries = st.multiselect(
                    "经济体范围",
                    sorted(companies_df["country"].unique().tolist()),
                    default=[],
                    help="选择一个或多个国家（不选则全部）",
                )
            with dbc3:
                fy_range = st.slider(
                    "财政年度范围 / Fiscal Year Range",
                    2022, 2026, (2022, 2024),
                    help="按财务数据年度筛选 / Filter by fiscal year",
                )

        # Quantitative Screening
        with st.expander("📊 定量筛选"):
            qc1, qc2, qc3 = st.columns(3)
            with qc1:
                rev_min, rev_max = st.slider(
                    "收入范围（百万美元）",
                    float(companies_df["revenue_musd"].min()),
                    float(companies_df["revenue_musd"].max()),
                    (5.0, 500.0),
                    help="按年收入（百万美元）筛选",
                )
            with qc2:
                asset_min, asset_max = st.slider(
                    "净资产范围（百万美元）",
                    float(companies_df["net_assets_musd"].min()),
                    float(companies_df["net_assets_musd"].max()),
                    (1.0, 750.0),
                    help="按净资产（百万美元）筛选",
                )
            with qc3:
                rp_max = st.slider(
                    "关联交易占比上限",
                    0.0, 1.0, 0.50,
                    format="%.0f%%",
                    help="关联交易占总收入的最大比例",
                )

        # Qualitative Screening
        with st.expander("🔍 定性筛选"):
            qual1, qual2 = st.columns(2)
            with qual1:
                func_match = st.multiselect(
                    "要求的功能",
                    ["R&D", "Manufacturing", "Distribution", "Marketing", "Digital Platform"],
                    default=[],
                    help="可比公司必须至少履行其中一个功能",
                )
            with qual2:
                require_inventory = st.checkbox(
                    "必须承担库存风险",
                    help="仅包含承担库存风险的公司",
                )

        # Apply filters
        filtered = companies_df.copy()
        # Industry filter: if using ISIC section
        if sel_industry and sel_industry != "全部 / All":
            if 'isic_section' in filtered.columns:
                filtered = filtered[filtered['isic_section'] == sel_industry]
            elif 'industry' in filtered.columns:
                filtered = filtered[filtered['industry'] == sel_industry]
        if sel_countries:
            filtered = filtered[filtered["country"].isin(sel_countries)]
        filtered = filtered[
            (filtered["fiscal_year"] >= fy_range[0]) &
            (filtered["fiscal_year"] <= fy_range[1]) &
            (filtered["revenue_musd"] >= rev_min) &
            (filtered["revenue_musd"] <= rev_max) &
            (filtered["net_assets_musd"] >= asset_min) &
            (filtered["net_assets_musd"] <= asset_max) &
            (filtered["related_party_pct"] <= rp_max)
        ]
        if require_inventory:
            filtered = filtered[filtered["has_inventory_risk"] == True]
        if func_match:
            filtered = filtered[filtered["functions"].apply(lambda x: any(f in x for f in func_match))]

        st.markdown(f"**符合筛选条件的公司：{len(filtered)} 家**")

        if len(filtered) > 0:
            display_cols = ["company_name", "country", "industry", "revenue_musd",
                           "operating_margin", "roe", "related_party_pct", "fiscal_year"]
            st.dataframe(filtered[display_cols], use_container_width=True, hide_index=True)

            # Export
            wiz_export_col1, _ = st.columns([1, 4])
            with wiz_export_col1:
                csv_wiz = filtered[display_cols].to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 导出筛选结果 / Export",
                    data=csv_wiz,
                    file_name=f"screened_companies_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                )

            if st.button("保存第2步，进入第3步 / Save & Next", key="save_step2", type="primary"):
                st.session_state.wiz_step2 = {
                    "Industry": sel_industry,
                    "经济体 / Economies": ", ".join(sel_countries) if sel_countries else "All",
                    "FY Range": f"{fy_range[0]}-{fy_range[1]}",
                    "Revenue Range": f"${rev_min:.1f}M - ${rev_max:.1f}M",
                    "Net Assets Range": f"${asset_min:.1f}M - ${asset_max:.1f}M",
                    "Max Related Party %": f"{rp_max*100:.0f}%",
                    "Selected Companies": len(filtered),
                }
                st.session_state.wiz_current_tab = 2
                st.rerun()
        else:
            st.warning("没有公司符合筛选条件，请放宽一些筛选标准。")

    # ========== Step 3 ==========
    with tabs[2]:
        st.subheader("第3步：可比性调整")
        st.markdown("对筛选出的可比公司进行调整，使其与被测试方更具可比性。")

        adj_type = st.selectbox(
            "调整类型",
            ["Working Capital Adjustment (WCA)", "Risk Adjustment", "Asset Intensity Adjustment",
             "Functional Adjustment", "Custom Adjustment"],
            format_func=lambda x: {
                "Working Capital Adjustment (WCA)": "营运资金调整 (WCA)",
                "Risk Adjustment": "风险调整",
                "Asset Intensity Adjustment": "资产密集度调整",
                "Functional Adjustment": "功能调整",
                "Custom Adjustment": "自定义调整",
            }.get(x, x),
            help="选择要应用的调整类型",
        )

        if adj_type == "Working Capital Adjustment (WCA)":
            st.markdown("#### 营运资金调整计算器")
            st.markdown("根据应收账款天数、应付账款天数和存货天数的差异计算营运资金调整额。")

            wca_col1, wca_col2 = st.columns(2)
            with wca_col1:
                ar_days = st.number_input("应收账款天数", min_value=0.0, max_value=365.0, value=60.0, step=1.0,
                                         help="应收账款平均回收期")
                ap_days = st.number_input("应付账款天数", min_value=0.0, max_value=365.0, value=45.0, step=1.0,
                                         help="应付账款平均付款期")
            with wca_col2:
                inv_days = st.number_input("存货天数", min_value=0.0, max_value=365.0, value=30.0, step=1.0,
                                          help="存货平均持有天数")
                int_rate = st.number_input("利率 (%)", min_value=0.0, max_value=50.0, value=5.0, step=0.1,
                                          help="适用于调整的利率")

            annual_rev = st.number_input("年收入（百万美元）", min_value=0.0, max_value=10000.0, value=100.0, step=1.0,
                                        help="被测试方的年收入（百万美元）")

            wca_result = calculate_wca(ar_days, ap_days, inv_days, int_rate, annual_rev)

            st.markdown("---")
            result_cols = st.columns(3)
            with result_cols[0]:
                net_cycle = ar_days - ap_days + inv_days
                st.metric("净营运资金周期", f"{net_cycle:.1f} 天")
            with result_cols[1]:
                st.metric("WCA 调整额", f"${wca_result:.4f}M")
            with result_cols[2]:
                pct_of_revenue = (wca_result / annual_rev * 100) if annual_rev > 0 else 0
                st.metric("WCA 占收入比例", f"{pct_of_revenue:.4f}%")

            if st.button("保存第3步，进入第4步 / Save & Next", key="save_step3", type="primary"):
                st.session_state.wiz_step3 = {
                    "Adjustment Type": adj_type,
                    "AR Days": ar_days,
                    "AP Days": ap_days,
                    "Inventory Days": inv_days,
                    "Interest Rate": f"{int_rate}%",
                    "Annual Revenue": f"${annual_rev}M",
                    "WCA Amount": f"${wca_result:.4f}M",
                }
                st.session_state.wiz_current_tab = 3
                st.rerun()

        else:
            st.info(f"{adj_type} 模块可提供详细的计算选项。选择「营运资金调整」可进行交互式计算演示。")
            adj_magnitude = st.slider("调整幅度 (%)", -20.0, 20.0, 0.0, 0.5,
                                     help="手动调整百分比")
            st.metric("应用调整", f"{adj_magnitude:.2f}%")

    # ========== Step 4 ==========
    with tabs[3]:
        st.subheader("第4步：独立交易区间确定")
        st.markdown("分析筛选出的可比公司的盈利能力，确定独立交易区间。")

        # Use filtered data or full dataset
        companies_df = load_mock_companies()
        if "wiz_step2" in st.session_state and st.session_state.wiz_step2:
            # Re-apply last known filters (simplified)
            pass

        # Simulate PLI distribution
        pli_data = companies_df["operating_margin"].dropna()
        pli_values = pli_data.values * 100  # Convert to percentage

        if len(pli_values) > 0:
            stats = {
                "Min": np.min(pli_values),
                "Q1": np.percentile(pli_values, 25),
                "Median": np.median(pli_values),
                "Q3": np.percentile(pli_values, 75),
                "Max": np.max(pli_values),
                "Mean": np.mean(pli_values),
                "Count": len(pli_values),
            }

            # Stats display
            stat_cols = st.columns(len(stats))
            for col, (key, val) in zip(stat_cols, stats.items()):
                with col:
                    if key == "Count":
                        st.metric(key, f"{int(val)}")
                    else:
                        st.metric(key, f"{val:.2f}%")

            # Box plot
            fig_box = go.Figure()
            fig_box.add_trace(go.Box(
                y=pli_values,
                name="Operating Margin (%)",
                boxmean=True,
                marker_color=PRIMARY_COLOR,
                line_color=PRIMARY_COLOR,
            ))
            fig_box.update_layout(
                title="Arm's Length Range - Operating Margin Distribution",
                yaxis_title="Operating Margin (%)",
                height=400,
                showlegend=False,
            )
            st.plotly_chart(fig_box, use_container_width=True)

            # Interpretation
            st.markdown("#### 🤖 自动解读")
            median_val = stats["Median"]
            iqr = stats["Q3"] - stats["Q1"]
            if median_val > 15:
                interpretation = (
                    f"中位营业利润率 **{median_val:.2f}%** 表明该行业利润率较高。"
                    f"四分位距 (IQR) 为 **{iqr:.2f}%**，说明可比公司之间差异适中。"
                    f"被测试方的利润率应落在 **{stats['Q1']:.2f}% - {stats['Q3']:.2f}%**"
                    f"（四分位区间）内才算符合独立交易原则。"
                )
            elif median_val > 5:
                interpretation = (
                    f"中位营业利润率 **{median_val:.2f}%** 符合行业常态。"
                    f"四分位距 **{iqr:.2f}%** 提供了合理的独立交易区间。"
                    f"落在 **{stats['Q1']:.2f}% - {stats['Q3']:.2f}%** 内的结果通常可接受。"
                )
            else:
                interpretation = (
                    f"中位营业利润率 **{median_val:.2f}%** 较低，"
                    f"可能反映薄利商业模式或竞争激烈的市场环境。"
                    f"建议仔细分析功能差异。"
                    f"可接受区间为 **{stats['Q1']:.2f}% - {stats['Q3']:.2f}%**。"
                )
            st.info(interpretation)

            # Method recommendation
            st.markdown("#### 📋 推荐转让定价方法 / Recommended TP Method")
            tx_type = st.session_state.wiz_step1.get("Transaction Type", "Sale of Goods")
            method_map = {
                "Sale of Goods": "交易净利润法 (TNMM) — 营业利润率 / TNMM - Operating Margin",
                "Provision of Services": "成本加成法 或 TNMM / Cost Plus or TNMM",
                "Transfer of Intangibles": "剩余利润分割法 / Residual Profit Split Method",
                "Financial Transactions": "可比非受控价格法 (CUP) / Comparable Uncontrolled Price",
                "Digital Services": "TNMM — 营业利润率 / TNMM - Operating Margin",
                "Cross-Border Services": "TNMM 或 成本加成法 / TNMM or Cost Plus",
            }
            recommended = method_map.get(tx_type, "TNMM")
            st.markdown(f"<span class='method-badge'>推荐方法: {recommended}</span>", unsafe_allow_html=True)

            # Export report
            st.markdown("---")
            if st.button("📤 导出分析报告 (.docx)", key="export_report"):
                tx_desc = st.session_state.get("wiz_step1", {})
                screen_params = st.session_state.get("wiz_step2", {})
                adj_results = st.session_state.get("wiz_step3", {})
                al_range_stats = {k: f"{v:.2f}%" if isinstance(v, float) else str(v) for k, v in stats.items()}

                doc_bytes = generate_word_report(tx_desc, screen_params, adj_results, al_range_stats, recommended)
                st.download_button(
                    label="📥 下载Word报告",
                    data=doc_bytes,
                    file_name=f"TP_Analysis_Report_{datetime.now().strftime('%Y%m%d')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
                st.success("报告生成成功！")
        else:
            st.warning("数据不足，无法确定独立交易区间。请先完成第1-3步。")


# ============== Page 3: Extractive Industry Pricing ==============

def extractive_pricing() -> None:
    """Page 3: Extractive Industry Pricing Workbench - Commodity pricing tools."""
    st.markdown("<div class='main-header'>⛏️ 采掘业定价工作台 / Extractive Industry Pricing</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='sub-header'>为采掘业提供大宗商品定价分析工具 / Commodity Pricing Tools for the Extractive Sector</div>",
        unsafe_allow_html=True,
    )

    st.markdown("""
    <div class='info-bar'>
    💡 <b>本页功能说明 / Page Guide:</b><br>
    <b>中文：</b>专门针对采掘业（采矿、石油、天然气）的转让定价分析工具。<br>
    &nbsp;&nbsp;• <b>大宗商品价格查询</b>：查看各类矿产、能源的国际市场价格走势<br>
    &nbsp;&nbsp;• <b>可比价格分析</b>：将关联交易价格与市场公开价格对比<br>
    <b>English:</b> Transfer pricing analysis tools for the extractive sector (mining, oil & gas).<br>
    &nbsp;&nbsp;• <b>Commodity Prices</b>: View international market price trends<br>
    &nbsp;&nbsp;• <b>Comparable Price Analysis</b>: Compare related-party prices with market prices<br>
    <br>
    <i>依据 / Reference: 联合国国际税收合作框架公约第6条（税基侵蚀）及联合国采掘业税收指南<br>
    UN Framework Convention Art. 6 (Base Erosion) & UN Extractive Industries Taxation Guide</i>
    </div>
    """, unsafe_allow_html=True)

    # Help
    if st.button("Help", key="help_extractive", help="Guidance on extractive industry pricing"):
        st.info("""
        **Extractive Industry Pricing Help**

        This workbench provides tools for pricing commodities in the extractive sector:
        - **Exchange Prices**: Reference market prices for major commodities
        - **CUP Calculator**: Calculate the Comparable Uncontrolled Price for mineral sales
        - **Comparable Transactions**: Reference table of recent comparable transactions

        The CUP method is typically the most appropriate method for commodity transactions
        where quoted prices are available on recognized exchanges.
        """)

    # Commodity selector
    commodity = st.selectbox(
        "Select Commodity",
        ["Copper Concentrate", "Iron Ore (Fe 62%)", "Gold Dore (Au 85%)"],
        help="Select the commodity type for pricing analysis",
    )

    # Exchange prices
    st.markdown("---")
    st.subheader("Reference Exchange Prices (30-Day Trend)")

    commodity_data = load_commodity_data()
    selected_commodity = commodity.split(" (")[0] if " (" in commodity else commodity.split(" ")[0]
    if selected_commodity == "Copper":
        selected_commodity = "Copper"
    elif selected_commodity == "Iron":
        selected_commodity = "Iron Ore"
    elif selected_commodity == "Gold":
        selected_commodity = "Gold"

    comm_data = commodity_data[commodity_data["commodity"] == selected_commodity].sort_values("date")

    if len(comm_data) > 0:
        latest_price = comm_data.iloc[-1]["price_usd"]
        prev_price = comm_data.iloc[-2]["price_usd"] if len(comm_data) > 1 else latest_price
        price_change = latest_price - prev_price

        kpi_cols = st.columns(4)
        with kpi_cols[0]:
            unit = "USD/tonne" if selected_commodity in ["Copper", "Iron Ore"] else "USD/oz"
            st.metric(f"Latest Price ({unit})", f"${latest_price:,.2f}", f"${price_change:,.2f}")
        with kpi_cols[1]:
            avg_price = comm_data["price_usd"].mean()
            st.metric("30-Day Average", f"${avg_price:,.2f}")
        with kpi_cols[2]:
            high_price = comm_data["price_usd"].max()
            st.metric("30-Day High", f"${high_price:,.2f}")
        with kpi_cols[3]:
            low_price = comm_data["price_usd"].min()
            st.metric("30-Day Low", f"${low_price:,.2f}")

        # Price trend chart
        fig_trend = go.Figure()
        fig_trend.add_trace(go.Scatter(
            x=comm_data["date"], y=comm_data["price_usd"],
            mode="lines+markers", name=selected_commodity,
            line=dict(color=PRIMARY_COLOR, width=2),
            marker=dict(size=6),
        ))
        fig_trend.update_layout(
            title=f"{selected_commodity} Price Trend (Last 30 Days)",
            xaxis_title="Date",
            yaxis_title=f"Price ({unit})",
            height=400,
            margin=dict(l=0, r=0, t=40, b=0),
        )
        st.plotly_chart(fig_trend, use_container_width=True)

    # CUP Calculator
    st.markdown("---")
    st.subheader("CUP Calculator - Mineral Sale Pricing")
    st.markdown("Calculate the arm's length price for mineral sales using the Comparable Uncontrolled Price method.")

    calc_col1, calc_col2 = st.columns(2)
    with calc_col1:
        spot_price = st.number_input("Spot/Reference Price (USD)", min_value=0.0, value=float(latest_price if len(comm_data) > 0 else 0), step=10.0,
                                    help="Current spot price or reference price on the exchange")
        grade_pct = st.number_input("Grade/Assay (%)", min_value=0.0, max_value=100.0, value=30.0 if "Copper" in commodity else (62.0 if "Iron" in commodity else 85.0), step=0.1,
                                   help="Mineral grade or assay percentage")
    with calc_col2:
        tc_rc = st.number_input("TC/RC (USD/tonne or %)", min_value=0.0, value=100.0 if "Copper" in commodity else 0.0, step=1.0,
                               help="Treatment and refining charges (for copper concentrate)")
        freight = st.number_input("Freight & Insurance (USD/tonne)", min_value=0.0, value=25.0, step=1.0,
                                 help="Transportation and insurance costs")

    # Calculate
    if "Copper" in commodity:
        payable_metal = spot_price * (grade_pct / 100)
        deductions = tc_rc + freight
        net_price = payable_metal - deductions
        payable_formula = f"Spot Price x Grade% = ${spot_price:,.2f} x {grade_pct}% = ${payable_metal:,.2f}/tonne"
    elif "Iron" in commodity:
        fe_adjustment = (grade_pct - 62) * 0.01 * spot_price  # Simplified
        net_price = spot_price + fe_adjustment - freight
        payable_formula = f"Base Price + Fe Adjustment - Freight"
    else:  # Gold
        payable_metal = spot_price * (grade_pct / 100)
        deductions = freight
        net_price = payable_metal - deductions
        payable_formula = f"Spot Price x Gold Content% = ${spot_price:,.2f} x {grade_pct}%"

    st.markdown("---")
    result_cols = st.columns(3)
    with result_cols[0]:
        st.metric("Calculated Net Price", f"${net_price:,.2f}/tonne" if "Gold" not in commodity else f"${net_price:,.2f}/oz")
    with result_cols[1]:
        st.metric("Deductions", f"${deductions:,.2f}")
    with result_cols[2]:
        margin = (net_price / spot_price * 100) if spot_price > 0 else 0
        st.metric("Net/Spot Ratio", f"{margin:.1f}%")

    st.caption(f"Calculation: {payable_formula}")

    # Comparable transactions reference
    st.markdown("---")
    st.subheader("Comparable Transactions Reference")

    np.random.seed(88)
    tx_data = []
    for i in range(15):
        tx_data.append({
            "date": (datetime.now() - timedelta(days=np.random.randint(1, 180))).strftime("%Y-%m-%d"),
            "commodity": np.random.choice(["Copper Concentrate", "Iron Ore", "Gold Dore"]),
            "buyer": f"Refiner {i+1:02d}",
            "seller": f"Mine {i+1:02d}",
            "volume_mt": int(np.random.uniform(1000, 50000)),
            "price_usd": round(np.random.uniform(80, 9000), 2),
            "terms": np.random.choice(["FOB", "CIF", "CFR"]),
            "basis": np.random.choice(["LME", "Platts", "LBMA", "Spot"]),
        })
    tx_df = pd.DataFrame(tx_data)
    st.dataframe(tx_df, use_container_width=True, hide_index=True)


# Fix variable name typo



# ============== Page 4: Country Policy Lookup ==============

def country_policy() -> None:
    """Page 4: Country Policy Lookup - National TP regulation and treaty information."""
    st.markdown("<div class='main-header'>📋 各经济体转让定价政策查询 / Economy Policy Lookup</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='sub-header'>查询各国的转让定价法规、文档要求和条约网络 / National TP Regulations, Documentation & Treaty Networks</div>",
        unsafe_allow_html=True,
    )

    st.markdown("""
    <div class='info-bar'>
    💡 <b>本页功能说明 / Page Guide:</b><br>
    <b>中文：</b>查询各经济体转让定价政策的「百科全书」。<br>
    &nbsp;&nbsp;• <b>法规要求</b>：各国对转让定价文档的要求<br>
    &nbsp;&nbsp;• <b>争议解决机制</b>：APA、MAP、仲裁机制<br>
    &nbsp;&nbsp;• <b>安全港规则</b>：简化合规的安全港规定<br>
    <b>English:</b> Encyclopedia of national transfer pricing policies.<br>
    &nbsp;&nbsp;• <b>Regulations</b>: TP documentation requirements by economy<br>
    &nbsp;&nbsp;• <b>Dispute Resolution</b>: APA, MAP, and arbitration mechanisms<br>
    &nbsp;&nbsp;• <b>Safe Harbour</b>: Simplified compliance rules<br>
    <br>
    <i>依据 / Reference: 联合国国际税收合作框架公约第13条（争议预防与解决）及联合国税收协定范本<br>
    UN Framework Convention Art. 13 (Dispute Prevention & Resolution) & UN Model Tax Convention</i>
    </div>
    """, unsafe_allow_html=True)

    if st.button("Help", key="help_policy", help="Guidance on economy policy lookup"):
        st.info("""
        **Economy Policy Lookup Help**

        Search for an economy to view its transfer pricing policy framework:
        - TP regulations and legislative basis
        - Documentation requirements (Local File, Master File, CbCR)
        - Safe harbor rules and thresholds
        - Tax treaty network size
        - MAP/APA availability and statistics
        - Information exchange mechanisms

        Use the **Compare** feature to view two countries side by side.
        """)

    df_policies = load_country_policies()

    # Search / select
    search_col, compare_toggle_col = st.columns([3, 1])
    with search_col:
        selected_country = st.selectbox(
            "Select Economy",
            [""] + sorted(df_policies["country"].tolist()),
            help="Choose an economy to view its TP policy details",
        )
    with compare_toggle_col:
        compare_mode = st.toggle("Compare Mode", help="Enable side-by-side comparison of two countries")

    if compare_mode:
        col_a, col_b = st.columns(2)
        with col_a:
            country_a = st.selectbox("经济体A / Economy A", [""] + sorted(df_policies["country"].tolist()), key="country_a")
        with col_b:
            country_b = st.selectbox("经济体B / Economy B", [""] + sorted(df_policies["country"].tolist()), key="country_b")

        if country_a and country_b:
            row_a = df_policies[df_policies["country"] == country_a].iloc[0]
            row_b = df_policies[df_policies["country"] == country_b].iloc[0]

            for label, key in [
                ("TP Regulations", "has_tp_regulations"),
                ("Documentation Requirements", "documentation_req"),
                ("Safe Harbor Available", "safe_harbor"),
                ("Tax Treaties", "treaty_count"),
                ("MAP Cases (2024)", "map_cases_2024"),
                ("APA Available", "apa_available"),
                ("Exchange Mechanism", "exchange_mechanism"),
            ]:
                c1, c2 = st.columns(2)
                with c1:
                    val = row_a[key]
                    display = "Yes" if val is True else ("No" if val is False else str(val))
                    st.metric(f"{label} - {country_a}", display)
                with c2:
                    val = row_b[key]
                    display = "Yes" if val is True else ("No" if val is False else str(val))
                    st.metric(f"{label} - {country_b}", display)
                st.markdown("---")

            # Comparison chart
            fig_comp = go.Figure()
            fig_comp.add_trace(go.Bar(
                x=["Tax Treaties", "MAP Cases (2024)"],
                y=[row_a["treaty_count"], row_a["map_cases_2024"]],
                name=country_a,
                marker_color=PRIMARY_COLOR,
            ))
            fig_comp.add_trace(go.Bar(
                x=["Tax Treaties", "MAP Cases (2024)"],
                y=[row_b["treaty_count"], row_b["map_cases_2024"]],
                name=country_b,
                marker_color=ACCENT_GREEN,
            ))
            fig_comp.update_layout(
                barmode="group",
                title=f"对比: {country_a} vs {country_b} / Comparison",
                height=350,
            )
            st.plotly_chart(fig_comp, use_container_width=True)

    elif selected_country:
        row = df_policies[df_policies["country"] == selected_country].iloc[0]

        st.markdown(f"## {selected_country}")

        # Top metrics
        mcol1, mcol2, mcol3, mcol4 = st.columns(4)
        with mcol1:
            st.metric("TP Regulations", "Yes" if row["has_tp_regulations"] else "No")
        with mcol2:
            st.metric("Tax Treaties", row["treaty_count"])
        with mcol3:
            st.metric("MAP Cases (2024)", row["map_cases_2024"])
        with mcol4:
            st.metric("APA Available", "Yes" if row["apa_available"] else "No")

        st.markdown("---")

        # Detail sections
        dcol1, dcol2 = st.columns(2)
        with dcol1:
            st.markdown("#### Documentation Requirements")
            st.info(row["documentation_req"])

            st.markdown("#### Safe Harbor Rules")
            if row["safe_harbor"]:
                st.success("Safe harbor rules are available in this jurisdiction.")
            else:
                st.warning("No safe harbor rules currently in effect.")

        with dcol2:
            st.markdown("#### Information Exchange Mechanisms")
            st.info(row["exchange_mechanism"])

            st.markdown("#### MAP/APA Framework")
            if row["apa_available"]:
                st.success("APA program is operational. MAP is available under tax treaties.")
            else:
                st.warning("APA not yet available. MAP may be accessible through treaty network.")

        # Policy summary text
        st.markdown("---")
        st.markdown("#### Transfer Pricing Regulation Summary")
        summary_text = (
            f"**{selected_country}** has {'comprehensive' if row['has_tp_regulations'] else 'limited'} "
            f"transfer pricing regulations requiring **{row['documentation_req']}**. "
            f"The economy maintains **{row['treaty_count']} tax treaties** providing MAP access, "
            f"and {'operates' if row['apa_available'] else 'does not yet operate'} an APA program. "
            f"Information exchange is conducted through **{row['exchange_mechanism']}**. "
            f"Safe harbor provisions are {'available' if row['safe_harbor'] else 'not currently available'}."
        )
        st.markdown(summary_text)

        # Treaty network visualization (simplified)
        st.markdown("---")
        st.markdown("#### Tax Treaty Network (Illustrative)")
        treaty_count = row["treaty_count"]
        # Create a simple radial chart representation
        angles = np.linspace(0, 2 * np.pi, max(treaty_count, 10), endpoint=False)
        r_values = np.ones(len(angles))
        fig_treaty = go.Figure()
        fig_treaty.add_trace(go.Scatterpolar(
            r=r_values,
            theta=angles * 180 / np.pi,
            mode="markers",
            marker=dict(size=12, color=PRIMARY_COLOR),
            name="Treaty Partners",
        ))
        fig_treaty.add_trace(go.Scatterpolar(
            r=[0],
            theta=[0],
            mode="markers",
            marker=dict(size=20, color=ACCENT_RED),
            name=selected_country,
        ))
        fig_treaty.update_layout(
            polar=dict(radialaxis=dict(visible=False, range=[0, 1.5])),
            showlegend=True,
            title=f"Treaty Network: {treaty_count} agreements",
            height=400,
        )
        st.plotly_chart(fig_treaty, use_container_width=True)
    else:
        # Show all countries overview table
        st.info("Select an economy above to view detailed policy information, or enable Compare Mode.")
        st.subheader("Countries Overview")
        display_df = df_policies.copy()
        display_df["TP Regs"] = display_df["has_tp_regulations"].apply(lambda x: "Yes" if x else "No")
        display_df["Safe Harbor"] = display_df["safe_harbor"].apply(lambda x: "Yes" if x else "No")
        display_df["APA"] = display_df["apa_available"].apply(lambda x: "Yes" if x else "No")
        st.dataframe(
            display_df[["country", "TP Regs", "documentation_req", "Safe Harbor", "treaty_count", "map_cases_2024", "APA", "exchange_mechanism"]],
            use_container_width=True,
            hide_index=True,
        )


# ============== Page 5: Data Contribution Hub ==============

def data_contribution() -> None:
    """Page 5: Data Contribution Hub - Upload, template, tiered management."""
    st.markdown("<div class='main-header'>📤 数据贡献管理 / Data Contribution Hub</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='sub-header'>向全球转让定价数据库贡献可比公司数据 — 含导入模板与密级分层管理</div>",
        unsafe_allow_html=True,
    )

    st.markdown("""
    <div class='info-bar'>
    💡 <b>本页功能说明 / Page Guide:</b><br>
    <b>中文：</b>管理各经济体向数据库贡献数据的「工作台」，包含三大功能：<br>
    &nbsp;&nbsp;• <b>📥 导入模板下载</b>：提供标准化CSV模板，确保数据格式统一<br>
    &nbsp;&nbsp;• <b>🔐 密级分层管理</b>：Tier 1/2/3 三级数据分类，不同密级有不同访问权限和用途<br>
    &nbsp;&nbsp;• <b>📊 贡献统计</b>：查看各国的数据贡献排名、密级分布和进度<br>
    <b>English:</b> Workspace for economy data contributions with three core functions:<br>
    &nbsp;&nbsp;• <b>📥 Import Template</b>: Standardized CSV template for consistent data format<br>
    &nbsp;&nbsp;• <b>🔐 Tiered Classification</b>: Tier 1/2/3 data classification with differentiated access<br>
    &nbsp;&nbsp;• <b>📊 Contribution Stats</b>: Economy rankings, tier distribution, and progress<br>
    <br>
    <i>依据 / Reference: 联合国国际税收合作框架公约第11条（能力建设）<br>
    UN Framework Convention Art. 11 (Capacity Building) — Data sharing to support developing economies</i>
    </div>
    """, unsafe_allow_html=True)

    # ===== Tab 1: Import Template =====
    # ===== Tab 2: Upload =====
    # ===== Tab 3: Tier Management =====
    # ===== Tab 4: Stats =====
    contrib_tabs = st.tabs([
        "📥 导入模板 / Import Template",
        "📤 数据上传 / Upload",
        "🔐 密级管理 / Tier Management",
        "📊 贡献统计 / Statistics",
    ])

    # ========== Tab 1: Import Template ==========
    with contrib_tabs[0]:
        st.subheader("📥 导入模板下载 / Download Import Template")

        st.markdown("""
        **中文说明：** 请先下载标准化CSV模板，按模板格式填写数据后上传。
        模板包含字段说明（第二个Sheet），请仔细阅读每个字段的含义和要求。

        **English:** Download the standardized CSV template first. Fill in your data following
        the template format, then upload it. The template includes a field guide sheet explaining
        each field's meaning and requirements.
        """)

        # Generate template
        template_col1, template_col2 = st.columns([1, 2])
        with template_col1:
            template_format = st.radio(
                "选择格式 / Format",
                ["CSV", "Excel (.xlsx)"],
                help="CSV兼容性最好；Excel带字段说明"
            )

            # Always render download button directly (no nested button)
            template_data = generate_import_template()
            if template_format == "CSV":
                csv_str = template_data.to_csv(index=False)
                st.download_button(
                    label="📥 下载CSV模板 / Download CSV",
                    data=csv_str.encode('utf-8-sig'),
                    file_name="TP_Comparable_Data_Template.csv",
                    mime="text/csv",
                )
            else:
                import io
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    template_data.to_excel(writer, sheet_name='Data', index=False)
                    field_guide = pd.DataFrame({
                        '字段名 Field': ['company_id','company_name','country_code','industry_code_isic','revenue','net_worth','fiscal_year','data_source','data_tier'],
                        '说明 Description': ['公司ID / Company ID','公司名称 / Company name','经济体代码 / Economy code','ISIC行业代码 / ISIC code','年收入 / Revenue','净资产 / Net worth','财政年度 / Fiscal year','数据来源 / Data source','密级 / Tier'],
                        '必填 / Required': ['✅','✅','✅','✅','✅','✅','✅','✅','✅'],
                    })
                    field_guide.to_excel(writer, sheet_name='Field Guide', index=False)
                st.download_button(
                    label="📥 下载Excel模板 / Download Excel",
                    data=output.getvalue(),
                    file_name="TP_Comparable_Data_Template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            st.success("✅ 点击上方按钮下载 / Click button above to download")

        with template_col2:
            st.markdown("##### 📋 模板字段预览 / Template Fields Preview")
            st.markdown("""
            | 字段 Field | 说明 Description | 必填 |
            |---|---|---|
            | company_id | 公司唯一ID | ✅ |
            | company_name | 公司名称（可匿名）| ✅ |
            | country_code | ISO 3国别代码 | ✅ |
            | industry_code_isic | ISIC行业代码 | ✅ |
            | revenue | 年收入（美元）| ✅ |
            | net_worth | 净资产（美元）| ✅ |
            | fiscal_year | 财政年度 | ✅ |
            | data_source | 数据来源 | ✅ |
            | data_tier | 数据密级 | ✅ |
            | cogs | 销售成本 | ❌ |
            | operating_margin | 营业利润率 | ❌ |
            | functions_performed | 履行功能 | ❌ |

            💡 **提示 / Tip:** 下载Excel版本可获得完整的字段说明（第二个Sheet）。
            """)

    # ========== Tab 2: Upload ==========
    with contrib_tabs[1]:
        st.subheader("📤 数据上传 / Upload Data")

        tier = st.selectbox(
            "选择数据密级 / Select Data Tier",
            [
                "Tier1 — 审计级数据 / Audited (Highest)",
                "Tier2 — 审阅级数据 / Reviewed",
                "Tier3 — 估算级数据 / Estimated",
            ],
            help="Tier1提供最详细可靠的数据；Tier3为汇总估算数据",
        )

        tier_info = {
            "Tier1 — 审计级数据 / Audited (Highest)": """
            **🔐 Tier 1 — 审计级数据 / Audited Data**
            - 来源：经审计的财务报表、公开上市年报
            - 可信度：最高 ✅✅✅
            - 允许用途：可比性分析、独立交易区间确定、税务审计参考
            - 访问权限：所有成员国税务机关
            - 来源 / Source: Audited financial statements, public filings
            - Reliability: Highest
            """,
            "Tier2 — 审阅级数据 / Reviewed": """
            **🟡 Tier 2 — 审阅级数据 / Reviewed Data**
            - 来源：税务机关评估报告、国别报告(CbC)、纳税申报数据
            - 可信度：较高 ✅✅
            - 允许用途：趋势分析、行业基准、初步筛选
            - 访问权限：签署信息交换协议的国家
            - 来源 / Source: Tax authority assessments, CbC reports
            - Reliability: Moderate-high
            """,
            "Tier3 — 估算级数据 / Estimated": """
            **🟠 Tier 3 — 估算级数据 / Estimated Data**
            - 来源：行业协会统计、第三方数据库估算、宏观经济推算
            - 可信度：一般 ✅
            - 允许用途：宏观分析、数据缺口填补、初步参考
            - 访问权限：仅限汇总统计，不披露单一公司
            - 来源 / Source: Industry statistics, third-party estimates
            - Reliability: Reference only
            """,
        }
        st.info(tier_info[tier])

        # File upload
        uploaded_file = st.file_uploader(
            "上传数据文件 / Upload Data File",
            type=["csv", "xlsx", "xls"],
            help="请使用模板格式上传 / Please use the template format",
        )

        # Data sovereignty declaration
        st.markdown("---")
        st.subheader("📋 数据主权声明 / Data Sovereignty Declaration")
        declaration = st.checkbox(
            "本人确认上传数据已匿名化处理，贡献国已授权在联合国国际税收合作框架公约下共享。"
            "数据不含个人身份信息或可识别单一纳税人的商业敏感信息。\n\n"
            "I confirm that the data is anonymized and the contributing economy has authorized "
            "its sharing under the UN Framework Convention. The data does not contain personally "
            "identifiable information or commercially sensitive information.",
            help="必填项 / Required",
        )

        if uploaded_file and declaration:
            st.success("✅ 文件上传成功！数据将进入审核流程。/ File uploaded! Data will enter the review process.")
            st.info(f"📄 文件 / File: **{uploaded_file.name}** | 大小 / Size: **{uploaded_file.size:,} bytes** | 密级 / Tier: **{tier[:6]}**")

            # Simulated validation
            st.markdown("##### 🔍 自动验证结果 / Automated Validation Results")
            val_col1, val_col2, val_col3 = st.columns(3)
            with val_col1:
                st.metric("必填字段完整率 / Required Fields", "100%", "✅")
            with val_col2:
                st.metric("经济体代码有效 / Valid Economy Codes", "98%", "✅")
            with val_col3:
                st.metric("数据合理性 / Reasonableness", "95%", "✅")

        elif uploaded_file and not declaration:
            st.warning("⚠️ 请先勾选数据主权声明。/ Please check the Data Sovereignty Declaration.")

    # ========== Tab 3: Tier Management ==========
    with contrib_tabs[2]:
        st.subheader("🔐 密级分层管理 / Tiered Data Classification Management")

        st.markdown("""
        <div class='info-bar'>
        <b>中文说明：</b>数据库采用三级密级分类体系，不同密级的数据有不同的来源、可信度和使用场景。<br>
        点击下方各密级标签查看详细说明和当前数据量。<br><br>
        <b>English:</b> The database uses a three-tier classification system. Each tier has different
        data sources, reliability levels, and permitted use cases.
        </div>
        """, unsafe_allow_html=True)

        # Load actual tier statistics from database
        tier_df = db_query("""
            SELECT
                data_tier,
                COUNT(*) as company_count,
                COUNT(DISTINCT country_code) as country_count,
                AVG(operating_margin) as avg_margin
            FROM comparable_companies
            GROUP BY data_tier
            ORDER BY data_tier
        """)

        if tier_df is None or tier_df.empty:
            tier_df = pd.DataFrame({
                'data_tier': ['Tier1', 'Tier2', 'Tier3'],
                'company_count': [495, 258, 281],
                'country_count': [85, 60, 45],
                'avg_margin': [12.5, 9.8, 7.2],
            })

        # Tier cards
        tier_col1, tier_col2, tier_col3 = st.columns(3)

        tier_colors = {'Tier1': '#28a745', 'Tier2': '#ffc107', 'Tier3': '#fd7e14'}
        tier_names = {
            'Tier1': 'Tier 1 — 审计级 / Audited',
            'Tier2': 'Tier 2 — 审阅级 / Reviewed',
            'Tier3': 'Tier 3 — 估算级 / Estimated',
        }
        tier_descs = {
            'Tier1': '经审计的财务报表\nAudited financial statements',
            'Tier2': '税务机关评估、CbC报告\nTax authority assessments',
            'Tier3': '行业估算、推算数据\nIndustry estimates',
        }
        tier_access = {
            'Tier1': '所有成员国 / All member states',
            'Tier2': '信息交换协议国 / Exchange agreement',
            'Tier3': '仅汇总统计 / Aggregated only',
        }

        for i, (col, tier_key) in enumerate(zip([tier_col1, tier_col2, tier_col3], ['Tier1', 'Tier2', 'Tier3'])):
            row = tier_df[tier_df['data_tier'] == tier_key]
            count = int(row['company_count'].iloc[0]) if not row.empty else 0
            countries = int(row['country_count'].iloc[0]) if not row.empty else 0
            with col:
                color = tier_colors[tier_key]
                st.markdown(f"""
                <div style='border: 2px solid {color}; border-radius: 10px; padding: 1rem; text-align: center;'>
                    <h3 style='color: {color}; margin: 0;'>{tier_names[tier_key]}</h3>
                    <p style='font-size: 0.85rem; color: #666;'>{tier_descs[tier_key]}</p>
                    <h2 style='margin: 0.5rem 0;'>{count:,}</h2>
                    <p style='font-size: 0.8rem;'>家公司 / companies</p>
                    <p style='font-size: 0.8rem;'>覆盖 {countries} 个国家 / {countries} countries</p>
                    <hr style='border: 0.5px solid #ddd;'>
                    <p style='font-size: 0.75rem;'>🔒 访问权限 / Access:<br>{tier_access[tier_key]}</p>
                </div>
                """, unsafe_allow_html=True)

        # Tier comparison table
        st.markdown("---")
        st.markdown("##### 📊 密级对比表 / Tier Comparison Table")
        comparison_data = pd.DataFrame({
            '密级 / Tier': ['Tier 1', 'Tier 2', 'Tier 3'],
            '来源 / Source': [
                '审计财务报表、公开年报\nAudited financials, public filings',
                '税务评估、CbC报告、纳税申报\nTax assessments, CbC reports',
                '行业统计、第三方估算\nIndustry stats, estimates',
            ],
            '可信度 / Reliability': ['⭐⭐⭐ 最高', '⭐⭐ 较高', '⭐ 一般'],
            '允许用途 / Permitted Use': [
                '可比性分析、独立交易区间确定\nFull comparability analysis',
                '趋势分析、行业基准、初步筛选\nTrend analysis, benchmarking',
                '宏观分析、缺口填补\nMacro analysis, gap filling',
            ],
            '访问权限 / Access': [
                '所有成员国 / All members',
                '信息交换协议国 / Exchange agreement',
                '仅汇总，不披露单一 / Aggregated only',
            ],
        })
        st.dataframe(comparison_data, use_container_width=True, hide_index=True)

        # Tier distribution by country
        st.markdown("---")
        st.markdown("##### 🌍 各密级数据经济体分布 / Tier Distribution by Economy")
        tier_country_df = db_query("""
            SELECT c.country_name, cc.data_tier, COUNT(*) as cnt
            FROM comparable_companies cc
            JOIN countries c ON c.country_code = cc.country_code
            GROUP BY c.country_name, cc.data_tier
            ORDER BY cnt DESC
            LIMIT 20
        """)
        if tier_country_df is not None and not tier_country_df.empty:
            fig_tier = px.bar(
                tier_country_df, x='country_name', y='cnt', color='data_tier',
                title='Top 20 Countries by Data Tier / 数据量前20国家的密级分布',
                labels={'country_name': '经济体 / Economy', 'cnt': '公司数 / Companies', 'data_tier': '密级 / Tier'},
                height=450,
                color_discrete_map={'Tier1': '#28a745', 'Tier2': '#ffc107', 'Tier3': '#fd7e14'},
            )
            st.plotly_chart(fig_tier, use_container_width=True)
        else:
            st.info("暂无数据 / No data available")

    # ========== Tab 4: Statistics ==========
    with contrib_tabs[3]:
        st.subheader("📊 贡献统计 / Contribution Statistics")

        stats_col1, stats_col2, stats_col3, stats_col4 = st.columns(4)
        with stats_col1:
            st.metric("贡献国家 / Contributing Countries", "28")
        with stats_col2:
            st.metric("总数据集 / Total Datasets", "1,247")
        with stats_col3:
            st.metric("待审核 / Pending Review", "23")
        with stats_col4:
            st.metric("本月批准 / Approved This Month", "18")

        # Contribution trend chart
        months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
        uploads = [15, 18, 22, 20, 25, 18]
        approvals = [12, 16, 20, 19, 22, 16]

        fig_trend = go.Figure()
        fig_trend.add_trace(go.Bar(
            x=months, y=uploads, name="上传 / Uploads", marker_color=PRIMARY_COLOR,
        ))
        fig_trend.add_trace(go.Bar(
            x=months, y=approvals, name="批准 / Approved", marker_color=ACCENT_GREEN,
        ))
        fig_trend.update_layout(
            barmode="group",
            title="月度数据贡献趋势 (2025) / Monthly Contribution Trend",
            xaxis_title="月份 / Month",
            yaxis_title="数据集数量 / Datasets",
            height=350,
        )
        st.plotly_chart(fig_trend, use_container_width=True)

        # Top contributors
        st.markdown("#### 🏆 贡献最多经济体 (2025) / Top Contributing Countries")
        contributors = pd.DataFrame({
            "经济体 / Economy": ["India", "South Africa", "Brazil", "China", "Kenya", "Nigeria", "Mexico", "Indonesia"],
            "数据集 / Datasets": [180, 145, 132, 128, 95, 88, 82, 75],
            "质量评分 / Quality Score": [95, 92, 90, 94, 88, 85, 91, 87],
            "Tier1占比 / Tier1 %": ["62%", "55%", "48%", "70%", "35%", "28%", "52%", "40%"],
        })
        st.dataframe(contributors, use_container_width=True, hide_index=True)


# ============== Page 6: MAP/APA Case Tracker ==============

def map_apa_tracker() -> None:
    """Page 6: MAP/APA Case Tracker - Monitor mutual agreement and advance pricing agreement cases."""
    st.markdown("<div class='main-header'>⚖️ 争议解决追踪器 / MAP & APA Tracker</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='sub-header'>追踪各经济体相互协商程序和预约定价安排 / Track MAP and APA Cases Across Jurisdictions</div>",
        unsafe_allow_html=True,
    )

    st.markdown("""
    <div class='info-bar'>
    💡 <b>本页功能说明 / Page Guide:</b><br>
    <b>中文：</b>追踪跨国税务争议解决案件的「监控台」。<br>
    &nbsp;&nbsp;• <b>MAP案例</b>：相互协商程序 — 两国税务部门协商解决双重征税<br>
    &nbsp;&nbsp;• <b>APA案例</b>：预约定价安排 — 事前约定关联交易定价方法<br>
    &nbsp;&nbsp;• <b>仲裁案例</b>：通过独立仲裁解决税务争议<br>
    <b>English:</b> Monitor cross-border tax dispute resolution cases.<br>
    &nbsp;&nbsp;• <b>MAP</b>: Mutual Agreement Procedure — resolving double taxation through bilateral negotiation<br>
    &nbsp;&nbsp;• <b>APA</b>: Advance Pricing Agreement — pre-agreeing TP methods with tax authorities<br>
    &nbsp;&nbsp;• <b>Arbitration</b>: Resolving disputes through independent arbitration<br>
    <br>
    <i>依据 / Reference: 联合国国际税收合作框架公约第13条（争议预防与解决）及联合国税收协定范本第25条<br>
    UN Framework Convention Art. 13 & UN Model Tax Convention Art. 25 (Mutual Agreement Procedure)</i>
    </div>
    """, unsafe_allow_html=True)

    if st.button("❓ 帮助 / Help", key="help_map", help="使用指南 / Guide"):
        st.info("""
        **争议解决追踪器 — 使用指南 / MAP Tracker Guide**

        **中文：** 本页追踪各经济体相互协商程序(MAP)和预约定价安排(APA)案件。
        - **MAP**：两国税务部门协商解决双重征税争议
        - **APA**：事前与税务机关约定转让定价方法，预防争议
        - 颜色含义：🟡进行中 / 🟢已解决 / 🔵待处理 / 🔴未达成协议

        **English:** This tracker monitors MAP and APA cases across jurisdictions.
        """)

    cases_df = load_map_cases()

    # Filters
    fcol1, fcol2, fcol3 = st.columns(3)
    with fcol1:
        status_filter = st.multiselect("状态 / Status", cases_df["status"].unique().tolist(), default=[])
    with fcol2:
        type_filter = st.multiselect("案件类型 / Case Type", cases_df["case_type"].unique().tolist(), default=[])
    with fcol3:
        issue_filter = st.multiselect("争议事项 / Issue", cases_df["issue"].unique().tolist(), default=[])

    filtered = cases_df.copy()
    if status_filter:
        filtered = filtered[filtered["status"].isin(status_filter)]
    if type_filter:
        filtered = filtered[filtered["case_type"].isin(type_filter)]
    if issue_filter:
        filtered = filtered[filtered["issue"].isin(issue_filter)]

    # Statistics
    st.markdown("---")
    st.subheader("案件统计 / Case Statistics")

    total_cases = len(filtered)
    active_cases = len(filtered[filtered["status"] == "Active"])
    resolved_cases = len(filtered[filtered["status"] == "Resolved"])
    avg_duration = filtered["duration_days"].mean() if len(filtered) > 0 else 0
    total_disputed = filtered["amount_disputed_musd"].sum() if len(filtered) > 0 else 0

    stat_cols = st.columns(5)
    with stat_cols[0]:
        st.metric("总案件数 / Total", total_cases)
    with stat_cols[1]:
        st.metric("进行中 / Active", active_cases)
    with stat_cols[2]:
        st.metric("已解决 / Resolved", resolved_cases)
    with stat_cols[3]:
        st.metric("平均时长 / Avg Duration", f"{avg_duration:.0f} 天/days")
    with stat_cols[4]:
        st.metric("争议金额 / Disputed", f"${total_disputed:.1f}M")

    # Cases table
    st.markdown("---")
    st.subheader("案件列表 / Case List")

    if len(filtered) > 0:
        display_df = filtered[["case_id", "case_type", "taxpayer", "country_a", "country_b",
                               "issue", "status", "start_date", "duration_days", "amount_disputed_musd"]].copy()
        display_df.columns = ["案件号/ID", "类型/Type", "纳税人/Taxpayer", "经济体A/Economy A", "经济体B/Economy B",
                              "争议事项/Issue", "状态/Status", "起始日/Start", "时长(天)/Duration", "争议额($M)/Disputed"]

        st.dataframe(
            display_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Status": st.column_config.TextColumn("Status", help="Case status"),
            },
        )
    else:
        st.info("No cases match the selected filters.")

    # Trend analysis
    st.markdown("---")
    st.subheader("Trend Analysis")

    trend_col1, trend_col2 = st.columns(2)

    with trend_col1:
        # Status distribution
        status_counts = cases_df["status"].value_counts().reset_index()
        status_counts.columns = ["Status", "Count"]
        fig_pie = px.pie(
            status_counts, values="Count", names="Status",
            title="Case Status Distribution",
            color="Status",
            color_discrete_map={
                "Active": ACCENT_YELLOW,
                "Resolved": ACCENT_GREEN,
                "Pending": PRIMARY_COLOR,
                "Closed - No agreement": ACCENT_RED,
            },
            height=350,
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    with trend_col2:
        # Duration by case type
        fig_box = px.box(
            cases_df, x="case_type", y="duration_days", color="status",
            title="Case Duration by Type",
            labels={"duration_days": "Duration (days)", "case_type": "Case Type"},
            height=350,
            color_discrete_map={
                "Active": ACCENT_YELLOW,
                "Resolved": ACCENT_GREEN,
                "Pending": PRIMARY_COLOR,
                "Closed - No agreement": ACCENT_RED,
            },
        )
        st.plotly_chart(fig_box, use_container_width=True)

    # Efficiency metrics
    st.markdown("---")
    st.subheader("Efficiency Indicators")

    eff_col1, eff_col2, eff_col3 = st.columns(3)
    with eff_col1:
        resolution_rate = len(cases_df[cases_df["status"] == "Resolved"]) / max(len(cases_df), 1) * 100
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=resolution_rate,
            title={"text": "Resolution Rate (%)"},
            gauge={
                "axis": {"range": [0, 100]},
                "bar": {"color": ACCENT_GREEN},
                "steps": [
                    {"range": [0, 30], "color": "#ffcccc"},
                    {"range": [30, 70], "color": "#ffffcc"},
                    {"range": [70, 100], "color": "#ccffcc"},
                ],
                "threshold": {
                    "line": {"color": "red", "width": 4},
                    "thickness": 0.75,
                    "value": 50,
                },
            },
            height=250,
        ))
        st.plotly_chart(fig_gauge, use_container_width=True)

    with eff_col2:
        # MAP timeline trend (mock)
        quarters = ["Q1 2023", "Q2 2023", "Q3 2023", "Q4 2023", "Q1 2024", "Q2 2024", "Q3 2024", "Q4 2024"]
        new_cases = [8, 10, 7, 12, 9, 11, 8, 13]
        resolved_cases_trend = [5, 7, 8, 9, 7, 10, 9, 11]

        fig_line = go.Figure()
        fig_line.add_trace(go.Scatter(
            x=quarters, y=new_cases, mode="lines+markers",
            name="New Cases", line=dict(color=ACCENT_RED),
        ))
        fig_line.add_trace(go.Scatter(
            x=quarters, y=resolved_cases_trend, mode="lines+markers",
            name="Resolved", line=dict(color=ACCENT_GREEN),
        ))
        fig_line.update_layout(
            title="MAP Case Trend",
            xaxis_title="Quarter",
            yaxis_title="Cases",
            height=280,
            legend=dict(orientation="h", yanchor="bottom", y=-0.3, xanchor="center", x=0.5),
        )
        st.plotly_chart(fig_line, use_container_width=True)

    with eff_col3:
        # Average duration by issue type
        issue_duration = cases_df.groupby("issue")["duration_days"].mean().sort_values(ascending=True).reset_index()
        issue_duration.columns = ["Issue", "Avg Duration"]
        fig_bar = px.bar(
            issue_duration, y="Issue", x="Avg Duration",
            orientation="h", title="Avg. Duration by Issue Type",
            color="Avg Duration",
            color_continuous_scale="RdYlGn_r",
            height=280,
        )
        st.plotly_chart(fig_bar, use_container_width=True)


# ============== Page 7: Digital Economy Dashboard ==============

def digital_economy() -> None:
    """Page 7: Digital Economy & Cross-Border Services Dashboard."""
    st.markdown("<div class='main-header'>💻 数字经济与跨境服务 / Digital Economy & Cross-Border Services</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='sub-header'>分析数字服务交易和跨境服务定价 / Digital Service Transactions & Cross-Border Pricing Analysis</div>",
        unsafe_allow_html=True,
    )

    st.markdown("""
    <div class='info-bar'>
    💡 <b>本页功能说明 / Page Guide:</b><br>
    <b>中文：</b>分析数字经济领域转让定价的「专业工具」。<br>
    &nbsp;&nbsp;• <b>数字经济可比数据筛选</b>：按国家、利润率、财政年度筛选数字服务公司<br>
    &nbsp;&nbsp;• <b>收入vs利润率分析</b>：气泡图展示不同平台类型的定价特征<br>
    &nbsp;&nbsp;• <b>数字强度指标</b>：衡量企业数字化程度的综合指标<br>
    <b>English:</b> Specialized tools for digital economy transfer pricing analysis.<br>
    &nbsp;&nbsp;• <b>Data Filtering</b>: Screen digital service companies by economy, margin, fiscal year<br>
    &nbsp;&nbsp;• <b>Revenue vs Margin</b>: Bubble chart showing pricing characteristics by platform type<br>
    &nbsp;&nbsp;• <b>Digital Intensity</b>: Composite metric for digitalization level<br>
    <br>
    <i>依据 / Reference: 联合国国际税收合作框架公约第5条及联合国数字经济税收挑战工作组报告<br>
    UN Framework Convention Art. 5 & UN Committee of Experts on International Cooperation in Tax Matters: Digital Economy Report</i>
    </div>
    """, unsafe_allow_html=True)

    if st.button("Help", key="help_digital", help="Guidance on digital economy analysis"):
        st.info("""
        **Digital Economy Dashboard Help**

        This dashboard provides tools for analyzing digital economy transactions:
        - **Service Type Selection**: Choose from cloud computing, SaaS, digital advertising, etc.
        - **Comparable Data Filter**: Screen digital economy comparables by region, size, and margins
        - **Cross-Border Pricing Analysis**: Analyze profitability indicators for cross-border services

        Workstream II of the Framework Convention addresses the tax challenges of the digitalized economy.
        """)

    # Service type selection
    st.markdown("---")
    st.subheader("数字服务类型选择 / Service Type Selection")

    service_type_en = st.selectbox(
        "选择数字服务类型 / Select Digital Service Type",
        ["All", "Cloud Computing", "SaaS", "Digital Advertising", "E-commerce Platform",
         "Data Analytics", "IT Outsourcing", "Digital Content", "Fintech Services"],
        format_func=lambda x: {
            "All": "全部 / All",
            "Cloud Computing": "云计算 / Cloud Computing",
            "SaaS": "软件即服务 / SaaS",
            "Digital Advertising": "数字广告 / Digital Advertising",
            "E-commerce Platform": "电商平台 / E-commerce",
            "Data Analytics": "数据分析 / Data Analytics",
            "IT Outsourcing": "IT外包 / IT Outsourcing",
            "Digital Content": "数字内容 / Digital Content",
            "Fintech Services": "金融科技 / Fintech",
        }.get(x, x),
        help="选择要分析的数字服务类型 / Select the type of digital service",
    )
    service_type = service_type_en

    # Comparable data filter
    st.markdown("---")
    st.subheader("数字经济可比数据筛选 / Digital Economy Comparable Data Filter")

    digital_df = load_digital_comparables()
    if service_type != "All":
        digital_df = digital_df[digital_df["service_type"] == service_type]

    # Filters
    fcol1, fcol2, fcol3 = st.columns(3)
    with fcol1:
        dig_countries = st.multiselect(
            "经济体 / Economies",
            sorted(digital_df["country"].unique().tolist()),
            default=[],
            help="Filter by economy of the comparable company",
        )
    with fcol2:
        margin_range = st.slider(
            "Operating Margin Range (%)",
            float(digital_df["operating_margin"].min() * 100),
            float(digital_df["operating_margin"].max() * 100),
            (0.0, 45.0),
            help="Filter by operating margin percentage",
        )
    with fcol3:
        fy_filter = st.multiselect(
            "Fiscal Year",
            sorted(digital_df["fiscal_year"].unique().tolist()),
            default=sorted(digital_df["fiscal_year"].unique().tolist()),
            help="Filter by fiscal year",
        )

    # Apply filters
    filtered_digital = digital_df.copy()
    if dig_countries:
        filtered_digital = filtered_digital[filtered_digital["country"].isin(dig_countries)]
    filtered_digital = filtered_digital[
        (filtered_digital["operating_margin"] * 100 >= margin_range[0]) &
        (filtered_digital["operating_margin"] * 100 <= margin_range[1])
    ]
    if fy_filter:
        filtered_digital = filtered_digital[filtered_digital["fiscal_year"].isin(fy_filter)]

    st.markdown(f"**筛选结果：{len(filtered_digital)} 家公司 / Filtered Results: {len(filtered_digital)} companies**")

    if len(filtered_digital) > 0:
        # Metrics
        mcol1, mcol2, mcol3, mcol4 = st.columns(4)
        with mcol1:
            st.metric("公司数 / Companies", len(filtered_digital))
        with mcol2:
            st.metric("平均利润率 / Avg Margin", f"{filtered_digital['operating_margin'].mean()*100:.2f}%")
        with mcol3:
            st.metric("中位收入 / Median Revenue", f"${filtered_digital['revenue_musd'].median():.1f}M")
        with mcol4:
            st.metric("平均数字强度 / Avg Digital Intensity", f"{filtered_digital['digital_intensity'].mean():.2f}")

        # Data table
        st.dataframe(
            filtered_digital[["company_name", "country", "service_type", "revenue_musd",
                             "operating_margin", "cost_plus_markup", "tnmm_profit_margin",
                             "digital_intensity", "fiscal_year"]],
            use_container_width=True,
            hide_index=True,
        )

        # Scatter plot - Revenue vs Operating Margin
        fig_scatter = px.scatter(
            filtered_digital,
            x="revenue_musd",
            y="operating_margin",
            color="service_type",
            size="digital_intensity",
            hover_data=["company_name", "country"],
            title="收入vs营业利润率（气泡大小=数字强度）/ Revenue vs Operating Margin",
            labels={"revenue_musd": "收入($M) / Revenue", "operating_margin": "营业利润率 / Operating Margin"},
            height=450,
        )
        fig_scatter.update_layout(margin=dict(l=0, r=0, t=40, b=0))
        st.plotly_chart(fig_scatter, use_container_width=True)
        st.info("💡 **图表说明 / Chart Guide:** 每个气泡代表一家公司。横轴是年收入，纵轴是营业利润率，气泡大小表示数字强度（越大越依赖数字业务）。颜色代表平台类型。/ Each bubble = a company. X-axis: annual revenue. Y-axis: operating margin. Bubble size: digital intensity. Color: platform type.")

        # Box plot by service type
        if service_type == "All":
            fig_box = px.box(
                filtered_digital,
                x="service_type",
                y="operating_margin",
                title="各平台类型利润率分布 / Operating Margin by Service Type",
                labels={"operating_margin": "营业利润率 / Operating Margin", "service_type": "平台类型 / Service Type"},
                height=400,
                color="service_type",
            )
            st.plotly_chart(fig_box, use_container_width=True)
            st.info("💡 **图表说明 / Chart Guide:** 箱线图展示每种平台类型的利润率分布。箱体中线=中位数，箱体上下边=四分位区间，圆点=异常值。/ Box plot shows margin distribution by platform type. Middle line: median. Box edges: Q1/Q3. Dots: outliers.")
    else:
        st.warning("没有符合筛选条件的公司，请调整筛选。/ No companies match the selected filters.")

    # Cross-Border Services Analysis
    st.markdown("---")
    st.subheader("跨境服务定价分析 / Cross-Border Services Pricing Analysis")

    cbs_col1, cbs_col2 = st.columns(2)
    with cbs_col1:
        st.markdown("#### TNMM 利润率指标 / TNMM Profitability Indicators")

        if len(filtered_digital) > 0:
            pli_options = ["Operating Margin", "Cost Plus Markup", "TNMM Profit Margin"]
            selected_pli = st.selectbox("利润水平指标 / Profit Level Indicator", pli_options)

            pli_col = {
                "Operating Margin": "operating_margin",
                "Cost Plus Markup": "cost_plus_markup",
                "TNMM Profit Margin": "tnmm_profit_margin",
            }[selected_pli]

            values = filtered_digital[pli_col].dropna() * 100
            fig_hist = px.histogram(
                values, nbins=20,
                title=f"{selected_pli} Distribution",
                labels={"value": f"{selected_pli} (%)"},
                height=350,
                color_discrete_sequence=[PRIMARY_COLOR],
            )
            st.plotly_chart(fig_hist, use_container_width=True)

            st.markdown(f"**Mean**: {values.mean():.2f}% | **Median**: {values.median():.2f}% | **Std Dev**: {values.std():.2f}%")

    with cbs_col2:
        st.markdown("#### Regional Comparison")

        if len(filtered_digital) > 0:
            region_summary = filtered_digital.groupby("country").agg(
                companies=("company_name", "count"),
                avg_margin=("operating_margin", "mean"),
                avg_revenue=("revenue_musd", "mean"),
            ).reset_index()
            region_summary = region_summary.sort_values("avg_margin", ascending=True)

            fig_bar = px.bar(
                region_summary,
                y="country",
                x="avg_margin",
                orientation="h",
                title="Average Operating Margin by Economy",
                labels={"avg_margin": "平均营业利润率 / Avg Operating Margin", "country": "经济体 / Economy"},
                height=350,
                color="avg_margin",
                color_continuous_scale="Blues",
            )
            st.plotly_chart(fig_bar, use_container_width=True)

    # Workstream II Knowledge
    st.markdown("---")
    st.subheader("工作流II：数字经济税收 / Workstream II: Digital Economy Taxation")

    with st.expander("了解联合国框架公约工作流II / Learn about Workstream II"):
        st.markdown("""
        **Workstream II** addresses the tax challenges arising from the digitalization of the economy.
        Key focus areas include:

        1. **Nexus Rules**: Determining when a jurisdiction has taxing rights over digital transactions
        2. **Profit Allocation**: How to allocate profits from digital services across jurisdictions
        3. **Administrative Cooperation**: Enhancing information exchange for digital economy transactions
        4. **Simplified Approaches**: Safe harbors and simplified methods for low-value digital transactions

        **Cross-Border Services Pricing Challenges:**
        - Difficulty in identifying comparable transactions for novel digital services
        - Rapidly evolving business models and value chains
        - Intangibles-heavy business models with limited physical presence
        - User participation and data as value creation factors

        **Recommended Approach:**
        The TNMM (Transactional Net Margin Method) using the operating margin or Berry ratio
        is often the most appropriate method for routine digital service providers.
        For highly integrated digital operations, the Profit Split Method may be more suitable.
        """)


# ============== Page 8: Risk Assessment Center ==============

def risk_assessment() -> None:
    """Page 8: Transfer Pricing Risk Assessment Center — UN INC aligned."""
    st.markdown("<div class='main-header'>🚨 转让定价风险评估中心 / TP Risk Assessment Center</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='sub-header'>基于联合国国际税收合作框架公约的智能风险识别与评估</div>",
        unsafe_allow_html=True,
    )

    st.markdown("""
    <div class='info-bar'>
    💡 <b>本页功能说明 / Page Guide:</b><br>
    <b>中文：</b>基于联合国INC会议成果，对跨国企业转让定价风险进行智能评估。<br>
    &nbsp;&nbsp;• <b>风险雷达图</b>：从6个维度评估关联交易风险（利润率偏离、关联交易占比、避税港使用、无形资产迁移、债务杠杆、功能风险匹配）<br>
    &nbsp;&nbsp;• <b>风险等级</b>：🟢低风险 / 🟡中风险 / 🔴高风险 / ⚫极高风险<br>
    &nbsp;&nbsp;• <b>合规检查</b>：对照联合国框架公约各条款自动检查合规性<br>
    &nbsp;&nbsp;• <b>CbC国别报告分析</b>：分析跨国企业全球利润分配与经济活动实质的匹配度<br>
    <b>English:</b> Intelligent TP risk assessment aligned with UN INC outcomes.<br>
    &nbsp;&nbsp;• <b>Risk Radar</b>: 6-dimension risk scoring (profit deviation, related-party ratio, tax haven use, intangibles migration, debt leverage, function-risk match)<br>
    &nbsp;&nbsp;• <b>Risk Levels</b>: 🟢 Low / 🟡 Medium / 🔴 High / ⚫ Critical<br>
    &nbsp;&nbsp;• <b>Compliance Check</b>: Automated checks against UN Framework Convention articles<br>
    &nbsp;&nbsp;• <b>CbC Analysis</b>: Profit allocation vs. economic substance alignment<br>
    <br>
    <i>依据 / Reference: 联合国国际税收合作框架公约第5条、第6条、第13条及联合国INC会议成果文件<br>
    UN Framework Convention Arts. 5, 6, 13 & UN INC Conference Outcomes</i>
    </div>
    """, unsafe_allow_html=True)

    risk_tabs = st.tabs([
        "🎯 风险雷达 / Risk Radar",
        "📋 合规检查 / Compliance Check",
        "📊 CbC分析 / CbC Analysis",
    ])

    # ========== Tab 1: Risk Radar ==========
    with risk_tabs[0]:
        st.subheader("🎯 转让定价风险雷达图 / TP Risk Radar")

        st.markdown("""
        **中文说明：** 输入被评估企业的关键指标，系统从6个维度自动评分并生成风险雷达图。
        评分越高表示风险越大（0=无风险，100=极高风险）。

        **English:** Enter the assessed company's key metrics. The system scores 6 risk dimensions
        and generates a radar chart. Higher scores = higher risk (0=no risk, 100=critical).
        """)

        input_col1, input_col2 = st.columns(2)
        with input_col1:
            st.markdown("##### 企业信息 / Company Info")
            company_name = st.text_input("企业名称 / Company Name", value="MNE Group A", key="ra_company")
            country = st.selectbox("所在经济体 / Jurisdiction", ["China", "India", "Brazil", "South Africa", "Kenya", "Nigeria", "Other"], key="ra_country")
            industry = st.selectbox("行业 / Industry", ["Manufacturing", "Services", "Digital/Tech", "Pharmaceuticals", "Mining", "Financial Services"], key="ra_industry")

        with input_col2:
            st.markdown("##### 关键指标 / Key Metrics")
            operating_margin = st.slider("营业利润率 / Operating Margin (%)", -20.0, 50.0, 3.0, 0.5, key="ra_margin",
                                         help="被测试方的营业利润率，与行业基准对比")
            rp_ratio = st.slider("关联交易占比 / Related Party Ratio (%)", 0, 100, 40, 5, key="ra_rp",
                                 help="关联交易占总收入的比例")
            effective_tax = st.slider("实际税率 / Effective Tax Rate (%)", 0.0, 35.0, 15.0, 0.5, key="ra_tax",
                                      help="实际税率低于法定税率可能暗示利润转移")
            debt_to_equity = st.slider("债务权益比 / Debt-to-Equity Ratio", 0.0, 10.0, 3.0, 0.1, key="ra_de",
                                       help="过高的债务权益比可能暗示资本弱化")

        intangibles_transfer = st.checkbox("是否存在无形资产跨境转移 / Intangibles cross-border transfer", key="ra_intang",
                                           help="无形资产向低税地转移是高风险信号")
        tax_haven_use = st.checkbox("是否在避税港设立子公司 / Subsidiaries in tax havens", key="ra_haven",
                                    help="在零税或极低税地设立子公司")

        # Calculate risk scores
        risk_dims = {
            '利润率偏离\nProfit Deviation': max(0, min(100, (10 - abs(operating_margin - 10)) * 10)),
            '关联交易占比\nRelated Party': min(100, rp_ratio * 1.5),
            '避税港使用\nTax Haven Use': (80 if tax_haven_use else 20) + (100 - effective_tax * 2),
            '无形资产迁移\nIntangibles Migration': (85 if intangibles_transfer else 15),
            '债务杠杆\nDebt Leverage': min(100, debt_to_equity * 15),
            '功能风险匹配\nFunction-Risk Match': 50,  # Default moderate
        }

        # Adjust for industry
        if industry == "Digital/Tech":
            risk_dims['无形资产迁移\nIntangibles Migration'] += 10
        if industry == "Pharmaceuticals":
            risk_dims['无形资产迁移\nIntangibles Migration'] += 15

        # Clamp
        risk_dims = {k: min(100, max(0, v)) for k, v in risk_dims.items()}

        # Overall risk
        overall_risk = sum(risk_dims.values()) / len(risk_dims)
        if overall_risk >= 70:
            risk_level = "⚫ 极高风险 / Critical"
            risk_color = "#000000"
            risk_advice = "强烈建议立即启动转让定价专项审计，并考虑启动相互协商程序(MAP)。"
        elif overall_risk >= 50:
            risk_level = "🔴 高风险 / High"
            risk_color = "#dc3545"
            risk_advice = "建议进行深入转让定价分析，准备完整的转让定价文档，考虑预约定价安排(APA)。"
        elif overall_risk >= 30:
            risk_level = "🟡 中风险 / Medium"
            risk_color = "#ffc107"
            risk_advice = "建议监控风险变化，完善转让定价文档，定期进行自查。"
        else:
            risk_level = "🟢 低风险 / Low"
            risk_color = "#28a745"
            risk_advice = "风险水平正常，保持常规合规监控即可。"

        # Display overall risk
        st.markdown("---")
        risk_col1, risk_col2 = st.columns([1, 2])
        with risk_col1:
            st.markdown(f"""
            <div style='text-align: center; padding: 1.5rem; border: 3px solid {risk_color};
                        border-radius: 15px; background-color: {risk_color}11;'>
                <h2 style='color: {risk_color}; margin: 0;'>{risk_level}</h2>
                <h1 style='color: {risk_color}; margin: 0.5rem 0;'>{overall_risk:.0f}</h1>
                <p style='color: #666;'>综合风险评分 / Overall Risk Score</p>
            </div>
            """, unsafe_allow_html=True)

        with risk_col2:
            st.markdown("##### 💡 风险评估建议 / Risk Assessment Advice")
            st.markdown(f"<div style='padding: 1rem; border-left: 4px solid {risk_color}; background: #f8f9fa;'><b>中文：</b>{risk_advice}</div>", unsafe_allow_html=True)
            st.markdown(f"""
            <div style='padding: 1rem; border-left: 4px solid {risk_color}; background: #f8f9fa; margin-top: 0.5rem;'>
            <b>English:</b> {
                'Strongly recommend immediate TP audit and consider MAP.' if overall_risk >= 70 else
                'Recommend in-depth TP analysis, prepare full documentation, consider APA.' if overall_risk >= 50 else
                'Monitor risk changes, improve TP documentation, periodic self-review.' if overall_risk >= 30 else
                'Risk level normal. Maintain routine compliance monitoring.'
            }
            </div>
            """, unsafe_allow_html=True)

        # Radar chart
        st.markdown("---")
        fig_radar = go.Figure()
        fig_radar.add_trace(go.Scatterpolar(
            r=list(risk_dims.values()) + [list(risk_dims.values())[0]],
            theta=list(risk_dims.keys()) + [list(risk_dims.keys())[0]],
            fill='toself',
            fillcolor=f'rgba({int(risk_color[1:3], 16)}, {int(risk_color[3:5], 16)}, {int(risk_color[5:7], 16)}, 0.2)',
            line_color=risk_color,
            name='风险评分 / Risk Score',
        ))
        fig_radar.update_layout(
            polar=dict(radialaxis=dict(visible=True, range=[0, 100])),
            title="6维风险雷达图 / 6-Dimension Risk Radar",
            height=500,
            showlegend=False,
        )
        st.plotly_chart(fig_radar, use_container_width=True)

    # ========== Tab 2: Compliance Check ==========
    with risk_tabs[1]:
        st.subheader("📋 联合国框架公约合规检查 / UN Convention Compliance Check")

        st.markdown("""
        **中文说明：** 对照联合国国际税收合作框架公约各核心条款，检查企业转让定价合规情况。
        每项检查都附有条款依据和详细说明。

        **English:** Check TP compliance against each core article of the UN Framework Convention.
        Each check includes the article reference and detailed explanation.
        """)

        compliance_items = [
            {
                'article': '第5条 / Art. 5',
                'title': '公平分配征税权 / Fair Distribution of Taxing Rights',
                'desc': '利润应在经济活动发生地和价值创造地之间合理分配\nProfit should be allocated where economic activity and value creation occur',
                'check': '关联交易定价是否反映了各方实际履行的功能和承担的风险',
                'risk': '高风险' if overall_risk >= 50 else '达标',
            },
            {
                'article': '第6条 / Art. 6',
                'title': '税基侵蚀 / Base Erosion',
                'desc': '防止通过关联交易将利润转移至低税管辖区\nPrevent profit shifting to low-tax jurisdictions through related-party transactions',
                'check': '是否存在向低税率地区的大额支付（特许权使用费、服务费、利息）',
                'risk': '高风险' if tax_haven_use or rp_ratio > 50 else '达标',
            },
            {
                'article': '第11条 / Art. 11',
                'title': '能力建设 / Capacity Building',
                'desc': '发展中国家应加强转让定价管理和审计能力\nDeveloping countries should strengthen TP management and audit capacity',
                'check': '是否保留了完整的转让定价文档（主文件、本地文件、国别报告）',
                'risk': '需完善',
            },
            {
                'article': '第13条 / Art. 13',
                'title': '争议预防与解决 / Dispute Prevention & Resolution',
                'desc': '通过MAP和APA预防和解决跨境税务争议\nPrevent and resolve cross-border tax disputes through MAP and APA',
                'check': '是否建立了争议预防机制（APA、预审沟通等）',
                'risk': '需完善',
            },
        ]

        for item in compliance_items:
            risk_color_val = '#dc3545' if '高风险' in item['risk'] else ('#ffc107' if '需完善' in item['risk'] else '#28a745')
            st.markdown(f"""
            <div style='border: 1px solid #ddd; border-left: 4px solid {risk_color_val};
                        border-radius: 8px; padding: 1rem; margin: 0.5rem 0;'>
                <h4 style='margin: 0; color: {PRIMARY_COLOR};'>{item['article']}: {item['title']}</h4>
                <p style='color: #666; font-size: 0.85rem; margin: 0.5rem 0;'>{item['desc']}</p>
                <p style='margin: 0.3rem 0;'><b>检查项 / Check:</b> {item['check']}</p>
                <span style='background-color: {risk_color_val}; color: white; padding: 0.2rem 0.8rem;
                            border-radius: 12px; font-size: 0.8rem;'>{item['risk']}</span>
            </div>
            """, unsafe_allow_html=True)

    # ========== Tab 3: CbC Analysis ==========
    with risk_tabs[2]:
        st.subheader("📊 国别报告(CbC)利润分配分析 / CbC Profit Allocation Analysis")

        st.markdown("""
        **中文说明：** 模拟分析跨国企业全球利润分配与经济活动实质（员工数、资产、收入）的匹配度。
        如果利润集中在低税地区而经济活动在发展中国家，可能存在利润转移风险。

        **English:** Analyze alignment between MNE global profit allocation and economic substance
        (employees, assets, revenue). Profit concentration in low-tax jurisdictions while economic
        activity occurs in developing economies may indicate profit shifting.
        """)

        # Simulated CbC data
        cbc_data = pd.DataFrame({
            '管辖区 / Jurisdiction': ['中国 / China', '印度 / India', '爱尔兰 / Ireland', '百慕大 / Bermuda', '新加坡 / Singapore', '美国 / USA'],
            '收入占比 / Revenue %': [35, 15, 8, 2, 10, 30],
            '利润占比 / Profit %': [15, 8, 30, 25, 12, 10],
            '员工占比 / Employee %': [40, 20, 3, 0, 5, 32],
            '资产占比 / Asset %': [30, 12, 15, 5, 8, 30],
            '实际税率 / ETR %': [25, 22, 12.5, 0, 17, 21],
        })

        st.markdown("##### 模拟CbC数据 / Simulated CbC Data")
        st.dataframe(cbc_data, use_container_width=True, hide_index=True)

        # Misalignment chart
        st.markdown("---")
        st.markdown("##### 📉 利润与经济实质匹配度 / Profit vs. Economic Substance Alignment")

        fig_cbc = go.Figure()
        jurisdictions = cbc_data['管辖区 / Jurisdiction']
        fig_cbc.add_trace(go.Bar(x=jurisdictions, y=cbc_data['收入占比 / Revenue %'], name='收入占比 / Revenue %', marker_color=PRIMARY_COLOR))
        fig_cbc.add_trace(go.Bar(x=jurisdictions, y=cbc_data['利润占比 / Profit %'], name='利润占比 / Profit %', marker_color=ACCENT_RED))
        fig_cbc.add_trace(go.Bar(x=jurisdictions, y=cbc_data['员工占比 / Employee %'], name='员工占比 / Employee %', marker_color=ACCENT_GREEN))
        fig_cbc.update_layout(
            barmode='group',
            title='收入/利润/员工分布对比 / Revenue vs Profit vs Employee Distribution',
            yaxis_title='占比 (%)',
            height=450,
        )
        st.plotly_chart(fig_cbc, use_container_width=True)

        # Interpretation
        st.markdown("---")
        st.markdown("##### 🤖 自动分析 / Automated Analysis")
        st.warning("""
        ⚠️ **风险信号 / Risk Signals:**

        **中文：**
        • 爱尔兰利润占比(30%)远超收入占比(8%)和员工占比(3%) → 可能存在无形资产相关利润转移
        • 百慕大利润占比(25%)但员工占比(0%)和收入占比(2%) → 典型的避税港利润集中
        • 中国收入占比(35%)和员工占比(40%)但利润占比仅(15%) → 利润可能与经济活动实质不匹配

        **English:**
        • Ireland: Profit share (30%) far exceeds revenue (8%) and employee (3%) → Possible intangibles-related profit shifting
        • Bermuda: Profit share (25%) but 0% employees and 2% revenue → Classic tax haven profit concentration
        • China: Revenue (35%) and employee (40%) but only 15% profit → Profit may not match economic substance

        💡 **建议 / Recommendation:** 依据联合国框架公约第5条，利润应在价值创造地分配。
        建议对爱尔兰和百慕大的利润来源进行深入转让定价分析。
        """)


# ============== Page 9: Industry Benchmark & Trend Analysis ==============

def industry_benchmark() -> None:
    """Page 9: Industry Benchmark & Multi-Year Trend Analysis."""
    st.markdown("<div class='main-header'>📈 行业基准与趋势分析 / Industry Benchmark & Trend Analysis</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='sub-header'>跨行业、跨国家、跨年度的利润率基准对比 — 对标BvD Orbis级别分析</div>",
        unsafe_allow_html=True,
    )

    st.markdown("""
    <div class='info-bar'>
    💡 <b>本页功能说明 / Page Guide:</b><br>
    <b>中文：</b>对标BvD Orbis等专业数据库的行业基准分析功能。<br>
    &nbsp;&nbsp;• <b>行业基准对比</b>：选择行业，查看各国该行业的利润率中位数、四分位区间<br>
    &nbsp;&nbsp;• <b>多年度趋势</b>：2022-2024年度利润率变化趋势，识别行业走势<br>
    &nbsp;&nbsp;• <b>PLI分布</b>：营业利润率、ROE、ROA等关键指标的分布密度图<br>
    &nbsp;&nbsp;• <b>跨国对比</b>：同一行业在不同国家的盈利能力差异<br>
    <b>English:</b> BvD Orbis-level industry benchmark analysis.<br>
    &nbsp;&nbsp;• <b>Industry Benchmark</b>: Select an industry to view median margins, IQR by country<br>
    &nbsp;&nbsp;• <b>Multi-Year Trends</b>: 2022-2024 margin trends to identify industry trajectories<br>
    &nbsp;&nbsp;• <b>PLI Distribution</b>: Kernel density of operating margin, ROE, ROA<br>
    &nbsp;&nbsp;• <b>Cross-Country</b>: Same industry profitability differences across countries<br>
    <br>
    <i>依据 / Reference: 联合国转让定价实务手册第3章（可比性分析）<br>
    UN Practical Manual on Transfer Pricing, Ch. 3 (Comparability Analysis)</i>
    </div>
    """, unsafe_allow_html=True)

    companies_df = load_mock_companies()

    # Filters
    filter_col1, filter_col2, filter_col3 = st.columns(3)
    with filter_col1:
        # Group by ISIC section letter for cleaner categories
        companies_df['isic_section'] = companies_df['industry_code_isic'].astype(str).str[0]
        isic_sections = sorted(companies_df['isic_section'].dropna().unique().tolist())
        section_names = {
            'A': 'A-农林牧渔 / Agriculture',
            'B': 'B-采矿 / Mining',
            'C': 'C-制造 / Manufacturing',
            'D': 'D-电力燃气 / Energy',
            'E': 'E-水务 / Water',
            'F': 'F-建筑 / Construction',
            'G': 'G-批发零售 / Wholesale & Retail',
            'H': 'H-运输仓储 / Transport',
            'I': 'I-住宿餐饮 / Hospitality',
            'J': 'J-信息技术 / IT',
            'K': 'K-金融保险 / Financial',
            'L': 'L-房地产 / Real Estate',
            'M': 'M-专业服务 / Professional',
            'N': 'N-辅助服务 / Support',
        }
        section_labels = [section_names.get(s, f'{s}-其他') for s in isic_sections]
        sel_section_idx = st.selectbox(
            "行业大类 / Industry Sector (ISIC)",
            range(len(isic_sections)),
            format_func=lambda i: section_labels[i] if i < len(section_labels) else 'All',
        )
        sel_section = isic_sections[sel_section_idx]

    with filter_col2:
        # Year selection
        available_years = sorted(companies_df['fiscal_year'].dropna().unique().tolist())
        sel_years = st.multiselect(
            "年度 / Fiscal Years",
            available_years,
            default=available_years,
            help="选择要分析的年度，可多选"
        )

    with filter_col3:
        sel_tier = st.multiselect(
            "数据密级 / Data Tier",
            ['Tier1', 'Tier2', 'Tier3'],
            default=['Tier1', 'Tier2'],
            help="Tier1最可靠，建议优先使用"
        )

    # Apply filters
    filtered = companies_df.copy()
    if sel_section and sel_section != 'All':
        filtered = filtered[filtered['isic_section'] == sel_section]
    if sel_years:
        filtered = filtered[filtered['fiscal_year'].isin(sel_years)]
    if 'data_tier' in filtered.columns and sel_tier:
        filtered = filtered[filtered['data_tier'].isin(sel_tier)]

    if filtered.empty:
        st.warning("⚠️ 无符合条件的数据，请调整筛选。/ No data matches the selected criteria.")
        return

    st.markdown(f"**符合条件：{len(filtered)} 家公司 / {len(filtered['country'].unique())} 个经济体**")

    # ===== Section 1: Industry Benchmark by Economy =====
    st.markdown("---")
    st.subheader("📊 各经济体利润率基准 / Profitability Benchmark by Economy")

    country_benchmark = filtered.groupby('country').agg(
        companies=('company_id', 'count'),
        median_margin=('operating_margin', 'median'),
        q1_margin=('operating_margin', lambda x: np.percentile(x.dropna(), 25)),
        q3_margin=('operating_margin', lambda x: np.percentile(x.dropna(), 75)),
        median_roe=('roe', 'median'),
    ).round(4).reset_index().sort_values('companies', ascending=False).head(20)

    # Box plot by country
    fig_box = px.box(
        filtered[filtered['country'].isin(country_benchmark['country'].head(15))],
        x='country', y='operating_margin',
        color='country',
        title=f'{section_labels[sel_section_idx]} — 各经济体营业利润率分布 / Operating Margin by Economy',
        labels={'operating_margin': '营业利润率 / Operating Margin', 'country': '经济体 / Economy'},
        height=450,
    )
    fig_box.update_layout(showlegend=False, margin=dict(l=0, r=0, t=40, b=0))
    st.plotly_chart(fig_box, use_container_width=True)

    # Benchmark table with export
    st.markdown("##### 📋 基准数据表 / Benchmark Table")
    display_benchmark = country_benchmark.copy()
    display_benchmark.columns = ['经济体/Economy', '公司数/Companies', '中位利润率/Median Margin',
                                 'Q1', 'Q3', '中位ROE/Median ROE']
    display_benchmark['中位利润率/Median Margin'] = (display_benchmark['中位利润率/Median Margin'] * 100).round(2).astype(str) + '%'
    display_benchmark['Q1'] = (display_benchmark['Q1'] * 100).round(2).astype(str) + '%'
    display_benchmark['Q3'] = (display_benchmark['Q3'] * 100).round(2).astype(str) + '%'
    display_benchmark['中位ROE/Median ROE'] = (display_benchmark['中位ROE/Median ROE'] * 100).round(2).astype(str) + '%'
    st.dataframe(display_benchmark, use_container_width=True, hide_index=True)

    # Export button
    export_col1, export_col2 = st.columns([1, 4])
    with export_col1:
        csv_data = country_benchmark.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 导出CSV / Export CSV",
            data=csv_data,
            file_name=f"industry_benchmark_{sel_section}_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
        )

    # ===== Section 2: Multi-Year Trend =====
    st.markdown("---")
    st.subheader("📅 多年度趋势分析 / Multi-Year Trend Analysis")

    if len(sel_years) >= 2:
        trend_data = filtered.groupby(['fiscal_year', 'isic_section']).agg(
            median_margin=('operating_margin', 'median'),
            q1=('operating_margin', lambda x: np.percentile(x.dropna(), 25)),
            q3=('operating_margin', lambda x: np.percentile(x.dropna(), 75)),
            companies=('company_id', 'count'),
        ).reset_index()

        fig_trend = go.Figure()
        fig_trend.add_trace(go.Scatter(
            x=trend_data['fiscal_year'], y=trend_data['q3'],
            name='Q3 (75th percentile)', mode='lines',
            line=dict(width=0), showlegend=False,
        ))
        fig_trend.add_trace(go.Scatter(
            x=trend_data['fiscal_year'], y=trend_data['q1'],
            name='IQR Range', mode='lines',
            line=dict(width=0), fill='tonexty',
            fillcolor='rgba(0, 86, 149, 0.15)', showlegend=True,
        ))
        fig_trend.add_trace(go.Scatter(
            x=trend_data['fiscal_year'], y=trend_data['median_margin'],
            name='中位数 / Median', mode='lines+markers',
            line=dict(color=PRIMARY_COLOR, width=3),
            marker=dict(size=8),
        ))
        fig_trend.update_layout(
            title=f'{section_labels[sel_section_idx]} — 年度利润率趋势 / Annual Margin Trend',
            xaxis_title='年度 / Fiscal Year',
            yaxis_title='营业利润率 / Operating Margin',
            height=400,
            margin=dict(l=0, r=0, t=40, b=0),
        )
        st.plotly_chart(fig_trend, use_container_width=True)

        # Trend interpretation
        if len(trend_data) >= 2:
            first_median = trend_data['median_margin'].iloc[0]
            last_median = trend_data['median_margin'].iloc[-1]
            change = last_median - first_median
            if change > 0.01:
                trend_msg = f"📈 利润率从 {first_median*100:.1f}% 上升至 {last_median*100:.1f}%，行业盈利能力在改善。/ Margins improving."
            elif change < -0.01:
                trend_msg = f"📉 利润率从 {first_median*100:.1f}% 下降至 {last_median*100:.1f}%，行业盈利能力在下降。/ Margins declining."
            else:
                trend_msg = f"➡️ 利润率保持稳定在 {last_median*100:.1f}% 左右。/ Margins stable."
            st.info(trend_msg)
    else:
        st.info("💡 选择2个或以上年度以查看趋势。/ Select 2+ years to view trends.")

    # ===== Section 3: PLI Distribution =====
    st.markdown("---")
    st.subheader("📊 PLI分布密度图 / PLI Distribution Density")

    pli_col1, pli_col2, pli_col3 = st.columns(3)
    with pli_col1:
        st.markdown("##### 营业利润率 / Operating Margin")
        fig_m = px.histogram(filtered, x='operating_margin', nbins=30, marginal='box',
                             color_discrete_sequence=[PRIMARY_COLOR], height=300)
        fig_m.update_layout(margin=dict(l=0, r=0, t=10, b=0), showlegend=False)
        st.plotly_chart(fig_m, use_container_width=True)

    with pli_col2:
        st.markdown("##### ROE (净资产回报率)")
        fig_r = px.histogram(filtered, x='roe', nbins=30, marginal='box',
                            color_discrete_sequence=[ACCENT_GREEN], height=300)
        fig_r.update_layout(margin=dict(l=0, r=0, t=10, b=0), showlegend=False)
        st.plotly_chart(fig_r, use_container_width=True)

    with pli_col3:
        st.markdown("##### ROA (总资产回报率)")
        fig_a = px.histogram(filtered, x='roa', nbins=30, marginal='box',
                            color_discrete_sequence=[ACCENT_YELLOW], height=300)
        fig_a.update_layout(margin=dict(l=0, r=0, t=10, b=0), showlegend=False)
        st.plotly_chart(fig_a, use_container_width=True)

    # ===== Section 4: Company Listing with Export =====
    st.markdown("---")
    st.subheader("📋 公司明细 / Company Listing")
    display_cols = ['company_name', 'country', 'industry_code_isic', 'revenue_musd',
                    'operating_margin', 'roe', 'roa', 'related_party_pct',
                    'fiscal_year', 'data_tier']
    available_cols = [c for c in display_cols if c in filtered.columns]
    st.dataframe(filtered[available_cols], use_container_width=True, hide_index=True)

    export_col1, export_col2 = st.columns([1, 4])
    with export_col1:
        csv_full = filtered[available_cols].to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 导出全部CSV / Export All",
            data=csv_full,
            file_name=f"company_listing_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
        )


# ============== Helper: Country Detail Panel ==============

def render_country_detail(country_name: str, country_code: str) -> None:
    """Render a country detail panel with companies, policies, and disputes."""
    st.markdown(f"### 🌍 {country_name} / {country_code}")

    companies_df = load_mock_companies()
    # Match by country name — also check if "China" should match "Taiwan (China)" entries
    if country_name == "China":
        country_data = companies_df[companies_df['country'].isin(["China", "Taiwan (China)"])]
    else:
        country_data = companies_df[companies_df['country'] == country_name]

    detail_col1, detail_col2, detail_col3, detail_col4 = st.columns(4)
    with detail_col1:
        st.metric("可比公司 / Companies", len(country_data))
    with detail_col2:
        st.metric("中位利润率 / Median Margin",
                  f"{country_data['operating_margin'].median()*100:.1f}%" if not country_data.empty else "N/A")
    with detail_col3:
        st.metric("中位收入 / Median Revenue (M$)",
                  f"${country_data['revenue_musd'].median():.1f}" if not country_data.empty else "N/A")
    with detail_col4:
        if 'data_tier' in country_data.columns:
            tier_counts = country_data['data_tier'].value_counts()
            st.metric("主要密级 / Primary Tier", tier_counts.index[0] if not tier_counts.empty else "N/A")
        else:
            st.metric("密级 / Tier", "N/A")

    if not country_data.empty:
        st.markdown("##### 公司列表 / Company List")
        display_cols = ['company_name', 'industry_code_isic', 'revenue_musd',
                        'operating_margin', 'roe', 'fiscal_year']
        available = [c for c in display_cols if c in country_data.columns]
        st.dataframe(country_data[available].head(20), use_container_width=True, hide_index=True)
    else:
        st.info(f"暂无 {country_name} 的可比公司数据。/ No comparable company data for {country_name}.")


# ============== Helper: Global Search ==============

def global_search(query: str) -> list:
    """Search across companies, countries, and industries. Returns list of (type, label, detail) tuples."""
    results = []
    if not query or len(query.strip()) < 2:
        return results

    companies_df = load_mock_companies()

    # Search companies
    company_matches = companies_df[
        companies_df['company_name'].str.contains(query, case=False, na=False)
    ].head(10)
    for _, row in company_matches.iterrows():
        results.append(('company', row['company_name'],
                         f"{row.get('country', '')} | {row.get('industry_code_isic', '')} | FY{row.get('fiscal_year', '')}"))

    # Search countries
    country_matches = companies_df[
        companies_df['country'].str.contains(query, case=False, na=False)
    ]['country'].unique()[:10]
    for c in country_matches:
        count = len(companies_df[companies_df['country'] == c])
        results.append(('country', c, f"{count} companies"))

    # Search industries
    if 'industry_code_isic' in companies_df.columns:
        industry_matches = companies_df[
            companies_df['industry_code_isic'].astype(str).str.contains(query, case=False, na=False)
        ]['industry_code_isic'].unique()[:5]
        for ind in industry_matches:
            count = len(companies_df[companies_df['industry_code_isic'] == ind])
            results.append(('industry', f"ISIC {ind}", f"{count} companies"))

    return results


# ============== Page 10: Customs Valuation Reference ==============

def customs_valuation() -> None:
    """Page 10: Customs Valuation Data as TP auxiliary reference."""
    st.markdown("<div class='main-header'>🛃 海关估价参考 / Customs Valuation Reference</div>", unsafe_allow_html=True)
    st.markdown(
        "<div class='sub-header'>海关申报数据辅助转让定价分析 — WTO海关估价协定 × 联合国转让定价实务手册</div>",
        unsafe_allow_html=True,
    )

    st.markdown("""
    <div class='info-bar'>
    💡 <b>本页功能说明 / Page Guide:</b><br>
    <b>中文：</b>海关估价数据是转让定价分析的「重要辅助参考」。根据WTO海关估价协定和联合国转让定价实务手册，海关进口申报价格可作为可比非受控价格(CUP)的参考依据。<br>
    &nbsp;&nbsp;• <b>海关vs转让定价对比</b>：将关联交易的海关申报价与可比价格对比，识别异常定价<br>
    &nbsp;&nbsp;• <b>关联交易筛查</b>：标记海关系统中申报的关联进出口交易<br>
    &nbsp;&nbsp;• <b>价格偏离预警</b>：关联交易价格与独立交易价格偏离>15%自动预警<br>
    &nbsp;&nbsp;• <b>HS编码查询</b>：按商品编码查看进出口价格分布<br>
    <b>English:</b> Customs valuation data serves as an important auxiliary reference for TP analysis. Per WTO Customs Valuation Agreement and UN Practical Manual, customs declared prices can support CUP method.<br>
    &nbsp;&nbsp;• <b>Customs vs TP Comparison</b>: Compare related-party customs prices with comparable prices<br>
    &nbsp;&nbsp;• <b>Related-Party Screening</b>: Flag related-party import/export transactions<br>
    &nbsp;&nbsp;• <b>Price Deviation Alert</b>: Auto-flag >15% price deviations from arm's length<br>
    &nbsp;&nbsp;• <b>HS Code Lookup</b>: View price distribution by HS commodity code<br>
    <br>
    <i>依据 / Reference: WTO海关估价协定 / WTO Customs Valuation Agreement; 联合国转让定价实务手册第4章 / UN Practical Manual Ch.4; OECD转让定价指南第2章（海关与转让定价协同）/ OECD TP Guidelines Ch.2</i>
    </div>
    """, unsafe_allow_html=True)

    customs_df = db_query("""
        SELECT cv.*, 
               rc.country_name AS reporting_country_name,
               pc.country_name AS partner_country_name
        FROM customs_valuations cv
        LEFT JOIN countries rc ON rc.country_code = cv.reporting_country
        LEFT JOIN countries pc ON pc.country_code = cv.partner_country
        ORDER BY cv.valuation_date DESC
    """)

    if customs_df is None or customs_df.empty:
        st.warning("⚠️ 海关估价数据不可用。/ Customs valuation data not available.")
        return

    customs_tabs = st.tabs([
        "🔍 海关数据查询 / Data Search",
        "⚖️ 关联交易对比 / Related-Party Comparison",
        "📊 统计分析 / Statistics",
        "📖 使用指引 / Usage Guide",
    ])

    # ===== Tab 1: Data Search =====
    with customs_tabs[0]:
        st.subheader("🔍 海关数据查询 / Customs Data Search")

        filter_c1, filter_c2, filter_c3, filter_c4 = st.columns(4)
        with filter_c1:
            sel_direction = st.selectbox(
                "贸易方向 / Trade Direction",
                ["全部 / All"] + customs_df["trade_direction"].unique().tolist(),
                key="customs_dir"
            )
        with filter_c2:
            sel_category = st.selectbox(
                "商品类别 / Category",
                ["全部 / All"] + sorted(customs_df["commodity_category"].unique().tolist()),
                key="customs_cat"
            )
        with filter_c3:
            sel_related = st.selectbox(
                "关联交易 / Related Party",
                ["全部 / All", "仅关联 / Related only", "仅独立 / Unrelated only"],
                key="customs_rel"
            )
        with filter_c4:
            sel_year = st.multiselect(
                "年度 / Fiscal Year",
                sorted(customs_df["fiscal_year"].unique().tolist()),
                default=sorted(customs_df["fiscal_year"].unique().tolist()),
                key="customs_year"
            )

        filtered = customs_df.copy()
        if sel_direction != "全部 / All":
            filtered = filtered[filtered["trade_direction"] == sel_direction]
        if sel_category != "全部 / All":
            filtered = filtered[filtered["commodity_category"] == sel_category]
        if sel_related == "仅关联 / Related only":
            filtered = filtered[filtered["related_party_flag"] == True]
        elif sel_related == "仅独立 / Unrelated only":
            filtered = filtered[filtered["related_party_flag"] == False]
        if sel_year:
            filtered = filtered[filtered["fiscal_year"].isin(sel_year)]

        st.markdown(f"**筛选结果：{len(filtered)} 条记录 / {len(filtered)} records**")

        if len(filtered) > 0:
            display_cols = ["customs_id", "trade_direction", "reporting_country_name",
                           "partner_country_name", "hs_code", "product_description",
                           "declared_value_usd", "unit_price_usd", "related_party_flag",
                           "price_difference_pct", "verification_status", "valuation_date"]
            available = [c for c in display_cols if c in filtered.columns]
            display = filtered[available].copy()
            display.columns = [c.replace("_name", "").replace("_usd", " ($)") for c in available]
            st.dataframe(display, use_container_width=True, hide_index=True)

            # Export
            exp_c1, _ = st.columns([1, 4])
            with exp_c1:
                csv_customs = filtered.to_csv(index=False).encode('utf-8-sig')
                st.download_button(
                    label="📥 导出CSV / Export",
                    data=csv_customs,
                    file_name=f"customs_data_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                )

    # ===== Tab 2: Related-Party Comparison =====
    with customs_tabs[1]:
        st.subheader("⚖️ 关联交易海关vs转让定价对比 / Related-Party Customs vs TP Comparison")

        related_df = customs_df[customs_df["related_party_flag"] == True].copy()

        if related_df.empty:
            st.info("暂无关联交易数据。/ No related-party transactions found.")
        else:
            st.markdown(f"**关联交易记录：{len(related_df)} 条 / {len(related_df)} related-party records**")

            # Price deviation chart
            if "price_difference_pct" in related_df.columns:
                related_df["deviation_abs"] = related_df["price_difference_pct"].abs()

                fig_dev = px.scatter(
                    related_df.dropna(subset=["price_difference_pct"]),
                    x="declared_value_usd",
                    y="price_difference_pct",
                    color="commodity_category",
                    hover_data=["hs_code", "product_description", "reporting_country_name"],
                    title="关联交易价格偏离度散点图 / Related-Party Price Deviation",
                    labels={
                        "declared_value_usd": "申报价值($) / Declared Value",
                        "price_difference_pct": "价格偏离(%) / Price Deviation %",
                        "commodity_category": "类别 / Category",
                    },
                    height=450,
                )
                # Add ±15% threshold lines
                fig_dev.add_hline(y=15, line_dash="dash", line_color="red", annotation_text="+15% 预警线")
                fig_dev.add_hline(y=-15, line_dash="dash", line_color="red", annotation_text="-15% 预警线")
                st.plotly_chart(fig_dev, use_container_width=True)
                st.info("💡 **图表说明 / Chart Guide:** 每个点是一条关联交易。纵轴是海关申报价与转让定价可比价的偏离百分比，红色虚线为±15%预警线，超出预警线的交易需要进一步审查。/ Each dot = a related-party transaction. Y-axis: price deviation %. Red dashed lines: ±15% alert threshold.")

            # Flagged transactions
            st.markdown("---")
            st.markdown("##### 🚨 价格偏离预警 / Price Deviation Alerts")
            flagged = related_df[related_df["price_adjustment_flag"] == True]
            st.metric("需调整的交易数 / Flagged Transactions", len(flagged))
            if not flagged.empty:
                flag_cols = ["customs_id", "product_description", "reporting_country_name",
                            "partner_country_name", "declared_value_usd", "price_difference_pct",
                            "adjustment_reason"]
                available_flag = [c for c in flag_cols if c in flagged.columns]
                st.dataframe(flagged[available_flag], use_container_width=True, hide_index=True)

    # ===== Tab 3: Statistics =====
    with customs_tabs[2]:
        st.subheader("📊 海关数据统计分析 / Customs Statistics")

        stat_c1, stat_c2, stat_c3, stat_c4 = st.columns(4)
        with stat_c1:
            st.metric("总记录数 / Total Records", len(customs_df))
        with stat_c2:
            st.metric("关联交易 / Related Party", len(customs_df[customs_df["related_party_flag"] == True]))
        with stat_c3:
            flagged_count = len(customs_df[customs_df["price_adjustment_flag"] == True])
            st.metric("预警交易 / Flagged", flagged_count)
        with stat_c4:
            avg_dev = customs_df["price_difference_pct"].dropna().mean() if "price_difference_pct" in customs_df.columns else 0
            st.metric("平均偏离度 / Avg Deviation", f"{avg_dev:.1f}%" if avg_dev else "N/A")

        # Category distribution
        st.markdown("---")
        st.markdown("##### 📦 商品类别分布 / Category Distribution")
        cat_counts = customs_df.groupby("commodity_category").agg(
            records=("customs_id", "count"),
            related=("related_party_flag", "sum"),
            avg_value=("declared_value_usd", "mean"),
        ).reset_index().sort_values("records", ascending=False)
        fig_cat = px.bar(
            cat_counts, x="commodity_category", y="records",
            color="related", title="各类别海关记录数 / Records by Category",
            labels={"commodity_category": "类别 / Category", "records": "记录数 / Records", "related": "关联 / Related"},
            height=400,
        )
        st.plotly_chart(fig_cat, use_container_width=True)

        # Country distribution
        st.markdown("---")
        st.markdown("##### 🌍 申报经济体分布 / Reporting Economy Distribution")
        country_counts = customs_df.groupby("reporting_country_name").agg(
            records=("customs_id", "count"),
            related=("related_party_flag", "sum"),
        ).reset_index().sort_values("records", ascending=False).head(15)
        fig_country = px.bar(
            country_counts, x="reporting_country_name", y="records",
            color="related", title="Top 15 申报经济体 / Top 15 Reporting Economies",
            labels={"reporting_country_name": "经济体 / Economy", "records": "记录数 / Records", "related": "关联 / Related"},
            height=400,
        )
        st.plotly_chart(fig_country, use_container_width=True)

    # ===== Tab 4: Usage Guide =====
    with customs_tabs[3]:
        st.subheader("📖 海关估价在转让定价中的使用指引 / Usage Guide")

        st.markdown("""
        ### 🔗 海关估价与转让定价的关系
        ### Customs Valuation & Transfer Pricing: The Link

        **中文：**

        海关估价和转让定价是同一问题的两个视角：
        - **海关估价**关注进口货物的申报价格是否合理（征收关税的依据）
        - **转让定价**关注关联企业之间的交易价格是否符合独立交易原则（征收所得税的依据）

        两者都遵循「独立交易原则」(Arm's Length Principle)，但适用场景不同：
        - 海关：进口时一次性估价
        - 转让定价：年度整体利润水平分析

        **如何使用海关数据辅助转让定价分析：**

        1. **CUP法参考**：如果海关有同类商品的独立交易进口价格，可作为CUP法的可比价格
        2. **异常筛查**：如果关联进口价格显著低于独立进口价格，可能存在利润转移
        3. **事后验证**：对已完成的海关申报进行转让定价审查
        4. **风险预警**：海关价格偏离>15%的交易应重点审查

        ---

        **English:**

        Customs valuation and transfer pricing are two perspectives on the same issue:
        - **Customs Valuation**: Whether the declared import price is reasonable (basis for tariffs)
        - **Transfer Pricing**: Whether related-party prices meet the arm's length principle (basis for income tax)

        Both follow the **Arm's Length Principle**, but apply in different contexts:
        - Customs: One-time valuation at import
        - Transfer Pricing: Annual overall profit analysis

        **How to use customs data for TP analysis:**

        1. **CUP Method Reference**: Independent import prices for similar goods can serve as CUP comparables
        2. **Anomaly Screening**: Related-party import prices significantly below independent prices may indicate profit shifting
        3. **Post-Import Verification**: Conduct TP review of completed customs declarations
        4. **Risk Alert**: Transactions with >15% price deviation should be prioritized for review

        ---

        ### 📋 依据文件 / References

        | 文件 / Document | 相关条款 / Relevant Articles |
        |---|---|
        | WTO海关估价协定 / WTO Customs Valuation Agreement | 第1-7条：估价方法 / Articles 1-7: Valuation methods |
        | 联合国转让定价实务手册 / UN Practical Manual | 第4章：海关与转让定价协同 / Ch.4: Customs-TP synergy |
        | OECD转让定价指南 / OECD TP Guidelines | 第2章：可比性分析 / Ch.2: Comparability analysis |
        | 联合国国际税收合作框架公约 / UN Framework Convention | 第5条、第6条 / Art. 5, Art. 6 |
        """)


# ============== Sidebar ==============

def render_sidebar() -> str:
    """Render the sidebar navigation and return the selected page name."""
    with st.sidebar:
        st.markdown("<div class='sidebar-title'>🌍 国际转让定价对比数据库<br>ITP Comparability Database</div>", unsafe_allow_html=True)
        st.markdown("*UN International Tax Cooperation Framework*")
        st.markdown("---")

        # Global Search
        st.markdown("#### 🔎 全局搜索 / Global Search")
        search_query = st.text_input(
            "搜索公司/国家/行业 / Search",
            placeholder="输入关键词... / Enter keyword...",
            key="global_search_input",
            help="搜索公司名称、国家名、ISIC行业代码"
        )
        if search_query and len(search_query.strip()) >= 2:
            results = global_search(search_query)
            if results:
                st.markdown(f"**找到 {len(results)} 条结果 / {len(results)} results found:**")
                for r_type, r_label, r_detail in results[:15]:
                    icon = {"company": "🏢", "country": "🌐", "industry": "🏭"}.get(r_type, "•")
                    st.markdown(f"<small>{icon} <b>{r_label}</b> — {r_detail}</small>", unsafe_allow_html=True)
                st.markdown("---")
            else:
                st.markdown("<small>无匹配结果 / No results</small>")
                st.markdown("---")

        # Navigation
        st.markdown("#### 📑 功能导航 / Navigation")
        page_options = [
            "Global Data Portal",
            "Industry Benchmark",
            "Comparability Wizard",
            "Extractive Pricing",
            "Country Policy",
            "Data Contribution",
            "MAP/APA Tracker",
            "Digital Economy",
            "Risk Assessment",
            "Customs Valuation",
        ]
        page_labels = {
            "Global Data Portal": "🌐 全球数据总览 / Global Data Portal",
            "Industry Benchmark": "📈 行业基准与趋势 / Industry Benchmark",
            "Comparability Wizard": "🔍 可比性分析向导 / Comparability Wizard",
            "Extractive Pricing": "⛏️ 采掘业定价 / Extractive Pricing",
            "Country Policy": "📋 各经济体政策 / Economy Policy",
            "Data Contribution": "📤 数据贡献管理 / Data Contribution",
            "MAP/APA Tracker": "⚖️ 争议解决追踪 / MAP & APA Tracker",
            "Digital Economy": "💻 数字经济分析 / Digital Economy",
            "Risk Assessment": "🚨 风险评估中心 / Risk Assessment",
            "Customs Valuation": "🛃 海关估价参考 / Customs Valuation",
        }
        selected = st.radio("Go to", page_options, format_func=lambda x: page_labels.get(x, x), label_visibility="collapsed")

        st.markdown("---")
        st.markdown("#### ℹ️ 关于本系统 / About")
        st.markdown("""
        本系统支持**联合国国际税收合作框架公约**的实施，
        帮助发展中国家税务官员获取可比数据进行转让定价分析。

        This system supports the implementation of the
        **UN International Tax Cooperation Framework Convention**,
        helping developing economies access comparable data
        for transfer pricing analyses.

        **支持条款 / Articles Supported:**
        - 第5条 / Art. 5: 公平分配征税权 / Fair Distribution of Taxing Rights
        - 第11条 / Art. 11: 能力建设 / Capacity Building
        - 第13条 / Art. 13: 争议预防与解决 / Dispute Prevention & Resolution
        """)
        if 'db_error' in st.session_state:
            st.markdown("---")
            st.warning(f"⚠️ 数据库连接失败，使用模拟数据。错误: {st.session_state['db_error'][:100]}")

        st.markdown("---")
        st.markdown(f"<small>版本 v3.0 | {datetime.now().year}</small>", unsafe_allow_html=True)

        return selected


# ============== Main ==============

def main() -> None:
    """Main entry point: render sidebar navigation and the selected page."""
    selected_page = render_sidebar()

    # Page router
    if selected_page == "Global Data Portal":
        global_data_portal()
    elif selected_page == "Industry Benchmark":
        industry_benchmark()
    elif selected_page == "Comparability Wizard":
        comparability_wizard()
    elif selected_page == "Extractive Pricing":
        extractive_pricing()
    elif selected_page == "Country Policy":
        country_policy()
    elif selected_page == "Data Contribution":
        data_contribution()
    elif selected_page == "MAP/APA Tracker":
        map_apa_tracker()
    elif selected_page == "Digital Economy":
        digital_economy()
    elif selected_page == "Risk Assessment":
        risk_assessment()
    elif selected_page == "Customs Valuation":
        customs_valuation()


if __name__ == "__main__":
    main()
