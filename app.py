# BUILD: SIMPLE-V40-CONFIDENCE-MACHINE-MATCHING-20260820\n# BUILD: SIMPLE-V39-SMART-MACHINE-DUPLICATE-REVIEW-20260820
# BUILD: SIMPLE-V38-COMPLIANCE-MACHINE-REVIEW-20260820
# BUILD: SIMPLE-V37-SINGLE-PHOTO-CONTROL-20260820
# BUILD: SIMPLE-V36-ALL-PORTALS-SPEED-OPTIMIZED-20260820
# BUILD: SIMPLE-V35-MANAGER-SPEED-FIX-20260820
# BUILD: SIMPLE-V34-CROSS-DEVICE-DUPLICATE-LOCK-20260820
# BUILD: SIMPLE-V33-DUPLICATE-SUBMIT-GUARD-20260820
# BUILD: SIMPLE-V32-SHIFT-REQUIRED-FIX-20260820
# BUILD: SIMPLE-V31-PHOTO-VISIBILITY-CAMERA-20260820
# BUILD: SIMPLE-V30-SUPABASE-PHOTO-STORAGE-20260820
# BUILD: SIMPLE-V29-EXCEL-EXPORT-FIX-20260820\n# BUILD: SIMPLE-V28-DRIVER-SPEED-POOL-20260820
# BUILD: SIMPLE-V27-PERFORMANCE-OPTIMIZED-20260820
# BUILD: SIMPLE-V26-ROLE-LOCATION-UI-FIX-20260820
# BUILD: SIMPLE-V25-ROLE-ACCESS-ENFORCED-20260820
# BUILD: SIMPLE-V24-FULL-MACHINE-HISTORY-20260820
# BUILD: SIMPLE-V23-STAFF-FIRST-LOGIN-FIX-20260820
# BUILD: SIMPLE-V22-SUPABASE-EMPTY-QUERY-FIX-20260820\n# BUILD: SIMPLE-V21-SUPABASE-DATABASE-20260820
# BUILD: SIMPLE-V20-CLOUD-STABLE-ALL-EXCEL-20260819
# BUILD: SIMPLE-V19-IT-TEST-PORTAL-20260819
# BUILD: SIMPLE-V18-IT-ADMIN-CLEANUP-20260819
# BUILD: SIMPLE-V17-CLOUD-EXCEL-FIX-20260819
# BUILD: SIMPLE-V16-CLOUD-SAFE
# BUILD: SIMPLE-V15-MANAGER-COMPLIANCE-TABS-20260819
# BUILD: SIMPLE-V14-MAINTENANCE-TABS-20260819
# BUILD: SIMPLE-V13-SHIFT-REPORTS-MACHINE-HISTORY-20260819
# BUILD: SIMPLE-V12-SHIFT-FIX-CLEAN-20260819
# BUILD: SIMPLE-V11-SHIFT-DB-FIXED-20260819
# BUILD: SIMPLE-V11-SHIFTS-MACHINE-HISTORY-20260819
# BUILD: SIMPLE-V10-HUMAN-LABELS-TIMELINE-20260819
# BUILD: SIMPLE-V9-SUBMIT-MESSAGE-UI-20260819
# BUILD: SIMPLE-V8-DOWNLOAD-TODAY-STATUS-20260818
# BUILD: SIMPLE-V7-DAILY-TASKS-CASE-HISTORY-20260818
# BUILD: SIMPLE-V5-COMPLIANCE-TABS-20260818
# BUILD: SIMPLE-V5-CLEAN-REPORTS-20260818
# BUILD: SIMPLE-V5-RESET-AND-DETAILED-VALIDATION-20260818
# BUILD: SIMPLE-V5-FIXED-COLUMNS-20260818
# BUILD: SIMPLE-V5-20260818
# BUILD: V4-AUTO-CLOSE-REPAIR-20260818
# BUILD: V4-TEMP-ROLE-PREVIEW-20260817
# BUILD: V4-STAFF-ROW-ACTIONS-20260817

import streamlit as st
import sqlite3
import psycopg2
from psycopg2 import errors
from psycopg2.extras import RealDictCursor
from psycopg2.pool import ThreadedConnectionPool
import json
import io
import re
import hashlib
import secrets
import smtplib
from pathlib import Path
from datetime import datetime, date, timedelta
from email.message import EmailMessage
import requests
import pandas as pd

BASE = Path(__file__).parent
DB = BASE / "forklift_role_portal_v4.db"
EQ_FILE = BASE / "equipment_master.csv"
UPLOAD = BASE / "uploads"
UPLOAD.mkdir(exist_ok=True)

# STAFF ROW ACTIONS ENABLED
st.set_page_config(
    page_title="18W Forklift Safety",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container{max-width:1280px;padding-top:1rem;padding-bottom:4rem}
.hero{padding:18px 22px;border-radius:18px;background:linear-gradient(135deg,#f3f7fb,#fff);border:1px solid #dfe5ec;margin-bottom:18px}
.hero h2{margin:0 0 5px 0}
.muted{color:#667085;font-size:.9rem}
.section-title{font-weight:800;font-size:1.08rem;margin:1rem 0 .4rem}
.issue-card{border-left:7px solid #d92d20;background:#fff6f5;padding:12px 14px;border-radius:12px;margin:8px 0}
.ok-card{border-left:7px solid #12b76a;background:#f2fff6;padding:12px 14px;border-radius:12px;margin:8px 0}
.purple-card{border-left:7px solid #7f56d9;background:#faf7ff;padding:12px 14px;border-radius:12px;margin:8px 0}
.orange-card{border-left:7px solid #f79009;background:#fffaeb;padding:12px 14px;border-radius:12px;margin:8px 0}
.machine-selected{padding:7px 10px;border-radius:9px;background:#f4f7fb;border:1px solid #dbe4ef;font-size:.90rem;color:#344054;margin:.35rem 0 .65rem 0}
.stButton>button{border-radius:10px;font-weight:750;min-height:2.55rem}
[data-testid="stMetric"]{border:1px solid #e5e7eb;border-radius:12px;padding:12px;background:white}
</style>
""", unsafe_allow_html=True)

# Approved inspection grouping — keep this order unchanged.
GROUPS = [
    ("🔋 Battery & Electrical", [
        "Battery", "Vent Caps", "Connector Covers", "Cables", "Battery Connectors"
    ]),
    ("🏗️ Front / Mast / Forks", [
