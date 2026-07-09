# -*- coding: utf-8 -*-
import streamlit as st

def inject_css():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'IBM Plex Sans', sans-serif;
    }

    .stApp {
        background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 50%, #bae6fd 100%);
        color: #0f172a;
    }

    section[data-testid="stSidebar"] {
        background: rgba(224, 242, 254, 0.9) !important;
        border-right: 1px solid #93c5fd;
    }

    section[data-testid="stSidebar"] .stSelectbox label,
    section[data-testid="stSidebar"] .stFileUploader label,
    section[data-testid="stSidebar"] h2,
    section[data-testid="stSidebar"] h3 {
        color: #0369a1 !important;
        font-size: 0.82rem;
        font-weight: 600;
        letter-spacing: 0.05em;
        text-transform: uppercase;
    }

    [data-testid="metric-container"] {
        background: rgba(255, 255, 255, 0.85);
        border: 1px solid #93c5fd;
        border-radius: 8px;
        padding: 12px 16px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }

    [data-testid="metric-container"] label {
        color: #0369a1 !important;
        font-size: 0.75rem !important;
        font-weight: 600;
    }

    [data-testid="metric-container"] [data-testid="stMetricValue"] {
        color: #0284c7 !important;
        font-family: 'IBM Plex Mono', monospace;
    }

    .stTabs [data-baseweb="tab-list"] {
        background: rgba(255, 255, 255, 0.7);
        border-radius: 8px;
        padding: 4px;
        border: 1px solid #cbd5e1;
    }

    .stTabs [data-baseweb="tab"] {
        color: #475569;
        font-weight: 600;
        font-size: 0.82rem;
    }

    .stTabs [aria-selected="true"] {
        color: #0284c7 !important;
        background: #ffffff !important;
        border-radius: 6px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }

    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #0284c7, #0369a1);
        color: white;
        border: none;
        border-radius: 6px;
        font-weight: 600;
        letter-spacing: 0.04em;
        transition: all 0.2s;
        box-shadow: 0 4px 6px -1px rgba(2, 132, 199, 0.2);
    }

    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #0ea5e9, #0284c7);
        transform: translateY(-1px);
        box-shadow: 0 6px 12px rgba(2, 132, 199, 0.3);
    }

    .stDataFrame {
        border: 1px solid #93c5fd;
        border-radius: 8px;
        background: #ffffff;
    }

    h1 {
        color: #0369a1 !important;
        font-weight: 700 !important;
        letter-spacing: -0.02em !important;
    }

    h3 {
        color: #0284c7 !important;
        font-weight: 600 !important;
        font-size: 1rem !important;
    }

    .streamlit-expanderHeader {
        font-size: 0.85rem !important;
        font-weight: 600;
        color: #0369a1 !important;
    }

    /* Custom download button container matching screenshot */
    div.stDownloadButton {
        background-color: transparent !important;
        border: 1.5px solid #d53f8c !important; /* bright pink/magenta border */
        border-radius: 10px !important;
        padding: 2px !important;
        text-align: center !important;
        width: 100% !important;
    }

    div.stDownloadButton > button {
        background-color: transparent !important;
        color: #d53f8c !important; /* pink/magenta text */
        border: none !important;
        font-size: 1rem !important;
        font-weight: 600 !important;
        width: 100% !important;
        height: 46px !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        gap: 8px !important;
        transition: all 0.3s ease !important;
    }

    div.stDownloadButton > button:hover {
        color: #fbb6ce !important; /* light pink */
        background-color: rgba(213, 63, 140, 0.08) !important;
        box-shadow: 0 0 15px rgba(213, 63, 140, 0.25) !important;
        border: none !important;
    }
    </style>
    """, unsafe_allow_html=True)
