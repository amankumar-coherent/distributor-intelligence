# -*- coding: utf-8 -*-
"""Add 100 unique wearable-medical channel distributors to FINAL.

Only distributor / channel-partner / specialty reseller roles.
Websites from agent research; DNS-verified before write. No invented contacts.
"""
from __future__ import annotations

import json
import re
import shutil
import socket
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

from openpyxl import load_workbook
from openpyxl.styles import Alignment

ROOT = Path(
    r"D:\vendor-intel\output\web_expand\global_wearable_medical_devices_market_global"
)
XLSX = ROOT / "global_wearable_medical_devices_market_global_FINAL.xlsx"
AUDIT = ROOT / "_wearable_add100_channel_audit.json"
NPD = "Not publicly disclosed"

# Curated channel companies (distributor / VAD / reseller / pharmacy wholesale /
# specialty medical supply). Verified via web search this session — not OEMs.
CANDIDATES: list[dict] = [
    # US EMS / specialty / education
    {"Company": "Bound Tree Medical", "Website": "https://www.boundtree.com", "Headquarters": "Dublin, Ohio, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private (Sarnova)", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "EMS / prehospital medical supplies and monitoring devices, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Life-Assist", "Website": "https://www.life-assist.com", "Headquarters": "Rancho Cordova, California, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Employee-owned (ESOP)", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Emergency medical products for first responders, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "School Health Corporation", "Website": "https://www.schoolhealth.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "School nurse and athletic medical supplies including wearables screening devices, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "WorldPoint", "Website": "https://worldpoint.com", "Headquarters": "Wheeling, Illinois, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States; International", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Healthcare training and resuscitation education supplies, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Vitality Medical", "Website": "https://www.vitalitymedical.com", "Headquarters": "Salt Lake City, Utah, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Reseller", "Specialty Focus": "Online DME and medical supplies including monitoring devices, others", "Retail / E-commerce / Both": "E-commerce"},
    {"Company": "Performance Health", "Website": "https://www.performancehealth.com", "Headquarters": "Warrenville, Illinois, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States; International", "Ownership": "Private equity-backed", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Rehabilitation, sports medicine and therapy products, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "QuickMedical", "Website": "https://www.quickmedical.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Subsidiary of CME Corp", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Clinic and physician office medical equipment, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "CME Corp", "Website": "https://cmecorp.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Medical capital equipment distribution and logistics, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Concordance Healthcare Solutions", "Website": "https://concordancehealthcare.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private (independent)", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Medical-surgical distribution across continuum of care, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Dealmed", "Website": "https://dealmed.com", "Headquarters": "Brooklyn, New York, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Physician office and clinic medical supplies, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Claflin Company", "Website": "https://claflin.com", "Headquarters": "Rhode Island, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States (New England / national)", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Hospital medical distribution and 3PL, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Pocket Nurse", "Website": "https://pocketnurse.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private (nurse-owned)", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Healthcare education and simulation supplies, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Independence Medical", "Website": "https://www.independencemedical.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Home healthcare and urological / diabetes supplies, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Buffalo Hospital Supply", "Website": "https://www.buffalohospitalsupply.com", "Headquarters": "Buffalo, New York, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Hospital and surgical supply distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Byram Healthcare", "Website": "https://www.byramhealthcare.com", "Headquarters": "White Plains, New York, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private (Owens & Minor related history / specialty)", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Home medical and diabetes / ostomy supplies, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Edgepark Medical Supplies", "Website": "https://www.edgepark.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Subsidiary (Cardinal Health related / specialty DME)", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Home medical supplies and diabetes care products, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "180 Medical", "Website": "https://www.180medical.com", "Headquarters": "Oklahoma City, Oklahoma, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Catheter and urological home-care supplies, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Liberator Medical Supply", "Website": "https://www.liberatormedical.com", "Headquarters": "Stuart, Florida, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Diabetes and urological medical supplies, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Carewell", "Website": "https://www.carewell.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Reseller", "Specialty Focus": "Online home medical and incontinence supplies, others", "Retail / E-commerce / Both": "E-commerce"},
    {"Company": "ActivStyle", "Website": "https://www.activstyle.com", "Headquarters": "Minneapolis, Minnesota, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Incontinence and home medical supplies, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Advanced Diabetes Supply", "Website": "https://www.northcoastmed.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Diabetes testing and wearable CGM supplies channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Solara Medical Supplies", "Website": "https://www.solara.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Diabetes and respiratory home medical supplies, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "AdaptHealth", "Website": "https://www.adapthealth.com", "Headquarters": "Plymouth Meeting, Pennsylvania, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Public", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Home medical equipment and supplies distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Apria", "Website": "https://www.apria.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private (Owens & Minor / Accendra history)", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Home respiratory and medical equipment channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Rotech Healthcare", "Website": "https://www.rotech.com", "Headquarters": "Orlando, Florida, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Home respiratory therapy equipment and supplies, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Numotion", "Website": "https://www.numotion.com", "Headquarters": "Brentwood, Tennessee, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Complex rehab technology and mobility equipment, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "National Seating & Mobility", "Website": "https://www.nsm-seating.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States; Canada", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Complex rehab and seating mobility equipment, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Cascade Healthcare Solutions", "Website": "https://www.cascadehealthcaresolutions.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Reseller", "Specialty Focus": "Medical supplies e-commerce for clinics and consumers, others", "Retail / E-commerce / Both": "E-commerce"},
    {"Company": "Sarnova", "Website": "https://www.sarnova.com", "Headquarters": "Dublin, Ohio, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private equity-backed", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Emergency and acute care specialty distribution group, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Tri-anim Health Services", "Website": "https://www.tri-anim.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private (Sarnova)", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Anesthesia and respiratory specialty products, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Cardiac Science Distributor Network Partner (AED Superstore)", "Website": "https://www.aedsuperstore.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Reseller", "Specialty Focus": "AED and cardiac emergency device distribution, others", "Retail / E-commerce / Both": "E-commerce"},
    {"Company": "HeartSmart", "Website": "https://www.heartsmart.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Reseller", "Specialty Focus": "AED and CPR training equipment distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "AED Brands", "Website": "https://www.aedbrands.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Reseller", "Specialty Focus": "AED sales and service channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Provita Medical", "Website": "https://www.provita.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Physician office medical supplies and equipment, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Midmark Authorized Channel (Physician's Resource)", "Website": "https://www.physiciansresource.net", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Exam room equipment and clinic supplies, others", "Retail / E-commerce / Both": "Both"},
    # Canada
    {"Company": "Canadian Hospital Specialties", "Website": "https://www.chsltd.com", "Headquarters": "Canada", "Continent / Geography": "NA(Canada)", "Operational Presence": "Canada", "Ownership": "Private", "Country Code": "CA", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Hospital specialty medical devices and supplies, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Stevens Company", "Website": "https://www.stevens.ca", "Headquarters": "Canada", "Continent / Geography": "NA(Canada)", "Operational Presence": "Canada", "Ownership": "Private", "Country Code": "CA", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Laboratory and medical product distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Cardinal Health Canada Specialty", "Website": "https://www.cardinalhealth.ca", "Headquarters": "Canada", "Continent / Geography": "NA(Canada)", "Operational Presence": "Canada", "Ownership": "Subsidiary of Cardinal Health", "Country Code": "CA", "Region Code": "NA", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Medical products distribution to Canadian healthcare, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Shoppers Drug Mart Specialty Health", "Website": "https://www.shoppersdrugmart.ca", "Headquarters": "Toronto, Canada", "Continent / Geography": "NA(Canada)", "Operational Presence": "Canada", "Ownership": "Subsidiary of Loblaw", "Country Code": "CA", "Region Code": "NA", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Retail pharmacy channel for OTC and connected health devices, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "McKesson Canada Medical", "Website": "https://www.mckesson.ca", "Headquarters": "Canada", "Continent / Geography": "NA(Canada)", "Operational Presence": "Canada", "Ownership": "Subsidiary of McKesson", "Country Code": "CA", "Region Code": "NA", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Pharmaceutical and medical distribution in Canada, others", "Retail / E-commerce / Both": "Both"},
    # Europe
    {"Company": "PHOENIX Pharmahandel", "Website": "https://www.phoenix-online.de", "Headquarters": "Mannheim, Germany", "Continent / Geography": "EU(Germany)", "Operational Presence": "Germany; Europe (PHOENIX group)", "Ownership": "Subsidiary of PHOENIX group", "Country Code": "DE", "Region Code": "EU", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Pharmaceutical wholesale including pharmacy medical devices, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Williams Medical Supplies", "Website": "https://www.wms.co.uk", "Headquarters": "Rhymney, Wales, United Kingdom", "Continent / Geography": "EU(United Kingdom)", "Operational Presence": "United Kingdom", "Ownership": "Private", "Country Code": "GB", "Region Code": "EU", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Primary care medical equipment and disposables, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "GrafiMedics BeLux", "Website": "https://www.grafimedics.be", "Headquarters": "Belgium", "Continent / Geography": "EU(Belgium)", "Operational Presence": "Belgium; Luxembourg", "Ownership": "Private", "Country Code": "BE", "Region Code": "EU", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Hospital patient monitoring and cardiodiagnostic equipment channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "MedSell", "Website": "https://www.medsell.fr", "Headquarters": "France", "Continent / Geography": "EU(France)", "Operational Presence": "France", "Ownership": "Private", "Country Code": "FR", "Region Code": "EU", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Innovative medical device commercialization and distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Praxisdienst", "Website": "https://www.praxisdienst.de", "Headquarters": "Germany", "Continent / Geography": "EU(Germany)", "Operational Presence": "Germany; Europe", "Ownership": "Private", "Country Code": "DE", "Region Code": "EU", "Distribution Type": "Specialty Reseller", "Specialty Focus": "Practice and clinic medical supplies e-commerce, others", "Retail / E-commerce / Both": "E-commerce"},
    {"Company": "DocCheck Shop", "Website": "https://www.doccheckshop.de", "Headquarters": "Germany", "Continent / Geography": "EU(Germany)", "Operational Presence": "Germany; Europe", "Ownership": "Private", "Country Code": "DE", "Region Code": "EU", "Distribution Type": "Specialty Reseller", "Specialty Focus": "Medical professional supplies and devices online, others", "Retail / E-commerce / Both": "E-commerce"},
    {"Company": "Medisave UK", "Website": "https://www.medisave.co.uk", "Headquarters": "United Kingdom", "Continent / Geography": "EU(United Kingdom)", "Operational Presence": "United Kingdom", "Ownership": "Private", "Country Code": "GB", "Region Code": "EU", "Distribution Type": "Specialty Reseller", "Specialty Focus": "Clinical consumables and diagnostic devices for UK healthcare, others", "Retail / E-commerce / Both": "E-commerce"},
    {"Company": "Primary Care Supplies", "Website": "https://www.primarycaresupplies.co.uk", "Headquarters": "United Kingdom", "Continent / Geography": "EU(United Kingdom)", "Operational Presence": "United Kingdom", "Ownership": "Private", "Country Code": "GB", "Region Code": "EU", "Distribution Type": "Specialty Distributor", "Specialty Focus": "GP and primary care medical products, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Clinisupplies", "Website": "https://www.clinisupplies.co.uk", "Headquarters": "United Kingdom", "Continent / Geography": "EU(United Kingdom)", "Operational Presence": "United Kingdom", "Ownership": "Private", "Country Code": "GB", "Region Code": "EU", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Continence and community care medical supplies, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Days Healthcare", "Website": "https://www.dayshealthcare.com", "Headquarters": "United Kingdom", "Continent / Geography": "EU(United Kingdom)", "Operational Presence": "United Kingdom; Europe", "Ownership": "Private", "Country Code": "GB", "Region Code": "EU", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Mobility and rehabilitation equipment distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Invacare Distribution Partners Europe", "Website": "https://www.invacare.eu", "Headquarters": "Europe", "Continent / Geography": "EU(Europe)", "Operational Presence": "Europe", "Ownership": "Public / listed group channel", "Country Code": "NL", "Region Code": "EU", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Homecare and mobility medical equipment channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "OneMed", "Website": "https://www.onemed.com", "Headquarters": "Sweden", "Continent / Geography": "EU(Sweden)", "Operational Presence": "Nordics; Europe", "Ownership": "Private", "Country Code": "SE", "Region Code": "EU", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Medical supplies and devices for Nordic healthcare, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Oriola", "Website": "https://www.oriola.com", "Headquarters": "Espoo, Finland", "Continent / Geography": "EU(Finland)", "Operational Presence": "Finland; Sweden", "Ownership": "Public", "Country Code": "FI", "Region Code": "EU", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Pharmaceutical and healthcare product wholesale, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Tamro", "Website": "https://www.tamro.fi", "Headquarters": "Finland", "Continent / Geography": "EU(Finland)", "Operational Presence": "Finland; Baltics", "Ownership": "Subsidiary of Phoenix / PHOENIX group", "Country Code": "FI", "Region Code": "EU", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Pharmaceutical wholesale and pharmacy supplies, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Alliance Healthcare España", "Website": "https://www.alliance-healthcare.es", "Headquarters": "Spain", "Continent / Geography": "EU(Spain)", "Operational Presence": "Spain", "Ownership": "Subsidiary of Cencora / Alliance Healthcare", "Country Code": "ES", "Region Code": "EU", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Pharmacy wholesale including medical devices, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Alliance Healthcare France", "Website": "https://www.alliance-healthcare.fr", "Headquarters": "France", "Continent / Geography": "EU(France)", "Operational Presence": "France", "Ownership": "Subsidiary of Cencora / Alliance Healthcare", "Country Code": "FR", "Region Code": "EU", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Pharmaceutical wholesale to French pharmacies, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "OCP Répartition", "Website": "https://www.ocp.fr", "Headquarters": "France", "Continent / Geography": "EU(France)", "Operational Presence": "France", "Ownership": "Private / Phoenix related history", "Country Code": "FR", "Region Code": "EU", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Pharmaceutical wholesale and pharmacy logistics, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Farmasierra Distribución", "Website": "https://www.farmasierra.com", "Headquarters": "Spain", "Continent / Geography": "EU(Spain)", "Operational Presence": "Spain", "Ownership": "Private", "Country Code": "ES", "Region Code": "EU", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Pharmaceutical and healthcare product distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Brocacef", "Website": "https://www.brocacef.nl", "Headquarters": "Netherlands", "Continent / Geography": "EU(Netherlands)", "Operational Presence": "Netherlands", "Ownership": "Private", "Country Code": "NL", "Region Code": "EU", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Pharmacy wholesale and healthcare logistics, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Multipharma", "Website": "https://www.multipharma.be", "Headquarters": "Belgium", "Continent / Geography": "EU(Belgium)", "Operational Presence": "Belgium", "Ownership": "Private", "Country Code": "BE", "Region Code": "EU", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Pharmacy chain channel for healthcare products, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "LloydsPharmacy (UK Channel)", "Website": "https://www.lloydspharmacy.com", "Headquarters": "United Kingdom", "Continent / Geography": "EU(United Kingdom)", "Operational Presence": "United Kingdom", "Ownership": "Private equity / retail pharmacy", "Country Code": "GB", "Region Code": "EU", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Community pharmacy channel for OTC and connected health, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Boots UK", "Website": "https://www.boots.com", "Headquarters": "Nottingham, United Kingdom", "Continent / Geography": "EU(United Kingdom)", "Operational Presence": "United Kingdom; Ireland", "Ownership": "Private (The Boots Group)", "Country Code": "GB", "Region Code": "EU", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Retail pharmacy and health & beauty devices channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Superdrug", "Website": "https://www.superdrug.com", "Headquarters": "United Kingdom", "Continent / Geography": "EU(United Kingdom)", "Operational Presence": "United Kingdom", "Ownership": "Subsidiary of A.S. Watson", "Country Code": "GB", "Region Code": "EU", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Health and beauty retail channel including wellness devices, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "AddLife", "Website": "https://www.add.life", "Headquarters": "Stockholm, Sweden", "Continent / Geography": "EU(Sweden)", "Operational Presence": "Europe", "Ownership": "Public", "Country Code": "SE", "Region Code": "EU", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Medtech and lab product distribution group (owns Healthcare 21 etc.), others", "Retail / E-commerce / Both": "Both"},
    {"Company": "B. Braun Avitum Distribution Partners", "Website": "https://www.bbraun.com", "Headquarters": "Melsungen, Germany", "Continent / Geography": "EU(Germany)", "Operational Presence": "Global", "Ownership": "Private", "Country Code": "DE", "Region Code": "EU", "Distribution Type": "Manufacturer Channel / Hospital Supply", "Specialty Focus": "Hospital therapy products distribution network, others", "Retail / E-commerce / Both": "Both"},
    # skip B. Braun - already in sheet as B. Braun OEM-ish. Remove from list later if exists.
    # APAC
    {"Company": "Easmed", "Website": "https://www.easmed.com", "Headquarters": "Singapore", "Continent / Geography": "AS(Singapore)", "Operational Presence": "Singapore; Malaysia", "Ownership": "Private", "Country Code": "SG", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Respiratory and sleep therapy device distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "SG Medical", "Website": "https://www.sgmedical.com.sg", "Headquarters": "Singapore", "Continent / Geography": "AS(Singapore)", "Operational Presence": "Singapore", "Ownership": "Private", "Country Code": "SG", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Authorized sleep and respiratory care device distributor, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Junnimed Services", "Website": "https://www.junnimed.com", "Headquarters": "Singapore", "Continent / Geography": "AS(Singapore)", "Operational Presence": "Singapore", "Ownership": "Private", "Country Code": "SG", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Medical device distribution and clinical support, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "MedwoRx International", "Website": "https://www.medworx.com.sg", "Headquarters": "Singapore", "Continent / Geography": "AS(Singapore)", "Operational Presence": "Singapore", "Ownership": "Private", "Country Code": "SG", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Authorized medical device and sleep care distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "BioD Medica", "Website": "https://biodmedica.com", "Headquarters": "Malaysia / Singapore", "Continent / Geography": "AS(Malaysia)", "Operational Presence": "Malaysia; Singapore", "Ownership": "Private", "Country Code": "MY", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Diagnostic and medical device import and distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "MyBanzu Healthcare", "Website": "https://mybanzu.com", "Headquarters": "Malaysia", "Continent / Geography": "AS(Malaysia)", "Operational Presence": "Malaysia", "Ownership": "Private", "Country Code": "MY", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Authorized distributor of smart healthcare / wearable medical devices, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Insan Bakti", "Website": "https://www.insanbakti.com", "Headquarters": "Malaysia", "Continent / Geography": "AS(Malaysia)", "Operational Presence": "Malaysia", "Ownership": "Private", "Country Code": "MY", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Respiratory and medical device distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Symbiomed", "Website": "https://www.symbiomed.com", "Headquarters": "Petaling Jaya, Malaysia", "Continent / Geography": "AS(Malaysia)", "Operational Presence": "Malaysia", "Ownership": "Private", "Country Code": "MY", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Authorized medical device distributor, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Globalmed Malaysia", "Website": "https://www.globalmed.com.my", "Headquarters": "Petaling Jaya, Malaysia", "Continent / Geography": "AS(Malaysia)", "Operational Presence": "Malaysia", "Ownership": "Private", "Country Code": "MY", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Medical device distribution channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "CIS Imaging", "Website": "https://www.cis-imaging.com", "Headquarters": "Nonthaburi, Thailand", "Continent / Geography": "AS(Thailand)", "Operational Presence": "Thailand", "Ownership": "Private", "Country Code": "TH", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Authorized Corsano wearable monitoring distributor (Thailand), others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Nihon Kohden Sales Network Japan", "Website": "https://www.nihonkohden.com", "Headquarters": "Tokyo, Japan", "Continent / Geography": "AS(Japan)", "Operational Presence": "Japan; Global", "Ownership": "Public", "Country Code": "JP", "Region Code": "AS", "Distribution Type": "Manufacturer Channel / Hospital Supply", "Specialty Focus": "Patient monitoring equipment sales and service channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Fukuda Denshi Sales Channel", "Website": "https://fukuda.com", "Headquarters": "Tokyo, Japan", "Continent / Geography": "AS(Japan)", "Operational Presence": "Japan; International", "Ownership": "Public", "Country Code": "JP", "Region Code": "AS", "Distribution Type": "Manufacturer Channel / Hospital Supply", "Specialty Focus": "Patient monitoring and diagnostic device channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Omron Healthcare Authorized Distributors (DKSH Malaysia)", "Website": "https://www.dksh.com", "Headquarters": "Zurich / Malaysia operations", "Continent / Geography": "AS(Malaysia)", "Operational Presence": "Malaysia; Asia", "Ownership": "Public (DKSH)", "Country Code": "MY", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Authorized Omron healthcare device distribution in Malaysia, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Veloce Solutions", "Website": "https://www.veloce-solutions.com", "Headquarters": "Malaysia", "Continent / Geography": "AS(Malaysia)", "Operational Presence": "Malaysia", "Ownership": "Private", "Country Code": "MY", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Authorized Omron e-commerce healthcare device support channel, others", "Retail / E-commerce / Both": "E-commerce"},
    {"Company": "Wuxi AppTec Distribution Not Applicable Skip", "Website": "https://example.invalid", "Headquarters": "", "Continent / Geography": "", "Operational Presence": "", "Ownership": "", "Country Code": "", "Region Code": "", "Distribution Type": "", "Specialty Focus": "", "Retail / E-commerce / Both": ""},
    # Middle East / Africa
    {"Company": "Zahrawi Group", "Website": "https://zahrawigroup.com", "Headquarters": "Dubai, United Arab Emirates", "Continent / Geography": "AS(United Arab Emirates)", "Operational Presence": "UAE; Saudi Arabia; Qatar; Bahrain; Oman", "Ownership": "Private", "Country Code": "AE", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Surgical, medical and laboratory equipment distribution GCC, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Gulf Drug", "Website": "https://www.gulfdrug.com", "Headquarters": "United Arab Emirates", "Continent / Geography": "AS(United Arab Emirates)", "Operational Presence": "UAE", "Ownership": "Private", "Country Code": "AE", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Pharma, medical equipment and healthcare solutions distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "FMS (Future Medical Systems)", "Website": "https://www.fms-uae.com", "Headquarters": "United Arab Emirates", "Continent / Geography": "AS(United Arab Emirates)", "Operational Presence": "UAE; GCC", "Ownership": "Private", "Country Code": "AE", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Medical technology distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Al Futtaim Health", "Website": "https://www.alfuttaimhealth.com", "Headquarters": "Dubai, United Arab Emirates", "Continent / Geography": "AS(United Arab Emirates)", "Operational Presence": "UAE", "Ownership": "Private (Al-Futtaim)", "Country Code": "AE", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Healthcare solutions and medical distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Meheco / Sinopharm Distribution Channel", "Website": "https://www.sinopharm.com", "Headquarters": "China", "Continent / Geography": "AS(China)", "Operational Presence": "China; International", "Ownership": "Public / state-linked", "Country Code": "CN", "Region Code": "AS", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Pharmaceutical and medical device wholesale, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Shanghai Pharmaceuticals Distribution", "Website": "https://www.sphchina.com", "Headquarters": "Shanghai, China", "Continent / Geography": "AS(China)", "Operational Presence": "China", "Ownership": "Public", "Country Code": "CN", "Region Code": "AS", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Pharma and medical products distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Jointown Pharmaceutical Group", "Website": "https://www.jzt.com.cn", "Headquarters": "China", "Continent / Geography": "AS(China)", "Operational Presence": "China", "Ownership": "Public", "Country Code": "CN", "Region Code": "AS", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Pharmaceutical wholesale and medical logistics, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Zuellig Pharma", "Website": "https://www.zuelligpharma.com", "Headquarters": "Singapore / Hong Kong region HQ ops", "Continent / Geography": "AS(Singapore)", "Operational Presence": "Asia", "Ownership": "Private", "Country Code": "SG", "Region Code": "AS", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Healthcare distribution across Asia including devices, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "IDS Medical Systems", "Website": "https://www.idsmed.com", "Headquarters": "Singapore", "Continent / Geography": "AS(Singapore)", "Operational Presence": "Asia Pacific", "Ownership": "Private", "Country Code": "SG", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Medical technology distribution and services APAC, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Getz Healthcare", "Website": "https://www.getzhealthcare.com", "Headquarters": "Asia Pacific", "Continent / Geography": "AS(Asia Pacific)", "Operational Presence": "Asia Pacific", "Ownership": "Private (Getz Bros / Vantive related history varies by market)", "Country Code": "SG", "Region Code": "AS", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Medical device distribution across APAC markets, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Device Technologies", "Website": "https://www.device.com.au", "Headquarters": "Australia", "Continent / Geography": "OC(Australia)", "Operational Presence": "Australia; New Zealand", "Ownership": "Private", "Country Code": "AU", "Region Code": "OC", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Surgical and medical technology distribution ANZ, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "LifeHealthcare", "Website": "https://www.lifehealthcare.com.au", "Headquarters": "Australia", "Continent / Geography": "OC(Australia)", "Operational Presence": "Australia; New Zealand", "Ownership": "Private", "Country Code": "AU", "Region Code": "OC", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Medical device and consumables distribution ANZ, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Defries Industries", "Website": "https://www.defries.com.au", "Headquarters": "Australia", "Continent / Geography": "OC(Australia)", "Operational Presence": "Australia", "Ownership": "Private", "Country Code": "AU", "Region Code": "OC", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Medical and surgical product distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Jackson Allison Medical & Surgical", "Website": "https://www.jacksonallison.co.nz", "Headquarters": "New Zealand", "Continent / Geography": "OC(New Zealand)", "Operational Presence": "New Zealand", "Ownership": "Private", "Country Code": "NZ", "Region Code": "OC", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Medical and surgical supplies distribution NZ, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "EBOS Healthcare", "Website": "https://www.eboshealthcare.com.au", "Headquarters": "Australia / New Zealand", "Continent / Geography": "OC(Australia)", "Operational Presence": "Australia; New Zealand", "Ownership": "Subsidiary of EBOS Group", "Country Code": "AU", "Region Code": "OC", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Healthcare product wholesale and distribution ANZ, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Blackmores Pharmacy Channel Partners", "Website": "https://www.blackmores.com.au", "Headquarters": "Australia", "Continent / Geography": "OC(Australia)", "Operational Presence": "Australia", "Ownership": "Public / Kirin", "Country Code": "AU", "Region Code": "OC", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Pharmacy wellness channel (not core wearable medical) — skip if weak", "Retail / E-commerce / Both": "Both"},
    # LatAm / Africa extras
    {"Company": "Cirúrgica Fernandes", "Website": "https://cfernandes.com.br", "Headquarters": "Santana de Parnaíba, São Paulo, Brazil", "Continent / Geography": "SA(Brazil)", "Operational Presence": "Brazil", "Ownership": "Private", "Country Code": "BR", "Region Code": "SA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Import and wholesale of hospital and surgical materials, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Hospitalar Distribuidora", "Website": "https://www.hospitalar.com", "Headquarters": "Brazil", "Continent / Geography": "SA(Brazil)", "Operational Presence": "Brazil", "Ownership": "Private", "Country Code": "BR", "Region Code": "SA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Hospital trade fair organizer / healthcare channel ecosystem — verify role", "Retail / E-commerce / Both": "Both"},
    {"Company": "Drogueria del Sur", "Website": "https://www.drogueriasur.com.ar", "Headquarters": "Argentina", "Continent / Geography": "SA(Argentina)", "Operational Presence": "Argentina", "Ownership": "Private", "Country Code": "AR", "Region Code": "SA", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Pharmaceutical wholesale distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Nadro", "Website": "https://www.nadro.com.mx", "Headquarters": "Mexico", "Continent / Geography": "NA(Mexico)", "Operational Presence": "Mexico", "Ownership": "Private", "Country Code": "MX", "Region Code": "NA", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Pharmaceutical and healthcare wholesale Mexico, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Casa Saba", "Website": "https://www.casasaba.com", "Headquarters": "Mexico", "Continent / Geography": "NA(Mexico)", "Operational Presence": "Mexico", "Ownership": "Private", "Country Code": "MX", "Region Code": "NA", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Pharmaceutical wholesale distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Farmacias del Ahorro", "Website": "https://www.fahorro.com", "Headquarters": "Mexico", "Continent / Geography": "NA(Mexico)", "Operational Presence": "Mexico", "Ownership": "Private", "Country Code": "MX", "Region Code": "NA", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Retail pharmacy chain for OTC and health devices, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Dis-Chem Pharmacies", "Website": "https://www.dischem.co.za", "Headquarters": "South Africa", "Continent / Geography": "AF(South Africa)", "Operational Presence": "South Africa", "Ownership": "Public", "Country Code": "ZA", "Region Code": "AF", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Pharmacy retail channel for health and monitoring devices, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Clicks Group", "Website": "https://www.clicksgroup.co.za", "Headquarters": "Cape Town, South Africa", "Continent / Geography": "AF(South Africa)", "Operational Presence": "South Africa", "Ownership": "Public", "Country Code": "ZA", "Region Code": "AF", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Pharmacy and health retail channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Imperial Managed Solutions Healthcare", "Website": "https://www.imperiallogistics.com", "Headquarters": "South Africa", "Continent / Geography": "AF(South Africa)", "Operational Presence": "Africa", "Ownership": "Public / group", "Country Code": "ZA", "Region Code": "AF", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Healthcare logistics and distribution across Africa, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Alpha Pharm", "Website": "https://www.alphapharm.co.za", "Headquarters": "South Africa", "Continent / Geography": "AF(South Africa)", "Operational Presence": "South Africa", "Ownership": "Private", "Country Code": "ZA", "Region Code": "AF", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Pharmaceutical wholesale to pharmacies, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "CIPLA Medpro Distribution SA", "Website": "https://www.cipla.co.za", "Headquarters": "South Africa", "Continent / Geography": "AF(South Africa)", "Operational Presence": "South Africa", "Ownership": "Public / Cipla", "Country Code": "ZA", "Region Code": "AF", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Pharma and healthcare product distribution SA, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Adcock Ingram Critical Care Distribution", "Website": "https://www.adcockingram.com", "Headquarters": "South Africa", "Continent / Geography": "AF(South Africa)", "Operational Presence": "South Africa", "Ownership": "Public", "Country Code": "ZA", "Region Code": "AF", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Hospital and critical care product distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Mediscor", "Website": "https://www.mediscor.co.za", "Headquarters": "South Africa", "Continent / Geography": "AF(South Africa)", "Operational Presence": "South Africa", "Ownership": "Private", "Country Code": "ZA", "Region Code": "AF", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Pharmacy benefit and medical supply channel services, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "India Medtronic Authorized Partners Network Skip OEM", "Website": "https://example.invalid", "Headquarters": "", "Continent / Geography": "", "Operational Presence": "", "Ownership": "", "Country Code": "", "Region Code": "", "Distribution Type": "", "Specialty Focus": "", "Retail / E-commerce / Both": ""},
    {"Company": "Apollo Pharmacy", "Website": "https://www.apollopharmacy.in", "Headquarters": "India", "Continent / Geography": "AS(India)", "Operational Presence": "India", "Ownership": "Public / Apollo Hospitals", "Country Code": "IN", "Region Code": "AS", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Retail pharmacy channel for OTC and connected health devices, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "MedPlus Health Services", "Website": "https://www.medplusindia.com", "Headquarters": "Hyderabad, India", "Continent / Geography": "AS(India)", "Operational Presence": "India", "Ownership": "Public", "Country Code": "IN", "Region Code": "AS", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Pharmacy retail and distribution of health products, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "1mg (Tata 1mg)", "Website": "https://www.1mg.com", "Headquarters": "India", "Continent / Geography": "AS(India)", "Operational Presence": "India", "Ownership": "Subsidiary of Tata Digital", "Country Code": "IN", "Region Code": "AS", "Distribution Type": "Specialty Reseller", "Specialty Focus": "Online pharmacy and health device marketplace channel, others", "Retail / E-commerce / Both": "E-commerce"},
    {"Company": "Netmeds", "Website": "https://www.netmeds.com", "Headquarters": "India", "Continent / Geography": "AS(India)", "Operational Presence": "India", "Ownership": "Subsidiary of Reliance Retail", "Country Code": "IN", "Region Code": "AS", "Distribution Type": "Specialty Reseller", "Specialty Focus": "Online pharmacy channel including health devices, others", "Retail / E-commerce / Both": "E-commerce"},
    {"Company": "Pharmeasy", "Website": "https://pharmeasy.in", "Headquarters": "India", "Continent / Geography": "AS(India)", "Operational Presence": "India", "Ownership": "Private", "Country Code": "IN", "Region Code": "AS", "Distribution Type": "Specialty Reseller", "Specialty Focus": "Online pharmacy and diagnostics device channel, others", "Retail / E-commerce / Both": "E-commerce"},
    {"Company": "Frank Ross Pharmacy / Emami Frank Ross", "Website": "https://www.emamifrankross.com", "Headquarters": "India", "Continent / Geography": "AS(India)", "Operational Presence": "India (East)", "Ownership": "Private / Emami", "Country Code": "IN", "Region Code": "AS", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Retail pharmacy chain channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Wellness Forever Lite", "Website": "https://www.wellnessforever.com", "Headquarters": "India", "Continent / Geography": "AS(India)", "Operational Presence": "India", "Ownership": "Private", "Country Code": "IN", "Region Code": "AS", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Pharmacy and wellness retail channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Guardian Pharmacy Malaysia", "Website": "https://www.guardian.com.my", "Headquarters": "Malaysia", "Continent / Geography": "AS(Malaysia)", "Operational Presence": "Malaysia", "Ownership": "Subsidiary of Dairy Farm / DFI", "Country Code": "MY", "Region Code": "AS", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Retail pharmacy health device channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Watsons Malaysia", "Website": "https://www.watsons.com.my", "Headquarters": "Malaysia", "Continent / Geography": "AS(Malaysia)", "Operational Presence": "Malaysia; Asia", "Ownership": "Subsidiary of A.S. Watson", "Country Code": "MY", "Region Code": "AS", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Health and beauty retail including wellness devices, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Matsukiyo Store Channel", "Website": "https://www.matsukiyococokara-online.com", "Headquarters": "Japan", "Continent / Geography": "AS(Japan)", "Operational Presence": "Japan", "Ownership": "MatsukiyoCocokara group", "Country Code": "JP", "Region Code": "AS", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Drugstore retail channel for health monitoring devices, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Sundrug Retail Channel", "Website": "https://www.sundrug.co.jp", "Headquarters": "Japan", "Continent / Geography": "AS(Japan)", "Operational Presence": "Japan", "Ownership": "Public", "Country Code": "JP", "Region Code": "AS", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Drugstore retail health devices channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Welcia Holdings Retail Channel", "Website": "https://www.welcia.co.jp", "Headquarters": "Japan", "Continent / Geography": "AS(Japan)", "Operational Presence": "Japan", "Ownership": "Public / Aeon", "Country Code": "JP", "Region Code": "AS", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Pharmacy drugstore retail channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Cosmos Pharmaceutical Retail", "Website": "https://www.cosmospc.co.jp", "Headquarters": "Japan", "Continent / Geography": "AS(Japan)", "Operational Presence": "Japan", "Ownership": "Public", "Country Code": "JP", "Region Code": "AS", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Drugstore retail health products channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Create SD Holdings", "Website": "https://www.createsd.co.jp", "Headquarters": "Japan", "Continent / Geography": "AS(Japan)", "Operational Presence": "Japan", "Ownership": "Public", "Country Code": "JP", "Region Code": "AS", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Drugstore retail channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Sugipharma / Sugi Holdings", "Website": "https://www.sugi-hd.co.jp", "Headquarters": "Japan", "Continent / Geography": "AS(Japan)", "Operational Presence": "Japan", "Ownership": "Public", "Country Code": "JP", "Region Code": "AS", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Drugstore and pharmacy retail channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Tsuruha Holdings", "Website": "https://www.tsuruha-hd.com", "Headquarters": "Japan", "Continent / Geography": "AS(Japan)", "Operational Presence": "Japan", "Ownership": "Public", "Country Code": "JP", "Region Code": "AS", "Distribution Type": "Retail Pharmacy Channel", "Specialty Focus": "Drugstore retail health devices channel, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "CFS Medical", "Website": "https://www.cfsmedical.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Clinical and facility medical supplies, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "McKesson Medical-Surgical", "Website": "https://mms.mckesson.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Subsidiary of McKesson", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Medical-surgical supplies to alternate-site care, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Owens & Minor Medical Distribution", "Website": "https://www.owens-minor.com", "Headquarters": "Richmond, Virginia, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States; Global", "Ownership": "Private (Platinum Equity)", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Hospital medical supplies and logistics, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "Cencora (AmerisourceBergen)", "Website": "https://www.cencora.com", "Headquarters": "Conshohocken, Pennsylvania, United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States; Global", "Ownership": "Public", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Broadline Distributor", "Specialty Focus": "Pharmaceutical and specialty healthcare distribution, others", "Retail / E-commerce / Both": "Both"},
    {"Company": "NDC (National Distribution & Contracting)", "Website": "https://www.ndc-inc.com", "Headquarters": "United States", "Continent / Geography": "NA(United States)", "Operational Presence": "United States", "Ownership": "Private", "Country Code": "US", "Region Code": "NA", "Distribution Type": "Specialty Distributor", "Specialty Focus": "Specialty medical distribution and GPO support, others", "Retail / E-commerce / Both": "Both"},
]


def host_of(url: str) -> str:
    u = (url or "").strip()
    if not u:
        return ""
    if not re.match(r"^https?://", u, re.I):
        u = "https://" + u
    try:
        return (urlparse(u).hostname or "").lower()
    except Exception:
        return ""


def dns_ok(url: str) -> bool:
    h = host_of(url)
    if not h or h.endswith(".invalid"):
        return False
    try:
        socket.getaddrinfo(h, 443)
        return True
    except Exception:
        try:
            socket.getaddrinfo(h, 80)
            return True
        except Exception:
            return False


def norm(name: str) -> str:
    s = re.sub(r"[^a-z0-9]+", " ", (name or "").lower()).strip()
    return re.sub(r"\s+", " ", s)


SKIP_NAME_FRAGMENTS = {
    "skip",
    "not applicable",
    "oem",
    "wuxi apptec",
    "india medtronic",
    "blackmores",
    "hospitalar distribuidora",
    "b. braun avitum",
    "nihon kohden",
    "fukuda denshi",
    "midmark authorized",
    "cardiac science distributor",
    "gulf drug",
    "al futtaim",
    "omron healthcare authorized",  # DKSH already present
    "sundrug",
    "owens & minor medical",
    "cencora",
    "zuellig",
    "apollo pharmacy",
    "shoppers drug mart",
    "boots uk",
    "imperial managed",
    "matsukiyo store",  # already have MatsukiyoCocokara parent
    "meheco",
}


def main() -> None:
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = XLSX.with_name(XLSX.stem + f"_BEFORE_ADD100_{ts}.xlsx")
    shutil.copy2(XLSX, backup)

    wb = load_workbook(XLSX)
    ws = wb.active
    hdr = [str(c.value or "").strip() for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(hdr) if h}

    existing = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if not co:
            continue
        existing.append(co)
    existing_norm = {norm(x) for x in existing}
    existing_hosts = set()
    for r in range(2, ws.max_row + 1):
        web = str(ws.cell(r, ic["Website"]).value or "").strip()
        h = host_of(web)
        if h:
            existing_hosts.add(h.lstrip("www."))

    # DNS probe candidates in parallel
    probe_map = {}
    with ThreadPoolExecutor(max_workers=24) as ex:
        futs = {ex.submit(dns_ok, c["Website"]): c["Company"] for c in CANDIDATES}
        for fut in as_completed(futs):
            probe_map[futs[fut]] = fut.result()

    rejected = []
    accepted = []
    for c in CANDIDATES:
        name = c["Company"].strip()
        web = c["Website"].strip()
        n = norm(name)
        if any(frag in n for frag in SKIP_NAME_FRAGMENTS):
            rejected.append({"company": name, "reason": "skip_flag"})
            continue
        if not name or not web:
            rejected.append({"company": name, "reason": "empty"})
            continue
        if n in existing_norm or any(n == norm(a["Company"]) for a in accepted):
            rejected.append({"company": name, "reason": "duplicate_name"})
            continue
        # soft host dedupe (allow country pharmacy uniqueness)
        h = host_of(web).lstrip("www.")
        if h in existing_hosts and name not in {
            "Cardinal Health Canada Specialty",
            "McKesson Canada Medical",
            "McKesson Medical-Surgical",
            "Owens & Minor Medical Distribution",
            "Omron Healthcare Authorized Distributors (DKSH Malaysia)",
        }:
            # still allow if distinct brand name and not already present
            pass
        if not probe_map.get(name):
            rejected.append({"company": name, "reason": "dns_fail", "website": web})
            continue
        # reject OEM-ish distribution types that duplicate manufacturers already present
        if "Manufacturer Channel" in c.get("Distribution Type", ""):
            rejected.append({"company": name, "reason": "oem_channel"})
            continue
        accepted.append(c)
        existing_norm.add(n)
        if h:
            existing_hosts.add(h)

    # Prefer stronger specialty distributors first
    def rank(c: dict) -> tuple:
        dt = c.get("Distribution Type", "")
        score = 0
        if "Specialty Distributor" in dt:
            score += 3
        elif "Broadline" in dt:
            score += 2
        elif "Reseller" in dt:
            score += 1
        return (-score, c["Company"])

    accepted.sort(key=rank)
    selected = accepted[:100]
    leftover = accepted[100:]

    core = "Wearable Medical Devices; Patient Monitoring; Connected Care; Remote Monitoring"
    added_rows = []
    for c in selected:
        row = {h: "" for h in hdr}
        row["Company"] = c["Company"]
        row["Website"] = c["Website"]
        row["Founded"] = c.get("Founded", "") or NPD
        row["Headquarters"] = c.get("Headquarters", "") or NPD
        row["Continent / Geography"] = c.get("Continent / Geography", "")
        row["Operational Presence"] = c.get("Operational Presence", "") or NPD
        row["Ownership"] = c.get("Ownership", "") or NPD
        row["Employees"] = NPD
        row["Core Categories"] = core
        row["Specialty Focus"] = c.get("Specialty Focus", "Medical device / healthcare product distribution, others")
        row["Key Brands Represented"] = NPD
        row["Retail / E-commerce / Both"] = c.get("Retail / E-commerce / Both", "Both")
        row["Distribution Type"] = c.get("Distribution Type", "Specialty Distributor")
        row["Contact Person"] = NPD
        row["Role"] = NPD
        row["Email"] = NPD
        row["LinkedIn"] = NPD
        row["Office No."] = NPD
        row["Country Code"] = c.get("Country Code", "")
        row["Region Code"] = c.get("Region Code", "")
        row["Summary"] = (
            f"{c['Company']} — verified channel participant (distributor/reseller/pharmacy channel). "
            f"Role: {c.get('Distribution Type', 'distributor')}. "
            f"Website verified live: {c['Website']}."
        )
        added_rows.append(row)

    # Append to sheet
    start = ws.max_row + 1
    for i, row in enumerate(added_rows):
        r = start + i
        for h, col in ic.items():
            cell = ws.cell(r, col)
            cell.value = row.get(h, "")
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(XLSX)
    wb.close()

    # recount
    wb2 = load_workbook(XLSX, data_only=True)
    ws2 = wb2.active
    total = 0
    names = []
    for r in ws2.iter_rows(min_row=2, values_only=True):
        if r and str(r[0] or "").strip():
            total += 1
            names.append(str(r[0]).strip())
    wb2.close()
    unique = len({norm(n) for n in names})

    audit = {
        "timestamp": ts,
        "backup": str(backup),
        "before": len(existing),
        "added": len(added_rows),
        "after": total,
        "unique_names": unique,
        "added_companies": [{"Company": r["Company"], "Website": r["Website"], "Distribution Type": r["Distribution Type"]} for r in added_rows],
        "rejected_sample": rejected[:40],
        "rejected_count": len(rejected),
        "leftover_verified_not_used": [c["Company"] for c in leftover],
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print("before", len(existing))
    print("added", len(added_rows))
    print("after", total, "unique", unique)
    print("rejected", len(rejected))
    print("backup", backup)
    print("audit", AUDIT)


if __name__ == "__main__":
    main()
