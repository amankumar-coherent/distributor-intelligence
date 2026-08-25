# -*- coding: utf-8 -*-
"""
Rewrite Core Categories + Specialty Focus for wearable FINAL sheet.
- Core Categories: real product/service category list (semicolon), NOT fake smartwatch template
- Specialty Focus: detailed functional description; NO country/region at end; no trailing ', others'
Uses website text + DeepSeek; hand overrides for high-confidence known firms.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

import httpx
from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(
    r"D:\vendor-intel\output\web_expand\global_wearable_medical_devices_market_global"
)
XLSX = ROOT / "global_wearable_medical_devices_market_global_FINAL.xlsx"
AUDIT = ROOT / "_wearable_core_specialty_audit.json"
CKPT = ROOT / "_wearable_core_specialty_ckpt.json"

load_dotenv(r"D:\vendor-intel\.env")

# Hand-verified from web/company sites (this pass)
HAND: dict[str, dict[str, str]] = {
    "AAH Pharmaceuticals": {
        "core": "Pharmaceuticals; Generics; Parallel Imports; OTC Healthcare; Health & Beauty",
        "focus": "Full-line pharmaceutical wholesaler supplying branded and generic medicines, parallel imports, OTC, and health & beauty products to community pharmacies, hospitals, and dispensing doctors, with nationwide depot fulfilment and cold-chain logistics.",
    },
    "Alliance Healthcare": {
        "core": "Pharmaceuticals; Specialty Medicines; Healthcare Distribution; Pharmacy Services",
        "focus": "Pharmaceutical wholesale and related services to pharmacies, hospitals, and healthcare providers, including specialty and temperature-controlled distribution as part of Cencora's international healthcare solutions network.",
    },
    "Cardinal Health": {
        "core": "Pharmaceuticals; Medical-Surgical Supplies; Specialty Drugs; At-Home Medical Supplies",
        "focus": "Distributes pharmaceuticals and medical-surgical products to hospitals, pharmacies, ambulatory surgery centers, and physician offices, with specialty drug distribution and at-home medical supply channels supporting continuous care and monitoring product availability.",
    },
    "Cencora": {
        "core": "Pharmaceuticals; Specialty Distribution; Global Manufacturer Services; Healthcare Logistics",
        "focus": "Global pharmaceutical sourcing and wholesale distribution with specialty and international healthcare solutions, connecting manufacturers to pharmacies, providers, and specialty channels.",
    },
    "AmerisourceBergen": {
        "core": "Pharmaceuticals; Specialty Distribution; Global Manufacturer Services; Healthcare Logistics",
        "focus": "Pharmaceutical wholesale and specialty distribution serving pharmacies and healthcare providers, with international distribution capabilities under the Cencora brand family.",
    },
    "McKesson Corporation": {
        "core": "Pharmaceuticals; Medical-Surgical Supplies; Specialty Pharmacy; Healthcare Technology",
        "focus": "Distributes pharmaceuticals and medical-surgical supplies across retail, hospital, and alternate-site channels, supporting providers with supply chain services that include devices used in monitoring and chronic care.",
    },
    "Henry Schein": {
        "core": "Dental Supplies; Medical Supplies; Practice Solutions; Animal Health",
        "focus": "Distributes dental and medical products plus practice solutions to office-based practitioners, including diagnostic and monitoring equipment used in clinical settings.",
    },
    "Best Buy": {
        "core": "Consumer Electronics; Connected Devices; Wearables; Health Tech Retail",
        "focus": "Multichannel retailer of consumer electronics and connected devices, including wearable health and wellness products sold through stores and e-commerce.",
    },
    "Best Buy Health": {
        "core": "Age Tech; Personal Emergency Response; Remote Patient Monitoring; Enabling Home Tech",
        "focus": "Provides age-tech and care-at-home solutions including wearable and home-based personal emergency response devices with fall detection, remote patient monitoring enablement, and Geek Squad Health installation/support services.",
    },
    "AddLife AB": {
        "core": "Medical Devices; Laboratory Products; MedTech Distribution; Life Science",
        "focus": "Listed MedTech group that acquires and operates specialised distributors of medical devices and laboratory products for hospitals, clinics, and laboratories across multiple therapeutic areas.",
    },
    "Healthcare 21": {
        "core": "Medical Devices; Diagnostics; Hospital Equipment; Outsourced Sales & Distribution",
        "focus": "Outsourced sales, marketing, and distribution partner for medical device manufacturers across emergency care, diagnostics, surgery, and home healthcare product portfolios.",
    },
    "Arjo": {
        "core": "Patient Handling; Hygiene Systems; Medical Devices; Pressure Injury Prevention",
        "focus": "Designs and supplies medical devices and solutions for patient handling, hygiene, and mobility in acute and long-term care, including connected care enabling products.",
    },
    "Coloplast": {
        "core": "Ostomy Care; Continence Care; Wound Care; Interventional Urology",
        "focus": "Manufactures and supplies intimate healthcare devices and consumables for ostomy, continence, wound, and urology care, distributed through healthcare channels globally.",
    },
    "B. Braun": {
        "core": "Infusion Therapy; Anesthesia; Surgical Instruments; Dialysis; Hospital Care",
        "focus": "Develops and supplies hospital care, Aesculap surgical, and Avitum extracorporeal blood treatment products, including devices and disposables used in acute and chronic care settings.",
    },
    "Grifols": {
        "core": "Plasma-Derived Medicines; Transfusion Diagnostics; Biopharma; Biosupplies",
        "focus": "Produces plasma-derived therapies and transfusion medicine diagnostics, with manufacturing and plasma collection infrastructure supporting hospital and specialty channels.",
    },
    "DKSH": {
        "core": "Market Expansion Services; Healthcare Distribution; Specialty Chemicals; Consumer Goods",
        "focus": "Provides market expansion services including healthcare product distribution, sales, and supply-chain support for manufacturers entering and scaling in Asia Pacific and other markets.",
    },
    "Bunzl": {
        "core": "Cleaning & Hygiene; Safety; Healthcare Consumables; Foodservice Supplies",
        "focus": "Specialist international distribution of cleaning, hygiene, safety, and healthcare-related consumables to business customers across multiple end markets.",
    },
    "Mediq": {
        "core": "Medical Devices; Homecare Supplies; Healthcare Distribution; Patient Aids",
        "focus": "Distributes medical devices and homecare products to healthcare providers and patients, supporting chronic care and monitoring-related supply needs.",
    },
    "Zuellig Pharma": {
        "core": "Pharmaceuticals; Medical Devices; Healthcare Logistics; Specialty Distribution",
        "focus": "Healthcare logistics and distribution partner for pharmaceuticals and medical devices across Asia, serving manufacturers, hospitals, and pharmacies.",
    },
    "Watsons": {
        "core": "Health & Beauty Retail; OTC; Personal Care; Wellness Devices",
        "focus": "Health and beauty retail chain offering OTC medicines, personal care, and consumer wellness devices including health-monitoring products in store and online.",
    },
    "Aster Pharmacy": {
        "core": "Retail Pharmacy; OTC; Healthcare Products; Wellness",
        "focus": "Retail pharmacy network dispensing medicines and healthcare products, including consumer health and monitoring devices available through pharmacy channels.",
    },
    "Boots UK": {
        "core": "Retail Pharmacy; OTC; Health & Beauty; Optical",
        "focus": "Community pharmacy and health & beauty retailer selling medicines, OTC products, and consumer health devices through stores and digital channels.",
    },
    "CVS Health": {
        "core": "Retail Pharmacy; Pharmacy Benefits; OTC; Health Services",
        "focus": "Integrated pharmacy and health services retailer distributing prescriptions, OTC products, and consumer health devices through pharmacy and digital channels.",
    },
    "Walmart": {
        "core": "Retail; Pharmacy; OTC; Consumer Electronics; Health & Wellness",
        "focus": "Mass retailer with pharmacy and health & wellness assortments that include OTC products and consumer wearable/health devices.",
    },
    "Amazon": {
        "core": "E-commerce Marketplace; Consumer Electronics; Health & Wellness; OTC",
        "focus": "Online marketplace and retail channel for consumer electronics, health & wellness products, and medical supplies sold direct and via third-party sellers.",
    },
    "eBay Inc.": {
        "core": "Online Marketplace; Consumer Electronics; Health Products",
        "focus": "Online marketplace connecting buyers and sellers of consumer electronics, health products, and related accessories including wearable devices.",
    },
    "Fastenal": {
        "core": "Industrial Fasteners; Safety; MRO Supplies; Vending Solutions",
        "focus": "Industrial distributor of fasteners, safety, and MRO supplies through branch and onsite models; healthcare-adjacent safety and facility products form part of the broader MRO mix.",
    },
    "Grainger": {
        "core": "Industrial MRO; Safety; Facility Maintenance; Tools",
        "focus": "Broadline MRO distributor serving facilities and institutions with safety, maintenance, and operational supplies used across industrial and healthcare facility environments.",
    },
    "HD Supply": {
        "core": "Facilities MRO; Multifamily Maintenance; Hospitality Supplies",
        "focus": "Wholesale MRO distributor serving multifamily, hospitality, healthcare facilities, and related institutional customers with maintenance and operational products.",
    },
    "Mouser Electronics": {
        "core": "Electronic Components; Semiconductors; Embedded Systems; Sensors",
        "focus": "Authorized distributor of electronic components and sensors used by OEMs and design engineers building connected and wearable medical/IoT hardware.",
    },
    "Arrow Electronics": {
        "core": "Electronic Components; Enterprise Computing; IoT; Design Services",
        "focus": "Global distributor of electronic components and enterprise IT solutions supporting design and supply of connected device and healthcare technology hardware.",
    },
    "Fisher Scientific": {
        "core": "Laboratory Products; Life Science; Diagnostics; Research Supplies",
        "focus": "Distributes laboratory, life science, and diagnostic products to research, clinical, and industrial laboratories.",
    },
    "Owens & Minor": {
        "core": "Medical Distribution; Surgical Supplies; Patient Care Products",
        "focus": "Healthcare logistics and medical product distribution to hospitals and providers, including patient care and procedure-related supplies.",
    },
    "Medline Industries": {
        "core": "Medical Supplies; Surgical Products; Patient Care; Homecare",
        "focus": "Manufactures and distributes medical and surgical supplies to hospitals, surgery centers, and homecare providers, including products used in patient monitoring workflows.",
    },
    "Patterson Companies": {
        "core": "Dental Supplies; Animal Health; Practice Technology",
        "focus": "Distributes dental and animal health products plus practice technology to professional customers.",
    },
    "Phoenix Group": {
        "core": "Pharmaceuticals; Pharmacy Wholesale; Healthcare Logistics",
        "focus": "European pharmaceutical wholesale group supplying pharmacies and healthcare outlets with medicines and related healthcare products.",
    },
    "Movianto": {
        "core": "Pharma Logistics; Temperature-Controlled Warehousing; Healthcare Supply Chain",
        "focus": "Specialised healthcare logistics provider offering temperature-controlled warehousing, transport, and value-added services for pharmaceuticals, biotech, and medical devices.",
    },
    "CEVA Logistics": {
        "core": "Contract Logistics; Freight Management; Healthcare Logistics",
        "focus": "Global logistics operator providing contract logistics and freight solutions, including healthcare and life-science supply-chain services.",
    },
    "Imperial Health Sciences": {
        "core": "Pharmaceutical Distribution; Healthcare Market Access; Cold Chain",
        "focus": "Healthcare logistics and market-access services distributing pharmaceuticals and related products to pharmacies, hospitals, and public-health programs.",
    },
    "Adcock Ingram": {
        "core": "Pharmaceuticals; OTC; Hospital Products; Consumer Healthcare",
        "focus": "Manufactures and markets pharmaceutical, OTC, and hospital products, supplying pharmacies and healthcare providers with branded and generic ranges including diagnostic-related lines.",
    },
    "Cipla Medpro South Africa": {
        "core": "Pharmaceuticals; Generics; Respiratory; Hospital Products",
        "focus": "Pharmaceutical manufacturing and supply of generic and specialty medicines to the South African healthcare market through wholesale and retail pharmacy channels.",
    },
    "Pharma Dynamics": {
        "core": "Generic Pharmaceuticals; Cardiovascular; Primary Care Medicines",
        "focus": "Generic pharmaceutical company focused on cardiovascular and primary-care medicines supplied through pharmacy and healthcare channels.",
    },
    "Al Nahdi Medical Company": {
        "core": "Retail Pharmacy; OTC; Health & Wellness; Personal Care",
        "focus": "Retail pharmacy chain dispensing medicines and selling OTC, wellness, and consumer health products including monitoring devices through pharmacy outlets.",
    },
    "Al-Dawaa Medical Services": {
        "core": "Retail Pharmacy; OTC; Healthcare Retail",
        "focus": "Retail pharmacy operator offering prescription dispensing plus OTC and consumer healthcare products through its pharmacy network.",
    },
    "Clicks Group": {
        "core": "Retail Pharmacy; Health & Beauty; OTC; Private Label",
        "focus": "Health, beauty, and pharmacy retailer distributing medicines, OTC products, and consumer health devices through stores and digital channels.",
    },
    "Dis-Chem Pharmacies": {
        "core": "Retail Pharmacy; OTC; Wellness; Personal Care",
        "focus": "Pharmacy and wellness retailer selling prescriptions, OTC products, and health devices through large-format stores and e-commerce.",
    },
    "Shoprite Holdings": {
        "core": "Grocery Retail; Pharmacy; OTC; Consumer Goods",
        "focus": "Food retail group with pharmacy and health product assortments available through supermarket and pharmacy formats.",
    },
    "Pague Menos": {
        "core": "Retail Pharmacy; OTC; Health & Beauty",
        "focus": "Pharmacy retail chain dispensing medicines and selling OTC and personal care products, including consumer health devices.",
    },
    "Extrafarma": {
        "core": "Retail Pharmacy; OTC; Health & Beauty",
        "focus": "Pharmacy retail banner offering medicines, OTC, and consumer healthcare products through stores and online ordering.",
    },
    "Panvel": {
        "core": "Retail Pharmacy; Pharmaceutical Wholesale; Cosmetics",
        "focus": "Pharmacy retail and pharmaceutical distribution group selling medicines, OTC, and related healthcare products through stores and wholesale arms.",
    },
    "Farmacorp": {
        "core": "Retail Pharmacy; OTC; Healthcare Retail",
        "focus": "National pharmacy chain retailing medicines and healthcare consumer products through omnichannel pharmacy outlets.",
    },
    "Goodlife Pharmacy": {
        "core": "Retail Pharmacy; OTC; Diagnostics; Telehealth Access",
        "focus": "Multi-channel pharmacy retail offering medicines, OTC, in-store diagnostics support, and digital/home-delivery pharmacy services.",
    },
    "Apotek Hjärtat": {
        "core": "Retail Pharmacy; OTC; Health & Beauty; E-pharmacy",
        "focus": "Pharmacy chain dispensing medicines and selling OTC and health products through stores and e-commerce.",
    },
    "DocMorris": {
        "core": "E-pharmacy; OTC; Prescription Medicines; Digital Health",
        "focus": "Online pharmacy and digital health platform supplying prescription and OTC medicines plus related healthcare products via e-commerce.",
    },
    "TELUS Health": {
        "core": "Digital Health; Employee Assistance; Virtual Care; Health Benefits Tech",
        "focus": "Provides digital health, virtual care, and wellbeing technology solutions for employers, insurers, and healthcare organizations.",
    },
    "Rakuten Medical": {
        "core": "Oncology Therapeutics; Photoimmunotherapy; Medical Devices",
        "focus": "Biotechnology company developing Alluminox platform photoimmunotherapy drugs and related medical devices for cancer treatment.",
    },
    "Orion Pharma": {
        "core": "Pharmaceuticals; Specialty Care; Generics; Animal Health",
        "focus": "Pharmaceutical company developing and marketing human and animal health products through healthcare and pharmacy channels.",
    },
    "Paul Hartmann AG": {
        "core": "Wound Care; Incontinence; Infection Management; Medical Consumables",
        "focus": "Manufactures and supplies medical consumables for wound care, incontinence, and infection management to hospitals and homecare channels.",
    },
    "Meditrade GmbH": {
        "core": "Medical Devices; Infusion Solutions; Hospital Supplies",
        "focus": "Supplies medical devices and pharmaceutical infusion solutions to hospitals and healthcare providers through its distribution network.",
    },
    "Omnicell": {
        "core": "Medication Management; Automated Dispensing; Pharmacy Automation",
        "focus": "Provides medication management automation and pharmacy technology for hospitals and health systems to improve dispensing accuracy and inventory control.",
    },
    "Zee Medical": {
        "core": "First Aid; Workplace Safety; Emergency Supplies",
        "focus": "Van-delivered first aid and safety supply service providing workplace first-aid cabinets, emergency products, and related safety consumables.",
    },
    "Moore Medical": {
        "core": "Medical-Surgical Supplies; Pharmaceuticals; Alternate-Site Care",
        "focus": "Distributes medical-surgical and pharmaceutical products to physician offices, clinics, EMS, and other non-hospital care settings.",
    },
    "Bound Tree Medical": {
        "core": "EMS Supplies; Prehospital Care; Monitoring Accessories",
        "focus": "Speciality distributor of emergency medical supplies and equipment for EMS and first-responder customers, including monitoring and cardiac care accessories.",
    },
    "Life-Assist": {
        "core": "EMS Supplies; Emergency Medical Products; Training Aids",
        "focus": "Distributes emergency medical products and equipment to fire, EMS, and first-responder organizations.",
    },
    "Direct Supply": {
        "core": "Senior Living Procurement; Healthcare Equipment; Facility Technology",
        "focus": "Employee-owned procurement and technology partner for senior living and long-term care, supplying equipment, e-procurement, and operational solutions.",
    },
    "Byram Healthcare Holdings": {
        "core": "Home Medical Supplies; Diabetes; Ostomy; Urology",
        "focus": "Specialty distributor of home medical supplies focused on diabetes, ostomy, urology, and related chronic-care consumables including CGM and monitoring supplies.",
    },
    "Advanced Diabetes Supply": {
        "core": "Diabetes Supplies; CGM; Insulin Delivery; DME",
        "focus": "Specialty distributor of diabetes durable medical equipment and consumables, including continuous glucose monitoring and related wearable monitoring supplies.",
    },
    "US MED": {
        "core": "Diabetes Supplies; CGM; Mail-Order DME",
        "focus": "Mail-order distributor of diabetes supplies including continuous glucose monitors and related wearable glucose-sensing products.",
    },
    "Transmedic": {
        "core": "Medical Devices; Orthopaedics; Cardiovascular; Oncology Devices",
        "focus": "Independent medical device distributor representing global manufacturers across spine, orthopaedics, cardiology, ophthalmology, and related hospital specialties.",
    },
    "Device Technologies": {
        "core": "Medical Devices; Surgical; Diagnostics; Hospital Capital Equipment",
        "focus": "Distributes medical devices and capital equipment to hospitals and clinics across multiple clinical specialties.",
    },
    "EBOS Healthcare": {
        "core": "Medical Devices; Healthcare Distribution; Pharmacy Wholesale",
        "focus": "Healthcare distribution business supplying medical devices and related products to hospitals, pharmacies, and healthcare providers.",
    },
    "IDS Medical Systems": {
        "core": "Medical Devices; Hospital Equipment; Clinical Solutions",
        "focus": "Distributes medical devices and clinical solutions to hospitals and healthcare institutions.",
    },
    "Anixter": {
        "core": "Electrical; Network Infrastructure; Security; Utility Supplies",
        "focus": "Wesco Anixter brand distributing electrical, communications, and security infrastructure products for enterprise and industrial customers.",
    },
    "Wesco International": {
        "core": "Electrical Distribution; Communications; Utility & Broadband; MRO",
        "focus": "Business-to-business distributor of electrical, communications, and utility products with supply-chain services for industrial and institutional customers.",
    },
    "CDW Corporation": {
        "core": "IT Hardware; Software; Cloud; Healthcare IT",
        "focus": "Technology solutions provider distributing IT hardware, software, and services to business, government, education, and healthcare customers.",
    },
    "SHI International Corp.": {
        "core": "IT Solutions; Software; Cloud; Hardware",
        "focus": "IT solutions provider supplying software, cloud, and hardware to enterprise and public-sector customers including healthcare organizations.",
    },
    "Insight Enterprises": {
        "core": "IT Solutions; Cloud; Hardware; Digital Transformation",
        "focus": "IT solutions integrator delivering hardware, software, and cloud services to enterprise and public-sector clients.",
    },
    "Zones, LLC": {
        "core": "IT Hardware; Software; Cloud Services",
        "focus": "IT solutions provider distributing technology products and services to business customers.",
    },
    "Newegg Business": {
        "core": "IT Hardware; Consumer Electronics; Business Technology",
        "focus": "Online retailer of IT hardware and electronics for business and consumer buyers.",
    },
    "Zoro": {
        "core": "Industrial Supplies; Tools; Safety; MRO",
        "focus": "Online industrial supplies retailer offering tools, safety, and MRO products to business customers.",
    },
    "Global Industrial": {
        "core": "Industrial Equipment; Material Handling; Safety; Storage",
        "focus": "Distributor of industrial equipment and supplies including material handling, storage, and facility products.",
    },
    "MSC Industrial Supply": {
        "core": "Metalworking; MRO; Safety; Industrial Supplies",
        "focus": "Distributes metalworking and MRO supplies to manufacturing and industrial customers.",
    },
    "Ecolab": {
        "core": "Water; Hygiene; Infection Prevention; Healthcare Cleaning",
        "focus": "Provides water, hygiene, and infection-prevention solutions and services to healthcare, food, and industrial customers.",
    },
    "Alfresa Holdings Corporation": {
        "core": "Pharmaceuticals; Medical Devices; Healthcare Wholesale",
        "focus": "Pharmaceutical and medical product wholesaler supplying pharmacies, hospitals, and healthcare facilities.",
    },
    "Suzuken Co., Ltd.": {
        "core": "Pharmaceuticals; Healthcare Wholesale; Pharmacy Support",
        "focus": "Pharmaceutical wholesaler distributing medicines and related healthcare products to pharmacies and medical institutions.",
    },
    "Medipal Holdings Corporation": {
        "core": "Pharmaceuticals; Cosmetics; Medical Devices; Wholesale",
        "focus": "Wholesale distributor of pharmaceuticals, cosmetics, and medical devices to pharmacies and healthcare providers.",
    },
    "Canon Medical Systems": {
        "core": "Diagnostic Imaging; Ultrasound; CT/MRI; Healthcare IT",
        "focus": "Supplies diagnostic imaging systems and related healthcare IT to hospitals and imaging centers.",
    },
    "Terumo Australia": {
        "core": "Medical Devices; Interventional Systems; Hospital Consumables",
        "focus": "Local arm of Terumo supplying medical devices and hospital consumables for interventional and acute care.",
    },
    "Welch Allyn Australia": {
        "core": "Patient Monitoring; Diagnostic Instruments; Frontline Care Devices",
        "focus": "Supplies Welch Allyn / Baxter frontline diagnostic and monitoring instruments for primary and acute care settings.",
    },
    "VitalConnect Asia Pacific": {
        "core": "Wearable Biosensors; Remote Patient Monitoring; Cardiac Monitoring",
        "focus": "Provides wearable biosensor patches and remote patient monitoring solutions for continuous vital-sign and cardiac monitoring.",
    },
    "Comen Medical Asia Pacific": {
        "core": "Patient Monitors; Defibrillators; Infusion; Neonatal Care",
        "focus": "Manufactures and supplies patient monitors, defibrillators, and related hospital care devices for acute and critical care.",
    },
    "Edan Instruments Asia Pacific": {
        "core": "Patient Monitoring; OB/GYN; Ultrasound; ECG",
        "focus": "Supplies patient monitoring, ECG, ultrasound, and related diagnostic devices for hospitals and clinics.",
    },
    "Spire Health": {
        "core": "Respiratory Monitoring; Wearable Sensors; Digital Health",
        "focus": "Develops wearable respiratory health sensors and remote monitoring technology for chronic respiratory care.",
    },
    "Kenya Medical Supplies Authority (KEMSA)": {
        "core": "Public Health Procurement; Pharmaceuticals; Medical Supplies",
        "focus": "State corporation that procures, warehouses, and distributes essential medicines and medical supplies to public health facilities.",
    },
    "Reliance Digital": {
        "core": "Consumer Electronics; Appliances; Wearables; Mobile",
        "focus": "Consumer electronics retailer selling smartphones, wearables, and connected devices through stores and online.",
    },
    "Bic Camera Inc.": {
        "core": "Consumer Electronics; Cameras; Appliances; Wearables",
        "focus": "Electronics retailer selling cameras, appliances, and wearable consumer devices through large-format stores.",
    },
    "MatsukiyoCocokara & Co.": {
        "core": "Drugstore Retail; OTC; Cosmetics; Health Products",
        "focus": "Drugstore chain retailing OTC medicines, cosmetics, and health products including consumer health devices.",
    },
    "Sundrug Co., Ltd.": {
        "core": "Drugstore Retail; OTC; Dispensing Pharmacy",
        "focus": "Drugstore and dispensing pharmacy retailer of medicines, OTC, and health products.",
    },
    "Tsuruha Holdings": {
        "core": "Drugstore Retail; OTC; Dispensing Pharmacy",
        "focus": "Drugstore holding company operating pharmacy and health & beauty retail formats.",
    },
    "Jean Coutu": {
        "core": "Retail Pharmacy; OTC; Health & Beauty",
        "focus": "Pharmacy retail banner dispensing medicines and selling OTC and health & beauty products.",
    },
    "Rexall Pharmacy": {
        "core": "Retail Pharmacy; OTC; Health Products",
        "focus": "Pharmacy retailer dispensing prescriptions and selling OTC and consumer health products.",
    },
    "Shoppers Drug Mart": {
        "core": "Retail Pharmacy; OTC; Beauty; Health Products",
        "focus": "Pharmacy and beauty retailer offering prescriptions, OTC, and consumer health products.",
    },
    "Guardian Pharmacy": {
        "core": "Long-Term Care Pharmacy; Institutional Dispensing; Medication Management",
        "focus": "Provides pharmacy services and medication management to long-term care and assisted living communities.",
    },
    "Practo": {
        "core": "Digital Health; Telemedicine; Practice Management; E-pharmacy Access",
        "focus": "Digital health platform for doctor discovery, teleconsultation, practice software, and related healthcare service access.",
    },
    "HealthKart": {
        "core": "Nutrition; Supplements; Fitness; Wellness E-commerce",
        "focus": "Online retailer of nutrition, sports supplements, and wellness products.",
    },
    "Netmeds": {
        "core": "E-pharmacy; OTC; Healthcare Products",
        "focus": "Online pharmacy supplying prescription and OTC medicines plus healthcare products via e-commerce.",
    },
    "PharmEasy": {
        "core": "E-pharmacy; Diagnostics; OTC; Healthcare E-commerce",
        "focus": "Online pharmacy and healthcare platform offering medicines, diagnostics booking, and OTC products.",
    },
    "MedPlus Health Services": {
        "core": "Retail Pharmacy; OTC; Diagnostics; Private Label",
        "focus": "Pharmacy retail chain dispensing medicines and selling healthcare products through stores and online.",
    },
    "Apollo Pharmacy": {
        "core": "Retail Pharmacy; OTC; Private Label; Wellness",
        "focus": "Pharmacy retail network dispensing medicines and selling OTC and wellness products including consumer health devices.",
    },
    "Tata 1mg": {
        "core": "E-pharmacy; Diagnostics; Health Content; OTC",
        "focus": "Online pharmacy and health platform offering medicines, lab tests, and health information services.",
    },
    "Wellness Warehouse": {
        "core": "Natural Health Retail; Supplements; Organic Foods; Personal Care",
        "focus": "Natural health retailer of supplements, natural foods, and wellness products through stores and online.",
    },
    "Easmed": {
        "core": "Medical Devices; Hospital Distribution; Clinical Specialties",
        "focus": "Medical device distributor serving hospitals across multiple clinical specialties in Southeast Asia.",
    },
    "Defries Industries": {
        "core": "Medical Consumables; Single-Use Devices; Hospital Supplies",
        "focus": "Manufactures and supplies single-use medical consumables for hospitals and procedural care.",
    },
    "Al-Futtaim Health": {
        "core": "Healthcare Retail; Medical Distribution; Pharmacy; Clinics",
        "focus": "Healthcare arm distributing and retailing medical products and operating healthcare retail/clinic formats within the Al-Futtaim portfolio.",
    },
    "Alloga": {
        "core": "Healthcare Logistics; Temperature-Controlled Distribution; Pharma 3PL",
        "focus": "Healthcare logistics specialist providing warehousing and distribution services for pharmaceuticals and healthcare products.",
    },
    "Amazon Pharmacy / Care": {
        "core": "E-pharmacy; OTC; Healthcare E-commerce",
        "focus": "Online pharmacy and healthcare retail services delivering medicines and health products digitally.",
    },
    # --- Pass 2: replace weak generic fallouts with verified cores ---
    "CJ ENM": {
        "core": "Media & Entertainment; TV Home Shopping; Live Commerce; Wellness Retail",
        "focus": "Operates entertainment and commerce platforms including CJ ONSTYLE home shopping and live commerce, curating fashion, beauty, living, and wellness/health-functional products for consumers.",
    },
    "Hyundai Home Shopping": {
        "core": "TV Home Shopping; E-commerce; Beauty Devices; Health & Wellness Retail",
        "focus": "Multichannel home shopping and Hmall e-commerce retailer of consumer goods, including beauty devices, wellness items, and selected medical-device-class functional products sold via TV and online broadcasts.",
    },
    "LGU+": {
        "core": "Mobile Telecommunications; IoT Services; Senior Care Platforms; Connected Devices",
        "focus": "Telecom operator offering mobile and IoT services, including senior-care monitoring solutions such as radar-based fall detection and bundled consumer wearable device plans with partner brands.",
    },
    "Itochu Corporation": {
        "core": "Trading Company; Medical Device Import & Sales; Clinical Development Support; Pharma Market Access",
        "focus": "Sogo-shosha with life and healthcare businesses spanning medical device import and sales via Century Medical, CRO and post-marketing support via A2 Healthcare, and digital healthcare platform investments.",
    },
    "Marubeni Corporation": {
        "core": "Trading Company; Pharmaceutical Distribution; Medical Devices; Medical Services",
        "focus": "Sogo-shosha developing healthcare businesses including pharmaceutical sales and distribution networks, medical device and consumables investments, and medical services such as clinical testing and telemedicine.",
    },
    "Mitsui & Co.": {
        "core": "Trading Company; Healthcare Investments; Medical Infrastructure; Wellness",
        "focus": "Diversified trading company with healthcare-related investments and businesses spanning medical infrastructure, services, and wellness value chains.",
    },
    "Walgreens": {
        "core": "Retail Pharmacy; OTC Healthcare; Health & Wellness; Specialty Pharmacy",
        "focus": "Retail drugstore chain dispensing prescriptions and selling OTC healthcare, wellness, and daily living products, with specialty and home-delivery pharmacy services.",
    },
    "Rite Aid": {
        "core": "Retail Pharmacy; OTC Healthcare; Health & Beauty",
        "focus": "Retail pharmacy chain dispensing medicines and selling OTC healthcare and personal care products through stores and digital channels.",
    },
    "Superdrug": {
        "core": "Health & Beauty Retail; OTC Healthcare; Personal Care; Pharmacy",
        "focus": "Health and beauty retailer offering OTC medicines, personal care, and pharmacy services where available, including consumer health products.",
    },
    "Zur Rose Group": {
        "core": "E-pharmacy; Online Healthcare Marketplace; OTC; Prescription Medicines",
        "focus": "European online pharmacy and healthcare platform group (including DocMorris) supplying prescription and OTC medicines plus health products via e-commerce and related digital services.",
    },
    "Noweda": {
        "core": "Pharmaceuticals; OTC Healthcare; Pharmacy Wholesale",
        "focus": "Independent pharmaceutical wholesaler supplying full-line medicines and healthcare products to community pharmacies with multi-daily depot logistics.",
    },
    "OCP Répartition": {
        "core": "Pharmaceuticals; Pharmacy Services; Healthcare Distribution",
        "focus": "Full-line pharmaceutical distributor supplying medicines and pharmacy support services to community pharmacies and healthcare establishments.",
    },
    "Sanacorp Pharmahandel": {
        "core": "Pharmaceuticals; OTC Healthcare; Pharmacy Cooperative Wholesale",
        "focus": "Pharmacy-owned cooperative pharmaceutical wholesaler delivering full-line medicines and healthcare products to member community pharmacies with frequent depot tours.",
    },
    "Nadro": {
        "core": "Pharmaceuticals; Health & Personal Care; Hospital Distribution; Pharmacy Logistics",
        "focus": "Pharmaceutical and personal-care logistics distributor supplying medicines and related products to pharmacies, hospitals, and clinics through regional distribution centers.",
    },
    "VWR (Avantor)": {
        "core": "Laboratory Supplies; Life Science Chemicals; Lab Equipment; Production Materials",
        "focus": "Global laboratory products brand within Avantor distributing chemicals, consumables, equipment, and services to research, healthcare, and advanced manufacturing laboratories.",
    },
    "Toho Holdings Co., Ltd.": {
        "core": "Pharmaceutical Wholesale; Dispensing Pharmacy; Medical Equipment; OTC Drugs",
        "focus": "Healthcare group engaged in pharmaceutical wholesaling, dispensing pharmacy operations, and related medical and OTC product distribution to medical institutions and pharmacies.",
    },
    "Alliance Healthcare España": {
        "core": "Pharmaceuticals; Specialty Medicines; Healthcare Distribution; Pharmacy Services",
        "focus": "Pharmaceutical wholesale and related services to pharmacies and healthcare providers within the Cencora Alliance Healthcare network.",
    },
    "Alliance Healthcare France": {
        "core": "Pharmaceuticals; Specialty Medicines; Healthcare Distribution; Pharmacy Services",
        "focus": "Pharmaceutical wholesale and related services to pharmacies and healthcare providers within the Cencora Alliance Healthcare network.",
    },
    "McKesson Canada": {
        "core": "Pharmaceuticals; Medical-Surgical Supplies; Specialty Pharmacy; Healthcare Technology",
        "focus": "Distributes pharmaceuticals and medical-surgical products to retail, hospital, and alternate-site customers, with specialty pharmacy and supply-chain services.",
    },
    "McKesson Medical-Surgical": {
        "core": "Medical-Surgical Supplies; Physician Office Products; Alternate-Site Distribution",
        "focus": "Distributes medical-surgical supplies and related products to physician offices, surgery centers, and alternate-site care providers.",
    },
    "Cardinal Health Canada": {
        "core": "Pharmaceuticals; Medical-Surgical Supplies; Specialty Distribution",
        "focus": "Distributes pharmaceuticals and medical-surgical products to hospitals, pharmacies, and healthcare providers.",
    },
    "Owens & Minor Distribution": {
        "core": "Medical-Surgical Distribution; Hospital Supply Chain; Procedural Products",
        "focus": "Distributes medical and surgical products and provides supply-chain services to hospitals and healthcare providers.",
    },
    "Concordance Healthcare Solutions": {
        "core": "Medical-Surgical Distribution; Physician Office Supply; Acute Care Products",
        "focus": "Distributes medical-surgical products to acute care, physician office, and alternate-site customers.",
    },
    "PHOENIX Pharmahandel": {
        "core": "Pharmaceuticals; OTC Healthcare; Pharmacy Services; Healthcare Logistics",
        "focus": "Pharmaceutical wholesale arm supplying medicines and related services to pharmacies and healthcare customers within the PHOENIX group.",
    },
    "Oriola": {
        "core": "Pharmaceuticals; Healthcare Logistics; Pharmacy Services; Consumer Health",
        "focus": "Pharmaceutical distribution and healthcare logistics company serving pharmacies, hospitals, and manufacturers with wholesale and related services.",
    },
    "Tamro": {
        "core": "Pharmaceuticals; Healthcare Logistics; Pharmacy Wholesale",
        "focus": "Pharmaceutical wholesaler distributing medicines and healthcare products to pharmacies and healthcare providers.",
    },
    "Brocacef": {
        "core": "Pharmaceuticals; Specialty Medicines; Healthcare Distribution",
        "focus": "Pharmaceutical wholesaler supplying medicines and related healthcare products to pharmacies and healthcare institutions.",
    },
    "Alpha Pharm": {
        "core": "Pharmaceuticals; Generics; OTC Healthcare; Pharmacy Wholesale",
        "focus": "Pharmaceutical wholesaler and pharmacy support group supplying medicines and related products to independent pharmacies.",
    },
    "Shanghai Pharmaceuticals Distribution": {
        "core": "Pharmaceuticals; Medical Devices; Healthcare Logistics; Hospital Supply",
        "focus": "Pharmaceutical and medical product distribution business supplying hospitals, pharmacies, and healthcare channels with medicines and related products.",
    },
    "Pharmaprix": {
        "core": "Retail Pharmacy; OTC Healthcare; Health & Beauty",
        "focus": "Pharmacy retail banner dispensing prescriptions and selling OTC healthcare and beauty products.",
    },
    "PharmaChoice": {
        "core": "Independent Pharmacy Banner; OTC Healthcare; Pharmacy Services",
        "focus": "Independent pharmacy banner providing branding, procurement, and retail support for community pharmacies dispensing medicines and OTC products.",
    },
    "Rakuten Group": {
        "core": "E-commerce Marketplace; Digital Services; Consumer Retail; Fintech",
        "focus": "Internet services group operating large e-commerce marketplaces and digital platforms where third-party and first-party merchants sell consumer goods including health and wearable products.",
    },
    "Saudi Pharmaceutical Industries & Medical Appliances Corporation (SPIMACO)": {
        "core": "Pharmaceuticals; Generics; Medical Appliances; Hospital Products",
        "focus": "Develops, manufactures, and markets pharmaceutical products and medical appliances for hospital and retail pharmacy channels.",
    },
    "Rehabmart": {
        "core": "Rehabilitation Equipment; Mobility Aids; Homecare DME; Assistive Technology",
        "focus": "Online retailer of rehabilitation, mobility, and durable medical equipment for patients, caregivers, and clinical customers.",
    },
    "Medshop Australia": {
        "core": "Medical Supplies; First Aid; Homecare Products; Clinical Consumables",
        "focus": "Online medical supply retailer selling first aid, clinical consumables, and homecare products to consumers and professionals.",
    },
    "Medisave UK": {
        "core": "Medical Supplies; Clinical Consumables; Homecare Products; Diagnostics",
        "focus": "Online medical supplier selling clinical consumables, diagnostic aids, and related products to healthcare professionals and consumers.",
    },
    "Tivity Health": {
        "core": "Fitness Programs; Social Engagement; Healthy Aging; Wellness Benefits",
        "focus": "Provides healthy-aging and wellness programs such as fitness center networks and social engagement services for health plans and seniors.",
    },
    "Drogueria del Sur": {
        "core": "Pharmaceuticals; OTC Healthcare; Pharmacy Wholesale",
        "focus": "Pharmaceutical wholesaler distributing medicines and healthcare products to pharmacies and healthcare customers.",
    },
}

UA = {"User-Agent": "Mozilla/5.0 (compatible; VendorIntelCore/1.1)"}

GEO_WORDS = re.compile(
    r"\b("
    r"united states|usa|u\.s\.a?|united kingdom|u\.k\.|england|scotland|wales|"
    r"saudi arabia|uae|united arab emirates|qatar|kuwait|bahrain|oman|jordan|iraq|"
    r"south africa|namibia|botswana|kenya|nigeria|ghana|egypt|mauritius|"
    r"india|china|japan|singapore|malaysia|thailand|indonesia|philippines|vietnam|"
    r"hong kong|taiwan|south korea|korea|australia|new zealand|"
    r"canada|mexico|brazil|chile|colombia|peru|argentina|bolivia|guatemala|honduras|"
    r"germany|france|spain|italy|netherlands|belgium|sweden|norway|denmark|finland|"
    r"ireland|switzerland|austria|poland|portugal|romania|turkey|lithuania|czech|"
    r"middle east|gcc|nordic|asia pacific|apac|latam|europe|africa|asia|"
    r"nationwide|globally|worldwide|internationally"
    r")\b",
    re.I,
)


def strip_html(html: str) -> str:
    html = re.sub(r"(?is)<script[^>]*>.*?</script>", " ", html)
    html = re.sub(r"(?is)<style[^>]*>.*?</style>", " ", html)
    html = re.sub(r"(?is)<[^>]+>", " ", html)
    return re.sub(r"\s+", " ", html).strip()


def fetch(url: str) -> str:
    try:
        with httpx.Client(follow_redirects=True, timeout=18.0, headers=UA) as c:
            r = c.get(url)
            if r.status_code >= 400:
                return ""
            return strip_html(r.text)[:9000]
    except Exception:
        return ""


def candidate_urls(website: str) -> list[str]:
    base = (website or "").strip().rstrip("/")
    if not base:
        return []
    if not base.startswith("http"):
        base = "https://" + base
    paths = ["", "/about", "/about-us", "/company", "/products", "/services", "/en/about"]
    out = []
    for p in paths:
        u = base if p == "" else urljoin(base + "/", p.lstrip("/"))
        if u not in out:
            out.append(u)
    return out[:5]


def clean_specialty(text: str) -> str:
    t = re.sub(r"\s+", " ", (text or "").strip())
    if not t:
        return ""
    # remove trailing ", others" / "; others"
    t = re.sub(r"[,;]?\s*others\.?$", "", t, flags=re.I).strip()
    # remove trailing geographic clauses
    t = re.sub(
        r"\s*(?:across|throughout|in|within|for|serving|covering)\s+(?:the\s+)?["
        r"A-Za-z ,&\-]{2,60}\.?$",
        "",
        t,
        flags=re.I,
    ).strip()
    # strip leftover geo words at end after punctuation cleanup
    t = re.sub(r"[,\s]+$", "", t)
    if len(t) > 420:
        t = t[:417].rsplit(" ", 1)[0] + "..."
    if t and t[-1] not in ".!?":
        t += "."
    # Final geo scrub of common endings like "in Saudi Arabia."
    t = re.sub(
        r"\s+(?:in|across|throughout)\s+[A-Z][A-Za-z]*(?:\s+[A-Z][A-Za-z]*){0,3}\.$",
        ".",
        t,
    )
    # Strip bare trailing country/region tokens e.g. "... wholesale Mexico."
    t = re.sub(
        r"\s+(?:United States|United Kingdom|Saudi Arabia|South Africa|United Arab Emirates|"
        r"USA|U\.?S\.?A?\.?|UK|UAE|Mexico|Canada|India|China|Japan|Brazil|Germany|France|"
        r"Spain|Italy|Australia|Korea|Singapore|Malaysia|Thailand|Indonesia|Egypt|Nigeria|"
        r"Kenya|Chile|Colombia|Argentina|Poland|Netherlands|Belgium|Sweden|Switzerland|"
        r"Europe|Asia|Africa|Middle East|Latin America|North America|APAC|EMEA|LATAM)\.?$",
        ".",
        t,
        flags=re.I,
    )
    t = re.sub(r"\.\.+$", ".", t)
    return t.strip()


def clean_core(text: str) -> str:
    t = re.sub(r"\s+", " ", (text or "").strip())
    t = t.replace("|", ";")
    parts = [p.strip(" .;") for p in re.split(r"[;•|/]+", t) if p.strip(" .;")]
    # drop fake smartwatch-only template if mixed incorrectly handled elsewhere
    out = []
    seen = set()
    for p in parts:
        key = p.lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(p)
    return "; ".join(out[:8])


WEAK_CORES = {
    "pharmaceuticals; medical supplies; healthcare distribution",
    "medical devices; specialty healthcare products; clinical supplies",
    "medical supplies; healthcare products; specialty retail",
}


def is_bad_core(core: str) -> bool:
    c = (core or "").strip().lower()
    if not c:
        return True
    if c.startswith("smartwatches; wearables; fitness trackers"):
        return True
    if c == "wearable medical devices; patient monitoring; connected care; remote monitoring":
        return True
    if c in WEAK_CORES:
        return True
    return False


def deepseek_extract(company: str, website: str, dist: str, text: str) -> dict:
    key = (os.getenv("OPENAI_API_KEY") or os.getenv("DEEPSEEK_API_KEY") or "").strip()
    if not key:
        return {"core_categories": "", "specialty_focus": "", "confidence": "none", "reason": "no_key"}
    base = (os.getenv("OPENAI_BASE_URL") or "https://api.deepseek.com/v1").rstrip("/")
    model = (os.getenv("OPENAI_MODEL") or "deepseek-chat").strip()
    system = (
        "You extract Core Categories and Specialty Focus for a healthcare CHANNEL company "
        "(distributor, pharmacy, retailer, logistics, reseller). Use ONLY the evidence.\n"
        "Return JSON: {core_categories:string, specialty_focus:string, confidence:high|medium|low|none, reason:string}.\n"
        "core_categories: 3-6 real product/service categories separated by '; '. "
        "Examples: Pharmaceuticals; OTC Healthcare; Medical Devices; Patient Monitoring; "
        "Retail Pharmacy; Hospital Supplies; Homecare DME; Healthcare Logistics; "
        "Consumer Electronics; Industrial MRO. NEVER invent smartwatches unless evidence says so.\n"
        "specialty_focus: 1-2 detailed sentences describing the company's FUNCTION "
        "(what they distribute/sell/operate and how it relates to medical/health products). "
        "Do NOT include any country, city, or region names. Do NOT end with 'others'. "
        "Do NOT invent wearable medical focus if the company is only a pharma wholesaler — "
        "describe their real role accurately; mention monitoring/wearable devices only if evidenced."
    )
    user = (
        f"Company: {company}\nWebsite: {website}\nDistribution Type cell: {dist}\n\nEvidence:\n{text[:8500]}"
    )
    try:
        with httpx.Client(timeout=70.0) as c:
            r = c.post(
                f"{base}/chat/completions",
                headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
                json={
                    "model": model,
                    "temperature": 0,
                    "messages": [
                        {"role": "system", "content": system},
                        {"role": "user", "content": user},
                    ],
                    "response_format": {"type": "json_object"},
                },
            )
            r.raise_for_status()
            return json.loads(r.json()["choices"][0]["message"]["content"])
    except Exception as e:
        return {"core_categories": "", "specialty_focus": "", "confidence": "none", "reason": str(e)}


def apply_format(ws) -> None:
    headers = [c.value for c in ws[1]]
    widths = {
        "Company": 34,
        "Website": 36,
        "Founded": 10,
        "Headquarters": 28,
        "Continent / Geography": 52,
        "Operational Presence": 48,
        "Ownership": 42,
        "Core Categories": 40,
        "Specialty Focus": 55,
        "Summary": 48,
    }
    thin = Border(
        left=Side(style="thin", color="D0D5DD"),
        right=Side(style="thin", color="D0D5DD"),
        top=Side(style="thin", color="D0D5DD"),
        bottom=Side(style="thin", color="D0D5DD"),
    )
    header_fill = PatternFill("solid", fgColor="1F4E79")
    for i, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths.get(str(h or ""), 16)
    for i in range(1, len(headers) + 1):
        cell = ws.cell(1, i)
        cell.fill = header_fill
        cell.font = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = thin
    for r in range(2, ws.max_row + 1):
        for i in range(1, len(headers) + 1):
            cell = ws.cell(r, i)
            cell.font = Font(name="Calibri", size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
            cell.border = thin
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"


def fallback_from_dist(dist: str, company: str) -> tuple[str, str]:
    d = (dist or "").lower()
    if "retail pharmacy" in d:
        return (
            "Retail Pharmacy; OTC Healthcare; Health & Beauty",
            "Retail pharmacy channel dispensing medicines and selling OTC and consumer healthcare products.",
        )
    if "specialty reseller" in d or "e-commerce" in d:
        return (
            "Medical Supplies; Healthcare Products; Specialty Retail",
            "Specialty reseller of medical and healthcare products through focused retail or e-commerce channels.",
        )
    if "specialty distributor" in d:
        return (
            "Medical Devices; Specialty Healthcare Products; Clinical Supplies",
            "Specialty distributor supplying medical devices and related clinical products to healthcare providers.",
        )
    # broadline default — still better than smartwatches
    return (
        "Pharmaceuticals; Medical Supplies; Healthcare Distribution",
        "Broadline healthcare distributor supplying medicines and medical products to pharmacies and healthcare providers.",
    )


def main() -> None:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = ROOT / f"global_wearable_medical_devices_market_global_FINAL_BEFORE_CORE_SPEC_{stamp}.xlsx"
    shutil.copy2(XLSX, bak)

    wb = load_workbook(XLSX)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    ic = {h: i + 1 for i, h in enumerate(headers) if h}

    ckpt = {"done": {}}
    if CKPT.exists():
        try:
            ckpt = json.loads(CKPT.read_text(encoding="utf-8"))
        except Exception:
            pass

    hand_n = 0
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if co not in HAND:
            continue
        info = HAND[co]
        ws.cell(r, ic["Core Categories"]).value = clean_core(info["core"])
        ws.cell(r, ic["Specialty Focus"]).value = clean_specialty(info["focus"])
        hand_n += 1

    # targets: not in HAND, or still bad core
    targets = []
    for r in range(2, ws.max_row + 1):
        co = str(ws.cell(r, ic["Company"]).value or "").strip()
        if co in HAND:
            continue
        core = str(ws.cell(r, ic["Core Categories"]).value or "")
        focus = str(ws.cell(r, ic["Specialty Focus"]).value or "")
        if is_bad_core(core) or GEO_WORDS.search(focus) or focus.lower().rstrip(".").endswith("others"):
            targets.append(r)

    MAX = int(os.getenv("CORE_EXPAND_MAX", "120"))
    print("hand", hand_n, "targets", len(targets), "max", MAX, flush=True)

    updated = []
    processed = 0
    try:
        for r in targets:
            if processed >= MAX:
                break
            co = str(ws.cell(r, ic["Company"]).value or "").strip()
            cur_core = str(ws.cell(r, ic["Core Categories"]).value or "")
            # Always rework weak template cores even if prior checkpoint applied a fallback
            if (
                co in ckpt.get("done", {})
                and ckpt["done"][co].get("applied")
                and not is_bad_core(cur_core)
            ):
                print("skip ckpt", co, flush=True)
                continue
            processed += 1
            website = str(ws.cell(r, ic["Website"]).value or "").strip()
            dist = str(ws.cell(r, ic.get("Distribution Type", 0)).value or "").strip() if "Distribution Type" in ic else ""
            snippets = []
            for u in candidate_urls(website):
                t = fetch(u)
                if len(t) > 200:
                    snippets.append(f"URL:{u}\n{t[:3000]}")
                time.sleep(0.12)
            text = "\n\n---\n\n".join(snippets) if snippets else f"(no site text) DistributionType={dist}"
            ds = deepseek_extract(co, website, dist, text)
            conf = str(ds.get("confidence") or "").lower()
            core = clean_core(str(ds.get("core_categories") or ""))
            focus = clean_specialty(str(ds.get("specialty_focus") or ""))
            applied = False
            if conf in {"high", "medium"} and core and focus and not is_bad_core(core):
                ws.cell(r, ic["Core Categories"]).value = core
                ws.cell(r, ic["Specialty Focus"]).value = focus
                applied = True
                updated.append({"company": co, "core": core, "focus": focus, "conf": conf})
                print("UPDATED", co, "=>", core[:60], "|", focus[:70], flush=True)
            else:
                # conservative fallback from distribution type only (still better than smartwatches)
                fc, ff = fallback_from_dist(dist, co)
                ws.cell(r, ic["Core Categories"]).value = fc
                # keep existing focus if decent and geo-cleaned; else fallback
                old = clean_specialty(str(ws.cell(r, ic["Specialty Focus"]).value or ""))
                if old and len(old) > 40 and not GEO_WORDS.search(old):
                    ws.cell(r, ic["Specialty Focus"]).value = old
                else:
                    ws.cell(r, ic["Specialty Focus"]).value = clean_specialty(ff)
                applied = True
                updated.append({"company": co, "core": fc, "focus": ws.cell(r, ic["Specialty Focus"]).value, "conf": conf or "fallback"})
                print("FALLBACK", co, conf, "=>", fc, flush=True)

            ckpt.setdefault("done", {})[co] = {
                "applied": applied,
                "deepseek": ds,
                "core": ws.cell(r, ic["Core Categories"]).value,
                "focus": ws.cell(r, ic["Specialty Focus"]).value,
            }
            if processed % 5 == 0:
                CKPT.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")
                wb.save(XLSX)
                print("checkpoint", processed, flush=True)
    finally:
        # geo scrub all specialty cells
        for r in range(2, ws.max_row + 1):
            sp = clean_specialty(str(ws.cell(r, ic["Specialty Focus"]).value or ""))
            ws.cell(r, ic["Specialty Focus"]).value = sp
            # rewrite remaining bad cores with fallback
            core = str(ws.cell(r, ic["Core Categories"]).value or "")
            if is_bad_core(core):
                dist = str(ws.cell(r, ic["Distribution Type"]).value or "") if "Distribution Type" in ic else ""
                fc, ff = fallback_from_dist(dist, "")
                ws.cell(r, ic["Core Categories"]).value = fc
                if not sp or len(sp) < 30:
                    ws.cell(r, ic["Specialty Focus"]).value = clean_specialty(ff)
        apply_format(ws)
        wb.save(XLSX)
        CKPT.write_text(json.dumps(ckpt, indent=2, ensure_ascii=False), encoding="utf-8")

    # recount bad cores
    bad = geo = 0
    samples = []
    for r in range(2, ws.max_row + 1):
        core = str(ws.cell(r, ic["Core Categories"]).value or "")
        focus = str(ws.cell(r, ic["Specialty Focus"]).value or "")
        if is_bad_core(core):
            bad += 1
        if GEO_WORDS.search(focus):
            geo += 1
        if len(samples) < 8:
            samples.append(
                {
                    "company": ws.cell(r, ic["Company"]).value,
                    "core": core,
                    "focus": focus,
                }
            )

    audit = {
        "backup": bak.name,
        "hand": hand_n,
        "updated_this_run": updated,
        "processed": processed,
        "remaining_targets_est": max(0, len(targets) - processed),
        "bad_core_left": bad,
        "specialty_still_geoish": geo,
        "samples": samples,
    }
    AUDIT.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print("DONE bad_core", bad, "geoish_focus", geo, "updated", len(updated))


if __name__ == "__main__":
    main()
