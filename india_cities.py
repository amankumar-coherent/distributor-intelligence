"""Individual Indian cities, for markets that have exhausted state-level sweeps.

A state query bundles 2-4 cities into one question: "Karnataka
(Bengaluru, Mysuru, Hubli)" gets answered mostly from Bengaluru, the same
saturation problem a country-level Dangjo query had against Vietnam. That
sweep took Dangjo's China from 1 company to 78 by asking city by city
instead of country by country -- the same escalation applies here once a
market has swept every state and stopped gaining.

Cities are the state-tier cities split apart, in the same industrial-
concentration order, so a market drilling down still hits its best
ground first.
"""
from __future__ import annotations

INDIA_CITIES: tuple[str, ...] = (
    # Tier 1 states, split into individual cities.
    "Bengaluru, Karnataka", "Mysuru, Karnataka", "Hubli, Karnataka",
    "Mumbai, Maharashtra", "Pune, Maharashtra", "Nagpur, Maharashtra",
    "Nashik, Maharashtra",
    "Hyderabad, Telangana", "Warangal, Telangana",
    "Chennai, Tamil Nadu", "Coimbatore, Tamil Nadu", "Hosur, Tamil Nadu",
    "New Delhi", "Noida, Uttar Pradesh", "Gurugram, Haryana",
    "Faridabad, Haryana",
    "Lucknow, Uttar Pradesh", "Kanpur, Uttar Pradesh",
    "Ghaziabad, Uttar Pradesh",
    "Ahmedabad, Gujarat", "Surat, Gujarat", "Vadodara, Gujarat",
    "Rajkot, Gujarat",
    "Panchkula, Haryana",
    # Tier 2/3 state capitals and industrial cities.
    "Jaipur, Rajasthan", "Jodhpur, Rajasthan", "Udaipur, Rajasthan",
    "Kolkata, West Bengal", "Bhubaneswar, Odisha",
    "Chandigarh", "Dehradun, Uttarakhand",
    "Bhopal, Madhya Pradesh", "Indore, Madhya Pradesh",
    "Patna, Bihar", "Ranchi, Jharkhand",
    "Thiruvananthapuram, Kerala", "Kochi, Kerala",
    "Visakhapatnam, Andhra Pradesh", "Vijayawada, Andhra Pradesh",
    "Guwahati, Assam", "Raipur, Chhattisgarh",
    "Amritsar, Punjab", "Ludhiana, Punjab",
)
