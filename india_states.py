"""India's states and union territories, ordered for market sweeps.

A market that saturates on "India" still has plenty left state by state.
Measured on anti_drone: 96 companies, of which 55 are in Bengaluru,
because only six cities had ever been asked. The other 22 states were
never queried at all.

Ordered by where industry actually concentrates, so the productive
states are reached first and a market that stops early still gets the
best of the list. Each entry names the state plus its main industrial
cities, because the model answers "Tamil Nadu (Chennai, Coimbatore)"
better than "Tamil Nadu" alone.
"""
from __future__ import annotations

# Tier 1: the states holding most of India's drone and defence industry.
# Tier 2: real industrial bases that a city-only sweep never reaches.
# Tier 3: smaller states and UTs, included so the sweep is exhaustive
#         rather than stopping at the obvious places.
INDIA_STATES: tuple[str, ...] = (
    # Tier 1
    "Karnataka (Bengaluru, Mysuru, Hubli)",
    "Maharashtra (Mumbai, Pune, Nagpur, Nashik)",
    "Telangana (Hyderabad, Warangal)",
    "Tamil Nadu (Chennai, Coimbatore, Hosur)",
    "Delhi NCR (New Delhi, Noida, Gurugram, Faridabad)",
    "Uttar Pradesh (Lucknow, Kanpur, Noida, Ghaziabad)",
    "Gujarat (Ahmedabad, Surat, Vadodara, Rajkot)",
    "Haryana (Gurugram, Faridabad, Panchkula)",
    # Tier 2
    "West Bengal (Kolkata, Durgapur, Siliguri)",
    "Kerala (Kochi, Thiruvananthapuram, Kozhikode)",
    "Rajasthan (Jaipur, Jodhpur, Udaipur)",
    "Madhya Pradesh (Indore, Bhopal, Jabalpur)",
    "Andhra Pradesh (Visakhapatnam, Vijayawada, Tirupati)",
    "Punjab (Ludhiana, Mohali, Amritsar)",
    "Odisha (Bhubaneswar, Cuttack, Rourkela)",
    "Bihar (Patna, Muzaffarpur)",
    "Jharkhand (Ranchi, Jamshedpur, Dhanbad)",
    "Chhattisgarh (Raipur, Bhilai)",
    "Assam and the North East (Guwahati, Shillong, Imphal)",
    "Uttarakhand (Dehradun, Haridwar, Rudrapur)",
    # Tier 3
    "Himachal Pradesh (Shimla, Baddi, Solan)",
    "Goa (Panaji, Verna, Margao)",
    "Jammu and Kashmir (Jammu, Srinagar)",
    "Chandigarh and Puducherry",
    "Tripura, Manipur, Meghalaya and Nagaland",
    "Sikkim, Mizoram and Arunachal Pradesh",
)


def states(limit: int | None = None) -> tuple[str, ...]:
    """The sweep list, optionally truncated to the top `limit` states."""
    return INDIA_STATES[:limit] if limit else INDIA_STATES
