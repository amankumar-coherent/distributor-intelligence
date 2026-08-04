"""Reject SEO page titles / platform brands mistaken for companies."""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from vendor_intel.pipeline.entity_gate import reject_bad_company_name, reject_reason


class TestSeoNameGate(unittest.TestCase):
    def test_rejects_platform_brand(self) -> None:
        self.assertEqual(
            reject_bad_company_name("Android", "android.com"),
            "platform_brand_not_company",
        )
        self.assertEqual(
            reject_reason("Android", "android.com"),
            "platform_brand_not_company",
        )

    def test_rejects_suppliers_in_geo_title(self) -> None:
        self.assertEqual(
            reject_bad_company_name(
                "Smartwatch Suppliers in Dubai", "whitecloudintl.ae"
            ),
            "seo_page_title",
        )

    def test_rejects_bare_product_label(self) -> None:
        self.assertEqual(
            reject_bad_company_name("Smartwatch", "smartwatch.partnermobile.com"),
            "generic_product_name",
        )

    def test_rejects_wholesale_listing_title(self) -> None:
        self.assertEqual(
            reject_bad_company_name(
                "Custom Wholesale Chain Smartwatch", "royalpeakdg.com"
            ),
            "seo_page_title",
        )

    def test_rejects_nav_page_support(self) -> None:
        self.assertEqual(
            reject_bad_company_name("Support", "support.paloaltonetworks.com"),
            "nav_page_not_company",
        )
        self.assertEqual(
            reject_reason("Support", "support.paloaltonetworks.com"),
            "nav_page_not_company",
        )

    def test_rejects_portal_subdomain(self) -> None:
        self.assertEqual(
            reject_bad_company_name("Palo Alto Networks", "support.paloaltonetworks.com"),
            "vendor_portal_subdomain",
        )
        self.assertEqual(
            reject_bad_company_name("Help Desk", "help.fortinet.com"),
            "vendor_portal_subdomain",
        )

    def test_rejects_product_support_title(self) -> None:
        self.assertEqual(
            reject_bad_company_name(
                "Software Firewall Platinum Support", "paloaltonetworks.sg"
            ),
            "seo_page_title",
        )

    def test_keeps_real_distributors(self) -> None:
        for name, domain in (
            ("Ingram Micro", "ingrammicro.com"),
            ("D&H Distributing", "dandh.com"),
            ("White Cloud International", "whitecloudintl.ae"),
            ("Smartwatchsolutions", "smartwatchsolutions.com"),
            ("Zesmartwatches", "zesmartwatches.com"),
            ("Mouser Electronics", "mouser.com"),
            ("Support Logistics Ltd", "supportlogistics.com"),
        ):
            self.assertIsNone(
                reject_bad_company_name(name, domain),
                msg=f"false positive: {name}",
            )


if __name__ == "__main__":
    unittest.main()
