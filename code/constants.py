import os
from pathlib import Path

MASAUSTU = str(Path.home() / "OneDrive" / "Desktop")
if not os.path.exists(MASAUSTU):
    MASAUSTU = str(Path.home() / "Desktop")

TARGET_VKN_FOR_ZERO_VAT = "7740042326"

# XML Namespace sabitleri
XML_NS = {
    "cbc": "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
    "cac": "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
}
