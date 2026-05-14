def normalize_sku(name):
    """Normalize product name to match SKU variant names."""
    n = str(name).lower()
    mapping = {
        "relaxing naturals": "Relaxing Naturals",
        "fresh bamboo": "Fresh Bamboo",
        "white tea": "White Tea",
        "cotton clean": "Cotton Clean",
        "quiet elegance": "Quiet Elegance",
        "charming home": "Charming Home",
        "bright mornings": "Bright Mornings",
        "mountain breeze": "Mountain Breeze",
        "beach escape": "Beach Escape",
        "candy cane": "Candy Cane",
        "apple cinnamon": "Apple Cinnamon",
        "easy drive": "Easy Drive",
        "creamy matcha": "Creamy Matcha",
        "white lavender": "White Lavender",
        "green tea": "Green Tea",
        "flushie": "Flushie 50ml",
        "lavender": "Lavender",
        "eucalyptus": "Eucalyptus",
        "peppermint": "Peppermint",
        "holiday hugs": "Holiday Hugs 200ml",
        "reed sticks": "Reed Sticks",
        "bamboo reed sticks": "Reed Sticks",
        "tropical paradise": "Tropical Paradise 200ml",
        "waterbased": "Waterbased",
        "clean air": "Clean Air",
        "fresh bamboo 500": "Fresh Bamboo 500ml",
        "white tea 500": "White Tea 500ml",
        "cotton clean 500": "Cotton Clean 500ml",
        "relaxing naturals 500": "Relaxing Naturals 500ml",
        "handwash": "Hand Wash Relaxing Naturals 500ml",
        "hand wash": "Hand Wash Relaxing Naturals 500ml",
        "reed diffuser refill": "White Lavender",  # mapped by variation
        "modern luxe": "Modern Luxe Reed Diffuser",
        "serene sanctuary": "Serene Sanctuary Reed Diffuser",
        "urban oasis": "Urban Oasis Reed Diffuser",
    }
    for key, val in mapping.items():
        if key in n:
            return val
    return name.strip()

def extract_valid_skus_from_template(ws_sample):
    """Reads the Offtake Report Summary per sku sheet to get the exact valid SKU names."""
    # Find the header row
    hdr_row = None
    for i, row in enumerate(ws_sample.iter_rows(values_only=True), 1):
        if row[0] == "VARIANT":
            hdr_row = i
            break
            
    if not hdr_row:
        return []
        
    valid_skus = []
    for i, row in enumerate(ws_sample.iter_rows(min_row=hdr_row+1, values_only=True), hdr_row+1):
        if row[0]:
            valid_skus.append(str(row[0]).strip())
    return valid_skus

def match_sku(normalized_name, valid_skus):
    """
    Attempts to match the normalized name to a valid SKU.
    Returns the valid SKU if matched, otherwise returns None.
    """
    for known_sku in valid_skus:
        if known_sku.lower() in normalized_name.lower() or normalized_name.lower() in known_sku.lower():
            return known_sku
    return None
