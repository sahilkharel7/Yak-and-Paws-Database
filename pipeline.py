import re
import sys
import time

import requests
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import os


# ---------------------------------------------------------------------------
# Validates that the given string is a 5-digit US ZIP code.
# ---------------------------------------------------------------------------
def validate_zip_code(zip_code):
    return bool(re.fullmatch(r"\d{5}", zip_code.strip()))


# ---------------------------------------------------------------------------
# Converts a US ZIP code to (latitude, longitude) via the Google Geocoding API.
# Returns None if the ZIP code cannot be geocoded.
# ---------------------------------------------------------------------------
def get_zip_coordinates(zip_code, api_key):
    url = "https://maps.googleapis.com/maps/api/geocode/json"
    params = {"address": zip_code, "key": api_key}

    try:
        response = requests.get(url, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()

        if data["status"] != "OK":
            print(f"[ERROR] Geocoding failed for ZIP {zip_code}: {data['status']}")
            return None

        location = data["results"][0]["geometry"]["location"]
        return (location["lat"], location["lng"])

    except requests.RequestException as e:
        print(f"[ERROR] Geocoding request error: {e}")
        return None


# ---------------------------------------------------------------------------
# Searches for nearby places using the Google Places Nearby Search API.
# Handles pagination via next_page_token (up to 3 pages max from Google).
# Returns a list of raw place result dictionaries.
# ---------------------------------------------------------------------------
def search_nearby_places(lat, lng, search_term, api_key):
    url = "https://maps.googleapis.com/maps/api/place/nearbysearch/json"
    params = {
        "location": f"{lat},{lng}",
        "radius": 5000,
        "keyword": search_term,
        "key": api_key,
    }

    all_results = []

    try:
        while True:
            response = requests.get(url, params=params, timeout=10)
            response.raise_for_status()
            data = response.json()

            if data["status"] not in ("OK", "ZERO_RESULTS"):
                print(
                    f"[ERROR] Nearby Search failed for '{search_term}': {data['status']}"
                )
                break

            all_results.extend(data.get("results", []))

            # Google provides next_page_token when more results are available
            next_token = data.get("next_page_token")
            if not next_token:
                break

            # Google requires a short delay before the token becomes valid
            time.sleep(2)
            params = {"pagetoken": next_token, "key": api_key}

        print(f"  Found {len(all_results)} results for '{search_term}'")
        return all_results

    except requests.RequestException as e:
        print(f"[ERROR] Nearby Search request error: {e}")
        return all_results


# ---------------------------------------------------------------------------
# Fetches full details for a single place using its place_id.
# Returns the detail dictionary, or None on failure.
# ---------------------------------------------------------------------------
def get_place_details(place_id, api_key):
    url = "https://maps.googleapis.com/maps/api/place/details/json"
    fields = (
        "name,place_id,formatted_address,address_components,"
        "formatted_phone_number,international_phone_number,website,url,"
        "rating,user_ratings_total,price_level,business_status,types,"
        "geometry,opening_hours"
    )
    params = {"place_id": place_id, "fields": fields, "key": api_key}

    try:
        response = requests.get(url, params=params, timeout=10)
        response.raise_for_status()
        data = response.json()

        if data["status"] != "OK":
            print(f"[ERROR] Place Details failed for {place_id}: {data['status']}")
            return None

        return data.get("result")

    except requests.RequestException as e:
        print(f"[ERROR] Place Details request error: {e}")
        return None


# ---------------------------------------------------------------------------
# Removes duplicate places based on place_id.
# ---------------------------------------------------------------------------
def deduplicate_places(places_list):
    seen = set()
    unique = []
    for place in places_list:
        pid = place.get("place_id")
        if pid and pid not in seen:
            seen.add(pid)
            unique.append(place)
    return unique


# ---------------------------------------------------------------------------
# Keeps only businesses whose formatted_address contains the target ZIP code.
# ---------------------------------------------------------------------------
def filter_by_zip(places_list, target_zip):
    return [
        p
        for p in places_list
        if target_zip in p.get("formatted_address", "")
    ]


# ---------------------------------------------------------------------------
# Helper: extracts a specific address component type from the
# address_components list returned by the Places API.
# ---------------------------------------------------------------------------
def _extract_component(components, target_type):
    for comp in components:
        if target_type in comp.get("types", []):
            return comp.get("long_name", "")
    return ""


# ---------------------------------------------------------------------------
# Helper: extracts the short_name for state from address_components.
# ---------------------------------------------------------------------------
def _extract_state(components):
    for comp in components:
        if "administrative_area_level_1" in comp.get("types", []):
            return comp.get("short_name", "")
    return ""


# ---------------------------------------------------------------------------
# Builds a flat record dictionary from a place details response.
# Missing fields default to an empty string.
# ---------------------------------------------------------------------------
def build_record(place_details, search_term, input_zip):
    if not place_details:
        return None

    components = place_details.get("address_components", [])
    geometry = place_details.get("geometry", {})
    location = geometry.get("location", {})
    opening_hours = place_details.get("opening_hours", {})
    periods_text = opening_hours.get("weekday_text", [])

    # Map each day name to its hours string from weekday_text
    # weekday_text looks like ["Monday: 9:00 AM – 6:00 PM", ...]
    day_hours = {
        "Monday": "",
        "Tuesday": "",
        "Wednesday": "",
        "Thursday": "",
        "Friday": "",
        "Saturday": "",
        "Sunday": "",
    }
    for entry in periods_text:
        for day in day_hours:
            if entry.startswith(day):
                # Everything after "Day: " is the hours portion
                day_hours[day] = entry.split(": ", 1)[1] if ": " in entry else entry

    types_list = place_details.get("types", [])

    return {
        "business_name": place_details.get("name", ""),
        "place_id": place_details.get("place_id", ""),
        "formatted_address": place_details.get("formatted_address", ""),
        "street_number": _extract_component(components, "street_number"),
        "street_name": _extract_component(components, "route"),
        "city": _extract_component(components, "locality"),
        "state": _extract_state(components),
        "zip_code": _extract_component(components, "postal_code"),
        "country": _extract_component(components, "country"),
        "latitude": location.get("lat", ""),
        "longitude": location.get("lng", ""),
        "phone_local": place_details.get("formatted_phone_number", ""),
        "phone_international": place_details.get("international_phone_number", ""),
        "website": place_details.get("website", ""),
        "google_maps_url": place_details.get("url", ""),
        "rating": place_details.get("rating", ""),
        "user_ratings_total": place_details.get("user_ratings_total", ""),
        "price_level": place_details.get("price_level", ""),
        "business_status": place_details.get("business_status", ""),
        "business_types": ", ".join(types_list),
        "open_now": opening_hours.get("open_now", ""),
        "hours_monday": day_hours["Monday"],
        "hours_tuesday": day_hours["Tuesday"],
        "hours_wednesday": day_hours["Wednesday"],
        "hours_thursday": day_hours["Thursday"],
        "hours_friday": day_hours["Friday"],
        "hours_saturday": day_hours["Saturday"],
        "hours_sunday": day_hours["Sunday"],
        "search_term": search_term,
        "input_zip_code": input_zip,
    }


# ---------------------------------------------------------------------------
# Saves a list of record dictionaries to a formatted Excel file.
# Returns the filename that was saved.
# ---------------------------------------------------------------------------
def save_to_excel(records, zip_code):
    filename = f"pet_stores_{zip_code}.xlsx"

    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Pet Businesses"

        if not records:
            wb.save(filename)
            return filename

        headers = list(records[0].keys())

        # Write bold header row
        bold_font = Font(bold=True)
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = bold_font

        # Write data rows
        for row_idx, record in enumerate(records, start=2):
            for col_idx, header in enumerate(headers, start=1):
                ws.cell(row=row_idx, column=col_idx, value=record.get(header, ""))

        # Auto-size columns based on content width
        for col_idx, header in enumerate(headers, start=1):
            max_length = len(str(header))
            for row_idx in range(2, len(records) + 2):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                if cell_value is not None:
                    max_length = max(max_length, len(str(cell_value)))
            adjusted_width = min(max_length + 2, 50)  # cap at 50 to stay readable
            ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

        wb.save(filename)
        print(f"\n[OK] Saved to {filename}")
        return filename

    except Exception as e:
        print(f"[ERROR] Failed to save Excel file: {e}")
        return None


# ---------------------------------------------------------------------------
# Main pipeline: ties all the steps together.
# ---------------------------------------------------------------------------
def main():
    # Step 1 — Get ZIP code from user
    zip_code = input("Enter ZIP code: ").strip()

    if not validate_zip_code(zip_code):
        print("[ERROR] Invalid ZIP code. Please enter a valid 5-digit US ZIP code.")
        sys.exit(1)

    # Step 2 — Load API key from .env
    load_dotenv()
    api_key = os.getenv("GOOGLE_PLACES_API_KEY")

    if not api_key:
        print("[ERROR] GOOGLE_PLACES_API_KEY not found in .env file.")
        sys.exit(1)

    # Step 3 — Convert ZIP to coordinates
    print(f"\nGeocoding ZIP code {zip_code}...")
    coords = get_zip_coordinates(zip_code, api_key)

    if coords is None:
        print("[ERROR] Could not geocode the ZIP code. Exiting.")
        sys.exit(1)

    lat, lng = coords
    print(f"  Coordinates: {lat}, {lng}")

    # Step 4 — Search for businesses using 3 search terms
    search_terms = ["pet store", "pet supplies", "pet shop"]
    all_raw_results = []

    print("\nSearching for pet businesses...")
    for term in search_terms:
        results = search_nearby_places(lat, lng, term, api_key)
        all_raw_results.extend(results)

    raw_count = len(all_raw_results)
    print(f"\nTotal raw results: {raw_count}")

    if raw_count == 0:
        print("[INFO] No businesses found. Exiting.")
        sys.exit(0)

    # Step 5 — Deduplicate by place_id
    unique_places = deduplicate_places(all_raw_results)
    dedup_count = len(unique_places)
    print(f"After deduplication: {dedup_count}")

    # Step 6 — Get full details for each unique place
    print("\nFetching place details...")
    detailed_places = []
    for i, place in enumerate(unique_places, start=1):
        pid = place.get("place_id")
        print(f"  [{i}/{dedup_count}] {place.get('name', 'Unknown')}")
        details = get_place_details(pid, api_key)
        if details:
            detailed_places.append(details)

    # Step 7 — Filter to keep only businesses in the target ZIP
    filtered = filter_by_zip(detailed_places, zip_code)
    filtered_count = len(filtered)
    print(f"\nAfter ZIP code filtering: {filtered_count}")

    if filtered_count == 0:
        print(f"[INFO] No businesses found in ZIP code {zip_code}. Exiting.")
        sys.exit(0)

    # Step 8 — Build clean records
    records = []
    for details in filtered:
        # Use the first matching search term found in the place's types/name
        # Default to the first search term
        matched_term = search_terms[0]
        record = build_record(details, matched_term, zip_code)
        if record:
            records.append(record)

    # Step 9 — Save to Excel
    filename = save_to_excel(records, zip_code)

    # Step 10 — Print summary
    print("\n" + "=" * 50)
    print("PIPELINE SUMMARY")
    print("=" * 50)
    print(f"  Raw results found:        {raw_count}")
    print(f"  After deduplication:       {dedup_count}")
    print(f"  After ZIP filtering:       {filtered_count}")
    print(f"  Records saved:             {len(records)}")
    if filename:
        print(f"  Output file:               {filename}")
    print("=" * 50)


if __name__ == "__main__":
    main()
