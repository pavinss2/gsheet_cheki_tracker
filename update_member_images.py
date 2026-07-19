import os
import re
import json
import requests
from bs4 import BeautifulSoup
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from googleapiclient.discovery import build

SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "10Zr4KwYlrbHL2Gh2c1st5uw9Bxmj3h5aHjney5bqTPI")
SHEET_NAME = "dim_member"

def get_credentials():
    # Try Service Account first
    sa_key = os.environ.get("GCP_SA_KEY")
    if sa_key:
        print("Using service account credentials from GCP_SA_KEY...")
        info = json.loads(sa_key)
        return service_account.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/spreadsheets"]
        )

    # Fallback to local clasp credentials for testing
    clasp_path = os.path.expanduser("~/.clasprc.json")
    if os.path.exists(clasp_path):
        print(f"Loading local clasp credentials from {clasp_path}...")
        with open(clasp_path, 'r') as f:
            data = json.load(f)
            token_info = data.get("token") or data.get("tokens", {}).get("default")
            if token_info:
                scopes = token_info.get("scope", "").split()
                # Ensure spreadsheets scope is present
                if "https://www.googleapis.com/auth/spreadsheets" not in scopes:
                    scopes.append("https://www.googleapis.com/auth/spreadsheets")
                return Credentials(
                    token=token_info.get("access_token"),
                    refresh_token=token_info.get("refresh_token"),
                    token_uri="https://oauth2.googleapis.com/token",
                    client_id=data.get("oauth2ClientSettings", {}).get("clientId") or token_info.get("client_id"),
                    client_secret=data.get("oauth2ClientSettings", {}).get("clientSecret") or token_info.get("client_secret"),
                    scopes=scopes
                )
    
    raise ValueError("No authentication credentials found. Set GCP_SA_KEY env variable or run 'clasp login' locally.")

# Helper to convert col index to letter
def get_column_letter(col_idx):
    letter = ""
    temp = col_idx
    while temp >= 0:
        letter = chr(temp % 26 + 65) + letter
        temp = temp // 26 - 1
    return letter

def extract_url(cell_value):
    if not cell_value:
        return None
    # Parse =HYPERLINK("url", "label") or =IMAGE("url")
    if isinstance(cell_value, str) and cell_value.startswith("="):
        # Look for hyperlink
        m = re.search(r'=HYPERLINK\(\s*["\']([^"\']+)["\']', cell_value, re.IGNORECASE)
        if m:
            return m.group(1)
        # Look for image
        m = re.search(r'=IMAGE\(\s*["\']([^"\']+)["\']', cell_value, re.IGNORECASE)
        if m:
            return m.group(1)
    if isinstance(cell_value, str) and cell_value.strip().startswith("http"):
        return cell_value.strip()
    return None

def is_url_active(url):
    if not url:
        return False
    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    try:
        r = requests.get(url, headers=headers, timeout=10)
        return r.status_code == 200
    except Exception as e:
        print(f"Error checking activity for {url}: {e}")
        return False

def clean_x_profile(input_val):
    if not input_val:
        return None
    s = str(input_val).strip()
    if not s or s.lower() == "n/a":
        return None
    if s.startswith("http"):
        if "twitter.com" in s:
            s = s.replace("twitter.com", "x.com")
        return s
    if s.startswith("@"):
        s = s[1:]
    return f"https://x.com/{s}"

def get_x_profile_image_url(profile_url):
    clean_url = clean_x_profile(profile_url)
    if not clean_url:
        return None
        
    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
    }
    
    # Try 1: Direct scrape
    try:
        r = requests.get(clean_url, headers=headers, timeout=10)
        if r.status_code == 200:
            soup = BeautifulSoup(r.text, 'html.parser')
            og_image = soup.find('meta', property='og:image') or soup.find('meta', name='twitter:image')
            if og_image and og_image.get('content'):
                avatar_url = og_image['content']
                suffix_pattern = r'_(?:normal|bigger|mini|reasonably_small|x96|\d+x\d+)(\.[a-zA-Z0-9]+)$'
                if re.search(suffix_pattern, avatar_url):
                    avatar_url = re.sub(suffix_pattern, r'_400x400\1', avatar_url)
                return avatar_url
    except Exception as e:
        print(f"Direct scrape failed for {clean_url}: {e}")
        
    # Try 2: Microlink API fallback
    try:
        print(f"Using Microlink fallback for {clean_url}...")
        api_url = f"https://api.microlink.io/?url={requests.utils.quote(clean_url)}"
        r = requests.get(api_url, timeout=10)
        if r.status_code == 200:
            data = r.json()
            if data.get("status") == "success" and data.get("data", {}).get("image", {}).get("url"):
                avatar_url = data["data"]["image"]["url"]
                suffix_pattern = r'_(?:normal|bigger|mini|reasonably_small|x96|\d+x\d+)(\.[a-zA-Z0-9]+)$'
                if re.search(suffix_pattern, avatar_url):
                    avatar_url = re.sub(suffix_pattern, r'_400x400\1', avatar_url)
                return avatar_url
    except Exception as e:
        print(f"Microlink fallback failed for {clean_url}: {e}")
        
    return None

def main():
    creds = get_credentials()
    service = build("sheets", "v4", credentials=creds)
    
    # Fetch values and formulas
    print(f"Fetching sheet {SHEET_NAME} data from spreadsheet {SPREADSHEET_ID}...")
    
    # Read values
    val_res = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{SHEET_NAME}'!A:Z"
    ).execute()
    values = val_res.get("values", [])
    
    # Read formulas (to preserve hyperlink formats)
    formula_res = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{SHEET_NAME}'!A:Z",
        valueRenderOption="FORMULA"
    ).execute()
    formulas = formula_res.get("values", [])
    
    if not values:
        print("No data found in the sheet.")
        return
        
    headers = [str(h).strip().lower() for h in values[0]]
    try:
        member_image_idx = headers.index("member_image")
        x_profile_idx = headers.index("x_profile")
    except ValueError as e:
        print(f"Error: Missing required columns in headers {headers}. {e}")
        return
        
    member_name_idx = headers.index("member_name") if "member_name" in headers else -1
    
    print(f"Columns found: member_image at index {member_image_idx}, x_profile at index {x_profile_idx}")
    
    updated_count = 0
    
    for i in range(1, len(values)):
        row_num = i + 1  # 1-indexed for Sheets
        member_name = values[i][member_name_idx] if (member_name_idx != -1 and member_name_idx < len(values[i])) else f"Row {row_num}"
        
        # Safely get current value and formula
        current_val = values[i][member_image_idx] if member_image_idx < len(values[i]) else ""
        current_formula = formulas[i][member_image_idx] if (i < len(formulas) and member_image_idx < len(formulas[i])) else ""
        
        x_profile_val = values[i][x_profile_idx] if x_profile_idx < len(values[i]) else ""
        
        # If formula is present, check that, otherwise value
        raw_image_url = extract_url(current_formula or current_val)
        
        print(f"\n[{i}/{len(values)-1}] Checking {member_name} (image: {raw_image_url})")
        
        if raw_image_url and is_url_active(raw_image_url):
            print("-> URL is active. Skipping.")
            continue
            
        print("-> URL is inactive, broken, or missing. Proceeding to update...")
        clean_profile = clean_x_profile(x_profile_val)
        if not clean_profile:
            print("-> x_profile is blank or 'N/A'. Skipping.")
            continue
            
        print(f"-> Resolving image for X profile: {clean_profile}...")
        new_img_url = get_x_profile_image_url(clean_profile)
        
        if new_img_url:
            # Reconstruct formula label if original had one
            label = "Link"
            if current_formula and current_formula.startswith("="):
                m = re.search(r'=HYPERLINK\(\s*["\'][^"\']+["\']\s*,\s*["\']([^"\']+)["\']\)', current_formula, re.IGNORECASE)
                if m:
                    label = m.group(1)
            
            new_cell_value = f'=HYPERLINK("{new_img_url}", "{label}")'
            col_letter = get_column_letter(member_image_idx)
            cell_range = f"'{SHEET_NAME}'!{col_letter}{row_num}"
            
            print(f"-> SUCCESS: Updating {cell_range} to {new_cell_value}")
            
            service.spreadsheets().values().update(
                spreadsheetId=SPREADSHEET_ID,
                range=cell_range,
                valueInputOption="USER_ENTERED",
                body={"values": [[new_cell_value]]}
            ).execute()
            
            updated_count += 1
        else:
            print(f"-> FAILED: Could not resolve new profile image URL for {member_name}")
            
    print(f"\nDone. Updated {updated_count} member images.")

if __name__ == "__main__":
    main()
