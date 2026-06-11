"""
DEBUG version - prints exactly what the scraper sees for Class 1, 2025-2026 only.
Run this in GitHub Actions and paste the output back.
"""
 
import re
import requests
from bs4 import BeautifulSoup
 
BASE_URL = "https://www.mshsaa.org/Activities/ClassAndDistrictAssignments.aspx"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.5",
    "Referer": "https://www.mshsaa.org/",
}
 
params = {"alg": 5, "class": 1}
r = requests.get(BASE_URL, params=params, headers=HEADERS, timeout=20)
print(f"HTTP Status: {r.status_code}")
print(f"Page length: {len(r.text)} chars")
print()
 
soup = BeautifulSoup(r.text, "lxml")
 
# ── 1. Show raw text around "Bucklin" ────────────────────────────────────────
raw = soup.get_text(" ", strip=True)
idx = raw.find("Bucklin")
if idx >= 0:
    print("=== RAW TEXT AROUND 'Bucklin' (±200 chars) ===")
    print(repr(raw[max(0, idx-200):idx+200]))
else:
    print("'Bucklin' NOT FOUND in page text at all")
print()
 
# ── 2. Show all anchor tags whose href contains Schedule.aspx ────────────────
anchors = soup.find_all("a", href=re.compile(r"Schedule\.aspx"))
print(f"=== SCHEDULE ANCHORS (first 20 of {len(anchors)}) ===")
for a in anchors[:20]:
    href = a.get("href", "")
    text = a.get_text(" ", strip=True)
    id_match = re.search(r"[?&]s=(\d+)", href)
    school_id = id_match.group(1) if id_match else "NO_ID"
    print(f"  ID={school_id:>6}  text={repr(text):<50}  href={href}")
print()
 
# ── 3. Show what build_id_to_canonical actually extracts ─────────────────────
start = raw.find("School ID School District")
print(f"Roster block found at position: {start}")
print()
 
pattern = re.compile(r'\b(\d+)\s+([\w][\w\s().,-]+?)\s+Logo\b')
id_map = {}
matches_found = 0
for m in pattern.finditer(raw[max(0,start):]):
    school_id = int(m.group(1))
    name = re.sub(r'\s+', ' ', m.group(2).strip())
    if name and len(name) >= 2:
        id_map[school_id] = name
        matches_found += 1
 
print(f"=== ID MAP ({matches_found} entries) ===")
for sid, name in sorted(id_map.items())[:30]:
    print(f"  {sid}: {repr(name)}")
print()
 
# ── 4. Specifically check Bucklin (ID 246) ───────────────────────────────────
print("=== BUCKLIN SPECIFIC CHECKS ===")
print(f"  ID 246 in id_map: {246 in id_map}")
if 246 in id_map:
    print(f"  canonical name: {repr(id_map[246])}")
 
# Check if "246" appears in raw text at all
idx246 = raw.find("246")
if idx246 >= 0:
    print(f"  '246' found in raw text at pos {idx246}:")
    print(f"  context: {repr(raw[max(0,idx246-50):idx246+100])}")
else:
    print("  '246' NOT found in raw text")
 
# Check what anchor has s=246
for a in anchors:
    href = a.get("href", "")
    if "s=246" in href:
        print(f"  Anchor with s=246: text={repr(a.get_text(' ', strip=True))}  href={href}")
print()
 
# ── 5. Show the first 2000 chars of raw page text ────────────────────────────
print("=== FIRST 2000 CHARS OF RAW PAGE TEXT ===")
print(raw[:2000])
