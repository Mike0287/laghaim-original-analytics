import os
import json
import re
from datetime import datetime, timezone
import pandas as pd
from playwright.sync_api import sync_playwright

OUTPUT_DIR = "output"
PET_JSON_FILE = os.path.join(OUTPUT_DIR, "pet_data.json")

os.makedirs(OUTPUT_DIR, exist_ok=True)

def parse_number(val_str):
    cleaned = re.sub(r"[^\d]", "", str(val_str))
    return int(cleaned) if cleaned else 0

def scrape_pet_rankings():
    all_pets = []
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        print("Navigating to Pet Rankings...")
        page.goto("https://www.laghaim-original.com/ranking_pet.xhtml#paginate-1", wait_until="networkidle")

        pagination_links = page.locator("a[href*='paginate-']").all()
        page_numbers = []
        for link in pagination_links:
            text = link.inner_text().strip()
            if text.isdigit():
                page_numbers.append(int(text))
        
        max_pages = max(page_numbers) if page_numbers else 1
        print(f"Detected {max_pages} pet ranking page(s). Scraping...")

        for current_page in range(1, max_pages + 1):
            if current_page > 1:
                page.goto(f"https://www.laghaim-original.com/ranking_pet.xhtml#paginate-{current_page}", wait_until="networkidle")
                page.wait_for_timeout(400)

            rows_data = page.evaluate('''() => {
                const rows = Array.from(document.querySelectorAll("table tbody tr, table tr"));
                return rows.map(r => {
                    const cells = Array.from(r.querySelectorAll("td"));
                    return cells.map(c => c.innerText.trim());
                }).filter(r => r.length >= 6);
            }''')

            for r in rows_data:
                rank = parse_number(r[0])
                name = r[1]
                owner = r[2]
                level = parse_number(r[3])
                pet_type = r[4]
                exp = parse_number(r[5])

                # Indented INSIDE the loop so every pet is captured
                if name and name.lower() != "name" and rank > 0:
                    all_pets.append({
                        "rank": rank,
                        "name": name,
                        "owner": owner,
                        "level": level,
                        "type": pet_type,
                        "exp": exp
                    })

            print(f" -> Scraped page {current_page}/{max_pages} ({len(all_pets)} total pets recorded)")

        browser.close()

    return all_pets

def process_pet_data(current_pets):
    today_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    
    if os.path.exists(PET_JSON_FILE):
        try:
            with open(PET_JSON_FILE, "r", encoding="utf-8") as f:
                stored_data = json.load(f)
        except Exception:
            stored_data = {"pets": {}}
    else:
        stored_data = {"pets": {}}

    previous_pets = stored_data.get("pets", {})
    updated_pets = {}

    for pet in current_pets:
        pet_key = f"{pet['name']}__{pet['owner']}"
        prev = previous_pets.get(pet_key)
        
        if prev:
            prev_rank = prev.get("rank", pet["rank"])
            prev_level = prev.get("level", pet["level"])
            
            rank_diff = prev_rank - pet["rank"]  # Positive = climbed ranks
            level_diff = pet["level"] - prev_level
            history = prev.get("history", {})
        else:
            rank_diff = 0
            level_diff = 0
            history = {}

        history[today_str] = {
            "rank": pet["rank"],
            "level": pet["level"],
            "exp": pet["exp"],
            "rank_change": rank_diff,
            "level_change": level_diff
        }

        pet_record = {
            "name": pet["name"],
            "owner": pet["owner"],
            "type": pet["type"],
            "level": pet["level"],
            "rank": pet["rank"],
            "exp": pet["exp"],
            "rank_change_today": rank_diff,
            "level_gain_today": level_diff,
            "history": history
        }

        updated_pets[pet_key] = pet_record

    output_payload = {
        "last_updated": datetime.now(timezone.utc).isoformat(),
        "total_pets": len(updated_pets),
        "pets": updated_pets
    }

    with open(PET_JSON_FILE, "w", encoding="utf-8") as f:
        json.dump(output_payload, f, indent=2, ensure_ascii=False)

    print(f"\n[OK] Processed {len(updated_pets)} pets -> saved to {PET_JSON_FILE}")

if __name__ == "__main__":
    pets = scrape_pet_rankings()
    if pets:
        process_pet_data(pets)