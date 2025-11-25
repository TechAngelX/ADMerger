from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
import csv
import json
import time
import os
import re

LOGIN_URL = "https://www.ucl.ac.uk/srs/user/login"
MAIN_URL = "https://www.ucl.ac.uk/srs/admissions-manual/qualifications/international-qualifications"

OUTPUT_CSV = "/Users/xeon2035/Desktop/ucl_degree_equivalencies_FINAL.csv"
OUTPUT_JSON = "/Users/xeon2035/Desktop/ucl_degree_equivalencies_FINAL.json"
PROGRESS_FILE = "/Users/xeon2035/Desktop/ucl_scrape_progress.json"

def parse_country_links(html):
    soup = BeautifulSoup(html, "html.parser")
    links = soup.find_all("a", href=True)
    countries = []
    
    for a in links:
        name = a.get_text(strip=True)
        href = a["href"].strip()
        
        if not name or href.startswith("#") or "back to top" in name.lower():
            continue
        
        if "/international-qualifications/" in href and not href.endswith("/international-qualifications"):
            if href.startswith("/"):
                full_url = f"https://www.ucl.ac.uk{href}"
            else:
                full_url = href
            
            countries.append({"name": name, "url": full_url})
    
    seen = set()
    unique = []
    for c in countries:
        if c["name"].lower() not in seen:
            seen.add(c["name"].lower())
            unique.append(c)
    
    return unique

def find_graduate_requirements_link(html):
    soup = BeautifulSoup(html, "html.parser")
    
    for a in soup.find_all("a", href=True):
        link_text = a.get_text(strip=True).lower()
        href = a["href"]
        
        if "graduate" in link_text and "requirement" in link_text and "undergraduate" not in link_text:
            if href.startswith("/"):
                return f"https://www.ucl.ac.uk{href}"
            else:
                return href
    
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if "graduate-requirements" in href.lower() and "undergraduate" not in href.lower():
            if href.startswith("/"):
                return f"https://www.ucl.ac.uk{href}"
            else:
                return href
    
    return None

def extract_equivalencies(html):
    soup = BeautifulSoup(html, "html.parser")
    
    h2 = soup.find("h2", string=lambda s: s and "UK degree classification" in s)
    
    if not h2:
        return {}
    
    results = {}
    current_key = None
    current_texts = []
    sections_found = []
    
    for tag in h2.find_all_next():
        if tag.name == "h2":
            if current_key and current_texts:
                results[current_key] = " ".join(current_texts)
            break
        
        if tag.name == "hr":
            continue
        
        if tag.name == "h3":
            h3_text = tag.get_text(strip=True)
            
            if current_key and current_texts:
                results[current_key] = " ".join(current_texts)
                sections_found.append(current_key)
                current_texts = []
            
            if "2:2" in h3_text or "2.2" in h3_text:
                current_key = "2:2 Hons equivalent"
            elif "2:1" in h3_text or "2.1" in h3_text:
                current_key = "2:1 Honours equivalent"
            elif "1st" in h3_text.lower() or "first" in h3_text.lower():
                current_key = "1st Honours equivalent"
            else:
                if len(sections_found) >= 3:
                    current_key = None
                    break
                current_key = None
        
        elif current_key and tag.name in ["p", "li", "ul"]:
            text = tag.get_text(" ", strip=True)
            if text and len(text) > 5:
                current_texts.append(text)
    
    if current_key and current_texts:
        results[current_key] = " ".join(current_texts)
    
    return results

def extract_numbers(text):
    """Extract numbers - PRIORITIZE CGPA"""
    if not text or text == "N/A":
        return "N/A"
    
    results = []
    
    # CGPA - HIGHEST PRIORITY
    cgpa_pattern = r'(?:CGPA|GPA)\s*(?:of\s*)?(\d+\.?\d*)\s*/\s*(\d+\.?\d*)'
    cgpa_matches = re.findall(cgpa_pattern, text, re.IGNORECASE)
    for match in cgpa_matches:
        numerator, denominator = match
        results.append(f"'{numerator}/{denominator}")
    
    if results:
        return results[0]
    
    # Percentages
    percent_pattern = r'(\d+\.?\d*)\s*%'
    percent_matches = re.findall(percent_pattern, text)
    if percent_matches:
        percentages = [float(p) for p in percent_matches]
        max_percent = max(percentages)
        if max_percent == int(max_percent):
            results.append(f"{int(max_percent)}%")
        else:
            results.append(f"{max_percent}%")
    
    # Fractions
    fraction_pattern = r'(?<!CGPA\s)(?<!GPA\s)(?<!\d\s)(\d+\.?\d*)\s*/\s*(\d+\.?\d*)'
    fraction_matches = re.findall(fraction_pattern, text)
    for match in fraction_matches:
        numerator, denominator = match
        fraction_str = f"'{numerator}/{denominator}"
        if not any(numerator in r and denominator in r for r in results):
            results.append(fraction_str)
    
    # Grade letters
    grade_pattern = r'\b([A-F][+-]?)\b'
    grade_matches = re.findall(grade_pattern, text)
    for match in grade_matches:
        if match not in results and match not in ['A', 'I', 'N']:
            results.append(match)
    
    # Class honors
    if "Upper Second Class" in text or "Second Class Division 1" in text or "Class II Division i" in text or "H2A" in text:
        if "2.1" not in results:
            results.append("2.1")
    elif "Lower Second Class" in text or "Second Class Division 2" in text or "Class II Division ii" in text or "H2B" in text:
        if "2.2" not in results:
            results.append("2.2")
    elif "First Class" in text or "1st Class" in text or "H1" in text:
        if "1.0" not in results:
            results.append("1.0")
    elif "Third Class" in text or "3rd Class" in text:
        if "3.0" not in results:
            results.append("3.0")
    
    # Descriptive grades
    descriptive_with_or = re.findall(r'(?:or|,)\s*["\']?([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)["\']?', text)
    for grade in descriptive_with_or:
        grade = grade.strip()
        if grade in ["Good", "Very Good", "Excellent", "Distinction", "Merit", "Credit"]:
            if grade not in results:
                results.append(grade)
    
    if results:
        seen = set()
        unique_results = []
        for r in results:
            key = r.replace("'", "")
            if key not in seen:
                seen.add(key)
                unique_results.append(r)
        return unique_results[0]
    
    if "contact" in text.lower():
        return "Contact admissions"
    
    return text[:100] + "..." if len(text) > 100 else text

def load_progress():
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_progress(data):
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def save_final_clean(data):
    """Save ONLY clean CSV and JSON"""
    cleaned_data = []
    for row in data:
        cleaned_row = {
            "country": row["country"],
            "2.2_requirement": extract_numbers(row.get("2:2 Hons equivalent", "N/A")),
            "2.1_requirement": extract_numbers(row.get("2:1 Honours equivalent", "N/A")),
            "1st_requirement": extract_numbers(row.get("1st Honours equivalent", "N/A"))
        }
        cleaned_data.append(cleaned_row)
    
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["country", "2.2_requirement", "2.1_requirement", "1st_requirement"]
        )
        writer.writeheader()
        writer.writerows(cleaned_data)
    
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(cleaned_data, f, ensure_ascii=False, indent=2)

def main():
    print("🌍 SCRAPING ALL COUNTRIES")
    print("🦁 Setting up Brave...")
    
    brave_options = Options()
    brave_options.binary_location = "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser"
    brave_options.add_argument("--no-sandbox")
    
    try:
        driver = webdriver.Chrome(options=brave_options)
        print("✅ Brave opened!")
        
    except Exception as e:
        print(f"❌ Failed: {e}")
        return
    
    try:
        print(f"📋 Login: {LOGIN_URL}")
        driver.get(LOGIN_URL)
        
        print("\n⏸️  LOG IN, then press ENTER\n")
        input("Press ENTER after login: ")
        
        print(f"\n🔍 Getting countries...")
        driver.get(MAIN_URL)
        time.sleep(3)
        
        html = driver.page_source
        countries = parse_country_links(html)
        
        print(f"🌍 Found {len(countries)} countries")
        
        all_data = load_progress()
        done_countries = {item["country"] for item in all_data}
        print(f"▶️ Resume: {len(done_countries)} done\n")
        
        for i, c in enumerate(countries, start=1):
            country_name = c["name"]
            
            if country_name in done_countries:
                print(f"({i}/{len(countries)}) ⏩ {country_name}")
                continue
            
            print(f"({i}/{len(countries)}) {country_name}...", end=" ")
            
            try:
                driver.get(c["url"])
                time.sleep(1.5)
                
                grad_url = find_graduate_requirements_link(driver.page_source)
                
                if not grad_url:
                    print("❌")
                    all_data.append({
                        "country": country_name,
                        "2:2 Hons equivalent": "N/A",
                        "2:1 Honours equivalent": "N/A",
                        "1st Honours equivalent": "N/A"
                    })
                    save_progress(all_data)
                    continue
                
                driver.get(grad_url)
                time.sleep(1.5)
                
                equivalencies = extract_equivalencies(driver.page_source)
                
                all_data.append({
                    "country": country_name,
                    "2:2 Hons equivalent": equivalencies.get("2:2 Hons equivalent", "N/A"),
                    "2:1 Honours equivalent": equivalencies.get("2:1 Honours equivalent", "N/A"),
                    "1st Honours equivalent": equivalencies.get("1st Honours equivalent", "N/A")
                })
                save_progress(all_data)
                print(f"✅")
                
            except Exception as e:
                print(f"⚠️ {e}")
                continue
        
        # Clean and save FINAL files
        print(f"\n🧹 Cleaning data...")
        save_final_clean(all_data)
        
        print(f"\n🎉 DONE! {len(all_data)} countries")
        print(f"✅ Saved ONLY:")
        print(f"  📄 {OUTPUT_CSV}")
        print(f"  📄 {OUTPUT_JSON}")
        print(f"\n💡 Delete {PROGRESS_FILE} to start fresh")
    
    except KeyboardInterrupt:
        print(f"\n\n⚠️ INTERRUPTED!")
        save_final_clean(all_data)
        print(f"✅ Saved progress: {len(all_data)} countries")
    
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
