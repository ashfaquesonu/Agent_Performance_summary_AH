from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd
from collections import defaultdict
import re

# --- CONFIGURATION ---
CATEGORIES = [
    "https://assethomezinternational.bitrix24.com/crm/deal/category/14/",
    "https://assethomezinternational.bitrix24.com/crm/deal/category/16/"
]

MEETING_STAGES = ["ONLINE - MEETINGS", "FIRST MEETING", "2ND SV/F2F", "3RD SV/F2F", "HOT PIPELINE", "DEAL WON"]
QUALIFIED_STAGE = "QUALIFIED"
CLAIM_STAGE = "CLAIMED MARKETING LEAD"
OUTPUT_PATH = r"C:\Users\ASUS\Documents\Performance_ProjectWise.xlsx"

SYSTEM_AGENTS = ["ASSET HOMEZ", "SYSTEM", "BITRIX24", "AUTOMATION"]
MANAGEMENT_AGENTS = ["ANVAR SADATH", "MUHAMMED ASHFAQUE M", "TEENA THOMAS"]

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
chrome_options.add_experimental_option("detach", True)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
wait = WebDriverWait(driver, 15)

master_report = defaultdict(lambda: defaultdict(lambda: {
    "claims": 0, "c_ids": set(), 
    "qual": 0, "q_ids": set(), 
    "meet_count": 0, "m_details": set() 
}))

def clean_sheet_name(name):
    if not name: return "Undefined Project"
    clean = re.sub(r'[\[\]\:\*\?\/\\]', '', name)
    return clean[:31]

def get_deal_id_from_url(url):
    try: return "".join(filter(str.isdigit, url.split('details/')[1].split('/')[0]))
    except: return None

def parse_history_row(row_element):
    try:
        cells = row_element.find_elements(By.CSS_SELECTOR, ".main-grid-cell-content")
        cell_texts = [c.text.strip() for c in cells if c.text.strip()]
        agent, event, desc = "", "", ""
        if len(cell_texts) >= 4:
            if any(char.isdigit() for char in cell_texts[0]):
                agent, event, desc = cell_texts[1], cell_texts[2], " ".join(cell_texts[3:])
            else:
                agent, event, desc = cell_texts[2], cell_texts[3], " ".join(cell_texts[4:])
        return agent.split('\n')[0].strip(), event, desc, row_element.text.strip()
    except: return "", "", "", ""

def fetch_all_history_entries(driver):
    history_entries = []
    while True:
        for _ in range(15):
            try:
                more_btns = driver.find_elements(By.CSS_SELECTOR, ".main-grid-more-btn, .main-grid-nav-more")
                visible = [b for b in more_btns if b.is_displayed()]
                if visible: driver.execute_script("arguments[0].click();", visible[0]); time.sleep(1.5)
                else: break
            except: break
        
        rows = driver.find_elements(By.CSS_SELECTOR, ".main-grid-row.main-grid-row-body")
        for row in rows:
            a, e, d, f = parse_history_row(row)
            if f: history_entries.append({"agent": a, "event": e, "desc": d, "full_text": f})
        
        try:
            next_btn = driver.find_elements(By.CSS_SELECTOR, "a.main-grid-page-next")
            if next_btn and next_btn[0].is_displayed() and "disabled" not in next_btn[0].get_attribute("class"):
                driver.execute_script("arguments[0].click();", next_btn[0]); time.sleep(2)
            else: break
        except: break
    return history_entries

def scrape_unified_master_report():
    try:
        all_unique_deal_ids = set()
        
        # --- BIT ACCURATE AUTO-SCROLL TO LOAD ALL DEALS ---
        for url in CATEGORIES:
            driver.get(url)
            input(f"\nApply filters on {url} and press ENTER...")
            
            last_count = 0
            retries = 0
            while retries < 4: # Increased retries for accuracy
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(2.5)
                
                try:
                    list_more = driver.find_elements(By.CSS_SELECTOR, ".main-grid-more-btn")
                    if list_more and list_more[0].is_displayed():
                        driver.execute_script("arguments[0].click();", list_more[0])
                        time.sleep(2)
                except: pass

                links = driver.find_elements(By.XPATH, "//a[contains(@href, '/crm/deal/details/')]")
                current_ids = list(dict.fromkeys([get_deal_id_from_url(l.get_attribute('href')) for l in links if get_deal_id_from_url(l.get_attribute('href'))]))
                
                if len(current_ids) > last_count:
                    last_count = len(current_ids)
                    retries = 0
                    print(f"Loading deals... Current Count: {last_count}", end="\r")
                else:
                    retries += 1
            
            all_unique_deal_ids.update(current_ids)
            print(f"\nFinal confirmed count for Category: {len(current_ids)}")

        # --- PROCESS DEALS ---
        for index, d_id in enumerate(all_unique_deal_ids):
            print(f"[{index+1}/{len(all_unique_deal_ids)}] Deal {d_id}", end=" ", flush=True)
            try:
                driver.get(f"https://assethomezinternational.bitrix24.com/crm/deal/details/{d_id}/")
                iframe = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "side-panel-iframe")))
                driver.switch_to.frame(iframe)
                
                # --- ACCURATE PROJECT DETECTION ---
                proj_name = "Undefined Project"
                # Strategy: Scroll inside iframe to trigger lazy-load of fields
                driver.execute_script("window.scrollTo(0, 500);") 
                time.sleep(1)

                # Search strategies for the field
                xpaths = [
                    "//div[contains(@class, 'ui-entity-editor-field-title')]//span[text()='Project Name-Facebook']/ancestor::div[contains(@class, 'ui-entity-editor-field-container')]//div[contains(@class, 'ui-entity-editor-field-content')]",
                    "//*[contains(text(), 'Project Name-Facebook')]/following::div[contains(@class, 'ui-entity-editor-field-content')][1]",
                    "//*[contains(text(), 'Project Name-Facebook')]/ancestor::div[1]/following-sibling::div[1]"
                ]
                
                for xp in xpaths:
                    try:
                        element = driver.find_element(By.XPATH, xp)
                        val = element.text.strip()
                        if val and val != "None" and val != "":
                            proj_name = val
                            break
                    except: continue
                
                print(f"({proj_name})", end=" ", flush=True)

                # SWITCH TO HISTORY
                try:
                    h_tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-id='tab_history']|//a[contains(., 'History')]")))
                    driver.execute_script("arguments[0].click();", h_tab)
                except:
                    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//*[contains(@class, 'more')]"))
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//span[contains(text(), 'History')]"))
                
                time.sleep(1.5)
                history_chronological = list(reversed(fetch_all_history_entries(driver)))

                current_resp = ""
                claimed_done = False

                for i, entry in enumerate(history_chronological):
                    actor = entry['agent'].strip()
                    actor_up = actor.upper()
                    full_up = entry['full_text'].upper()
                    desc_up = entry['desc'].upper()

                    if "RESPONSIBLE" in entry['event'].upper() or "RESPONSIBLE" in desc_up:
                        if "→" in entry['desc']: current_resp = entry['desc'].split("→")[-1].strip()
                        elif "->" in entry['desc']: current_resp = entry['desc'].split("->")[-1].strip()

                    if not claimed_done and (CLAIM_STAGE in full_up or "CLAIMED" in desc_up):
                        credited = actor
                        if actor_up in [m.upper() for m in MANAGEMENT_AGENTS]:
                            new_assign = ""
                            for look in history_chronological[i+1 : i+4]:
                                if "RESPONSIBLE" in look['event'].upper():
                                    new_assign = look['desc'].split("→")[-1].strip() if "→" in look['desc'] else ""
                                    break
                            credited = new_assign if new_assign else current_resp
                        
                        if credited and not any(s in credited.upper() for s in SYSTEM_AGENTS):
                            master_report[proj_name][credited]["claims"] += 1
                            master_report[proj_name][credited]["c_ids"].add(d_id)
                            claimed_done = True

                    if any(s in actor_up for s in SYSTEM_AGENTS): continue
                    
                    if QUALIFIED_STAGE in full_up:
                        master_report[proj_name][actor]["q_ids"].add(d_id)
                    
                    for stg in MEETING_STAGES:
                        if stg.upper() in full_up:
                            if (d_id, stg.upper()) not in master_report[proj_name][actor]["m_details"]:
                                master_report[proj_name][actor]["m_details"].add((d_id, stg.upper()))
                                master_report[proj_name][actor]["meet_count"] += 1
                
                print("-> Done")
            except: print("-> Error")
            driver.switch_to.default_content()

        # --- EXPORT ---
        if master_report:
            with pd.ExcelWriter(OUTPUT_PATH, engine='openpyxl') as writer:
                # Sort project names so 'Undefined' is last if it exists
                sorted_projects = sorted(master_report.keys(), key=lambda x: (x == "Undefined Project", x))
                for project in sorted_projects:
                    agents = master_report[project]
                    data_list = []
                    for agent, data in agents.items():
                        if not agent: continue
                        data_list.append({
                            "Agent Name": agent, "Claimed": data["claims"],
                            "Claimed IDs": ", ".join(sorted(list(data["c_ids"]))),
                            "Qualified": len(data["q_ids"]), "Qualified IDs": ", ".join(sorted(list(data["q_ids"]))),
                            "Meetings": data["meet_count"], 
                            "Meeting IDs": ", ".join(sorted(list(set([m[0] for m in data["m_details"]]))))
                        })
                    if data_list:
                        df = pd.DataFrame(data_list).sort_values(by="Claimed", ascending=False)
                        df.to_excel(writer, sheet_name=clean_sheet_name(project), index=False)
            print(f"\nReport successfully saved to {OUTPUT_PATH}")

    finally: driver.quit()

if __name__ == "__main__":
    scrape_unified_master_report()