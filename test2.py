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

# --- CONFIGURATION ---
CATEGORIES = [
    "https://assethomezinternational.bitrix24.com/crm/deal/category/14/",
    "https://assethomezinternational.bitrix24.com/crm/deal/category/16/"
]

MEETING_STAGES = ["ONLINE - MEETINGS", "FIRST MEETING", "2ND SV/F2F", "3RD SV/F2F", "HOT PIPELINE", "DEAL WON"]
QUALIFIED_STAGE = "QUALIFIED"
CLAIM_STAGE = "CLAIMED MARKETING LEAD"
OUTPUT_PATH = r"C:\Users\ASUS\Documents\Performance.xlsx"

SYSTEM_AGENTS = ["ASSET HOMEZ", "SYSTEM", "BITRIX24", "AUTOMATION"]
MANAGEMENT_AGENTS = ["ANVAR SADATH", "MUHAMMED ASHFAQUE M", "TEENA THOMAS"]

chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_experimental_option("prefs", {"profile.managed_default_content_settings.images": 2})
chrome_options.add_experimental_option("detach", True)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
driver.implicitly_wait(0.5)
wait = WebDriverWait(driver, 12)

master_report = defaultdict(lambda: {
    "claims": 0, "c_ids": set(), 
    "qual": 0, "q_ids": set(), 
    "meet_count": 0, "m_details": set() 
})

def get_deal_id_from_url(url):
    try: return "".join(filter(str.isdigit, url.split('details/')[1].split('/')[0]))
    except: return None

def parse_history_row(row_element):
    row_text = row_element.text.strip()
    cells = row_element.find_elements(By.CSS_SELECTOR, ".main-grid-cell-content")
    cell_texts = [c.text.strip() for c in cells if c.text.strip()]
    agent, event, desc = "", "", ""
    if len(cell_texts) >= 4:
        if any(char.isdigit() for char in cell_texts[0]):
            agent, event, desc = cell_texts[1], cell_texts[2], " ".join(cell_texts[3:])
        else:
            agent, event, desc = cell_texts[2], cell_texts[3], " ".join(cell_texts[4:])
    if agent: agent = agent.split('\n')[0].strip()
    return agent, event, desc, row_text

def fetch_all_history_entries(driver):
    history_entries = []
    while True:
        for _ in range(20):
            try:
                more_btns = driver.find_elements(By.CSS_SELECTOR, ".main-grid-more-btn, .main-grid-nav-more")
                visible = [b for b in more_btns if b.is_displayed()]
                if visible: driver.execute_script("arguments[0].click();", visible[0]); time.sleep(1.2)
                else: break
            except: break
        rows = driver.find_elements(By.CSS_SELECTOR, ".main-grid-row.main-grid-row-body")
        for row in rows:
            try:
                a, e, d, f = parse_history_row(row)
                if f: history_entries.append({"agent": a, "event": e, "desc": d, "full_text": f})
            except: continue
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
        for url in CATEGORIES:
            driver.get(url)
            input(f"\nApply filters on {url} and press ENTER...")
            last_count = 0
            while True:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                links = driver.find_elements(By.XPATH, "//a[contains(@href, '/crm/deal/details/')]")
                current_ids = list(dict.fromkeys([get_deal_id_from_url(l.get_attribute('href')) for l in links]))
                if len(current_ids) == last_count: break
                last_count = len(current_ids); time.sleep(1.5)
            all_unique_deal_ids.update(current_ids)

        for index, d_id in enumerate(all_unique_deal_ids):
            print(f"[{index+1}/{len(all_unique_deal_ids)}] Deal {d_id}", end=" ", flush=True)
            try:
                driver.get(f"https://assethomezinternational.bitrix24.com/crm/deal/details/{d_id}/")
                iframe = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "side-panel-iframe")))
                driver.switch_to.frame(iframe)
                try:
                    h_tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-id='tab_history']|//a[contains(., 'History')]")))
                    driver.execute_script("arguments[0].click();", h_tab)
                except:
                    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//*[contains(@class, 'more')]"))
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//span[contains(text(), 'History')]"))
                
                time.sleep(1.5)
                # Read oldest entries first to track responsible person changes correctly
                history_chronological = list(reversed(fetch_all_history_entries(driver)))

                current_responsible = ""
                claimed_done = False

                for i, entry in enumerate(history_chronological):
                    actor = entry['agent'].strip()
                    actor_upper = actor.upper()
                    event_upper = entry['event'].upper()
                    desc_upper = entry['desc'].upper()
                    full_upper = entry['full_text'].upper()

                    # TRACK RESPONSIBLE PERSON (Always, even if done by SYSTEM/ASSET HOMEZ)
                    if "RESPONSIBLE" in event_upper or "RESPONSIBLE" in desc_upper:
                        if "→" in entry['desc']: current_responsible = entry['desc'].split("→")[-1].strip()
                        elif "->" in entry['desc']: current_responsible = entry['desc'].split("->")[-1].strip()

                    # Logic for Claimed Stage
                    if not claimed_done and (CLAIM_STAGE in full_upper or "CLAIMED" in desc_upper):
                        credited_agent = actor
                        
                        # Apply Management Logic
                        if actor_upper in [m.upper() for m in MANAGEMENT_AGENTS]:
                            # Case 1: Look ahead for a NEW assignment within next 2 actions
                            new_assignee = ""
                            for look_ahead in history_chronological[i+1 : i+3]:
                                if "RESPONSIBLE" in look_ahead['event'].upper():
                                    if "→" in look_ahead['desc']: new_assignee = look_ahead['desc'].split("→")[-1].strip()
                                    break
                            
                            # Case 2: No new assignment, credit the existing responsible person (e.g. Sampat Shetty)
                            credited_agent = new_assignee if new_assignee else current_responsible
                        
                        if credited_agent and not any(sys in credited_agent.upper() for sys in SYSTEM_AGENTS):
                            master_report[credited_agent]["claims"] += 1
                            master_report[credited_agent]["c_ids"].add(d_id)
                            claimed_done = True

                    # Regular Metrics (Skip system agents for these)
                    if any(sys in actor_upper for sys in SYSTEM_AGENTS): continue
                    
                    if QUALIFIED_STAGE in full_upper:
                        master_report[actor]["q_ids"].add(d_id)
                    
                    for stage in MEETING_STAGES:
                        if stage.upper() in full_upper:
                            if (d_id, stage.upper()) not in master_report[actor]["m_details"]:
                                master_report[actor]["m_details"].add((d_id, stage.upper()))
                                master_report[actor]["meet_count"] += 1
                
                print("-> Done")
            except Exception as e: print(f"-> Error: {str(e)[:30]}")
            driver.switch_to.default_content()

        # Export
        final_list = []
        for agent, data in master_report.items():
            if not agent: continue
            ids = sorted(list(set([item[0] for item in data["m_details"]])))
            final_list.append({
                "Agent Name": agent, "Claimed Count": data["claims"],
                "Claimed Deal IDs": ", ".join(sorted(list(data["c_ids"]))),
                "Qualified Count": len(data["q_ids"]), "Qualified Deal IDs": ", ".join(sorted(list(data["q_ids"]))),
                "Total Meeting Transitions": data["meet_count"], "Meeting Deal IDs": ", ".join(ids)
            })
        if final_list:
            pd.DataFrame(final_list).sort_values(by="Claimed Count", ascending=False).to_excel(OUTPUT_PATH, index=False)
            print(f"\nReport Saved to {OUTPUT_PATH}")

    finally: driver.quit()

if __name__ == "__main__":
    scrape_unified_master_report()