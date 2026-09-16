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

MEETING_STAGES = [
    "ONLINE - MEETINGS", 
    "FIRST MEETING", 
    "2ND SV/F2F", 
    "3RD SV/F2F", 
    "HOT PIPELINE",
    "DEAL WON"
]

QUALIFIED_STAGE = "QUALIFIED"
CLAIM_STAGE = "CLAIMED MARKETING LEAD"
OUTPUT_PATH = r"C:\Users\ASUS\Documents\Performance.xlsx"

SYSTEM_AGENTS = ["ASSET HOMEZ", "SYSTEM", "BITRIX24", "AUTOMATION"]

# --- NEW LOGIC: MANAGEMENT AGENTS ---
# These agents "pass through" claims to the actual working agents
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
    "meet_count": 0, 
    "m_details": set() 
})

def get_deal_id_from_url(url):
    try:
        return "".join(filter(str.isdigit, url.split('details/')[1].split('/')[0]))
    except:
        return None

def parse_history_row(row_element):
    row_text = row_element.text.strip()
    cells = row_element.find_elements(By.CSS_SELECTOR, ".main-grid-cell-content")
    cell_texts = [c.text.strip() for c in cells if c.text.strip()]
    
    agent, event, desc = "", "", ""
    if len(cell_texts) >= 4:
        if any(char.isdigit() for char in cell_texts[0]):
            agent, event = cell_texts[1], cell_texts[2]
            desc = " ".join(cell_texts[3:])
        else:
            agent, event = cell_texts[2], cell_texts[3]
            desc = " ".join(cell_texts[4:])
    
    if agent: agent = agent.split('\n')[0].strip()
    return agent, event, desc, row_text

def fetch_all_history_entries(driver):
    history_entries = []
    while True:
        for _ in range(25):
            try:
                more_btns = driver.find_elements(By.CSS_SELECTOR, ".main-grid-more-btn, .main-grid-nav-more, a.main-grid-more-btn-text")
                visible_btns = [b for b in more_btns if b.is_displayed()]
                if visible_btns:
                    driver.execute_script("arguments[0].click();", visible_btns[0])
                    time.sleep(1.2)
                else: break
            except: break
        
        rows = driver.find_elements(By.CSS_SELECTOR, ".main-grid-row.main-grid-row-body")
        for row in rows:
            try:
                agent, event, desc, full_txt = parse_history_row(row)
                if full_txt:
                    history_entries.append({"agent": agent, "event": event, "desc": desc, "full_text": full_txt})
            except: continue

        try:
            next_btns = driver.find_elements(By.CSS_SELECTOR, "a.main-grid-page-next")
            visible_next = [b for b in next_btns if b.is_displayed() and "main-grid-page-disabled" not in (b.get_attribute("class") or "")]
            if visible_next:
                driver.execute_script("arguments[0].click();", visible_next[0])
                time.sleep(2)
            else: break
        except: break
    return history_entries

def scrape_unified_master_report():
    try:
        all_unique_deal_ids = set()
        for url in CATEGORIES:
            driver.get(url)
            input(f"\n--- Category: {url} ---\nLog in/Filter and press ENTER...")
            last_deal_count, same_count_retries = 0, 0
            while True:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(1.5)
                links = driver.find_elements(By.XPATH, "//a[contains(@href, '/crm/deal/details/')]")
                current_ids = list(dict.fromkeys([get_deal_id_from_url(l.get_attribute('href')) for l in links if "details/" in l.get_attribute('href')]))
                try:
                    more_btn = driver.find_elements(By.CSS_SELECTOR, ".main-grid-more-btn")
                    if more_btn and more_btn[0].is_displayed():
                        driver.execute_script("arguments[0].click();", more_btn[0])
                        time.sleep(2)
                    elif len(current_ids) == last_deal_count: same_count_retries += 1
                    else: same_count_retries, last_deal_count = 0, len(current_ids)
                except: same_count_retries += 1
                if same_count_retries >= 2:
                    all_unique_deal_ids.update(current_ids)
                    break

        print(f"\nProcessing {len(all_unique_deal_ids)} Deals...")

        for index, d_id in enumerate(all_unique_deal_ids):
            print(f"[{index+1}/{len(all_unique_deal_ids)}] Deal {d_id}", end=" ", flush=True)
            try:
                driver.get(f"https://assethomezinternational.bitrix24.com/crm/deal/details/{d_id}/")
                iframe = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "side-panel-iframe")))
                driver.switch_to.frame(iframe)

                try:
                    h_tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@data-id='tab_history']|//a[contains(., 'History')]|//span[contains(text(), 'History')]")))
                    driver.execute_script("arguments[0].click();", h_tab)
                except:
                    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//*[contains(@class, 'main-buttons-item-more')]"))
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'History')]"))))

                time.sleep(1.5)
                history_chronological = list(reversed(fetch_all_history_entries(driver)))

                # --- TRACKING VARIABLES ---
                claimed_found_for_this_deal = False
                current_responsible = ""

                for i, entry in enumerate(history_chronological):
                    agent_acting = entry['agent'].strip()
                    agent_upper = agent_acting.upper()
                    if not agent_acting or any(sys_a in agent_upper for sys_a in SYSTEM_AGENTS): continue

                    full_text_upper = entry['full_text'].upper()
                    event_upper = entry['event'].upper()
                    desc_upper = entry['desc'].upper()

                    # Update Current Responsible person tracking
                    if "RESPONSIBLE" in event_upper or "RESPONSIBLE" in desc_upper:
                        if "→" in entry['desc']: current_responsible = entry['desc'].split("→")[-1].strip()
                        elif "->" in entry['desc']: current_responsible = entry['desc'].split("->")[-1].strip()

                    # 1. UPDATED CLAIM LOGIC
                    is_stage_event = ("STAGE CHANGED" in event_upper) or ("STAGE CHANGED" in full_text_upper) or ("→" in desc_upper)
                    if is_stage_event and not claimed_found_for_this_deal:
                        target_stage = desc_upper.split("→")[-1].strip() if "→" in desc_upper else desc_upper
                        
                        if CLAIM_STAGE in target_stage or "CLAIMED" in target_stage:
                            final_agent_to_credit = agent_acting
                            
                            # Check if the person claiming is Management
                            if agent_acting.upper() in [m.upper() for m in MANAGEMENT_AGENTS]:
                                # Case 1: Look ahead for a "Responsible" change in next 2 entries (Re-assign case)
                                found_new_agent = False
                                for look_ahead in history_chronological[i+1 : i+3]:
                                    if "RESPONSIBLE" in look_ahead['event'].upper() or "RESPONSIBLE" in look_ahead['desc'].upper():
                                        if "→" in look_ahead['desc']:
                                            final_agent_to_credit = look_ahead['desc'].split("→")[-1].strip()
                                            found_new_agent = True
                                            break
                                
                                # Case 2: No re-assignment found, credit the person who was assigned BEFORE the claim
                                if not found_new_agent and current_responsible:
                                    final_agent_to_credit = current_responsible

                            # Attribute credit
                            master_report[final_agent_to_credit]["claims"] += 1
                            master_report[final_agent_to_credit]["c_ids"].add(d_id)
                            claimed_found_for_this_deal = True

                    # 2. QUALIFIED LOGIC
                    if QUALIFIED_STAGE in full_text_upper or QUALIFIED_STAGE in desc_upper:
                        master_report[agent_acting]["q_ids"].add(d_id)
                    
                    # 3. MEETING LOGIC
                    if "ASSIGNED MEETING" not in full_text_upper:
                        for stage in MEETING_STAGES:
                            if stage.upper() in full_text_upper or stage.upper() in desc_upper:
                                if (d_id, stage.upper()) not in master_report[agent_acting]["m_details"]:
                                    master_report[agent_acting]["m_details"].add((d_id, stage.upper()))
                                    master_report[agent_acting]["meet_count"] += 1
                
                print("-> Done")
            except Exception as e:
                print(f"-> Error: {str(e)[:30]}")
            driver.switch_to.default_content()

        # --- EXPORT ---
        final_list = []
        for agent, data in master_report.items():
            if not agent: continue
            involved_m_ids = sorted(list(set([item[0] for item in data["m_details"]])))
            final_list.append({
                "Agent Name": agent,
                "Claimed Count": data["claims"],
                "Claimed Deal IDs": ", ".join(sorted(list(data["c_ids"]))),
                "Qualified Count": len(data["q_ids"]),
                "Qualified Deal IDs": ", ".join(sorted(list(data["q_ids"]))),
                "Total Meeting Transitions": data["meet_count"],
                "Meeting Deal IDs": ", ".join(involved_m_ids)
            })

        if final_list:
            df = pd.DataFrame(final_list).sort_values(by=["Claimed Count"], ascending=False)
            df.to_excel(OUTPUT_PATH, index=False)
            print(f"\nSUCCESS: Saved to {OUTPUT_PATH}")

    finally: driver.quit()

if __name__ == "__main__":
    scrape_unified_master_report()