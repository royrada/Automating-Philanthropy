import csv
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 🧭 Set up Chrome WebDriver
options = Options()
# Uncomment below to run headlessly
# options.add_argument('--headless')
options.add_argument('--disable-gpu')
driver = webdriver.Chrome(options=options)

# 💾 Create new CSV file and write header
with open("scholarships.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow([
        "ID", "Award Name", "Organization", "Purpose",
        "Level of Study", "Award Type", "Award Amount", "Deadline"
    ])

    # 🔁 Loop through 20 pages
    for page_num in range(1, 21):
        print(f"Scraping page {page_num}...")
        url = f"https://www.careeronestop.org/Toolkit/Training/find-scholarships.aspx?&curpage={page_num}&pagesize=500"
        driver.get(url)

        try:
            WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CLASS_NAME, "cos-table-responsive"))
            )
        except Exception as e:
            print(f"[!] Timeout loading table on page {page_num}: {e}")
            continue

        # 📦 Scrape rows
        rows = driver.find_elements(By.CSS_SELECTOR, "table.cos-table-responsive tbody tr")
        print(f" → Found {len(rows)} rows.")

        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) < 5:
                continue

            award_td = cells[0]

            try:
                award_link = award_td.find_element(By.CSS_SELECTOR, ".detailPageLink a")
                award_name = award_link.text.strip()
                href = award_link.get_attribute("href")
                award_id = href.split("scholarshipId=")[-1]
            except:
                award_name = ""
                award_id = ""

            inner_divs = award_td.find_elements(By.XPATH, "./div/div")
            org_text = ""
            purpose_text = ""
            for div in inner_divs:
                div_text = div.text.strip()
                if div_text.startswith("Organization:"):
                    org_text = div_text.replace("Organization:", "").strip()
                elif div_text.startswith("Purposes:"):
                    purpose_text = div_text.replace("Purposes:", "").strip()

            level_of_study = cells[1].text.strip()
            award_type = cells[2].text.strip()
            award_amount = cells[3].text.strip()
            deadline = cells[4].text.strip()

            writer.writerow([
                award_id, award_name, org_text, purpose_text,
                level_of_study, award_type, award_amount, deadline
            ])

        # 💤 Pause briefly before next page
        time.sleep(2)

driver.quit()
print("✅ All pages scraped. Data saved to scholarships.csv")
