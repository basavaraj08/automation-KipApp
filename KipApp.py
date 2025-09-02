import gspread
from google.oauth2.service_account import Credentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
import time
import re
import datetime
import difflib
from difflib import SequenceMatcher
from selenium.webdriver.chrome.service import Service


def normalize_text(text):
    """Hilangkan spasi berlebih dan kecilkan huruf."""
    if not isinstance(text, str):
        return ""
    return re.sub(r'\s+', ' ', text.strip()).lower()

def find_best_match(target, options):
    """Cari match terbaik dan kembalikan dengan rasio kemiripan."""
    best_match = None
    highest_ratio = 0

    for option in options:
        ratio = SequenceMatcher(None, target, option).ratio()
        if ratio > highest_ratio:
            highest_ratio = ratio
            best_match = option

    return best_match, highest_ratio

# ✅ Define Google Sheets API scopes
scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# ✅ Load Google Sheets credentials
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

# ✅ Open the Google Sheet
spreadsheet_url = "https://docs.google.com/spreadsheets/d/19dP73kPkAvKLqYbOlaWYevdtsBH4dcAPPovhW2Fp7SM/edit?usp=sharing"
sheet = client.open_by_url(spreadsheet_url).worksheet("dashboard")

# ✅ Fetch data from columns A to J (modify range if needed)
range_values = sheet.get("A:J")

# ✅ Skip the header row (Start from index 1)
data_rows = range_values[1:]

# ✅ Show total records
total_records = len(data_rows)
print(f"✅ Total Records: {total_records}")

# ✅ Skip the header row (Start from index 1)
data_rows = range_values[1:]  # This removes the first row (header)

# ✅ Show total records
total_records = len(data_rows)
print(f"✅ Total Records: {total_records}")    

# ✅ Set up Selenium WebDriver using your **existing Chrome profile**
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")  # Maximize window
options.add_argument("--log-level=3")

#options.add_argument("user-data-dir=C:/Users/gumel/AppData/Local/Google/Chrome/User Data")  
#options.add_argument("profile-directory=Profile 2")  # Use your profile

# ✅ Start Chrome with your profile
chrome_driver_path = "D:/WISNU/_dev/automation/chromedriver.exe"
service = Service(executable_path=chrome_driver_path)
driver = webdriver.Chrome(service=service, options=options)

# ✅ Open KiPapp page
driver.execute_script("document.body.style.zoom='80%'")
driver.get("https://webapps.bps.go.id/kipapp/#/pelaksanaan-aksi")
time.sleep(30)  # Wait for page to load

# ✅ Wait and Click "OK" in the Notification Modal
try:
    time.sleep(2)  # Give some time for the modal to appear
    ok_button = driver.find_element(By.XPATH, '//button[contains(@class, "ant-btn-primary") and span[text()="OK"]]')
    driver.execute_script("arguments[0].click();", ok_button)  # Click using JavaScript to bypass overlay issues
    time.sleep(1)  # Small delay after clicking
except Exception as e:
    print("⚠️ No notification modal detected, continuing...")

# ✅ Click the "Pilih SKP" dropdown to open the options
try:
    # Locate and click the "Pilih SKP" dropdown
    skp_dropdown = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, '//div[contains(@class, "ant-select-selection--single") and .//div[text()="Pilih SKP"]]'))
    )
    skp_dropdown.click()
    time.sleep(1)  # Wait for options to load

    # ✅ Select the "Juli" option
    bulan_option = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, '//li[contains(text(), "1 Januari - 31 Desember (Bulan Juli)")]')) #ubah bulan disini
    )
    bulan_option.click()

    time.sleep(1)  # Small delay after selection
except Exception as e:
    print(f"⚠️ Error selecting 'Juli' from dropdown: {e}")

for idx, row in enumerate(data_rows, start=1):
    print("=" * 50)
    print(f"➡️ Processing row {idx}/{len(data_rows)}")

    # Ambil ulang nilai dari setiap baris
    rencana_kinerja = row[1] if len(row) > 1 else ""  # Column B
    tanggal_awal = row[4] if len(row) > 4 else ""  # Column E
    tanggal_akhir = row[5] if len(row) > 5 else ""  # Column F
    deskripsi_kegiatan = row[0] if len(row) > 0 else ""  # Column A
    deskripsi_capaian = row[6] if len(row) > 6 else ""  # Column G
    link_data_dukung = row[7] if len(row) > 7 else ""  # Column H
    
    # ✅ Pilih "Rencana Kinerja SKP" sebelum klik tombol "Add"
    try:
        dropdown = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, '//div[contains(@class, "ant-select-selection--single") and .//div[text()="Pilih rencana kinerja SKP"]]'))
        )
        driver.execute_script("arguments[0].click();", dropdown)

        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "ant-select-dropdown") and not(contains(@style, "display: none"))]'))
        )
        time.sleep(1)

        options = driver.find_elements(By.XPATH, '//ul[@class="ant-select-dropdown-menu-item-group-list"]/li')
        dropdown_texts = [normalize_text(opt.text) for opt in options]
        normalized_input = normalize_text(rencana_kinerja)

        # ✅ Cek apakah match persis
        if normalized_input in dropdown_texts:
            idx = dropdown_texts.index(normalized_input)
            option = options[idx]
            driver.execute_script("arguments[0].scrollIntoView(true);", option)
            time.sleep(1)
            option.click()

            print(f"✅ GSheet: '{rencana_kinerja.strip()}'  →  Selected: '{option.text.strip()}' (100%)")

        else:
            best_match, similarity = find_best_match(normalized_input, dropdown_texts)
            percentage = round(similarity * 100, 2)

            if best_match and similarity > 0.6:  # threshold aman
                idx = dropdown_texts.index(best_match)
                option = options[idx]
                driver.execute_script("arguments[0].scrollIntoView(true);", option)
                time.sleep(1)
                option.click()

                print(f"⚠️ GSheet: '{rencana_kinerja.strip()}'  →  Selected: '{option.text.strip()}' ({percentage}%) [approx match]")
            else:
                print(f"❌ GSheet: '{rencana_kinerja.strip()}'  →  No suitable option found! (Best match: {percentage}%)")
                continue

    except Exception as e:
        print(f"⚠️ Error selecting 'Rencana Kinerja SKP': {e}")
        continue


    
    # ✅ Click the "Add" button to open the form
    try:
        add_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, '//button[contains(@class, "ant-btn-primary") and .//span[text()="Add"]]'))
        )
        add_button.click()
        time.sleep(2)  # Wait for form to load
    except Exception as e:
        print(f"⚠️ Error clicking 'Add' button: {e}")
        continue  # Skip this row if adding fails

    # ✅ Wait for the modal form to be fully loaded
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CLASS_NAME, "ant-modal-content"))
        )
    except Exception as e:
        print(f"⚠️ Modal form did not load: {e}")
        continue  # Skip if modal fails to load


    # ✅ Check "Gunakan Periode Tanggal"
    try:
        periode_checkbox = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, '//label[contains(.,"Gunakan periode tanggal")]'))
        )
        driver.execute_script("arguments[0].click();", periode_checkbox)
    except Exception as e:
        print(f"⚠️ Error checking 'Gunakan Periode Tanggal': {e}")

    # ✅ Check "Masukkan ke Capaian SKP"
    try:
        capaian_checkbox = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, '//input[@id="form-add_isCapaianSKP"]'))
        )

        # Cek apakah sudah dicentang
        is_checked = driver.execute_script("return arguments[0].checked;", capaian_checkbox)

        if not is_checked:
            driver.execute_script("arguments[0].click();", capaian_checkbox)
          
    except Exception as e:
        print(f"⚠️ Error checking 'Masukkan ke Capaian SKP': {e}")

    try:
        
        # ✅ Ambil tanggal (dd) dari string tanggal
        start_date_dd = str(int(tanggal_awal.split("/")[1]))  # Ambil bagian "DD"
        end_date_dd = str(int(tanggal_akhir.split("/")[1]))  # Ambil bagian "DD"

        # ✅ Klik input tanggal untuk membuka kalender
        tanggal_picker = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "form-add_tanggal"))
        )
        tanggal_picker.click()

        # ✅ Tunggu sampai kalender muncul
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CLASS_NAME, "ant-calendar-picker-container"))
        )

        time.sleep(1)  # Beri jeda agar kalender sepenuhnya dirender

        # ✅ Pilih tanggal awal
        start_date_xpath = f"//td[not(contains(@class, 'ant-calendar-disabled-cell'))]/div[text()='{start_date_dd}']"
        start_date_element = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, start_date_xpath))
        )
        start_date_element.click()

        time.sleep(1)

        # ✅ Pindah ke input tanggal akhir dengan "TAB"
        tanggal_picker.send_keys(Keys.TAB)
        time.sleep(1)

        # ✅ Pilih tanggal akhir
        end_date_xpath = f"//td[not(contains(@class, 'ant-calendar-disabled-cell'))]/div[text()='{end_date_dd}']"
        end_date_element = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, end_date_xpath))
        )
        end_date_element.click()
  
        # ✅ Klik di luar untuk menutup date picker
        body = driver.find_element(By.TAG_NAME, "body")
        ActionChains(driver).move_to_element(body).click().perform()

    except Exception as e:
        print(f"⚠️ Error selecting date range: {e}")
        driver.save_screenshot("error_general.png") 

    def clean_text(text):
        """Cleans text by removing extra spaces, tabs, and newlines."""
        if not isinstance(text, str):
            return ""  # Return empty if the data is not a string

        cleaned_text = re.sub(r'\s+', ' ', text).strip()  # Replace multiple spaces, remove tabs/newlines
        return cleaned_text

    def fast_input_text(driver, element_id, text_value, field_name):
        """Fast input text using Selenium, verifies accuracy, and retries once if necessary."""
        try:
            input_field = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, element_id))
            )

            # ✅ Clear existing text
            input_field.clear()
            time.sleep(0.5)  # Small delay for UI refresh

            # ✅ Fast input
            input_field.send_keys(text_value)

            # ✅ Verify input
            time.sleep(0.5)  # Let the UI update
            actual_value = input_field.get_attribute("value")
            actual_value = actual_value.strip() if actual_value else ""

            if actual_value == text_value.strip():
                return

            # 🚨 Retry once if the value didn't stick
            print(f"⚠️ Mismatch detected in '{field_name}', retrying...")
            input_field.clear()
            input_field.send_keys(text_value)
            time.sleep(0.5)

            # Final verification
            actual_value_retry = input_field.get_attribute("value")
            actual_value_retry = actual_value_retry.strip() if actual_value_retry else ""

            if actual_value_retry == text_value.strip():
                print(f"✅ Successfully entered '{field_name}' after retry: {actual_value_retry}")
            else:
                print(f"❌ Still failed! Saving screenshot...")
                driver.save_screenshot(f"error_{field_name.replace(' ', '_')}.png")

        except Exception as e:
            print(f"⚠️ Error entering '{field_name}': {e}")

    deskripsi_kegiatan = clean_text(deskripsi_kegiatan)
    deskripsi_capaian = clean_text(deskripsi_capaian)
    link_data_dukung = clean_text(link_data_dukung)

    # ✅ Fill "Deskripsi Kegiatan"
    fast_input_text(driver, "form-add_kegiatan", deskripsi_kegiatan, "Deskripsi Kegiatan")

    # ✅ Fill "Deskripsi Capaian"
    fast_input_text(driver, "form-add_capaian", deskripsi_capaian, "Deskripsi Capaian")

    # ✅ Fill "Link Data Dukung"
    fast_input_text(driver, "form-add_dataDukung", link_data_dukung, "Link Data Dukung")
        
    # ✅ Click "Save" button
    try:
        save_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, '//button[contains(@class, "ant-btn-primary") and .//span[text()="Save"]]'))
        )
        save_button.click()

        time.sleep(2)  # Wait for the form to submit
    except Exception as e:
        print(f"⚠️ Error clicking 'Save' button: {e}")
        break  # Stop if saving fails

    print("✅ Entry completed. Moving to next row...\n")

print("✅ Form filling completed for all rows!")
driver.quit()
