# SSACODE-CASTE-AND-SCHOOL-NAME
# this code provide caste and school name using ssa code
import time
import pandas as pd

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager

# =====================================================
# CONFIGURATION
# =====================================================

INPUT_FILE = r"C:\Users\acer\Downloads\RptStudentList.xlsx"
OUTPUT_FILE = r"C:\Users\acer\Downloads\RptStudentList_Updated.xlsx"

SEARCH_URL = "https://www.ssgujarat.org/CTELogin.aspx"

UID_COLUMN = "Student UID"
CASTE_COLUMN = "Caste"
SCHOOL_COLUMN = "School Name"

SAVE_EVERY = 25

# =====================================================
# FUNCTIONS
# =====================================================

def normalize_caste(value):

    value = str(value).strip().upper()

    if value in ["", "-", "--", "NAN"]:
        return "No Record Found"

    if value == "GENERAL":
        return "GENERAL"

    if value == "SC":
        return "SC"

    if value == "ST":
        return "ST"

    if "SEBC" in value:
        return "OBC"

    if "OBC" in value:
        return "OBC"

    return value


# =====================================================
# LOAD EXCEL
# =====================================================

print("=" * 60)
print("Loading Excel...")
print("=" * 60)

df = pd.read_excel(
    INPUT_FILE,
    header=1
)

# Remove unwanted spaces from headers

df.columns = df.columns.astype(str).str.strip()

print("\nColumns Found\n")

for c in df.columns:
    print(c)

# Verify required columns

if UID_COLUMN not in df.columns:
    raise Exception(
        f"\nColumn '{UID_COLUMN}' not found.\n\nAvailable Columns:\n{list(df.columns)}"
    )

if CASTE_COLUMN not in df.columns:
    df[CASTE_COLUMN] = ""

if SCHOOL_COLUMN not in df.columns:
    df[SCHOOL_COLUMN] = ""

print("\nTotal Records :", len(df))

# =====================================================
# EDGE DRIVER
# =====================================================

options = Options()

options.add_argument("--start-maximized")

driver = webdriver.Edge(
    service=Service(
        EdgeChromiumDriverManager().install()
    ),
    options=options
)

wait = WebDriverWait(driver, 60)

driver.get(SEARCH_URL)

print("\nWebsite Opened Successfully")
print(driver.title)

# =====================================================
# START LOOP
# =====================================================

for index, row in df.iterrows():

    uid = str(row[UID_COLUMN]).strip()

    if uid == "" or uid.lower() == "nan":

        df.at[index, CASTE_COLUMN] = "No Record Found"
        df.at[index, SCHOOL_COLUMN] = "No Record Found"

        continue

    print("\n")
    print("=" * 60)
    print(f"Record {index+1} / {len(df)}")
    print("UID :", uid)

    try:

        # always return to first tab

        driver.switch_to.window(driver.window_handles[0])

        driver.get(SEARCH_URL)

        # Search Box

        search_box = wait.until(
            EC.visibility_of_element_located(
                (By.ID, "TxtSearch")
            )
        )

        search_box.clear()

        search_box.send_keys(uid)

        # Click GO

        driver.find_element(
            By.ID,
            "BtnSearch"
        ).click()

        # Wait for new tab

        wait.until(
            lambda d: len(d.window_handles) > 1
        )

        main_tab = driver.window_handles[0]

        profile_tab = driver.window_handles[-1]

        driver.switch_to.window(profile_tab)

        # Wait for Profile Page

        wait.until(
            EC.presence_of_element_located(
                (By.ID, "lblsocialcategory")
            )
        )
                # ==========================================
        # GET SCHOOL NAME
        # ==========================================

        try:
            school_name = driver.find_element(
                By.ID,
                "lblschoolname"
            ).text.strip()

            if school_name in ["", "-", "--"]:
                school_name = "No Record Found"

        except:
            school_name = "No Record Found"

        # ==========================================
        # GET SOCIAL CATEGORY
        # ==========================================

        try:
            social_category = driver.find_element(
                By.ID,
                "lblsocialcategory"
            ).text.strip()

        except:
            social_category = "-"

        caste = normalize_caste(social_category)

        print("School :", school_name)
        print("Caste  :", caste)

        # ==========================================
        # UPDATE DATAFRAME
        # ==========================================

        df.at[index, SCHOOL_COLUMN] = school_name
        df.at[index, CASTE_COLUMN] = caste

        # ==========================================
        # CLOSE PROFILE TAB
        # ==========================================

        driver.close()

        driver.switch_to.window(main_tab)

        # ==========================================
        # SAVE EVERY 25 RECORDS
        # ==========================================

        if (index + 1) % SAVE_EVERY == 0:

            df.to_excel(
                OUTPUT_FILE,
                index=False
            )

            print()
            print("=" * 60)
            print(f"Progress Saved : {index + 1}")
            print("=" * 60)

        time.sleep(1)

    except Exception as e:

        print()
        print("ERROR")
        print(e)

        df.at[index, CASTE_COLUMN] = "No Record Found"
        df.at[index, SCHOOL_COLUMN] = "No Record Found"

        try:

            while len(driver.window_handles) > 1:

                driver.switch_to.window(
                    driver.window_handles[-1]
                )

                driver.close()

            driver.switch_to.window(
                driver.window_handles[0]
            )

        except:
            pass

        time.sleep(2)
        # =====================================================
# FINAL SAVE
# =====================================================

print("\n")
print("=" * 60)
print("Saving Excel...")
print("=" * 60)

df.to_excel(
    OUTPUT_FILE,
    index=False
)

driver.quit()

print("\n")
print("=" * 60)
print("PROCESS COMPLETED")
print("=" * 60)

print(f"Output File : {OUTPUT_FILE}")
print(f"Total Records : {len(df)}")

print("\nDone.")
