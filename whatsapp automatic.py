"""
Bulk WhatsApp Sender — Fully Automatic
---------------------------------------
pip install selenium webdriver-manager pandas openpyxl
"""

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import threading
import time
import os

COUNTRY_CODES = {
    "pakistan": "92", "india": "91", "usa": "1", "uk": "44",
    "uae": "971", "saudi arabia": "966", "bangladesh": "880",
    "afghanistan": "93", "iran": "98", "turkey": "90",
    "germany": "49", "france": "33", "canada": "1", "australia": "61",
}

PROFILE_DIR = os.path.join(os.path.expanduser("~"), "whatsapp_chrome_profile")
excel_file = ""


def format_phone_number(phone: str, country: str) -> str:
    digits = "".join(filter(str.isdigit, phone)).lstrip("0")
    code = COUNTRY_CODES.get(country.lower().strip(), "")
    if not code:
        raw = "".join(filter(str.isdigit, country))
        code = raw if raw else "92"
    return ("+" + digits) if digits.startswith(code) else ("+" + code + digits)


def create_driver():
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from webdriver_manager.chrome import ChromeDriverManager

    options = Options()
    options.add_argument(f"--user-data-dir={PROFILE_DIR}")
    options.add_argument("--profile-directory=Default")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver


def wait_for_login(driver):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    log_message("⏳ WhatsApp Web load ho raha hai...\n")
    driver.get("https://web.whatsapp.com")
    driver.maximize_window()

    try:
        WebDriverWait(driver, 120).until(
            EC.presence_of_element_located((By.XPATH, '//div[@aria-label="Chat list"]'))
        )
        log_message("✅ Login successful!\n\n")
        return True
    except Exception:
        log_message("✗ Login timeout. Dobara try karein.\n")
        return False


def find_input_box(driver):
    """Try multiple XPaths — WhatsApp Web updates karta rehta hai."""
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    # Multiple selectors — pehla jo kaam kare use karo
    xpaths = [
        '//div[@contenteditable="true"][@data-tab="10"]',          # old
        '//div[@contenteditable="true"][@data-tab="1"]',           # some versions
        '//div[@title="Type a message"]',                          # title based
        '//footer//div[@contenteditable="true"]',                  # footer area
        '//div[@role="textbox"][@data-tab="10"]',                  # role based
        '//div[@spellcheck="true"][@contenteditable="true"]',      # spellcheck
        '//div[contains(@class,"copyable-text") and @contenteditable="true"]',
        '//div[@aria-placeholder="Type a message"]',               # aria placeholder
        '//div[@aria-label="Type a message"]',                     # aria label
    ]

    wait = WebDriverWait(driver, 35)
    for xpath in xpaths:
        try:
            box = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            if box and box.is_displayed():
                return box
        except Exception:
            continue

    # Last resort — CSS selector
    try:
        box = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, 'div[contenteditable="true"][data-lexical-editor="true"]')
        ))
        return box
    except Exception:
        pass

    raise Exception("Message input box nahi mila. WhatsApp Web update ho gaya hoga.")


def send_whatsapp_message(driver, phone: str, message: str):
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    import urllib.parse

    encoded_msg = urllib.parse.quote(message)
    url = f"https://web.whatsapp.com/send?phone={phone}&text={encoded_msg}"
    driver.get(url)

    # Wait for page to settle
    time.sleep(4)

    # Check for "invalid number" popup
    try:
        ok_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, '//div[@role="button"][.//text()="OK"]'))
        )
        ok_btn.click()
        raise Exception(f"Invalid number: {phone}")
    except Exception as e:
        if "Invalid number" in str(e):
            raise

    # Find input and send
    input_box = find_input_box(driver)
    time.sleep(1)
    input_box.send_keys(Keys.ENTER)
    time.sleep(3)

    # Verify message was sent (chat should still be open)
    try:
        WebDriverWait(driver, 8).until(
            EC.presence_of_element_located((By.XPATH, '//div[@data-id]'))
        )
    except Exception:
        pass  # non-critical check


def send_bulk_messages():
    global excel_file

    if not excel_file:
        root.after(0, lambda: messagebox.showerror("Error", "Please select an Excel file."))
        return

    template         = message_box.get("1.0", tk.END).strip()
    selected_country = country_var.get().strip()

    try:
        from selenium import webdriver
    except ImportError:
        root.after(0, lambda: messagebox.showerror(
            "Missing", "pip install selenium webdriver-manager"
        ))
        return

    try:
        df = pd.read_excel(excel_file)
        df.columns = df.columns.str.strip()

        for col in ["Name", "Phone", "From"]:
            if col not in df.columns:
                root.after(0, lambda c=col: messagebox.showerror(
                    "Error", f"Column '{c}' nahi mili. Required: Name, Phone, From"
                ))
                return

        total = len(df)
        driver = None

        try:
            log_message("🌐 Chrome open ho raha hai...\n")
            driver = create_driver()

            if not wait_for_login(driver):
                return

            success_count = 0
            fail_count    = 0

            for index, row in df.iterrows():
                name = "Unknown"
                try:
                    name   = str(row["Name"]).strip()
                    phone  = str(row["Phone"]).strip()
                    sender = str(row["From"]).strip()

                    formatted = format_phone_number(phone, selected_country)
                    msg = template.replace("{Name}", name).replace("{From}", sender)

                    log_message(f"[{index+1}/{total}] {name} → {formatted}\n")

                    send_whatsapp_message(driver, formatted, msg)

                    log_message(f"✓ Sent to {name}\n\n")
                    success_count += 1
                    time.sleep(8)

                except Exception as e:
                    err_short = str(e).split("Stacktrace")[0].strip()
                    log_message(f"✗ Failed for {name}: {err_short}\n\n")
                    fail_count += 1
                    time.sleep(3)

            root.after(0, lambda: messagebox.showinfo(
                "Done!",
                f"Complete!\n\n✓ Success: {success_count}\n✗ Failed: {fail_count}\nTotal: {total}"
            ))

        finally:
            if driver:
                driver.quit()
                log_message("🔒 Browser band.\n")

    except Exception as e:
        root.after(0, lambda err=e: messagebox.showerror("Error", str(err)))


def log_message(msg):
    root.after(0, lambda: (log.insert(tk.END, msg), log.see(tk.END)))

def select_file():
    global excel_file
    excel_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
    if excel_file:
        file_label.config(text=excel_file)

def start_thread():
    send_btn.config(state=tk.DISABLED, text="⏳ Sending...")
    def run():
        send_bulk_messages()
        root.after(0, lambda: send_btn.config(state=tk.NORMAL, text="▶  Start Sending"))
    threading.Thread(target=run, daemon=True).start()


# ── GUI ──────────────────────────────────────
root = tk.Tk()
root.title("Bulk WhatsApp Sender")
root.geometry("800x730")

tk.Label(root, text="Bulk WhatsApp Messaging System", font=("Arial", 16, "bold")).pack(pady=10)
tk.Label(root,
    text="✦ Pehli baar QR scan karo — session save ho jaega, baad mein auto login",
    font=("Arial", 9), fg="#007700", bg="#eaffea", relief="groove", padx=6, pady=3
).pack(fill="x", padx=20)

tk.Button(root, text="📂  Select Excel File", command=select_file, width=22).pack(pady=(12,0))
file_label = tk.Label(root, text="No file selected", wraplength=720, fg="gray")
file_label.pack(pady=3)

cf = tk.Frame(root); cf.pack(pady=6)
tk.Label(cf, text="Default Country:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=5)
country_var = tk.StringVar(value="Pakistan")
tk.OptionMenu(cf, country_var,
    "Pakistan","India","USA","UK","UAE","Saudi Arabia","Bangladesh",
    "Afghanistan","Iran","Turkey","Germany","France","Canada",
    "Australia","China","Russia","Qatar","Kuwait","Bahrain","Oman"
).pack(side=tk.LEFT)
tk.Label(cf, text="(used when number has no country code)",
         font=("Arial", 8), fg="gray").pack(side=tk.LEFT, padx=6)

tk.Label(root, text="Message Template", font=("Arial", 10, "bold")).pack()
message_box = tk.Text(root, width=88, height=8)
message_box.pack()
message_box.insert(tk.END,
"""Assalamualaikum {Name},

Hope you are doing well.

Regards,
{From}""")

send_btn = tk.Button(root, text="▶  Start Sending", command=start_thread,
                     font=("Arial", 12), width=22, bg="#25D366", fg="white")
send_btn.pack(pady=10)

tk.Label(root, text="Activity Log", font=("Arial", 10, "bold")).pack()
log = scrolledtext.ScrolledText(root, width=95, height=16)
log.pack(pady=5)

root.mainloop()