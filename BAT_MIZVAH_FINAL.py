
import time
from datetime import datetime
from selenium.webdriver.support import expected_conditions as EC

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys

def format_phone_number(phone_number: str) -> str:
    phone_number = str(phone_number).strip().replace(" ", "").replace("-", "")
    if phone_number.startswith("0"):
        phone_number = "972" + phone_number[1:]
    elif phone_number.startswith("+972"):
        phone_number = phone_number[1:]
    return phone_number

def send_whatsapp_message(phone_number: str, message: str,image_path: str):
    formatted_phone = format_phone_number(phone_number)
    url = f"https://web.whatsapp.com/send?phone={formatted_phone}&text={message}"

    options = Options()
    options.add_experimental_option("debuggerAddress", "localhost:9222")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    driver.get(url)
    time.sleep(12)  #LOADING CHAT....

    try:
        
        input_box = driver.find_element(By.XPATH, "//div[@contenteditable='true'][@data-tab='10']")
        input_box.send_keys(Keys.ENTER)
        print(f"Message sent to: {phone_number}")
    except Exception as e:
        print(f"ERROR IN SENDING - {phone_number}: {e}")
        return False

    time.sleep(5)

    try:
        attach_btn = driver.find_element(By.XPATH, '//div[@aria-label="Attach"]')
        attach_btn.click()
        time.sleep(5)
        image_box = driver.find_element(By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
        image_box.send_keys(image_path)
        time.sleep(5)
        send_button = driver.find_element(By.XPATH, '//span[@data-icon="wds-ic-send-filled"]')
        time.sleep(5)
        send_button.click()
        time.sleep(5)

        print(f"Message sent to: {phone_number}")
    except Exception as e:
        print(f"ERROR IN SENDING - {phone_number}: {e}")
        return False

    time.sleep(5)
    driver.quit()
    return True


if __name__ == "__main__":
   
   # READING FROM EXCEL
   file_path = "sample.xlsx"
   df = pd.read_excel("sample.xlsx", header=[0, 1])  

   print(df.columns.tolist())
   
   phones = df[('פרטי התקשרות', 'סלולרי')].dropna().astype(str)#.tolist()
   names = df[('Unnamed: 0_level_0', 'הזמנה לכבוד')]

   message = "🎉 BAT MITZVAH PARTY 🎉\nShani Levy is celebrating! ✨\nDon't forget to bring good vibes! 💃🕺\nPhoto attached 💌"
   image_path = r""#FILL IN HERE YOUR IMAGE PATH!
   log_rows = []

  
   for idx, phone in phones.items():
       success = send_whatsapp_message(phone, message, image_path)
       if success:
           send_time = datetime.now().strftime("%Y-%m-%d %H:%M")
           name = names.get(idx, "לא ידוע")
           log_rows.append({
               'שם': name,
               'מספר טלפון': phone,
               'תאריך שליחה': send_time
           })

   if log_rows:
       log_df = pd.DataFrame(log_rows)
       log_df.to_excel("sent_log.xlsx", index=False)
       print("FILE SAVED sent_log.xlsx ")
   else:
       print("NO NEW MESSAGES WERE SENT")





