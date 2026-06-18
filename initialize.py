import logging, tkinter as tk, winsound, io, time
from tkinter import ttk
from PIL import Image, ImageTk
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

#user validation gui pop up
def user_validation_popup(driver):
    def show_gui_and_handle():
        nonlocal driver
        popup = tk.Tk()
        popup.title("Captcha Required")
        popup.attributes('-topmost', True)
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", lambda: None)  # Disable window close or (X)
        screen_width = popup.winfo_screenwidth()
        screen_height = popup.winfo_screenheight()
        x_position = (screen_width - 230) // 2
        y_position = (screen_height - 215) // 2
        popup.geometry(f'230x215+{x_position}+{y_position}') #make gui always center
        ttk.Label(popup, text="User Validation Required", font=("Calibri", 11, "bold")).pack()
        img_label = ttk.Label(popup)
        img_label.pack(pady=2)
        attempts_label = ttk.Label(popup, text="")
        attempts_label.pack(pady=2)
        feedback_label = ttk.Label(popup, text="")
        feedback_label.pack()
        resample = getattr(Image, "Resampling", Image)

        def load_captcha(): #capture captcha image from browser(screeshot)
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//img[contains(@src, "captcha")]'))
                )
                captcha_elem = driver.find_element(By.XPATH, '//img[contains(@src, "captcha")]')
                captcha_png = captcha_elem.screenshot_as_png #screenshots the captcha form the browser
                image = Image.open(io.BytesIO(captcha_png)).resize((200, 70), resample.LANCZOS)
                photo = ImageTk.PhotoImage(image)
                img_label.configure(image=photo)
                img_label.image = photo

                try:
                    em_elem = driver.find_element(By.XPATH, '//em[contains(text(), "Number of attempts left")]')
                    text = em_elem.text
                    attempts_label.config(text=text)
                    logging.info(text)
                except NoSuchElementException:
                    attempts_label.config(text="")

                logging.info("Captcha image loaded.")

            except Exception as e:
                logging.error(f"Error loading captcha: {e}")
                feedback_label.config(text="Failed to load captcha.")
                with open("debug_captcha_page.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)

        def on_submit():
            user_input = entry.get().strip()
            if not user_input:
                feedback_label.config(text="Please enter the captcha.", foreground="red")
                return

            try:
                popup.withdraw()
                driver.find_element(By.NAME, "captcha_resp_txt").clear()
                driver.find_element(By.NAME, "captcha_resp_txt").send_keys(user_input)
                driver.find_element(By.XPATH, '//input[@type="submit" and @value="Submit"]').click()
                time.sleep(2)
                try:
                    err_div = driver.find_element(By.XPATH, '//div[contains(text(), "The specified URL is inaccessible")]')
                    if err_div:
                        logging.warning("Page inaccessible. Retrying..")
                        driver.back()
                        time.sleep(2)
                        popup.deiconify()
                        load_captcha()
                        feedback_label.config(text="Site inaccessible. Try again.", foreground="red")
                        return
                except NoSuchElementException:
                    pass
                try:
                    header = driver.find_element(By.XPATH, '/html/body/h3').text
                    if "User validation required" in header:
                        feedback_label.config(text="Captcha incorrect. Try again.", foreground="red")
                        popup.deiconify()
                        load_captcha()
                        return
                except NoSuchElementException:
                    pass
                logging.info("Captcha accepted.")
                popup.destroy()
            except Exception as e:
                logging.error(f"Unexpected error: {e}")
                feedback_label.config(text="Unexpected error. Try again.", foreground="red")

        def on_refresh():
            try:
                driver.refresh()
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//img[contains(@src, "captcha")]'))
                )
                load_captcha()
                feedback_label.config(text="Captcha refreshed.", foreground="blue")
            except Exception as e:
                logging.error(f"Refresh failed: {e}")
                feedback_label.config(text="Refresh failed.", foreground="red")

        load_captcha()
        entry_frame = ttk.Frame(popup)
        entry_frame.pack(pady=5)
        entry = ttk.Entry(entry_frame)
        entry.pack(side="left", padx=(0, 0))
        ttk.Button(entry_frame, text="⟳", command=on_refresh, width=3).pack(side="left")
        ttk.Button(popup, text="Submit", command=on_submit, width=20).pack()
        winsound.MessageBeep(winsound.MB_ICONEXCLAMATION)
        popup.transient()
        popup.grab_set()
        popup.wait_window(popup)

    show_gui_and_handle()
    return True