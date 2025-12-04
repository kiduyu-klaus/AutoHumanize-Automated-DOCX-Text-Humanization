import time
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import random
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from docx import Document
import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock
from docx.shared import Cm

from typing import Optional, Tuple, List
import pyperclip

LIST_OF_USER_AGENTS = [
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36',
    'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.90 Safari/537.36',
    'Mozilla/5.0 (Windows NT 5.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.90 Safari/537.36',
    'Mozilla/5.0 (Windows NT 6.2; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.90 Safari/537.36',
    'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.157 Safari/537.36',
    'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.133 Safari/537.36',
    'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.133 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36',
    'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36',
]

WEBSITE_URL = "https://texttohuman.com"
# Thread-safe print lock
print_lock = Lock()

def thread_safe_print(*args, **kwargs):
    """Thread-safe print function"""
    with print_lock:
        print(*args, **kwargs)



def get_random_user_agent():
    return random.choice(LIST_OF_USER_AGENTS)

def read_docx_with_spacing(file_path):
    """
    Read a DOCX file and return text while maintaining spacing and formatting.
    
    Args:
        file_path: str - Path to the DOCX file
        
    Returns:
        str: The extracted text with preserved spacing and line breaks
    """
    try:
        
        # Load the document
        doc = Document(file_path)
        
        # List to store all text elements
        text_elements = []
        
        # Iterate through all paragraphs
        for paragraph in doc.paragraphs:
            # Get the paragraph text
            para_text = paragraph.text
            
            # Preserve empty lines (paragraphs with no text)
            if not para_text.strip():
                text_elements.append('')
            else:
                text_elements.append(para_text)
        
        # Join all elements with newlines to maintain structure
        full_text = '\n'.join(text_elements)
        
        return full_text
        
    except ImportError:
        print("Error: python-docx library not installed. Install it using: pip install python-docx")
        return None
    except FileNotFoundError:
        print(f"Error: File not found at path: {file_path}")
        return None
    except Exception as e:
        print(f"Error reading DOCX file: {e}")
        return None
    
def get_huminizer_chrome_driver():
    # Launch undetected Chrome
    options = uc.ChromeOptions()
    custom_user_agent = get_random_user_agent()
    options.add_argument(f"--user-agent={custom_user_agent}")
    # Grant clipboard permissions automatically
    prefs = {
        "profile.default_content_setting_values.clipboard": 1,  # 1=allow, 2=block
        "profile.content_settings.exceptions.clipboard": {
            "[*.]texttohuman.com,*": {"setting": 1}
        }
    }
    options.add_experimental_option("prefs", prefs)
    
    # Additional clipboard permission via command line
    options.add_argument("--disable-features=ClipboardPrompt")
    #options.headless = True
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-blink-features=AutomationControlled")
    driver = uc.Chrome(options=options)
    

    driver.get(WEBSITE_URL)
    driver.set_page_load_timeout(60)
    return driver

def split_text_preserve_paragraphs_and_newlines(text, chunk_size=2000):
    """
    Split text into chunks while preserving paragraph boundaries and all newlines.
    
    Args:
        text: str - The text to split
        chunk_size: int - Target number of words per chunk (default: 2000)
        
    Returns:
        list: List of text chunks with preserved formatting
    """
    # Split by newlines but keep the newlines
    lines = text.split('\n')
    
    chunks = []
    current_chunk = []
    current_word_count = 0
    
    for i, line in enumerate(lines):
        # Count words in the line
        line_words = line.split()
        line_word_count = len(line_words)
        
        # If adding this line would exceed chunk_size and we have content
        if current_word_count + line_word_count > chunk_size and current_chunk:
            # Save current chunk (join with newlines)
            chunks.append('\n'.join(current_chunk))
            current_chunk = [line]
            current_word_count = line_word_count
        else:
            # Add line to current chunk
            current_chunk.append(line)
            current_word_count += line_word_count
    
    # Add the last chunk if it has content
    if current_chunk:
        chunks.append('\n'.join(current_chunk))
    
    return chunks

def get_Zero_Human_Alternative(dialog, driver):
    """
    Get the alternative button with "Human" type and 0% score.
    Retries up to 3 times by clicking reload if not found.
    
    Args:
        dialog: WebElement - The dialog containing alternatives
        driver: WebDriver instance for interactions
        
    Returns:
        str: The text of the best alternative, or None if not found
    """
    max_retries = 6
    
    for attempt in range(max_retries):
        print(f"   Attempt {attempt + 1}/{max_retries} to find 0% Human alternative...")
        
        try:
            # Get alternatives container
            alternatives_container = WebDriverWait(dialog, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div.space-y-2'))
            )
            alternative_buttons = alternatives_container.find_elements(By.TAG_NAME, 'button')
            
            if not alternative_buttons:
                print(f"   ✗ No alternative buttons found on attempt {attempt + 1}")
            else:
                # Process each button to find 0% Human alternative
                for button in alternative_buttons:
                    try:
                        # Get the spans inside the button
                        spans_container = button.find_element(By.CSS_SELECTOR, 'div.flex.items-center.gap-2.text-xs')
                        spans = spans_container.find_elements(By.TAG_NAME, 'span')
                        
                        if len(spans) >= 2:
                            alternative_type = spans[0].text  # "AI" or "Human"
                            alternative_score_text = spans[1].text  # "100%", "48%", etc.
                            
                            # Only process if alternative_type is "Human"
                            if alternative_type == "Human":
                                # Convert score to float (remove % sign)
                                try:
                                    alternative_score = float(alternative_score_text.replace('%', ''))
                                except ValueError:
                                    print(f"   ⚠ Could not parse score: {alternative_score_text}")
                                    continue
                                
                                # Get the alternative text
                                alternative_text_elem = button.find_element(By.CSS_SELECTOR, 'p.text-sm.text-foreground.flex-1')
                                alternative_text = driver.execute_script("return arguments[0].innerText;", alternative_text_elem)
                                
                                print(f"   Found Human alternative: {alternative_score}% - {alternative_text[:50]}...")
                                
                                # Check if this is 0% Human alternative
                                if alternative_score < 10.0: # less than 10% to account for rounding
                                    print(f"   ✓ Found 0% Human alternative!")
                                    return alternative_text
                        
                    except Exception as e:
                        print(f"   ⚠ Error processing button: {e}")
                        continue
            
            # If not found and not the last attempt, try reloading
            if attempt < max_retries - 1:
                print(f"   ⚠ 0% Human alternative not found, attempting reload...")
                try:
                    # Get reload button
                    reload_container = dialog.find_element(By.CSS_SELECTOR, 'div.flex.justify-end')
                    reload_alternatives_button = reload_container.find_element(By.TAG_NAME, 'button')
                    
                    # Click reload and wait
                    reload_alternatives_button.click()
                    print(f"   ✓ Clicked reload button, waiting 30 seconds...")
                    dialog = WebDriverWait(driver, 30).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, 'div[role="dialog"]'))
                    )
                    
                except Exception as e:
                    print(f"   ✗ Failed to reload alternatives: {e}")
                    break
            else:
                print(f"   ✗ Max retries reached, no 0% Human alternative found")
        
        except Exception as e:
            print(f"   ✗ Error on attempt {attempt + 1}: {e}")
            if attempt < max_retries - 1:
                time.sleep(2)  # Brief pause before retry
            continue
    
    # Return None if no 0% Human alternative found after all retries
    return None

def get_texttohuman_humanizer_final(humanize_text, driver, timeout=15):
    # Increase page load timeout if needed
    processing_timeout = 60
    try:
        # Wait until textarea is ready and locate it
        textarea_box = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'textarea[data-slot="textarea"]'))
        )

        textarea = textarea_box
        textarea.clear()
        time.sleep(1)
        textarea.click()

        # 🔽 Scroll textarea into view BEFORE interacting
        driver.execute_script("arguments[0].scrollIntoView({ behavior: 'smooth', block: 'center' });", textarea_box)

        # Copy text to clipboard
        pyperclip.copy(humanize_text)
        
        # Wait for paste button
        paste_button = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button.bg-primary\\/10")
            )
        )

        # Click paste
        paste_button.click()

        # Wait a moment for the text to register
        time.sleep(1)
        
        # Locate and click the "Humanize Now" button
        humanize_button = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'button[data-slot="button"]:not([disabled])'))
        )
        humanize_button.click()

        start_time = time.time()
        max_wait_time = processing_timeout
        check_interval = 2
        last_status = ""
        wait = WebDriverWait(driver, timeout)

        while True:
            elapsed_time = time.time() - start_time
            
            if elapsed_time > max_wait_time:
                thread_safe_print(f"Timeout after {elapsed_time:.1f} seconds")
                break
            
            try:
                status_div = driver.find_element(By.CSS_SELECTOR, 'div.flex.items-center.gap-4.text-xs.text-primary')
                status_text = status_div.text.strip()
                
                if status_text and status_text != last_status:
                    thread_safe_print(f"⚡ Autopilot: {status_text} ({int(elapsed_time)}s elapsed)")
                    last_status = status_text
                    
            except (NoSuchElementException, Exception):
                try:
                    spinner = driver.find_element(By.CSS_SELECTOR, '.animate-spin')
                    if spinner.is_displayed():
                        if int(elapsed_time) % 10 == 0 and int(elapsed_time) > 0:
                            thread_safe_print(f"Processing... ({int(elapsed_time)}s elapsed)")
                except (NoSuchElementException, Exception):
                    pass
            
            try:
                output_element = driver.find_element(By.CSS_SELECTOR, 'div.p-4.overflow-y-auto.rounded-lg.h-full.text-foreground.bg-background')
                if output_element and output_element.text.strip():
                    break
            except (NoSuchElementException, Exception):
                pass
            
            time.sleep(check_interval)
        
        # Find the output textarea/div (adjust selector as needed)
        output_element = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div.p-4.overflow-y-auto.rounded-lg.h-full.text-foreground.bg-background'))
        )
        
        # Get text using innerText to preserve newlines
        humanized_text = driver.execute_script("return arguments[0].innerText;", output_element)
        print(humanized_text)
        humanize_text1 = humanized_text

        marks = output_element.find_elements(By.TAG_NAME, 'mark')
        if marks:
            mark_data = []  # List of (mark_element, class, text, category)
            
            for i, mark in enumerate(marks):
                # ensure mark_class is always a string to avoid `in` checks on None
                mark_class = mark.get_attribute('class') or ""
                
                # Get mark text using innerText to preserve formatting
                mark_text = driver.execute_script("return arguments[0].innerText;", mark)
                
                if ('bg-yellow-100' in mark_class) or ('bg-yellow-900' in mark_class) or \
                ('bg-red-100' in mark_class) or ('bg-red-900' in mark_class):
                    
                    mark_type = "yellow" if 'yellow' in mark_class else "red"
                    print(f"\n🔄 Processing {mark_type} mark {i+1}/{len(marks)}")
                    print(f"   Original text: {mark_text[:80]}...")

                    try:
                        driver.execute_script("arguments[0].scrollIntoView(true);", mark)
                        time.sleep(1)
                        driver.execute_script("arguments[0].click();", mark)
                        
                        # Wait for dialog to load using while loop with timeout
                        dialog = None
                        start_time = time.time()
                        timeout_dialog = 30
                        alternatives_container = None
                        
                        while (time.time() - start_time) < timeout_dialog:
                            try:
                                dialog = driver.find_element(By.CSS_SELECTOR, 'div[role="dialog"]')
                                # Check if space-y-2 div is present (indicates dialog is fully loaded)
                                alternatives_container = dialog.find_element(By.CSS_SELECTOR, 'div.space-y-2')
                                print("   ✓ Dialog loaded with alternatives")
                                break
                            except:
                                time.sleep(0.5)
                                continue
                        
                        if dialog is None:
                            print("   ✗ Dialog failed to load within timeout")
                            continue
                        
                        # If mark_text is empty, get text from textarea
                        if mark_text.strip() == "":
                            try:
                                textarea = dialog.find_element(By.TAG_NAME, 'textarea')
                                mark_text = textarea.get_attribute('value') or driver.execute_script("return arguments[0].value;", textarea)
                                print(f"   Retrieved text from textarea: {mark_text[:80]}...")
                            except Exception as e:
                                print(f"   ✗ Failed to get textarea text: {e}")
                                continue
                        
                        # Use the function to get 0% Human alternative
                        best_alternative_text = get_Zero_Human_Alternative(dialog, driver)
                        
                        if best_alternative_text is not None:
                            print(f"   ✓ Best alternative text: {best_alternative_text[:80]}...")
                            
                            # Replace mark_text with best_alternative_text
                            # Use a more robust replacement that handles newlines
                            humanize_text1 = humanize_text1.replace(mark_text, best_alternative_text, 1)
                            print(f"   ✓ Replaced text in humanize_text1")
                        else:
                            print("   ✗ No 0% Human alternative found after all retries")
                        
                        # Close dialog (click the X button)
                        try:
                            close_button = dialog.find_element(By.CSS_SELECTOR, 'button[data-slot="dialog-close"]')
                            close_button.click()
                            time.sleep(1)
                        except Exception as e:
                            print(f"   ⚠ Failed to close dialog: {e}")
                            
                    except Exception as e:
                        print(f"   ✗ Failed to process mark: {e}")
                        continue
                
        #save to txt
        with open("humanized_text.txt", "w", encoding="utf-8") as f:
            f.write(humanize_text1)
         
        return humanize_text1 

    except Exception as e:
        print(f"Error occurred: {e}")
        return None
    
    finally:
        driver.quit()




