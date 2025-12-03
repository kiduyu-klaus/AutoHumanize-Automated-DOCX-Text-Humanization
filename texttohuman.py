import random
import pyperclip
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys  # Added missing import
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
from docx import Document
import os

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

def get_random_user_agent():
    return random.choice(LIST_OF_USER_AGENTS)


def read_docx_with_formatting(file_path):
    """
    Read a DOCX file and preserve formatting including:
    - Bold text as headings (with proper markdown)
    - Numbered lists
    - Regular paragraphs
    - Ignores placeholder text in [brackets]
    
    Args:
        file_path: Path to the DOCX file
        
    Returns:
        String containing formatted text with preserved structure
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File not found: {file_path}")
    
    doc = Document(file_path)
    formatted_text = []
    list_counter = {}
    
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        
        # Skip empty paragraphs
        if not text:
            continue
        
        # Skip placeholder text in brackets
        if text.startswith('[') and text.endswith(']'):
            continue
        
        style_name = paragraph.style.name
        
        # Handle headings by style
        if style_name.startswith('Heading'):
            try:
                level = int(style_name.split()[-1])
                formatted_text.append(f"\n{'#' * (level + 1)} {text}\n")
            except (ValueError, IndexError):
                formatted_text.append(f"\n## {text}\n")
        
        # Handle numbered lists
        elif style_name.startswith('List Number') or paragraph._element.xpath('.//w:numPr'):
            try:
                num_pr = paragraph._element.xpath('.//w:numPr')[0]
                ilvl = num_pr.xpath('.//w:ilvl')[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                level = int(ilvl) if ilvl else 0
            except (IndexError, ValueError):
                level = 0
            
            if level not in list_counter:
                list_counter[level] = 1
            else:
                list_counter[level] += 1
            
            indent = '  ' * level
            formatted_text.append(f"{indent}{list_counter[level]}. {text}")
        
        # Handle bulleted lists
        elif style_name.startswith('List Bullet'):
            try:
                num_pr = paragraph._element.xpath('.//w:numPr')[0]
                ilvl = num_pr.xpath('.//w:ilvl')[0].get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                level = int(ilvl) if ilvl else 0
            except (IndexError, ValueError):
                level = 0
            
            indent = '  ' * level
            formatted_text.append(f"{indent}• {text}")
        
        # Handle bold text (potential heading)
        elif any(run.bold for run in paragraph.runs if run.text.strip()):
            bold_chars = sum(len(run.text) for run in paragraph.runs if run.bold)
            total_chars = len(text)
            
            if bold_chars > total_chars * 0.8:
                formatted_text.append(f"\n## {text}\n")
            else:
                formatted_para = ""
                for run in paragraph.runs:
                    if run.bold and run.text.strip():
                        formatted_para += f"**{run.text}**"
                    else:
                        formatted_para += run.text
                formatted_text.append(formatted_para)
        
        # Regular paragraph
        else:
            formatted_text.append(text)
            list_counter = {}
    
    return '\n'.join(formatted_text)


def read_docx(file_path):
    """
    Read a DOCX file with formatting preservation.
    
    Args:
        file_path: Path to the DOCX file
        
    Returns:
        String containing all text from the document with preserved formatting
    """
    return read_docx_with_formatting(file_path)


def split_text_by_words(text, max_words=1200):
    """
    Split text into chunks of maximum word count.
    
    Args:
        text: The text to split
        max_words: Maximum number of words per chunk (default: 1200)
        
    Returns:
        List of text chunks
    """
    words = text.split()
    chunks = []
    current_chunk = []
    current_count = 0
    
    for word in words:
        current_chunk.append(word)
        current_count += 1
        
        if current_count >= max_words:
            chunks.append(' '.join(current_chunk))
            current_chunk = []
            current_count = 0
    
    if current_chunk:
        chunks.append(' '.join(current_chunk))
    
    return chunks


def get_texttohuman_humanizer(humanize_text, timeout=15, processing_timeout=60):
    """
    Humanize text using TextToHuman website.
    
    Args:
        humanize_text: Text to humanize
        timeout: Timeout for element loading (seconds)
        processing_timeout: Timeout for text processing (seconds)
        
    Returns:
        Humanized text or None if failed
    """
    options = uc.ChromeOptions()
    custom_user_agent = get_random_user_agent()
    options.add_argument(f"--user-agent={custom_user_agent}")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-blink-features=AutomationControlled")
    driver = uc.Chrome(options=options)

    try:
        print("Loading website...")
        driver.get(WEBSITE_URL)
        time.sleep(3)

        print("Scrolling down 25%...")
        driver.execute_script("window.scrollTo(0, Math.floor(document.documentElement.scrollHeight * 0.25));")
        time.sleep(1)
        
        textarea_selectors = [
            'textarea[data-slot="textarea"]',
            'textarea[placeholder*="Paste your AI-generated"]',
            'textarea.resize-none',
            'textarea'
        ]
        
        textarea_box = None
        for selector in textarea_selectors:
            try:
                textarea_box = WebDriverWait(driver, timeout).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                )
                print(f"Found textarea with selector: {selector}")
                break
            except TimeoutException:
                continue
        
        if not textarea_box:
            raise Exception("Could not find textarea element")

        print("Entering text...")
        try:
            pyperclip.copy(humanize_text)
            textarea_box.click()
            textarea_box.send_keys(Keys.CONTROL, 'v')
            print(f"✓ Text inserted ({len(humanize_text)} characters)")
        except Exception as e:
            try:
                print("Attempting to set text via JavaScript...")
                textarea_box.click()
                
                driver.execute_script("""
                    arguments[0].value = arguments[1];
                    var event = new Event('input', { bubbles: true });
                    arguments[0].dispatchEvent(event);
                    var changeEvent = new Event('change', { bubbles: true });
                    arguments[0].dispatchEvent(changeEvent);
                """, textarea_box, humanize_text)
                print("Text set via JavaScript")
                
                current_value = driver.execute_script("return arguments[0].value", textarea_box)
                if len(current_value) == len(humanize_text):
                    print("Text successfully set via JavaScript")
                else:
                    raise Exception("JavaScript set didn't work completely")
                    
            except Exception as js_error:
                print(f"JavaScript method failed, falling back to chunked send_keys: {js_error}")
                
                CHUNK_SIZE = 1000
                chunks = [humanize_text[i:i+CHUNK_SIZE] for i in range(0, len(humanize_text), CHUNK_SIZE)]
                
                textarea_box.clear()
                for i, chunk in enumerate(chunks):
                    textarea_box.send_keys(chunk)
                    if i % 5 == 0:
                        time.sleep(0.1)
                
                print("Text entered using chunked send_keys")

            final_text = textarea_box.get_attribute('value')
            print(f"Text length verification: Original: {len(humanize_text)}, Final: {len(final_text)}")
                
        time.sleep(2)
        
        button_selectors = [
            'button[data-slot="button"]:not([disabled])',
            'button:has-text("Humanize Now")',
            'button:contains("Humanize")',
            'div.flex.flex-col.gap-2.items-end button'
        ]
        
        humanize_button = None
        for selector in button_selectors:
            try:
                humanize_button = WebDriverWait(driver, timeout).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
                )
                print(f"Found button with selector: {selector}")
                break
            except:
                continue
        
        if not humanize_button:
            print("Trying to find button by text...")
            buttons = driver.find_elements(By.TAG_NAME, 'button')
            for btn in buttons:
                if 'Humanize' in btn.text:
                    humanize_button = btn
                    break
        
        if not humanize_button:
            raise Exception("Could not find Humanize button")
        
        print("Checking button state...")
        disabled_attr = humanize_button.get_attribute('disabled')
        
        if disabled_attr is not None:
            print("Button is disabled, waiting for it to be enabled...")
            wait_start = time.time()
            max_button_wait = 30
            
            while time.time() - wait_start < max_button_wait:
                disabled_attr = humanize_button.get_attribute('disabled')
                if disabled_attr is None:
                    print("✓ Button enabled!")
                    break
                time.sleep(0.5)
            else:
                raise Exception("Button remained disabled after 30 seconds")
        
        time.sleep(1)
        
        print("Clicking Humanize Now button...")
        try:
            humanize_button.click()
        except Exception as e:
            print(f"Regular click failed: {e}")
            try:
                print("Trying JavaScript click...")
                driver.execute_script("arguments[0].click();", humanize_button)
            except Exception as e2:
                print(f"JavaScript click failed: {e2}")
                print("Trying scroll and click...")
                driver.execute_script("arguments[0].scrollIntoView(true);", humanize_button)
                time.sleep(0.5)
                driver.execute_script("arguments[0].click();", humanize_button)
        
        try:
            print("Waiting for processing to start...")
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '.animate-spin'))
            )
            print("Processing started...")
        except TimeoutException:
            print("Loading spinner not detected, continuing...")
        
        print("Waiting for results (this may take up to 3 minutes)...")
        output_container = None
        start_time = time.time()
        max_wait_time = processing_timeout
        check_interval = 2
        
        output_selectors = [
            'div.overflow-y-auto.rounded-lg',
            'div[class*="overflow-y-auto"]',
            'div.p-4 div.w-full',
            'mark[data-chunk-type]'
        ]
        
        last_status = ""
        while True:
            elapsed_time = time.time() - start_time
            
            if elapsed_time > max_wait_time:
                print(f"Timeout after {elapsed_time:.1f} seconds")
                break
            
            try:
                status_div = driver.find_element(By.CSS_SELECTOR, 'div.flex.items-center.gap-4.text-xs.text-primary')
                status_text = status_div.text.strip()
                
                if status_text and status_text != last_status:
                    print(f"⚡ Autopilot: {status_text} ({int(elapsed_time)}s elapsed)")
                    last_status = status_text
                    
            except (NoSuchElementException, Exception):
                try:
                    spinner = driver.find_element(By.CSS_SELECTOR, '.animate-spin')
                    if spinner.is_displayed():
                        if int(elapsed_time) % 10 == 0 and int(elapsed_time) > 0:
                            print(f"Processing... ({int(elapsed_time)}s elapsed)")
                except (NoSuchElementException, Exception):
                    pass
            
            found_output = False
            for selector in output_selectors:
                try:
                    output_container = driver.find_element(By.CSS_SELECTOR, selector)
                    
                    if output_container and len(output_container.text.strip()) > 0:
                        print(f"✓ Found output with selector: {selector}")
                        print(f"✓ Results loaded after {elapsed_time:.1f} seconds")
                        found_output = True
                        break
                    else:
                        output_container = None
                except NoSuchElementException:
                    continue
            
            if found_output and output_container:
                break
            
            time.sleep(check_interval)
        
        if not output_container:
            print("Output container not found, checking page content...")
            time.sleep(3)
            
            try:
                page_text = driver.find_element(By.TAG_NAME, 'body').text
                
                if 'Humanizing your text' in page_text:
                    raise Exception(f"Still processing after {max_wait_time} seconds timeout")
                
                marks = driver.find_elements(By.TAG_NAME, 'mark')
                if marks:
                    humanized_text = ' '.join([mark.text for mark in marks if mark.text.strip()])
                    if humanized_text:
                        print("✓ Found results in mark tags")
                        return humanized_text
                
                raise Exception("Could not find results in page")
            except Exception as e:
                raise Exception(f"Failed to retrieve results: {str(e)}")
        
        print("Results loaded successfully!")
        
        print("\nAnalyzing results...")
        marks = output_container.find_elements(By.TAG_NAME, 'mark')
        
        humanized_text = output_container.text
        if not humanized_text or len(humanized_text) < 10:
            marks = output_container.find_elements(By.TAG_NAME, 'mark')
            if marks:
                humanized_text = ' '.join([mark.text for mark in marks if mark.text.strip()])
        
        return humanized_text

    except Exception as e:
        print(f"Error occurred: {e}")
        try:
            driver.save_screenshot("error_screenshot.png")
            print("Screenshot saved as error_screenshot.png")
            print(f"Current URL: {driver.current_url}")
            print(f"Page title: {driver.title}")
        except:
            pass
        return None
    
    finally:
        driver.quit()


def process_docx_file(docx_path, output_path=None, max_words=1200):
    """
    Read a DOCX file, split it into chunks, and humanize each chunk.
    
    Args:
        docx_path: Path to the input DOCX file
        output_path: Path to save the humanized text (optional)
        max_words: Maximum words per chunk (default: 1200)
        
    Returns:
        String containing all humanized text
    """
    print(f"\n{'='*60}")
    print(f"Processing DOCX file: {docx_path}")
    print(f"{'='*60}\n")
    
    print("Reading DOCX file...")
    original_text = read_docx(docx_path)
    print(f"Total characters: {len(original_text)}")
    print(f"Total words: {len(original_text.split())}")
    
    print(f"\nSplitting text into {max_words}-word chunks...")
    chunks = split_text_by_words(original_text, max_words)
    print(f"Created {len(chunks)} chunks")
    
    humanized_chunks = []
    failed_chunks = []
    
    for i, chunk in enumerate(chunks, 1):
        print(f"\n{'='*60}")
        print(f"Processing chunk {i}/{len(chunks)} ({len(chunk.split())} words)")
        print(f"{'='*60}")
        
        humanized = get_texttohuman_humanizer(chunk)
        
        if humanized:
            humanized_chunks.append(humanized)
            print(f"✓ Chunk {i} humanized successfully")
        else:
            humanized_chunks.append(chunk)
            failed_chunks.append(i)
            print(f"✗ Chunk {i} failed - keeping original text")
        
        if i < len(chunks):
            print("\nWaiting 5 seconds before next chunk...")
            time.sleep(5)
    
    final_text = '\n\n'.join(humanized_chunks)
    
    if output_path is None:
        base_name = os.path.splitext(docx_path)[0]
        output_path = f"{base_name}_humanized.txt"
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(final_text)
    
    print(f"\n{'='*60}")
    print("PROCESSING COMPLETE")
    print(f"{'='*60}")
    print(f"Total chunks: {len(chunks)}")
    print(f"Successful: {len(chunks) - len(failed_chunks)}")
    print(f"Failed: {len(failed_chunks)}")
    if failed_chunks:
        print(f"Failed chunk numbers: {failed_chunks}")
    print(f"\nOutput saved to: {output_path}")
    print(f"{'='*60}\n")
    
    return final_text


if __name__ == "__main__":
    docx_file = "your_document.docx"
    
    if os.path.exists(docx_file):
        result = process_docx_file(docx_file, max_words=1200)
    else:
        print(f"File not found: {docx_file}")
        print("\nTo use this script:")
        print("1. Place your DOCX file in the same directory")
        print("2. Update the 'docx_file' variable with your filename")
        print("3. Run the script again")