import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox, ttk, font
import asyncio
from docx import Document
from lxml import etree
import re
import os
import aiohttp
import threading
from datetime import datetime
import sys
import logging
# pyinstaller --onefile --windowed --add-data "azure.tcl:." --add-data "theme:theme" gptTMXalign.py

class CustomInputDialog(tk.Toplevel):
    def __init__(self, parent, title="", prompt=""):
        super().__init__(parent)
        self.var = tk.StringVar()

        self.title(title)
        self.geometry("250x200")

        # Configure the dialog layout
        self.label = ttk.Label(self, text=prompt)
        self.label.pack(pady=10)

        self.entry = ttk.Entry(self, textvariable=self.var)
        self.entry.pack(pady=10)

        self.button_frame = ttk.Frame(self)
        self.button_frame.pack(fill=tk.X, pady=10)

        self.ok_button = ttk.Button(self.button_frame, text="OK", command=self.on_ok)
        self.ok_button.pack(side=tk.RIGHT, padx=(0,10), pady=10)

        self.cancel_button = ttk.Button(self.button_frame, text="Cancel", command=self.on_cancel)
        self.cancel_button.pack(side=tk.RIGHT, padx=(0,10), pady=10)

        self.entry.bind("<Return>", lambda event: self.on_ok())
        self.entry.bind("<Escape>", lambda event: self.on_cancel())

        self.protocol("WM_DELETE_WINDOW", self.on_cancel)

        self.parent = parent
        self.result = None

        self.transient(parent)
        self.wait_visibility()
        self.grab_set()
        self.entry.focus_set()
        self.wait_window(self)

    def on_ok(self):
        self.result = self.var.get()
        self.destroy()

    def on_cancel(self):
        self.destroy()


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def read_docx(file_path):
    doc = Document(file_path)
    paragraphs = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    return paragraphs

async def process_paragraphs(api_key, output_file, english_paragraphs, khmer_paragraphs):
    async with aiohttp.ClientSession() as session:
        tasks = []
        for en_text, km_text in zip(english_paragraphs, khmer_paragraphs):
            if en_text.strip() and km_text.strip():
                task = asyncio.ensure_future(align_paragraphs(session, en_text, km_text, api_key))
                tasks.append(task)

        aligned_texts = await asyncio.gather(*tasks)
        for aligned_text in aligned_texts:
            aligned_pairs = parse_aligned_text(aligned_text)
            create_tmx(aligned_pairs, output_file)

def read_file(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        return file.read().split('\n')

async def align_paragraphs(session, en_text, km_text, api_key, src_lang_name, tgt_lang_name):
    print("Aligning: ")
    print(en_text + "\n\n")
    print(km_text + "\n\n")

    # Use the provided source and target language names in the prompt
    prompt = (f"Please align this {src_lang_name} paragraph with provided {tgt_lang_name} translation so it would be useful as Translation Memory. "
              f"Here is the {src_lang_name} Text in paragraph form:\n{en_text}\n"
              f"Here is the {tgt_lang_name} translation of that same paragraph:\n{km_text}\n\n"
              "Do not comment, do not provide me new translation but use what I gave you in the paragraphs, do not modify or cut short the provided sentences, "
              f"only give me the sentences or phrases I provided aligned in json like format as specified like this: {{\"{src_lang_name.lower()}\": \"introduction\", \"{tgt_lang_name.lower()}\": \"សេចក្ដី​ផ្ដើម៖\"}}.")


    headers = {
        'Authorization': f'Bearer {api_key}',
        'Content-Type': 'application/json'
    }
    # Adjust the system content message to include dynamic language names
    system_content = f"You are a highly proficient bilingual assistant capable of aligning {src_lang_name} and {tgt_lang_name} text for use in a Translation Memory file."
    # gpt-3.5-turbo
    # gpt-4-turbo-preview
    data = {
        "model": "gpt-4-turbo-preview",
        "messages": [
            {"role": "system", "content": system_content},
            {"role": "user", "content": prompt}
        ]
    }

    async with session.post('https://api.openai.com/v1/chat/completions', headers=headers, json=data, timeout=600) as resp:
        if resp.status == 200:
            response = await resp.json()
            aligned_text = response['choices'][0]['message']['content']
            return aligned_text
        else:
            return ''


# Async function to align paragraphs and create TMX file
async def process_paragraphs(api_key, output_file, english_paragraphs, khmer_paragraphs, src_lang_name, tgt_lang_name, progress_callback):
    async with aiohttp.ClientSession() as session:
        tasks = []
        for en_text, km_text in zip(english_paragraphs, khmer_paragraphs):
            if en_text and km_text:
                tasks.append(asyncio.ensure_future(align_paragraphs(session, en_text, km_text, api_key, src_lang_name, tgt_lang_name)))
        
        aligned_texts = await asyncio.gather(*tasks)
        for index, aligned_text in enumerate(aligned_texts):
            aligned_pairs = parse_aligned_text(aligned_text)
            create_tmx(aligned_pairs, output_file)
            progress_callback(index + 1, len(aligned_texts))  # Update progress bar after each paragraph is processed

def parse_aligned_text(aligned_text):
    print(aligned_text)
    aligned_pairs = []
    # Updated regex pattern to include curly braces
    pattern = r'\{\s*"english":\s*"(.*?)"\s*,\s*"khmer":\s*"(.*?)"\s*\}'

    matches = re.findall(pattern, aligned_text, re.DOTALL)
    for match in matches:
        # Ensure both English and Khmer texts are present
        if len(match) == 2:
            english_text, khmer_text = match
            aligned_pairs.append((english_text.strip(), khmer_text.strip()))
    print(aligned_pairs)
    return aligned_pairs






def create_tmx(aligned_pairs, output_file):
    NSMAP = {'xml': 'http://www.w3.org/XML/1998/namespace'}  # Namespace map for XML

    # Read existing TMX file if it exists and has content
    if os.path.exists(output_file) and os.path.getsize(output_file) > 0:
        tree = etree.parse(output_file)
        root = tree.getroot()
        body = root.find('body')
    else:
        # Create new TMX structure if file doesn't exist or is empty
        root = etree.Element('tmx', version="1.4", nsmap=NSMAP)
        etree.SubElement(root, 'header', {
            'creationtool': "ChatGPT TMX Generator",
            'creationtoolversion': "1.0",
            'datatype': "PlainText",
            'segtype': "sentence",
            'adminlang': "en",
            'srclang': "EN",
            'o-tmf': "ABCTransMem"
        })
        body = etree.SubElement(root, 'body')

    # Add new translation units
    for en_text, km_text in aligned_pairs:
        tu = etree.SubElement(body, 'tu')
        tuv_en = etree.SubElement(tu, 'tuv', {f"{{{NSMAP['xml']}}}lang": "EN"})
        etree.SubElement(tuv_en, 'seg').text = en_text.strip()
        tuv_km = etree.SubElement(tu, 'tuv', {f"{{{NSMAP['xml']}}}lang": "KM"})
        etree.SubElement(tuv_km, 'seg').text = km_text.strip()

    # Rewrite the updated TMX file
    tree = etree.ElementTree(root)
    tree.write(output_file, pretty_print=True, xml_declaration=True, encoding='UTF-8')

def get_user_data_directory():
    """Return the path to the user data directory for the application."""
    home_dir = os.path.expanduser("~")
    app_data_dir = os.path.join(home_dir, ".tmxGUI4")  # Use your app's name
    if not os.path.exists(app_data_dir):
        os.makedirs(app_data_dir)
    return app_data_dir

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def get_user_data_directory():
    """Return the path to the user data directory for the application."""
    home_dir = os.path.expanduser("~")
    app_data_dir = os.path.join(home_dir, ".tmxGeneratorAppData")
    if not os.path.exists(app_data_dir):
        os.makedirs(app_data_dir)
    return app_data_dir



# GUI Application
class TMXGeneratorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('ChatGPT TMX Generator')
        self.geometry('500x700')  # Adjust size as needed

        # Initialize Azure theme with corrected path
        azure_tcl_path = resource_path('azure.tcl')
        self.tk.call('source', azure_tcl_path)
        self.tk.call('set_theme', 'dark')
        self.output_directory = None

        self.generated_tmx_files = []  # Initialize the list to store paths of generated TMX files

        # Initialize variables
        self.file_pairs = []
        self.api_key = ""  # Attempt to load the API key on startup
        self.src_lang_code = tk.StringVar(value='EN')  # Default to English
        self.tgt_lang_code = tk.StringVar(value='KM')  # Default to Khmer
        # Initialize variables for language names
        self.src_lang_name = tk.StringVar(value='English')  # Default to English
        self.tgt_lang_name = tk.StringVar(value='Khmer')  # Default to Khmer

        self.create_widgets()
    
    def enter_api_key(self):
        dialog = CustomInputDialog(self, title="API Key", prompt="Enter your API Key:")
        if dialog.result:
            self.api_key = dialog.result
        else:
            messagebox.showwarning("API Key Required", "An API key is required to proceed.")
    
    def prompt_for_output_directory(self):
        """Ask the user to select an output directory for the TMX files."""
        self.output_directory = filedialog.askdirectory(title="Select Output Directory")
        if not self.output_directory:
            messagebox.showwarning("Output Directory Required", "Please select a valid output directory.")
            return False
        return True


    def create_widgets(self):
        # Create a top frame for the API Key button and other top-aligned widgets
        top_frame = ttk.Frame(self)
        top_frame.pack(fill='x', side='top')
        
        # API Key input - placed at the right of the top frame
        self.btn_api_key = ttk.Button(top_frame, text='Enter API Key', command=self.enter_api_key, style='Accent.TButton')
        self.btn_api_key.grid(row=0, column=1, padx=(10, 10), pady=(10, 10), sticky='ne')
    
        # Fill the rest of the top frame with a spacer label to push the API Key button to the right
        ttk.Label(top_frame, text='').grid(row=0, column=0, sticky='we')
        top_frame.grid_columnconfigure(0, weight=1)  # This makes the spacer label expand and push the API Key button to the right
    
    
        # Source Language Name input
        ttk.Label(self, text='Source Language Name:').pack(pady=(10,0))
        src_lang_name_entry = ttk.Entry(self, textvariable=self.src_lang_name, style='TEntry')
        src_lang_name_entry.pack()
    
        # Target Language Name input
        ttk.Label(self, text='Target Language Name:').pack(pady=(10,0))
        tgt_lang_name_entry = ttk.Entry(self, textvariable=self.tgt_lang_name, style='TEntry')
        tgt_lang_name_entry.pack()
    
        # Language code input fields
        ttk.Label(self, text='Source Language Code:').pack(pady=(10,0))
        src_lang_entry = ttk.Entry(self, textvariable=self.src_lang_code, style='TEntry')
        src_lang_entry.pack()
    
        ttk.Label(self, text='Target Language Code:').pack(pady=(10,0))
        tgt_lang_entry = ttk.Entry(self, textvariable=self.tgt_lang_code, style='TEntry')
        tgt_lang_entry.pack()
    
        # File selection button
        self.btn_select_files = ttk.Button(self, text='Select File Pairs', command=self.select_file_pairs, style='Accent.TButton')
        self.btn_select_files.pack(pady=10)
    
        # File pairs list display
        self.file_pairs_list = tk.Listbox(self, height=4)
        self.file_pairs_list.pack(pady=(0, 10), fill=tk.BOTH, expand=True)  # Reduce padding to bring the next widget closer
        
        # File clearing button
        self.btn_clear_files = ttk.Button(self, text='Clear Files', command=self.clear_file_pairs, style='Accent.TButton')
        self.btn_clear_files.pack(pady=(10, 20))  # Adjust padding to position the button closer to the Listbox
    
        # Progress bar - Azure theme should automatically style it
        self.progress = ttk.Progressbar(self, orient='horizontal', length=300, mode='determinate')
        self.progress.pack(pady=20)
    
        # Start processing button
        self.btn_start = ttk.Button(self, text='Start Processing', command=self.start_processing, style='Accent.TButton')
        self.btn_start.pack()

    def select_file_pairs(self):
        english_file_path = filedialog.askopenfilename(title="Select English File", filetypes=[("Word documents", "*.docx")])
        if english_file_path:
            khmer_file_path = filedialog.askopenfilename(title="Select Khmer File", filetypes=[("Word documents", "*.docx")])
            if khmer_file_path:
                self.file_pairs.append((english_file_path, khmer_file_path))
                self.file_pairs_list.insert(tk.END, f"{os.path.basename(english_file_path)} - {os.path.basename(khmer_file_path)}")

    def clear_file_pairs(self):
        self.file_pairs.clear()
        self.file_pairs_list.delete(0, tk.END)

    def update_progress(self, current, total):
        # Calculate the progress value
        value = (current / total) * 100
        # Schedule the progress bar update on the main GUI thread
        self.after(0, lambda: self.progress.configure(value=value))

    async def async_start_processing(self):
        total_files = len(self.file_pairs)
        if total_files == 0:
            messagebox.showwarning("No Files Selected", "Please select file pairs to process.")
            return
    
        for index, (english_file_path, khmer_file_path) in enumerate(self.file_pairs):
            print(f"Processing pair {index + 1}: {english_file_path} and {khmer_file_path}")  # Debug print
    
            # Generate unique output filename based on the input files
            output_file = f"translation_memory_{os.path.basename(english_file_path)}_{os.path.basename(khmer_file_path)}.tmx"
            
            # Ensure the output directory exists
            output_dir = self.output_directory
            output_file_path = os.path.join(output_dir, output_file)
            self.generated_tmx_files.append(output_file_path)
    
            # Process each pair of files
            english_paragraphs = read_docx(english_file_path)
            khmer_paragraphs = read_docx(khmer_file_path)
            await process_paragraphs(self.api_key, output_file_path, english_paragraphs, khmer_paragraphs, self.src_lang_name.get(), self.tgt_lang_name.get(), lambda current, total: self.update_progress(current, total_files))
    
            # Update the progress bar based on the number of processed files
            self.progress['value'] = ((index + 1) / total_files) * 100
            self.update_idletasks()

    def create_master_tmx(self):
        # Generate a unique filename with a timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        master_file_name = f"master_translation_memory_{timestamp}.tmx"
    
        NSMAP = {'xml': 'http://www.w3.org/XML/1998/namespace'}
        root = etree.Element('tmx', version="1.4", nsmap=NSMAP)
        header = etree.SubElement(root, 'header', {
            'creationtool': "ChatGPT TMX Generator",
            'creationtoolversion': "1.0",
            'datatype': "PlainText",
            'segtype': "sentence",
            'adminlang': self.src_lang_code.get(),
            'srclang': self.src_lang_code.get(),
            'o-tmf': "ABCTransMem"
        })
        body = etree.SubElement(root, 'body')
    
        for tmx_file in self.generated_tmx_files:
            try:
                tree = etree.parse(tmx_file)
                for tu in tree.xpath('//tu'):
                    body.append(tu)
            except Exception as e:
                print(f"Error processing {tmx_file}: {e}")
    
        output_dir = self.output_directory  # Use the selected output directory
        master_file_path = os.path.join(output_dir, master_file_name)
    
        # Save the master TMX file
        tree = etree.ElementTree(root)
        tree.write(master_file_path, pretty_print=True, xml_declaration=True, encoding='UTF-8')
        print(f"Master TMX file created at {master_file_path}")

    def start_processing(self):
        if not self.file_pairs or not self.api_key:
            messagebox.showwarning("Missing Information", "Please select file pairs and enter the API key.")
            return
        
        if not self.prompt_for_output_directory():  # Check if the user selected an output directory
            return  # Exit the method if no directory was selected
        
        self.btn_start.config(text="Processing...", state="disabled")
        processing_thread = threading.Thread(target=self.run_async_start_processing, daemon=True)
        processing_thread.start()

    def run_async_start_processing(self):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
    
        try:
            loop.run_until_complete(self.async_start_processing())
        except Exception as e:
            print(f"Error during processing: {e}")
        finally:
            loop.close()
    
        # Update UI in a thread-safe way
        self.after(0, self.finalize_processing)
    
    def finalize_processing(self):
        self.btn_start.config(text="Start Processing", state="normal")  # Reset button text and re-enable it
        self.progress['value'] = 0  # Reset the progress bar
        # Create the master TMX file
        self.create_master_tmx()
        self.show_completion_message()
        
    
    def show_completion_message(self):
        messagebox.showinfo("Processing Complete", "All file pairs have been processed successfully.")

if __name__ == "__main__":
    app = TMXGeneratorApp()
    app.mainloop()
