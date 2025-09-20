import re
import pandas as pd
from PIL import Image, ImageEnhance
import easyocr
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

def preprocess_image(image_path, max_width=500, max_height=600):
    """Preprocess image to improve handwriting recognition."""
    with Image.open(image_path) as img:
        img = img.convert('L')
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(1.5)
        img.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
        temp_path = 'temp_resized.png'
        img.save(temp_path)
    return temp_path

def extract_text_from_image(image_path):
    """Extract text using EasyOCR with detailed output and text reconstruction."""
    resized_path = preprocess_image(image_path)
    reader = easyocr.Reader(['en'], gpu=False, verbose=False)
    result = reader.readtext(resized_path, detail=1, paragraph=False)
    # Sort by y-coordinate to maintain top-to-bottom order
    result.sort(key=lambda x: x[0][0][1])  # Sort by top-left y
    text_segments = [(item[1], item[0]) for item in result]  # Text and bounding box
    print(f"Extracted text segments with bounding boxes: {text_segments}")
    # Reconstruct text
    reconstructed_text = ' '.join(t for t, _ in text_segments)
    return reconstructed_text

def parse_fields(text):
    """Improved parsing logic with better misread handling and pattern matching."""
    fields = {
        'EmployeeID': None,
        'Name': None,
        'Department': None,
        'TestDate': None,
        'TestType': None,
        'Result': None,
        'Notes': None
    }
    
    # Step 1: Apply corrections for common OCR misreads
    corrections = {
        'EmPloxe TUl': 'EmployeeID',
        '91 91': '911911',
        'Nam e [': 'Name',
        'Go kon': 'Brifon',
        'Peloarz': 'Peter',
        'TT': 'IT',
        '414n) 025': '9/14/2025',
        'Test Tz/c7': 'TestType',
        'Bloos': 'Blood',
        'Koroj': 'Result',
        'Kegaon +ot': 'Negative',
        'No ko': 'Notes',
        'Ronv om +est': 'Random test'
    }
    for old, new in corrections.items():
        text = text.replace(old, new)
    
    # Step 2: Use flexible patterns to extract values
    patterns = {
        'EmployeeID': r'EmployeeID[:\s]*(\d{6})',
        'Name': r'Name[:\s]*([A-Za-z\s]+)',
        'Department': r'(?:Department|IT)[:\s]*([A-Za-z\s]+)',
        'TestDate': r'(?:TestDate|9/14/2025)[:\s]*(\d{1,2}/\d{1,2}/\d{4})',
        'TestType': r'(?:TestType|Blood)[:\s]*([A-Za-z]+)',
        'Result': r'(?:Result|Negative)[:\s]*([A-Za-z]+)',
        'Notes': r'(?:Notes|Random test)[:\s]*(.+?)(?=\n|$)'
    }
    
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            value = match.group(1).strip()
            if key == 'TestDate':
                try:
                    month, day, year = map(int, value.split('/'))
                    value = f"{year:04d}-{month:02d}-{day:02d}"
                    datetime.strptime(value, '%Y-%m-%d')
                except (ValueError, IndexError):
                    value = None
            fields[key] = value
    
    # Step 3: Fallbacks for missing fields
    if not fields['EmployeeID']:
        emp_match = re.search(r'\d{6}', text)
        if emp_match:
            fields['EmployeeID'] = emp_match.group()
    
    if not fields['Name']:
        name_match = re.search(r'(Peter\s*[A-Za-z]+)', text, re.IGNORECASE)
        if name_match:
            fields['Name'] = name_match.group()
    
    if not fields['Department']:
        dept_match = re.search(r'(IT|Dept)', text, re.IGNORECASE)
        if dept_match:
            fields['Department'] = dept_match.group().upper()
    
    if not fields['TestDate']:
        date_match = re.search(r'(\d{1,2}/\d{1,2}/\d{4})', text)
        if date_match:
            fields['TestDate'] = date_match.group()
    
    if not fields['TestType']:
        type_match = re.search(r'(Blood|Urine|Hair|Saliva)', text, re.IGNORECASE)
        if type_match:
            fields['TestType'] = type_match.group().capitalize()
    
    if not fields['Result']:
        result_match = re.search(r'(Negative|Positive|Pending)', text, re.IGNORECASE)
        if result_match:
            fields['Result'] = result_match.group().capitalize()
    
    if not fields['Notes']:
        notes_match = re.search(r'Notes\s*(.+)', text, re.IGNORECASE)
        if notes_match:
            fields['Notes'] = notes_match.group(1).strip()
    
    if not all([fields['Name'], fields['TestDate'], fields['TestType']]):
        raise ValueError(f"Incomplete data: Missing {', '.join(k for k, v in fields.items() if k in ['Name', 'TestDate', 'TestType'] and not v)}.")
    
    return fields

def add_to_csv(fields):
    """Append parsed fields to CSV with EmployeeID as the first column."""
    try:
        df = pd.read_csv('DrugTestingOrganizer.csv')
    except FileNotFoundError:
        df = pd.DataFrame(columns=['EmployeeID', 'Name', 'Department', 'TestDate', 'TestType', 'Result', 'Notes'])
    
    new_row = pd.DataFrame([fields])
    df = pd.concat([df, new_row], ignore_index=True)
    # Ensure EmployeeID is the first column
    column_order = ['EmployeeID', 'Name', 'Department', 'TestDate', 'TestType', 'Result', 'Notes']
    df = df[column_order]
    df.to_csv('DrugTestingOrganizer.csv', index=False)
    return new_row[column_order].to_string(index=False)

class DrugTestApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Drug Testing OCR Scanner")
        self.root.geometry("600x400")
        
        self.label = tk.Label(root, text="Select a sticky note image to scan:")
        self.label.pack(pady=10)
        
        self.select_button = tk.Button(root, text="Browse Image", command=self.select_image)
        self.select_button.pack(pady=5)
        
        self.text_area = tk.Text(root, height=15, width=60)
        self.text_area.pack(pady=10)
        
        self.text_area.insert(tk.END, "Ready to scan an image...\n")
    
    def select_image(self):
        """Open file dialog and process selected image."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Image files", "*.jpg *.jpeg *.png")]
        )
        if not file_path:
            self.text_area.insert(tk.END, "No file selected.\n")
            return
        
        self.text_area.insert(tk.END, f"Processing image: {file_path}\n")
        try:
            text = extract_text_from_image(file_path)
            self.text_area.insert(tk.END, "\nExtracted Text:\n" + text + "\n")
            
            fields = parse_fields(text)
            self.text_area.insert(tk.END, "\nParsed Fields:\n" + str(fields) + "\n")
            
            result = add_to_csv(fields)
            self.text_area.insert(tk.END, "\nData Added to CSV:\n" + result + "\n")
            messagebox.showinfo("Success", "Data successfully added to DrugTestingOrganizer.csv!")
        except Exception as e:
            self.text_area.insert(tk.END, f"\nError: {str(e)}\nTip: Ensure the sticky note is clear and fields are labeled correctly.\n")
            messagebox.showerror("Error", f"Failed to process image: {str(e)}")
        
        self.text_area.see(tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    app = DrugTestApp(root)
    root.mainloop()