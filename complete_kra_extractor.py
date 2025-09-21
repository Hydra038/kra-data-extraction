"""
Complete KRA Bulk Data Extractor
===============================

Processes all KRA tax notices from the 60-page PDF and extracts:
- Date
- PIN (Tax Identification Number)  
- Taxpayer Name (from "To:" field)
- Subject (Notice type)
- Year (of computation)
- Station (KRA office location)
- Officer (Contact officer name)
- Total Tax (Amount due)

Author: GitHub Copilot
Date: September 19, 2025
"""
import fitz
import pytesseract
from PIL import Image
import pandas as pd
import re
import io
import os
from datetime import datetime

# Configure Tesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def extract_notice_data(page1, page2, notice_num):
    """Extract all required data from a 2-page KRA notice."""
    
    # Convert pages to high-quality images for OCR
    mat = fitz.Matrix(2.5, 2.5)  # Higher zoom for better accuracy
    
    pix1 = page1.get_pixmap(matrix=mat)
    pix2 = page2.get_pixmap(matrix=mat)
    
    img1 = Image.open(io.BytesIO(pix1.tobytes('ppm')))
    img2 = Image.open(io.BytesIO(pix2.tobytes('ppm')))
    
    # Extract text using OCR
    text1 = pytesseract.image_to_string(img1, lang='eng')
    text2 = pytesseract.image_to_string(img2, lang='eng')
    
    full_text = text1 + " " + text2
    
    # Initialize data structure
    data = {
        'Notice_Number': notice_num,
        'Date': '',
        'PIN': '',
        'Taxpayer_Name': '',
        'Subject': '',
        'Year': '',
        'Station': '',
        'Officer': '',
        'Total_Tax': ''
    }
    
    try:
        # Extract Date (multiple patterns for flexibility)
        date_patterns = [
            r'(\d{1,2}[A-Z]{2}\s+[A-Z]{3,9},?\s*\d{4})',  # 4TH SEPTEMBER, 2025
            r'(\d{1,2}\s+[A-Z]{3,9}\s+\d{4})',            # 04 SEP 2025
            r'(\d{1,2}[-/]\d{1,2}[-/]\d{4})',             # 04/09/2025
        ]
        
        for pattern in date_patterns:
            date_match = re.search(pattern, full_text, re.IGNORECASE)
            if date_match:
                data['Date'] = date_match.group(1).strip()
                break
        
        # Extract PIN (Tax ID format: Letter + 9 digits + Letter)
        pin_match = re.search(r'PIN:\s*([A-Z]\d{9}[A-Z])', full_text)
        if pin_match:
            data['PIN'] = pin_match.group(1)
        
        # Extract Taxpayer Name (company/individual name before P.O.BOX)
        name_patterns = [
            r'([A-Z][A-Z\s&.,()LTD-]{15,}?)\s*P\.\s*O\.\s*BOX',     # Before P.O.BOX
            r'PIN:\s*[A-Z]\d{9}[A-Z]\s+[^A-Z]*([A-Z][A-Z\s&.,()LTD-]{15,})', # After PIN
        ]
        
        for pattern in name_patterns:
            name_match = re.search(pattern, text1, re.IGNORECASE)
            if name_match:
                name = name_match.group(1).strip()
                # Clean up the name
                name = re.sub(r'\s+', ' ', name)  # Multiple spaces to single
                name = re.sub(r'[^\w\s&.,()LTD-]', '', name)  # Remove invalid chars
                if len(name) > 8 and 'LIMITED' in name.upper() or 'LTD' in name.upper() or len(name.split()) >= 2:
                    data['Taxpayer_Name'] = name
                    break
        
        # Extract Subject (the RE: notice line)
        subject_patterns = [
            r'RE:\s*([A-Z\s,\d]+TAX\s+PROCEDURES\s+ACT[^A-Z]*\d{4})',  # Full subject
            r'RE:\s*([A-Z\s,\d]+SECTION\s+\d+[^A-Z]*TAX[^A-Z]*ACT)',  # Section reference
        ]
        
        for pattern in subject_patterns:
            subject_match = re.search(pattern, full_text, re.IGNORECASE)
            if subject_match:
                subject = subject_match.group(1).strip()
                subject = re.sub(r'\s+', ' ', subject)  # Clean spaces
                data['Subject'] = subject
                break
        
        # Extract Year (from computation table)
        year_patterns = [
            r'Year\s*[|\[\]]*\s*(\d{4})',      # Year [2024]
            r'year\s*(\d{4})',                 # year 2024
            r'for\s*(\d{4})',                  # for 2024
            r'(\d{4})\s*tax',                  # 2024 tax
        ]
        
        for pattern in year_patterns:
            year_match = re.search(pattern, full_text, re.IGNORECASE)
            if year_match:
                year = year_match.group(1)
                if 2018 <= int(year) <= 2030:  # Reasonable year range
                    data['Year'] = year
                    break
        
        # Extract Station (KRA office location - improved accuracy)
        station_patterns = [
            # Primary: Extract from P.O.BOX address (most reliable)
            r'P\.\s*O\.\s*BOX\s*\d+[-\s]*\d*,?\s*([A-Z]{3,})',       # P.O.BOX 364-30500, LODWAR
            # Secondary: Direct station names
            r'\b(LODWAR|NAIROBI|MOMBASA|KISUMU|NAKURU|ELDORET|NYERI|MERU|MACHAKOS|KITALE|GARISSA|ISIOLO|MALINDI|KILIFI|EMBU|THIKA|KIAMBU)\b',
            # Tertiary: From region context
            r'(NORTH\s+RIFT|CENTRAL|COAST|WESTERN|EASTERN|NYANZA)\s+REGION',
        ]
        
        for pattern in station_patterns:
            station_match = re.search(pattern, full_text, re.IGNORECASE)
            if station_match:
                station = station_match.group(1).upper().strip()
                # Clean up station name
                if 'REGION' in station:
                    station = station.replace(' REGION', '').strip()
                # Map common variations
                if station in ['NORTH RIFT', 'NORTHRIFT']:
                    station = 'LODWAR'  # North Rift region typically uses Lodwar
                data['Station'] = station
                break
        
        # Extract Officer Name (contact person - improved patterns)
        officer_patterns = [
            # Primary pattern: Mr. Name Surname in contact section
            r'Mr\.?\s+([A-Z][a-z]+\s+[A-Z][a-z]+):\s*Tel:',           # Mr. Lochilia Emmanuel: Tel:
            r'Mr\.?\s+([A-Z][a-z]+\s+[A-Z][a-z]+)\s+on\s+Tel:',      # Mr Jefferson Mutavi on Tel:
            r'contact\s+Mr\.?\s+([A-Z][a-z]+\s+[A-Z][a-z]+)',        # contact Mr Lochilia Emmanuel
            r'([A-Z]{2,}\s+[A-Z]{2,})\s*Regional',                   # KENNETH RUTTO Regional (fallback)
            r'Mr\.?\s+([A-Z][a-z]+\s+[A-Z][a-z]+)\s*Email:',         # Mr Name Email: (alternative)
        ]
        
        for pattern in officer_patterns:
            officer_match = re.search(pattern, full_text, re.IGNORECASE)
            if officer_match:
                officer = officer_match.group(1).strip()
                # Properly capitalize the name
                officer = ' '.join(word.capitalize() for word in officer.split())
                data['Officer'] = officer
                break
        
        # Extract Total Tax Amount (improved patterns for various formats)
        tax_patterns = [
            # Primary: Table format with "Total Tax"
            r'Total\s+Tax\s+(\d{1,3}(?:,\d{3})*)',                    # Total Tax 78,936
            r'Total\s+Tax[:\s]+(\d{1,3}(?:,\d{3})*)',                 # Total Tax: 78,936
            # Secondary: "Tax Due" variations
            r'Tax\s+Due[:\s]*(\d{1,3}(?:,\d{3})*)',                   # Tax Due: 78,936
            r'Amount\s+Due[:\s]*(\d{1,3}(?:,\d{3})*)',                # Amount Due: 78,936
            # Tertiary: Summary table patterns
            r'(?:Total|Sum)[:\s]*(?:KES|Kshs?)?[:\s]*(\d{1,3}(?:,\d{3})*)', # Total: KES 78,936
            # Quaternary: Payment demand patterns
            r'(?:pay|remit)[:\s]*(?:KES|Kshs?)?[:\s]*(\d{1,3}(?:,\d{3})*)', # pay KES 78,936
            # Fallback: Large numbers in context
            r'\b(\d{1,3}(?:,\d{3})+)\b(?=\s*(?:KES|Kshs?|shillings))', # 78,936 KES
        ]
        
        for pattern in tax_patterns:
            tax_match = re.search(pattern, full_text, re.IGNORECASE)
            if tax_match:
                tax_amount = tax_match.group(1).strip()
                # Validate the amount (should be reasonable tax amount)
                numeric_value = int(tax_amount.replace(',', ''))
                if 1000 <= numeric_value <= 50000000:  # Reasonable tax range
                    data['Total_Tax'] = tax_amount
                    break
        
    except Exception as e:
        print(f"  ⚠️ Error extracting fields from notice {notice_num}: {e}")
    
    return data

def main():
    print("🔄 COMPLETE KRA BULK DATA EXTRACTOR")
    print("=" * 60)
    
    pdf_path = r"c:\Users\wisem\OneDrive\Desktop\Data Extraction\input6.pdf"
    
    if not os.path.exists(pdf_path):
        print(f"❌ PDF file not found: {pdf_path}")
        return
    
    try:
        # Open PDF
        doc = fitz.open(pdf_path)
        total_pages = doc.page_count
        total_notices = total_pages // 2
        
        print(f"📄 PDF Analysis:")
        print(f"   • Total pages: {total_pages}")
        print(f"   • Total notices: {total_notices}")
        print(f"   • Processing all {total_notices} notices...")
        print()
        
        results = []
        errors = []
        
        # Process ALL notices
        for notice_num in range(1, total_notices + 1):
            page1_idx = (notice_num - 1) * 2
            page2_idx = page1_idx + 1
            
            if page2_idx >= total_pages:
                break
            
            print(f"📋 Processing Notice {notice_num:2}/{total_notices}...", end=" ")
            
            try:
                page1 = doc[page1_idx]
                page2 = doc[page2_idx]
                
                notice_data = extract_notice_data(page1, page2, notice_num)
                results.append(notice_data)
                
                # Count successfully extracted fields
                filled_fields = len([v for v in notice_data.values() if v and str(v) != str(notice_num)])
                print(f"✅ {filled_fields}/8 fields extracted")
                
                # Show progress every 5 notices
                if notice_num % 5 == 0:
                    print(f"   📊 Progress: {notice_num}/{total_notices} notices processed")
                
            except Exception as e:
                print(f"❌ Error processing notice {notice_num}: {e}")
                errors.append(f"Notice {notice_num}: {str(e)}")
        
        doc.close()
        
        # Create comprehensive DataFrame
        df = pd.DataFrame(results)
        
        # Generate output filename with timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = rf"C:\Users\wisem\OneDrive\Desktop\KRA DATA EXTRACTION\KRA_Complete_Extract_{timestamp}.xlsx"
        
        # Save to Excel with formatting
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='KRA_Tax_Notices', index=False)
            
            # Auto-adjust column widths
            worksheet = writer.sheets['KRA_Tax_Notices']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        print("\n" + "=" * 60)
        print("🎉 EXTRACTION COMPLETE!")
        print(f"📊 Total notices processed: {len(results)}")
        print(f"💾 Data saved to: {os.path.basename(output_file)}")
        
        if errors:
            print(f"⚠️  {len(errors)} errors occurred during processing")
        
        # Show detailed extraction statistics
        print("\n📈 Field Extraction Statistics:")
        field_stats = {}
        for col in ['Date', 'PIN', 'Taxpayer_Name', 'Subject', 'Year', 'Station', 'Officer', 'Total_Tax']:
            filled = len([x for x in df[col] if x])
            percentage = (filled / len(df)) * 100 if len(df) > 0 else 0
            field_stats[col] = {'filled': filled, 'total': len(df), 'percentage': percentage}
            print(f"   {col:15}: {filled:2}/{len(df)} ({percentage:5.1f}%)")
        
        # Show sample of extracted data
        print(f"\n📋 Sample Extracted Data (First 5 Records):")
        sample_cols = ['Notice_Number', 'PIN', 'Taxpayer_Name', 'Station', 'Total_Tax']
        print(df[sample_cols].head().to_string(index=False))
        
        print(f"\n📋 Sample Extracted Data (Last 5 Records):")
        print(df[sample_cols].tail().to_string(index=False))
        
        # Summary statistics
        total_tax_amount = 0
        valid_tax_entries = 0
        for tax in df['Total_Tax']:
            if tax:
                try:
                    # Remove commas and convert to number
                    tax_num = float(tax.replace(',', ''))
                    total_tax_amount += tax_num
                    valid_tax_entries += 1
                except:
                    pass
        
        print(f"\n💰 Tax Summary:")
        print(f"   • Records with tax amounts: {valid_tax_entries}")
        print(f"   • Total tax amount: {total_tax_amount:,.0f}")
        print(f"   • Average tax per notice: {total_tax_amount/valid_tax_entries:,.0f}" if valid_tax_entries > 0 else "   • Average: N/A")
        
        print(f"\n✅ COMPLETE! All data extracted and saved to Excel file.")
        
        return df, output_file
        
    except Exception as e:
        print(f"❌ Fatal Error: {e}")
        import traceback
        traceback.print_exc()
        return None, None

if __name__ == "__main__":
    df, output_file = main()
    
    if df is not None:
        print(f"\n🎯 SUCCESS: {len(df)} KRA tax notices processed successfully!")
        print(f"📁 File location: {output_file}")
    else:
        print("❌ Extraction failed. Please check the error messages above.")