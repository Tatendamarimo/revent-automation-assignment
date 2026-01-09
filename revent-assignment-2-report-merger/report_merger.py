"""
Amazon & Noon Report Merger
Automatically merges Amazon and Noon data based on Column Relations Sheet
No hardcoding - fully reusable script
"""

import pandas as pd
from pathlib import Path
import sys
import glob
from tqdm import tqdm



class ReportMerger:
    """Merges Amazon and Noon reports based on dynamic column mapping"""
    
    def __init__(self, excel_file_path):
        """
        Initialize the Report Merger
        
        Args:
            excel_file_path: Path to the Excel file containing all sheets
        """
        self.excel_file_path = excel_file_path
        self.excel_file = pd.ExcelFile(excel_file_path)
        self.column_relations = None
        self.amazon_data = None
        self.noon_data = None
        self.summary_data = None
        
    def _find_sheet(self, keyword):
        """Helper to find sheet by keyword"""
        for sheet in self.excel_file.sheet_names:
            if keyword in sheet.lower():
                return sheet
        return None
    
    def load_sheets(self):
        """Load all required sheets from the Excel file"""
        print("Loading sheets...")
        print(f"Available sheets: {self.excel_file.sheet_names}")
        
        # Find and load Column Relations Sheet
        col_rel_sheet = self._find_sheet('column relations')
        if not col_rel_sheet:
            raise ValueError("Column Relations Sheet not found")
        self.column_relations = pd.read_excel(self.excel_file, sheet_name=col_rel_sheet, header=0)
        
        # Find and load Amazon sheet
        amazon_sheet = self._find_sheet('amazon')
        if not amazon_sheet:
            raise ValueError("Amazon sheet not found")
        self.amazon_data = pd.read_excel(self.excel_file, sheet_name=amazon_sheet)
        
        # Find and load Noon sheet
        noon_sheet = self._find_sheet('noon')
        if not noon_sheet:
            raise ValueError("Noon sheet not found")
        self.noon_data = pd.read_excel(self.excel_file, sheet_name=noon_sheet)
        
        print(f"✓ Loaded {len(self.column_relations)} mappings, {len(self.amazon_data)} Amazon rows, {len(self.noon_data)} Noon rows")
        
    def parse_column_mapping(self):
        """Parse the column relations sheet to create mapping dictionaries"""
        # Clean column names
        self.column_relations.columns = self.column_relations.columns.str.strip()
        
        # Find column names with flexible matching
        def find_col(keywords):
            for col in self.column_relations.columns:
                if all(k in col.lower() for k in keywords):
                    return col
            return None
        
        summary_col_name = find_col(['summary', 'column'])
        noon_col_name = find_col(['noon', 'column'])
        amazon_col_name = find_col(['amazon']) and (find_col(['amazon', 'column']) or find_col(['amazon', 'colum']))
        
        if not summary_col_name:
            raise ValueError("Could not find Summary Sheet column")
        
        print(f"\nMappings: Summary='{summary_col_name}', Noon='{noon_col_name}', Amazon='{amazon_col_name}'")
        
        # Create mappings
        noon_mapping = {}
        amazon_mapping = {}
        noon_remarks = {}
        amazon_remarks = {}
        
        for _, row in self.column_relations.iterrows():
            summary_col = str(row.get(summary_col_name, '')).strip()
            
            # Skip if summary column is empty or NaN
            if pd.isna(summary_col) or summary_col == '' or summary_col == 'nan':
                continue
            
            # Noon mapping
            if noon_col_name:
                noon_col = str(row.get(noon_col_name, '')).strip()
                if not pd.isna(noon_col) and noon_col != '' and noon_col != 'nan':
                    noon_mapping[summary_col] = noon_col
                    # Get first Remarks column for Noon
                    noon_remark = str(row.get('Remarks', '')).strip() if 'Remarks' in row.index else ''
                    if not pd.isna(noon_remark) and noon_remark != '' and noon_remark != 'nan':
                        noon_remarks[summary_col] = noon_remark
            
            # Amazon mapping
            if amazon_col_name:
                amazon_col = str(row.get(amazon_col_name, '')).strip()
                if not pd.isna(amazon_col) and amazon_col != '' and amazon_col != 'nan':
                    amazon_mapping[summary_col] = amazon_col
                    # Get second Remarks column for Amazon (Remarks.1)
                    amazon_remark = ''
                    for col in self.column_relations.columns:
                        if col == 'Remarks.1' or (col.startswith('Remarks') and col != 'Remarks'):
                            amazon_remark = str(row.get(col, '')).strip()
                            break
                    if amazon_remark == '' or pd.isna(amazon_remark) or amazon_remark == 'nan':
                        amazon_remark = str(row.get('Remarks', '')).strip() if 'Remarks' in row.index else ''
                    if not pd.isna(amazon_remark) and amazon_remark != '' and amazon_remark != 'nan':
                        amazon_remarks[summary_col] = amazon_remark
        
        return noon_mapping, amazon_mapping, noon_remarks, amazon_remarks
    
    def extract_date_component(self, date_value, component):
        """
        Extract day, month, or year from a date value
        
        Args:
            date_value: Date value (can be string, datetime, or number)
            component: 'day', 'month', or 'year'
        """
        if pd.isna(date_value):
            return None
            
        try:
            # Try to parse as datetime
            if isinstance(date_value, str):
                dt = pd.to_datetime(date_value, errors='coerce')
            elif isinstance(date_value, (int, float)):
                # Excel date number
                dt = pd.to_datetime('1899-12-30') + pd.Timedelta(days=date_value)
            else:
                dt = pd.to_datetime(date_value, errors='coerce')
            
            if pd.isna(dt):
                return None
                
            if component == 'day':
                return dt.day
            elif component == 'month':
                return dt.month
            elif component == 'year':
                return dt.year
        except:
            return None
    
    def apply_transformation(self, value, remark, source_col=None):
        """
        Apply transformation based on remarks
        
        Args:
            value: Original value
            remark: Transformation instruction from remarks
            source_col: Source column name for context
        """
        if pd.isna(value):
            return None
        
        remark_lower = str(remark).lower()
        
        # Check for "mark it NA" instruction
        if 'mark it "na"' in remark_lower or 'mark it "na"' in remark_lower:
            return "NA"
        
        # Check for date component extraction
        if 'day number from date' in remark_lower:
            return self.extract_date_component(value, 'day')
        elif 'month from date' in remark_lower:
            return self.extract_date_component(value, 'month')
        elif 'year from date' in remark_lower:
            return self.extract_date_component(value, 'year')
        
        # Check for multiplication (e.g., "multiply Price Including VAT with quantity")
        if 'multiply' in remark_lower:
            # This will be handled at row level, not single value
            return value
        
        # Check for contract/channel specific logic
        if 'if the contract is' in remark_lower or 'if sales channel is' in remark_lower:
            # This requires row-level context, return as-is for now
            return value
        
        # Default: return value as-is
        return value
    
    def process_noon_data(self, noon_mapping, noon_remarks):
        """Process Noon data according to mapping and remarks"""
        print("\nProcessing Noon data...")
        
        processed_rows = []
        
        for _, row in tqdm(self.noon_data.iterrows(), total=len(self.noon_data), desc="Processing Noon", unit="row"):
            processed_row = {'Source': 'Noon'}
            
            for summary_col, noon_col in noon_mapping.items():
                if noon_col in row.index:
                    value = row[noon_col]
                    
                    # Apply transformation if there's a remark
                    if summary_col in noon_remarks:
                        value = self.apply_transformation(value, noon_remarks[summary_col], noon_col)
                    
                    processed_row[summary_col] = value
                else:
                    processed_row[summary_col] = None
            
            # Handle special cases that need row-level context
            processed_row = self.handle_special_cases(processed_row, row, noon_remarks, 'noon')
            
            processed_rows.append(processed_row)
        
        print(f"✓ Processed {len(processed_rows)} Noon rows")
        return processed_rows
    
    def process_amazon_data(self, amazon_mapping, amazon_remarks):
        """Process Amazon data according to mapping and remarks"""
        print("\nProcessing Amazon data...")
        
        processed_rows = []
        
        for _, row in tqdm(self.amazon_data.iterrows(), total=len(self.amazon_data), desc="Processing Amazon", unit="row"):
            processed_row = {'Source': 'Amazon'}
            
            for summary_col, amazon_col in amazon_mapping.items():
                if amazon_col in row.index:
                    value = row[amazon_col]
                    
                    # Apply transformation if there's a remark
                    if summary_col in amazon_remarks:
                        value = self.apply_transformation(value, amazon_remarks[summary_col], amazon_col)
                    
                    processed_row[summary_col] = value
                else:
                    processed_row[summary_col] = None
            
            # Handle special cases that need row-level context
            processed_row = self.handle_special_cases(processed_row, row, amazon_remarks, 'amazon')
            
            processed_rows.append(processed_row)
        
        print(f"✓ Processed {len(processed_rows)} Amazon rows")
        return processed_rows
    
    def handle_special_cases(self, processed_row, original_row, remarks, source):
        """
        Handle special transformation cases that require row-level context
        
        Args:
            processed_row: The row being built for summary
            original_row: Original row from source data
            remarks: Remarks dictionary
            source: 'noon' or 'amazon'
        """
        # Handle Value (Including VAT) calculation
        # "multiply Price Including VAT with quantity for respective order id"
        for col, remark in remarks.items():
            remark_lower = str(remark).lower()
            
            if 'multiply' in remark_lower and 'price' in remark_lower and 'quantity' in remark_lower:
                # Find price and quantity columns
                price_col = None
                qty_col = None
                
                if source == 'noon':
                    # Look for price in original row
                    for key in original_row.index:
                        if 'price' in str(key).lower() and 'vat' in str(key).lower():
                            price_col = key
                        if 'quantity' in str(key).lower() or 'qty' in str(key).lower():
                            qty_col = key
                elif source == 'amazon':
                    for key in original_row.index:
                        if 'item price' in str(key).lower():
                            price_col = key
                        if 'quantity' in str(key).lower():
                            qty_col = key
                
                if price_col and qty_col:
                    try:
                        price = pd.to_numeric(original_row[price_col], errors='coerce')
                        qty = pd.to_numeric(original_row[qty_col], errors='coerce')
                        if not pd.isna(price) and not pd.isna(qty):
                            processed_row[col] = price * qty
                    except:
                        pass
            
            # Handle channel/contract specific logic
            if 'if the contract is' in remark_lower or 'if sales channel is' in remark_lower:
                # Extract the condition and action
                # Example: "if the contract is MPABANC\BKLSA, & document type is Invoice then it should be backated as noon KSA"
                if 'backated as noon ksa' in remark_lower or 'consider noon ksa' in remark_lower:
                    # Check conditions in original row
                    # This is a placeholder - actual implementation depends on exact column names
                    processed_row[col] = processed_row.get(col, None)
                elif 'consider amazon ksa' in remark_lower:
                    processed_row[col] = processed_row.get(col, None)
        
        return processed_row
    
    def merge_data(self):
        """Merge Amazon and Noon data into summary sheet"""
        print("\n" + "="*60)
        print("MERGING REPORTS")
        print("="*60)
        
        # Parse column mappings
        noon_mapping, amazon_mapping, noon_remarks, amazon_remarks = self.parse_column_mapping()
        
        print(f"\nColumn Mappings:")
        print(f"  Noon columns: {len(noon_mapping)}")
        print(f"  Amazon columns: {len(amazon_mapping)}")
        
        # Process both datasets
        noon_processed = self.process_noon_data(noon_mapping, noon_remarks)
        amazon_processed = self.process_amazon_data(amazon_mapping, amazon_remarks)
        
        # Combine into single DataFrame
        all_rows = noon_processed + amazon_processed
        self.summary_data = pd.DataFrame(all_rows)
        
        # Get all unique columns from column relations
        all_summary_cols = list(set(list(noon_mapping.keys()) + list(amazon_mapping.keys())))
        
        # Ensure all columns exist
        for col in all_summary_cols:
            if col not in self.summary_data.columns:
                self.summary_data[col] = None
        
        # Add Source column at the beginning
        cols = ['Source'] + [col for col in all_summary_cols if col in self.summary_data.columns]
        self.summary_data = self.summary_data[cols]
        
        print(f"\n✓ Merged data: {len(self.summary_data)} total rows")
        print(f"  - Noon: {len(noon_processed)} rows")
        print(f"  - Amazon: {len(amazon_processed)} rows")
        
    def save_summary(self, output_path=None):
        """Save the summary sheet to Excel file"""
        if output_path is None:
            # Create output filename based on input
            input_path = Path(self.excel_file_path)
            output_path = input_path.parent / f"{input_path.stem}_MERGED.xlsx"
        
        print(f"\nSaving summary to: {output_path}")
        
        # Create Excel writer
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # Write summary sheet
            self.summary_data.to_excel(writer, sheet_name='Summary Sheet', index=False)
            
            # Optionally copy original sheets
            self.amazon_data.to_excel(writer, sheet_name='Amazon', index=False)
            self.noon_data.to_excel(writer, sheet_name='Noon', index=False)
            self.column_relations.to_excel(writer, sheet_name='Column Relations Sheet', index=False)
        
        print(f"✓ Summary saved successfully!")
        print(f"  Output file: {output_path}")
        
        return output_path
    
    def run(self, output_path=None):
        """Run the complete merge process"""
        print("\n" + "="*60)
        print("AMAZON & NOON REPORT MERGER")
        print("="*60)
        print(f"Input file: {self.excel_file_path}\n")
        
        # Load sheets
        self.load_sheets()
        
        # Merge data
        self.merge_data()
        
        # Save summary
        output_file = self.save_summary(output_path)
        
        print("\n" + "="*60)
        print("MERGE COMPLETE!")
        print("="*60)
        
        return output_file


def main():
    """Main function to run the report merger"""
    # Get Excel file path from command line or auto-detect
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
    else:
        excel_files = glob.glob("*.xlsx")
        if not excel_files:
            print("Error: No Excel file found!\nUsage: python report_merger.py <excel_file_path>")
            sys.exit(1)
        excel_file = excel_files[0]
        print(f"Using: {excel_file}")
    
    output_path = sys.argv[2] if len(sys.argv) > 2 else None
    ReportMerger(excel_file).run(output_path)


if __name__ == "__main__":
    main()
