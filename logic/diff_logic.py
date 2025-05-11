import pandas as pd
from datetime import datetime
import os

class DataComparator:
    def __init__(self):
        self.report_dir = os.path.join('results', 'compare_reports')
        os.makedirs(self.report_dir, exist_ok=True)
        
    def validate_format(self, files):
        """Validate if multiple table files have consistent format"""
        if len(files) < 2:
            return False, "At least 2 files are required for comparison"
            
        # Get the first file's format as baseline
        base_df = pd.read_excel(files[0]) if files[0].endswith('.xlsx') else pd.read_csv(files[0])
        base_columns = list(base_df.columns)
        base_shape = base_df.shape
        
        for file in files[1:]:
            df = pd.read_excel(file) if file.endswith('.xlsx') else pd.read_csv(file)
            if list(df.columns) != base_columns:
                return False, f"File {os.path.basename(file)} has inconsistent headers"
            if df.shape[1] != base_shape[1]:
                return False, f"File {os.path.basename(file)} has inconsistent column count"
                
        return True, "Format validation passed"

    def manual_compare(self, files):
        """Manually compare multiple files"""
        is_valid, msg = self.validate_format(files)
        if not is_valid:
            return False, msg, None
            
        # Execute comparison logic
        diffs = []
        base_df = pd.read_excel(files[0]) if files[0].endswith('.xlsx') else pd.read_csv(files[0])
        
        for i in range(1, len(files)):
            comp_df = pd.read_excel(files[i]) if files[i].endswith('.xlsx') else pd.read_csv(files[i])
            diff = self._find_differences(base_df, comp_df, files[0], files[i])
            diffs.extend(diff)
            
        # Generate report
        report = self._generate_report(diffs, 'manual')
        return True, "Comparison completed", report
        
    def db_compare(self, db_file, input_file):
        """Compare with database file"""
        # Validate format
        db_df = pd.read_excel(db_file) if db_file.endswith('.xlsx') else pd.read_csv(db_file)
        input_df = pd.read_excel(input_file) if input_file.endswith('.xlsx') else pd.read_csv(input_file)
        
        if list(input_df.columns) != list(db_df.columns):
            return False, "Input file has inconsistent headers with database", None
            
        # Execute comparison
        updates = []
        new_items = []
        matches = 0
        
        for idx, row in input_df.iterrows():
            product_id = row[1]  # Column B is product ID
            db_row = db_df[db_df.iloc[:,1] == product_id]
            
            if db_row.empty:
                # New product
                new_items.append(row)
                db_df = db_df.append(row, ignore_index=True)
            else:
                # Compare product info
                diff = self._compare_rows(db_row.iloc[0], row)
                if diff:
                    updates.append({
                        'product_id': product_id,
                        'differences': diff,
                        'old_data': db_row.iloc[0],
                        'new_data': row
                    })
                    # Update database
                    db_df.loc[db_df.iloc[:,1] == product_id] = row
                else:
                    matches += 1
                    
        # Save updated database
        if db_file.endswith('.xlsx'):
            writer = pd.ExcelWriter(db_file, engine='openpyxl')
            db_df.to_excel(writer, index=False)
            writer.save()
        else:
            db_df.to_csv(db_file, index=False)
            
        # Generate report
        report = {
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'new_items': new_items,
            'updates': updates,
            'matches': matches,
            'total_compared': len(input_df)
        }
        
        return True, "Database comparison completed", report
        
    def _find_differences(self, df1, df2, file1, file2):
        """Find differences between two dataframes"""
        diffs = []
        for idx in range(len(df1)):
            row1 = df1.iloc[idx]
            row2 = df2.iloc[idx]
            
            diff_cols = []
            for col in df1.columns:
                if row1[col] != row2[col]:
                    diff_cols.append({
                        'column': col,
                        'file1_value': row1[col],
                        'file2_value': row2[col]
                    })
                    
            if diff_cols:
                diffs.append({
                    'row': idx+2,  # Starting from row 2
                    'file1': os.path.basename(file1),
                    'file2': os.path.basename(file2),
                    'differences': diff_cols
                })
                
        return diffs
        
    def _compare_rows(self, row1, row2):
        """Compare differences between two rows"""
        diffs = []
        for col in row1.index:
            if row1[col] != row2[col]:
                diffs.append({
                    'column': col,
                    'old_value': row1[col],
                    'new_value': row2[col]
                })
        return diffs
        
    def _generate_report(self, data, compare_type):
        """Generate comparison report"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_path = os.path.join(self.report_dir, f"{compare_type}_compare_{timestamp}.txt")
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(f"Comparison Report - {timestamp}\n")
            f.write("="*50 + "\n")
            
            if compare_type == 'manual':
                for diff in data:
                    f.write(f"Difference location: Row {diff['row']}\n")
                    f.write(f"File 1: {diff['file1']}\n")
                    f.write(f"File 2: {diff['file2']}\n")
                    for item in diff['differences']:
                        f.write(f"Column '{item['column']}':\n")
                        f.write(f"  File 1 value: {item['file1_value']}\n")
                        f.write(f"  File 2 value: {item['file2_value']}\n")
                    f.write("-"*50 + "\n")
            else:
                f.write(f"Total products compared: {data['total_compared']}\n")
                f.write(f"Matching products: {data['matches']}\n")
                f.write(f"New products: {len(data['new_items'])}\n")
                f.write(f"Updated products: {len(data['updates'])}\n\n")
                
                if data['new_items']:
                    f.write("New products:\n")
                    for item in data['new_items']:
                        f.write(f"Product ID: {item[1]}\n")
                        
                if data['updates']:
                    f.write("\nUpdated products:\n")
                    for update in data['updates']:
                        f.write(f"Product ID: {update['product_id']}\n")
                        for diff in update['differences']:
                            f.write(f"Column '{diff['column']}':\n")
                            f.write(f"  Old value: {diff['old_value']}\n")
                            f.write(f"  New value: {diff['new_value']}\n")
                        f.write("\n")
                        
        return report_path
