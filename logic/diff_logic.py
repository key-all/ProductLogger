import pandas as pd
from datetime import datetime
import os

class DataComparator:
    def __init__(self):
        self.report_dir = os.path.join('results', 'compare_reports')
        os.makedirs(self.report_dir, exist_ok=True)
        
    def validate_format(self, files):
        """验证多个表格文件格式是否一致"""
        if len(files) < 2:
            return False, "至少需要2个文件进行比对"
            
        # 获取第一个文件的格式作为基准
        base_df = pd.read_excel(files[0]) if files[0].endswith('.xlsx') else pd.read_csv(files[0])
        base_columns = list(base_df.columns)
        base_shape = base_df.shape
        
        for file in files[1:]:
            df = pd.read_excel(file) if file.endswith('.xlsx') else pd.read_csv(file)
            if list(df.columns) != base_columns:
                return False, f"文件{os.path.basename(file)}表头不一致"
            if df.shape[1] != base_shape[1]:
                return False, f"文件{os.path.basename(file)}列数不一致"
                
        return True, "格式验证通过"

    def manual_compare(self, files):
        """手动比对多个文件"""
        is_valid, msg = self.validate_format(files)
        if not is_valid:
            return False, msg, None
            
        # 执行比对逻辑
        diffs = []
        base_df = pd.read_excel(files[0]) if files[0].endswith('.xlsx') else pd.read_csv(files[0])
        
        for i in range(1, len(files)):
            comp_df = pd.read_excel(files[i]) if files[i].endswith('.xlsx') else pd.read_csv(files[i])
            diff = self._find_differences(base_df, comp_df, files[0], files[i])
            diffs.extend(diff)
            
        # 生成报告
        report = self._generate_report(diffs, 'manual')
        return True, "比对完成", report
        
    def db_compare(self, db_file, input_file):
        """与数据库文件比对"""
        # 验证格式
        db_df = pd.read_excel(db_file) if db_file.endswith('.xlsx') else pd.read_csv(db_file)
        input_df = pd.read_excel(input_file) if input_file.endswith('.xlsx') else pd.read_csv(input_file)
        
        if list(input_df.columns) != list(db_df.columns):
            return False, "输入文件与数据库表头不一致", None
            
        # 执行比对
        updates = []
        new_items = []
        matches = 0
        
        for idx, row in input_df.iterrows():
            product_id = row[1]  # B列是商品ID
            db_row = db_df[db_df.iloc[:,1] == product_id]
            
            if db_row.empty:
                # 新增商品
                new_items.append(row)
                db_df = db_df.append(row, ignore_index=True)
            else:
                # 比对商品信息
                diff = self._compare_rows(db_row.iloc[0], row)
                if diff:
                    updates.append({
                        'product_id': product_id,
                        'differences': diff,
                        'old_data': db_row.iloc[0],
                        'new_data': row
                    })
                    # 更新数据库
                    db_df.loc[db_df.iloc[:,1] == product_id] = row
                else:
                    matches += 1
                    
        # 保存更新后的数据库
        if db_file.endswith('.xlsx'):
            writer = pd.ExcelWriter(db_file, engine='openpyxl')
            db_df.to_excel(writer, index=False)
            writer.save()
        else:
            db_df.to_csv(db_file, index=False)
            
        # 生成报告
        report = {
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'new_items': new_items,
            'updates': updates,
            'matches': matches,
            'total_compared': len(input_df)
        }
        
        return True, "数据库比对完成", report
        
    def _find_differences(self, df1, df2, file1, file2):
        """找出两个数据框之间的差异"""
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
                    'row': idx+2,  # 从第2行开始
                    'file1': os.path.basename(file1),
                    'file2': os.path.basename(file2),
                    'differences': diff_cols
                })
                
        return diffs
        
    def _compare_rows(self, row1, row2):
        """比较两行数据的差异"""
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
        """生成比对报告"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        report_path = os.path.join(self.report_dir, f"{compare_type}_compare_{timestamp}.txt")
        
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(f"比对报告 - {timestamp}\n")
            f.write("="*50 + "\n")
            
            if compare_type == 'manual':
                for diff in data:
                    f.write(f"差异位置: 行 {diff['row']}\n")
                    f.write(f"文件1: {diff['file1']}\n")
                    f.write(f"文件2: {diff['file2']}\n")
                    for item in diff['differences']:
                        f.write(f"列 '{item['column']}':\n")
                        f.write(f"  文件1值: {item['file1_value']}\n")
                        f.write(f"  文件2值: {item['file2_value']}\n")
                    f.write("-"*50 + "\n")
            else:
                f.write(f"比对商品总数: {data['total_compared']}\n")
                f.write(f"匹配商品数: {data['matches']}\n")
                f.write(f"新增商品数: {len(data['new_items'])}\n")
                f.write(f"更新商品数: {len(data['updates'])}\n\n")
                
                if data['new_items']:
                    f.write("新增商品:\n")
                    for item in data['new_items']:
                        f.write(f"商品ID: {item[1]}\n")
                        
                if data['updates']:
                    f.write("\n更新商品:\n")
                    for update in data['updates']:
                        f.write(f"商品ID: {update['product_id']}\n")
                        for diff in update['differences']:
                            f.write(f"列 '{diff['column']}':\n")
                            f.write(f"  旧值: {diff['old_value']}\n")
                            f.write(f"  新值: {diff['new_value']}\n")
                        f.write("\n")
                        
        return report_path
