import pandas as pd
import os

def generate():
    docs_dir = 'Docs'
    output_dir = 'output'
    output_path = os.path.join(output_dir, '새시트_생성결과.xlsx')

    # Load Files
    print("Loading Files...")
    # Template: Row 1 is header. shape: 229x34 (1 header + 228 data)
    df_template = pd.read_csv(os.path.join(docs_dir, '템플릿.csv'))
    
    # Sheet21: No header. shape: 37x8
    df_sheet21 = pd.read_csv(os.path.join(docs_dir, '시트21.csv'), header=None)
    
    # Base Prices: Header is at line 7 (index 6). shape: 234x30
    df_prices = pd.read_csv(os.path.join(docs_dir, '기준가격.csv'), header=6)

    # Preprocessing Base Prices for lookup
    # Column indexing for df_prices:
    # K: 상품명 (col 10), D: 단가 (col 3), N: YL-2차액 (col 13), P: YL-3차액 (col 15), R: YL-4차액 (col 17), T: YL-5차액 (col 19), V: YL-6차액 (col 21)
    price_map = {}
    for _, row in df_prices.iterrows():
        name = str(row.iloc[10]).strip() # Changed from col 1 to col 10 (K열)
        if not name or name == 'nan': continue
        price_map[name] = {
            'base_price': row.iloc[3],
            'YL-2': row.iloc[13],
            'YL-3': row.iloc[15],
            'YL-4': row.iloc[17],
            'YL-5': row.iloc[19],
            'YL-6': row.iloc[21]
        }

    # Initial Header from Template
    header = list(df_template.columns)
    output_data = []

    # Iterate through Sheet21 (37 products)
    print(f"Processing {len(df_sheet21)} products...")
    for s_idx in range(len(df_sheet21)):
        product_row_21 = df_sheet21.iloc[s_idx]
        template_name = str(product_row_21.iloc[7]).strip() # Col H
        
        # Lookup prices with fallback (Retry with ◆)
        p_info = price_map.get(template_name, None)
        if not p_info:
            p_info = price_map.get(template_name + " ◆", None) # Retry with ◆
            
        if not p_info:
            print(f"Warning: Template name '{template_name}' not found in 기준가격.csv (even with fallback)")
            # fallback to empty/original
            p_info = {'base_price': None, 'YL-2': 0, 'YL-3': 0, 'YL-4': 0, 'YL-5': 0, 'YL-6': 0}

        # Iterate through Template rows (228 rows)
        for t_idx in range(len(df_template)):
            # Start with template row data
            row_data = list(df_template.iloc[t_idx])
            
            # Rule 2: Fill A~H for the first row of each block
            if t_idx == 0:
                for col_i in range(8): # A~H
                    row_data[col_i] = product_row_21.iloc[col_i]
            
            # Rule 3: W column (index 22) Basic Price
            if t_idx == 0:
                row_data[22] = p_info['base_price']
            else:
                row_data[22] = None
            
            # Rule 4: Z column (index 25) Color Extra Charge
            if t_idx <= 30:
                row_data[25] = "-"
            else:
                # Row 33 onwards (t_idx >= 31)
                yl_grade = str(row_data[14]).strip() # Col O
                if yl_grade in p_info and yl_grade != 'YL-1':
                    diff = p_info[yl_grade]
                    row_data[25] = diff # Number only
                else:
                    row_data[25] = "-"

            output_data.append(row_data)

    # Create DataFrame and save to Excel
    print("Saving to Excel...")
    df_result = pd.DataFrame(output_data, columns=header)
    
    # Formatting adjustment: VAT포함 컬럼 might need to be numeric
    # Cleaning commas or types if necessary, though pandas handles most.
    
    df_result.to_excel(output_path, index=False)
    print(f"Successfully created {output_path}")
    print(f"Total rows: {len(df_result) + 1} (including header)")

if __name__ == "__main__":
    generate()
