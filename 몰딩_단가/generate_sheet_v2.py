import pandas as pd
import os

def generate_v2():
    docs_dir = 'Docs/예림'
    output_dir = 'output'
    output_path = os.path.join(output_dir, '새시트_생성결과_v2.xlsx')

    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Load Files
    print("Loading Files (Yerim)...")
    df_template = pd.read_csv(os.path.join(docs_dir, '템플릿.csv'))
    df_sheet21 = pd.read_csv(os.path.join(docs_dir, '시트21.csv'), header=None)
    df_prices = pd.read_csv(os.path.join(docs_dir, '기준가격.csv'), header=6)

    # Preprocessing Base Prices for lookup
    # Product lookup: B열 (col 1), Basic Price: J열 (col 9) "공급가"
    # Diffs: Basic/Solid (11), Metallic/HP (13), Matte (15)
    price_map = {}
    for _, row in df_prices.iterrows():
        name = str(row.iloc[1]).strip() # Product Name in Column B (index 1)
        if not name or name == 'nan': continue
        
        # If product already exists and new row has NaN for diffs, skip it.
        # Or if the existing one is already solid, skip.
        # Here we prioritize the row that has '공급가' (col 9) and '베이직 차액' (col 11)
        base_price = row.iloc[9]
        if name in price_map and pd.isna(row.iloc[11]):
            continue

        price_map[name] = {
            'common_base': base_price, # 공급가 (Col J)
            '베이직': row.iloc[11],
            '솔리드': row.iloc[11],
            '메탈릭': row.iloc[13],
            'HP': row.iloc[13],
            '매트': row.iloc[15]
        }

    # Initial Header from Template
    header = list(df_template.columns)
    output_data = []

    # Iterate through Sheet21 (39 products)
    print(f"Processing {len(df_sheet21)} products...")
    for s_idx in range(len(df_sheet21)):
        product_row_21 = df_sheet21.iloc[s_idx]
        template_name = str(product_row_21.iloc[7]).strip() # Col H
        
        # Lookup prices with fallback (Retry with ◆)
        p_info = price_map.get(template_name, None)
        if not p_info:
            p_info = price_map.get(template_name + " ◆", None)
            
        if not p_info:
            print(f"Warning: Template name '{template_name}' not found in 기준가격.csv (even with fallback)")
            p_info = {'common_base': None, '베이직': 0, '솔리드': 0, '메탈릭': 0, 'HP': 0, '매트': 0}

        # Iterate through Template data rows (185 rows)
        for t_idx in range(len(df_template)):
            row_data = list(df_template.iloc[t_idx])
            
            # Rule 2: Fill A~H for the first row of each block
            if t_idx == 0:
                for col_i in range(8): # A~H
                    row_data[col_i] = product_row_21.iloc[col_i]
            
            # Rule 3: W column (index 22) Basic Price
            if t_idx == 0:
                row_data[22] = p_info['common_base']
            else:
                row_data[22] = None
            
            # Rule 4: Z column (index 25) Color Extra Charge
            # Using Row 33 onwards (t_idx >= 31)
            if t_idx <= 30:
                row_data[25] = "-"
            else:
                # Grade is in col 14 (O열)
                grade = str(row_data[14]).strip() # Changed to col 14 as per instruction
                if grade in p_info and grade != '베스트':
                    row_data[25] = p_info[grade]
                else:
                    row_data[25] = "-"

            output_data.append(row_data)

    # Save to Excel
    print("Saving to Excel (v2)...")
    df_result = pd.DataFrame(output_data, columns=header)
    df_result.to_excel(output_path, index=False)
    print(f"Successfully created {output_path}")
    print(f"Total rows: {len(df_result) + 1} (including header)")

if __name__ == "__main__":
    generate_v2()
