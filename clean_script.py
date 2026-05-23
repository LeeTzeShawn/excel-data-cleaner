import pandas as pd
import sys
import os

def data_cleaning(input_file, output_file=None):
    """
    Clean an Excel file containing player data and win/loss records.
    """
    
    print(f"📁 Loading: {input_file}")
    
    # STEP 1: Load the Excel file
    df = pd.read_excel(input_file)
    print(f"📊 Raw data shape: {df.shape}")
    
    # STEP 2: Extract Table 1 (Player Info) - rows 0-4, columns 0-5
    print("\n📋 Extracting Table 1 (Player Information)...")
    df1 = df.iloc[0:4, 0:5].copy()
    
    # Only drop columns if they exist
    cols_to_drop = ['Random_Column', 'Random_Column_2']
    existing_cols = [col for col in cols_to_drop if col in df1.columns]
    if existing_cols:
        df1 = df1.drop(columns=existing_cols)
    
    df1 = df1.reset_index(drop=True)
    print(f"   ✅ Table 1 shape: {df1.shape}")
    print(f"   📝 Players: {df1['Name'].tolist()}")
    
    # STEP 3: Extract Table 2 (Win/Loss Records) - rows 10-15, columns 6-9
    print("\n📋 Extracting Table 2 (Win/Loss Records)...")
    df2 = df.iloc[10:15, 6:9].copy()
    df2.columns = df2.iloc[0]
    df2 = df2.iloc[1:, :]
    df2 = df2.reset_index(drop=True)
    df2['Wins'] = pd.to_numeric(df2['Wins'], errors='coerce').astype('Int64')
    df2['Losses'] = pd.to_numeric(df2['Losses'], errors='coerce').astype('Int64')
    print(f"   ✅ Table 2 shape: {df2.shape}")
    
    # STEP 4: Reorder Table 2 to match Table 1 order
    print("\n🔄 Reordering Table 2...")
    desired_order = df1['Name'].tolist()
    df2 = df2.set_index('Name').reindex(desired_order).reset_index()
    print(f"   ✅ Reordered: {df2['Name'].tolist()}")
    
    # STEP 5: Combine tables
    print("\n🔗 Combining tables...")
    combined = pd.concat([df1, df2[['Wins', 'Losses']]], axis=1)
    
    # STEP 6: Add calculated columns
    print("\n🧮 Adding calculated columns...")
    # Avoid division by zero
    total_matches = combined['Wins'] + combined['Losses']
    combined['Win_Percentage'] = ((combined['Wins'] / total_matches) * 100).round(2)
    combined['Total_Matches'] = total_matches
    combined['Win_Loss_Ratio'] = (combined['Wins'] / combined['Losses']).round(2)
    combined['Rank_by_Wins'] = combined['Wins'].rank(ascending=False, method='min').astype('Int64')
    print(f"   ✅ Added: Win_Percentage, Total_Matches, Win_Loss_Ratio, Rank_by_Wins")
    
    # STEP 7: Save the cleaned data - FIXED VERSION
    if output_file is None:
        base_name = os.path.splitext(input_file)[0]
        output_file = f"{base_name}_cleaned.xlsx"
    
    # Make sure output_file has .xlsx extension
    if not output_file.endswith('.xlsx'):
        output_file += '.xlsx'
    
    # Save using openpyxl engine (explicitly)
    try:
        combined.to_excel(output_file, index=False, engine='openpyxl')
        print(f"\n✅ File saved successfully: {output_file}")
    except Exception as e:
        print(f"\n❌ Error saving with openpyxl: {e}")
        # Fallback to xlsxwriter
        try:
            combined.to_excel(output_file, index=False, engine='xlsxwriter')
            print(f"✅ File saved with xlsxwriter: {output_file}")
        except Exception as e2:
            print(f"❌ Both engines failed: {e2}")
            # Last resort: save as CSV
            csv_file = output_file.replace('.xlsx', '.csv')
            combined.to_csv(csv_file, index=False)
            print(f"✅ Saved as CSV instead: {csv_file}")
            return combined
    
    print("\n" + "="*50)
    print("✅ CLEANING COMPLETE!")
    print("="*50)
    print(f"📁 Output: {output_file}")
    
    # Preview the data
    print("\n📊 CLEANED DATA PREVIEW:")
    print(combined.to_string(index=False))
    
    return combined

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("="*50)
        print("📁 Excel Data Cleaner")
        print("="*50)
        print("\nUsage:")
        print("  python clean_script.py <excel_file>")
        print("\nExample:")
        print("  python clean_script.py Random_Test_UserInput.xlsx")
        print("="*50)
        sys.exit(1)
    
    input_file = sys.argv[1]
    
    if not os.path.exists(input_file):
        print(f"❌ Error: File '{input_file}' not found!")
        sys.exit(1)
    
    data_cleaning(input_file)
