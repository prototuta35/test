import pandas as pd
import numpy as np
import os

class ExcelProcessor:
    def __init__(self, input_folder, output_folder):
        self.input_folder = input_folder
        self.output_folder = output_folder
        os.makedirs(self.output_folder, exist_ok=True)

    def load_excel_files(self):
        """Load all Excel files from input folder"""
        files = [f for f in os.listdir(self.input_folder) if f.endswith('.xlsx')]
        return [pd.read_excel(os.path.join(self.input_folder, f)) for f in files]

    def clean_text_columns(self, df):
        """Clean and normalize text columns"""
        text_cols = df.select_dtypes(include=['object']).columns
        for col in text_cols:
            df[col] = df[col].str.strip().str.lower()
        return df

    def calculate_basic_stats(self, df):
        """Calculate basic statistics for numeric columns"""
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        return df[numeric_cols].agg(['mean', 'median', 'std', 'min', 'max'])

    def filter_features(self, df, columns=None, conditions=None):
        """Filter dataframe based on selected columns and conditions"""
        if columns:
            df = df[columns]
        if conditions:
            df = df.query(conditions)
        return df

    def process_files(self):
        """Main processing function"""
        dfs = self.load_excel_files()
        for i, df in enumerate(dfs):
            # Process each file
            df = self.clean_text_columns(df)
            stats = self.calculate_basic_stats(df)
            processed_df = self.filter_features(df)
            
            # Save results
            output_path = os.path.join(self.output_folder, f'processed_{i+1}.xlsx')
            with pd.ExcelWriter(output_path) as writer:
                processed_df.to_excel(writer, sheet_name='Processed Data', index=False)
                stats.to_excel(writer, sheet_name='Statistics')

if __name__ == "__main__":
    processor = ExcelProcessor('input', 'output')
    processor.process_files()
