import pandas as pd
import yaml
from pathlib import Path
from typing import Dict, List, Any, Union
import numpy as np


class ExcelTransformer:
    def __init__(self, config_path: str):
        """Initialize transformer with configuration file."""
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = yaml.safe_load(f)
        self.df = None
        self.input_file_path = None

    def load_data(self, file_path: str) -> pd.DataFrame:
        """Load data from various formats (HTML, Excel, CSV)."""
        file_path = Path(file_path)
        self.input_file_path = file_path  # Store input file path for later use
        load_options = self.config.get('load_options', {})
        
        if file_path.suffix.lower() in ['.html', '.xls']:
            # Handle HTML files with .xls extension (common from legacy systems)
            self.df = pd.read_html(file_path, encoding='utf-8')[0]
        elif file_path.suffix.lower() in ['.xlsx']:
            # Support sheet selection and data_only mode (values without formulas)
            sheet_name = load_options.get('sheet_name', 0)  # Default first sheet
            data_only = load_options.get('data_only', False)  # Load formulas by default
            
            # Load with pandas first, then convert formulas to values if needed
            self.df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            if data_only:
                # Convert formula results to values by re-evaluating the DataFrame
                # This handles most common Excel formulas automatically
                for col in self.df.columns:
                    # Try to evaluate any Excel-like expressions to numeric values
                    self.df[col] = pd.to_numeric(self.df[col], errors='ignore')
        elif file_path.suffix.lower() in ['.csv']:
            self.df = pd.read_csv(file_path)
        else:
            raise ValueError(f"Unsupported file format: {file_path.suffix}")
        
        return self.df

    def apply_transformations(self) -> pd.DataFrame:
        """Apply all transformations from config."""
        if self.df is None:
            raise ValueError("No data loaded. Call load_data() first.")
        
        for transformation in self.config.get('transformations', []):
            self._apply_single_transformation(transformation)
        
        return self.df

    def _apply_single_transformation(self, rule: Dict[str, Any]):
        """Apply a single transformation rule."""
        rule_type = rule['type']
        
        if rule_type == 'delete_columns':
            self._delete_columns(rule)
        elif rule_type == 'delete_rows':
            self._delete_rows(rule)
        elif rule_type == 'round_numbers':
            self._round_numbers(rule)
        elif rule_type == 'rename_columns':
            self._rename_columns(rule)
        elif rule_type == 'reorder_columns':
            self._reorder_columns(rule)
        elif rule_type == 'select_columns':
            self._select_columns(rule)
        elif rule_type == 'format_numbers':
            self._format_numbers(rule)
        elif rule_type == 'date_format':
            self._format_dates(rule)
        elif rule_type == 'format_accounting':
            self._format_accounting(rule)
        elif rule_type == 'round_last_column':
            self._round_last_column(rule)
        elif rule_type == 'replace_decimal_separator':
            self._replace_decimal_separator(rule)
        else:
            print(f"Warning: Unknown transformation type: {rule_type}")

    def _delete_columns(self, rule: Dict[str, Any]):
        """Delete specified columns."""
        columns_to_delete = []
        for col in rule['columns']:
            if isinstance(col, int):
                # Delete by index
                if col < len(self.df.columns):
                    columns_to_delete.append(self.df.columns[col])
            elif isinstance(col, str):
                # Delete by name or pattern
                if col in self.df.columns:
                    columns_to_delete.append(col)
                else:
                    # Try pattern matching
                    matching_cols = [c for c in self.df.columns if col.lower() in str(c).lower()]
                    columns_to_delete.extend(matching_cols)
        
        self.df = self.df.drop(columns=columns_to_delete, errors='ignore')

    def _delete_rows(self, rule: Dict[str, Any]):
        """Delete rows based on conditions."""
        for condition in rule.get('conditions', []):
            if 'contains' in condition:
                # Delete rows containing specific text
                mask = True
                for text in condition['contains']:
                    mask &= ~self.df.astype(str).apply(
                        lambda x: x.str.contains(str(text), case=False, na=False)
                    ).any(axis=1)
                self.df = self.df[mask]
            
            elif 'empty' in condition and condition['empty']:
                # Delete empty rows
                self.df = self.df.dropna(how='all')
            
            elif 'range' in condition:
                # Delete specific row range
                start, end = condition['range']
                indices_to_drop = list(range(start, min(end + 1, len(self.df))))
                self.df = self.df.drop(self.df.index[indices_to_drop])

    def _round_numbers(self, rule: Dict[str, Any]):
        """Round numbers in specified columns."""
        decimals = rule.get('decimals', 2)
        for col in rule['columns']:
            if col in self.df.columns:
                self.df[col] = pd.to_numeric(self.df[col], errors='coerce').round(decimals)

    def _rename_columns(self, rule: Dict[str, Any]):
        """Rename columns according to mapping."""
        self.df = self.df.rename(columns=rule['mapping'])

    def _reorder_columns(self, rule: Dict[str, Any]):
        """Reorder columns according to specified order."""
        new_order = []
        for col in rule['order']:
            if col in self.df.columns:
                new_order.append(col)
        
        # Add remaining columns
        remaining = [col for col in self.df.columns if col not in new_order]
        self.df = self.df[new_order + remaining]

    def _select_columns(self, rule: Dict[str, Any]):
        """Select only specified columns."""
        columns_to_keep = [col for col in rule['columns'] if col in self.df.columns]
        self.df = self.df[columns_to_keep]

    def _format_numbers(self, rule: Dict[str, Any]):
        """Format numbers according to specified format."""
        format_type = rule.get('format', 'standard')
        decimals = rule.get('decimals', 2)
        
        for col in rule['columns']:
            if col in self.df.columns:
                numeric_col = pd.to_numeric(self.df[col], errors='coerce')
                
                if format_type == 'fixed_point':
                    # Format for legacy system import (no thousands separator)
                    self.df[col] = numeric_col.map(lambda x: f"{x:.{decimals}f}" if pd.notna(x) else "")
                elif format_type == 'currency':
                    currency = rule.get('currency', 'EUR')
                    self.df[col] = numeric_col.map(lambda x: f"{x:,.{decimals}f} {currency}" if pd.notna(x) else "")

    def _format_dates(self, rule: Dict[str, Any]):
        """Format dates according to specified formats."""
        from_format = rule.get('from_format')
        to_format = rule.get('to_format')
        
        for col in rule['columns']:
            if col in self.df.columns:
                if from_format:
                    self.df[col] = pd.to_datetime(self.df[col], format=from_format, errors='coerce')
                else:
                    self.df[col] = pd.to_datetime(self.df[col], errors='coerce')
                
                if to_format:
                    self.df[col] = self.df[col].dt.strftime(to_format)

    def _format_accounting(self, rule: Dict[str, Any]):
        """Format columns as accounting numbers with decimals."""
        decimals = rule.get('decimals', 2)
        columns = rule['columns']
        
        # Convert column letters to indices if needed
        column_indices = []
        for col in columns:
            if isinstance(col, str) and len(col) == 1 and col.isalpha():
                # Convert letter (A, B, C...) to index (0, 1, 2...)
                col_index = ord(col.upper()) - ord('A')
                if col_index < len(self.df.columns):
                    column_indices.append(col_index)
            else:
                column_indices.append(col)
        
        # Apply formatting to the specified columns
        for col_idx in column_indices:
            if col_idx < len(self.df.columns):
                col_name = self.df.columns[col_idx]
                # Convert to numeric and round
                self.df[col_name] = pd.to_numeric(self.df[col_name], errors='coerce').round(decimals)
                # Store formatting info for later use in export
                if not hasattr(self, '_accounting_columns'):
                    self._accounting_columns = []
                self._accounting_columns.append((col_name, decimals))

    def _round_last_column(self, rule: Dict[str, Any]):
        """Round values in the last column to specified decimals."""
        decimals = rule.get('decimals', 2)
        
        if len(self.df.columns) > 0:
            last_col = self.df.columns[-1]  # Get last column name
            # Convert to numeric and round
            self.df[last_col] = pd.to_numeric(self.df[last_col], errors='coerce').round(decimals)

    def _replace_decimal_separator(self, rule: Dict[str, Any]):
        """Replace decimal separator in specified columns."""
        columns = rule.get('columns', [])
        from_sep = rule.get('from_separator', '.')
        to_sep = rule.get('to_separator', ',')
        
        target_columns = []
        if columns == 'last':
            # Apply to last column only
            if len(self.df.columns) > 0:
                target_columns = [self.df.columns[-1]]
        else:
            # Apply to specified columns
            target_columns = columns
        
        for col in target_columns:
            if col in self.df.columns:
                # Convert column to string, replace separator, keep numeric precision
                self.df[col] = self.df[col].astype(str).str.replace(from_sep, to_sep, regex=False)

    def _generate_output_filename(self, format_type: str, export_config: Dict[str, Any]) -> str:
        """Generate unique output filename based on input file with _mod suffix."""
        if self.input_file_path is None:
            # Fallback to config filename if no input file path
            filename_template = export_config.get('filename', f'output.{format_type}')
            from datetime import datetime
            return filename_template.replace('{date}', datetime.now().strftime('%Y%m%d'))
        
        # Get input file directory and name
        input_dir = self.input_file_path.parent
        input_stem = self.input_file_path.stem  # filename without extension
        
        # Determine output extension
        if format_type in ['txt', 'csv']:
            extension = '.txt' if format_type == 'txt' else '.csv'
        else:
            extension = '.xlsx'
        
        # Generate unique filename with _mod suffix
        counter = 1
        while True:
            suffix = '_mod' if counter == 1 else f'_mod{counter}'
            output_filename = f"{input_stem}{suffix}{extension}"
            output_path = input_dir / output_filename
            
            # Check if file exists or is locked (file in use)
            if not output_path.exists():
                try:
                    # Try to create and immediately delete a test file to check if path is writable
                    with open(output_path, 'w') as test_file:
                        pass
                    output_path.unlink()  # Delete the test file
                    return str(output_path)
                except (OSError, PermissionError):
                    # File might be locked or permission denied, try next number
                    counter += 1
                    continue
            else:
                # File exists, try next number
                counter += 1
                continue

    def export(self, output_path: str = None) -> str:
        """Export transformed data according to config."""
        export_config = self.config.get('export', {})
        format_type = export_config.get('format', 'xlsx')
        
        if output_path is None:
            # Generate automatic filename based on input file
            output_path = self._generate_output_filename(format_type, export_config)
        
        output_path = Path(output_path)
        
        if format_type == 'xlsx':
            # Check if we need to apply Excel formatting
            if export_config.get('apply_excel_formatting', False) and hasattr(self, '_accounting_columns'):
                self._export_with_formatting(output_path)
            else:
                self.df.to_excel(output_path, index=False)
        elif format_type == 'txt' or format_type == 'csv':
            delimiter = export_config.get('delimiter', '\t')
            headers = export_config.get('headers', True)
            self.df.to_csv(output_path, sep=delimiter, index=False, header=headers)
        
        return str(output_path)

    def _export_with_formatting(self, output_path: Path):
        """Export to Excel with accounting number formatting."""
        from openpyxl import Workbook
        from openpyxl.utils.dataframe import dataframe_to_rows
        
        # Create workbook and worksheet
        wb = Workbook()
        ws = wb.active
        
        # Add data to worksheet
        for r in dataframe_to_rows(self.df, index=False, header=True):
            ws.append(r)
        
        # Apply accounting format to specified columns
        if hasattr(self, '_accounting_columns'):
            for col_name, decimals in self._accounting_columns:
                if col_name in self.df.columns:
                    col_index = list(self.df.columns).index(col_name) + 1  # Excel is 1-indexed
                    col_letter = chr(ord('A') + col_index - 1)
                    
                    # Apply accounting number format to entire column (skip header)
                    accounting_format = f'_-* #,##0.{"0" * decimals}_-;-* #,##0.{"0" * decimals}_-;_-* "-"??_-;_-@_-'
                    for row in range(2, len(self.df) + 2):
                        cell = ws[f'{col_letter}{row}']
                        cell.number_format = accounting_format
        
        wb.save(output_path)

    def process(self, input_path: str, output_path: str = None) -> str:
        """Complete processing pipeline."""
        self.load_data(input_path)
        self.apply_transformations()
        return self.export(output_path)