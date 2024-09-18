import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl.styles import numbers


def export_to_excel(df, file_name):
    try:
        # Create a BytesIO object to store the Excel file
        output = BytesIO()

        # Use ExcelWriter to write the DataFrame to Excel
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')

            # Get the workbook and the active sheet
            workbook = writer.book
            worksheet = workbook['Sheet1']

            # Iterate through all cells and format appropriately
            for row in worksheet.iter_rows():
                for cell in row:
                    if isinstance(cell.value, (int, np.int64)):
                        cell.number_format = '0'
                    elif isinstance(cell.value, float):
                        cell.number_format = '0.00'
                    else:
                        cell.number_format = '@'

        # Seek to the beginning of the BytesIO object
        output.seek(0)

        # Streamlit download button
        st.download_button(
            label="Download filtered data as Excel",
            data=output,
            file_name=f"{file_name}_filtered.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Error processing Excel export: {str(e)}")


def apply_filters(df, filters):
    for filter_data in filters:
        col = filter_data['column']
        if filter_data['filter_type'] == 'Standard':
            threshold = filter_data['threshold']
            is_percentage = filter_data['is_percentage']
            if is_percentage and pd.api.types.is_numeric_dtype(df[col]):
                threshold = threshold / 100  # Convert to decimal for percentages

            if filter_data['operator'] == 'Remove rows with empty or 0 values':
                df = df[(df[col] != 0) & (df[col].notna())]
            elif filter_data['operator'] == 'Show only rows with empty or 0 values':
                df = df[(df[col] == 0) | (df[col].isna())]
            elif filter_data['operator'] == 'Equals':
                df = df[df[col] == threshold]
            elif filter_data['operator'] == 'Not Equals':
                df = df[df[col] != threshold]
            elif filter_data['operator'] == 'Contains':
                df = df[df[col].astype(str).str.contains(str(threshold), na=False)]
            elif filter_data['operator'] in ['<', '>', '<=', '>=']:
                if pd.api.types.is_numeric_dtype(df[col]):
                    threshold = pd.to_numeric(threshold, errors='coerce')
                    if filter_data['operator'] == '<':
                        df = df[df[col] < threshold]
                    elif filter_data['operator'] == '>':
                        df = df[df[col] > threshold]
                    elif filter_data['operator'] == '<=':
                        df = df[df[col] <= threshold]
                    elif filter_data['operator'] == '>=':
                        df = df[df[col] >= threshold]
        elif filter_data['filter_type'] == 'Calculation':
            threshold = filter_data['threshold']
            operator = filter_data['operator']
            is_percentage = filter_data['is_percentage']
            if is_percentage and pd.api.types.is_numeric_dtype(df[col]):
                threshold = threshold / 100
            if pd.api.types.is_numeric_dtype(df[col]):
                threshold = pd.to_numeric(threshold, errors='coerce')
                if operator == '>':
                    df = df[df[col] > threshold]
                elif operator == '<':
                    df = df[df[col] < threshold]
    return df


def apply_comparison(df, col_1, col_2, operator, calculation_option, percentage_value=None, apply_condition=True):
    if apply_condition:
        if operator == '<':
            condition = df[col_1] < df[col_2]
        elif operator == '>':
            condition = df[col_1] > df[col_2]

        if calculation_option == 'Adjust by percentage':
            df.loc[condition, 'Forecasted Value'] = df[col_1] * (1 + (percentage_value / 100))
        elif calculation_option == 'Make equal to Column 2':
            df.loc[condition, 'Forecasted Value'] = df[col_2]
    else:
        if calculation_option == 'Adjust by percentage':
            df['Forecasted Value'] = df[col_1] * (1 + (percentage_value / 100))
        elif calculation_option == 'Make equal to Column 2':
            df['Forecasted Value'] = df[col_2]
    return df


def apply_column_calculation(df, column, percentage, add_subtract):
    if add_subtract == 'Add':
        df['Calculated Value'] = df[column] * (1 + (percentage / 100))
    elif add_subtract == 'Subtract':
        df['Calculated Value'] = df[column] * (1 - (percentage / 100))
    return df


def delete_columns(df, columns_to_delete):
    df = df.drop(columns=columns_to_delete)
    return df


def main():
    st.set_page_config(page_title="Excel Viewer and Filter", layout="wide")
    st.title("Excel Tab Filter and Viewer")

    # Upload the Excel file
    uploaded_file = st.file_uploader("Upload your Excel file", type="xlsx")
    if uploaded_file is not None:
        try:
            # Load the Excel file
            xls = pd.ExcelFile(uploaded_file)

            # Select a tab
            selected_tab = st.selectbox("Select a tab to work with", xls.sheet_names)

            # Load DataFrame and preserve data types
            df = pd.read_excel(uploaded_file, sheet_name=selected_tab, dtype=object)

            # Convert numeric columns to appropriate types while preserving long integers
            for col in df.columns:
                try:
                    df[col] = pd.to_numeric(df[col], errors='raise', downcast='integer')
                except (ValueError, TypeError):
                    pass  # Keep as object if conversion fails

            # Display original data
            st.write("### Original Data:")
            st.dataframe(df)

            # Filter management
            st.write("### Filter Options:")
            filters = []
            add_filter = st.checkbox("Add a filter")

            while add_filter:
                with st.expander(f"Filter {len(filters) + 1}"):
                    col = st.selectbox("Select a column to filter on", df.columns, key=f"col_{len(filters)}")
                    filter_type = st.radio("Filter type", ['Standard', 'Calculation'],
                                           key=f"filter_type_{len(filters)}")

                    if filter_type == 'Standard':
                        operator = st.selectbox(
                            "Choose operator",
                            ['Show all rows', 'Remove rows with empty or 0 values',
                             'Show only rows with empty or 0 values', '<', '>', '<=', '>=', 'Equals', 'Not Equals',
                             'Contains'],
                            key=f"operator_{len(filters)}"
                        )

                        is_percentage = False
                        if operator in ['Equals', 'Not Equals', 'Contains']:
                            threshold = st.text_input(f"Enter a text value for {col}", key=f"threshold_{len(filters)}")
                        elif operator in ['<', '>', '<=', '>=']:
                            threshold = st.text_input(f"Enter a numeric threshold value for {col}",
                                                      key=f"threshold_{len(filters)}")
                            is_percentage = st.checkbox("Is this a percentage?", key=f"percentage_{len(filters)}")
                        else:
                            threshold = None

                        filters.append({
                            'column': col,
                            'filter_type': 'Standard',
                            'operator': operator,
                            'threshold': threshold,
                            'is_percentage': is_percentage
                        })

                    elif filter_type == 'Calculation':
                        operator = st.selectbox("Choose comparison operator", ['<', '>'], key=f"comp_op_{len(filters)}")
                        threshold = st.text_input(f"Enter a threshold value for {col}", key=f"threshold_{len(filters)}")
                        is_percentage = st.checkbox("Is this a percentage?", key=f"percentage_{len(filters)}")
                        filters.append({
                            'column': col,
                            'filter_type': 'Calculation',
                            'operator': operator,
                            'threshold': threshold,
                            'is_percentage': is_percentage
                        })

                add_filter = st.checkbox("Add another filter", key=f"add_filter_{len(filters)}")

            # Apply filters
            filtered_df = apply_filters(df, filters)

            # Column calculation (add or subtract percentage)
            st.write("### Column Calculation:")
            apply_calc = st.checkbox("Apply percentage calculation to a single column")
            if apply_calc:
                calc_col = st.selectbox("Select a column for calculation", df.columns)
                percentage = st.number_input("Enter percentage for adjustment", value=0.0)
                add_subtract = st.radio("Would you like to add or subtract the percentage?", ["Add", "Subtract"])

                # Apply column calculation
                filtered_df = apply_column_calculation(filtered_df, calc_col, percentage, add_subtract)

            # Delete Columns
            st.write("### Delete Columns:")
            delete_cols = st.multiselect("Select columns to delete", df.columns)
            if delete_cols:
                filtered_df = delete_columns(filtered_df, delete_cols)
                st.write("Deleted columns and updated data:")

            # Comparison section
            st.write("### Column Comparison and Calculation:")
            compare_columns = st.checkbox("Compare two columns and calculate")
            if compare_columns:
                col1, col2 = st.columns(2)
                with col1:
                    col_1 = st.selectbox("Select the first column", df.columns)
                    comparison_operator = st.selectbox("Choose comparison operator", ["<", ">"])
                with col2:
                    col_2 = st.selectbox("Select the second column", df.columns)
                    calculation_option = st.radio("Calculation Type",
                                                  ['Adjust by percentage', 'Make equal to Column 2'])

                percentage_value = None
                if calculation_option == 'Adjust by percentage':
                    percentage_value = st.number_input("Enter adjustment percentage", value=0.0)

                apply_condition = st.radio("Apply calculation:", ["Only if condition is met", "Apply to all rows"])

                # Apply column comparison and calculate the forecasted value
                filtered_df = apply_comparison(filtered_df, col_1, col_2, comparison_operator, calculation_option,
                                               percentage_value, apply_condition == "Only if condition is met")

            # Display filtered data
            st.write("### Filtered Data:")
            st.dataframe(filtered_df)

            # Excel Export
            export_to_excel(filtered_df, selected_tab)

        except Exception as e:
            st.error(f"Error reading the Excel file: {str(e)}")


if __name__ == "__main__":
    main()
