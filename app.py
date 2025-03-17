from flask import Flask, request, send_file, render_template
import pandas as pd
import io

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files.get('file')
        if not file:
            return "No file uploaded", 400
        
        try:
            # Read the entire workbook (all sheets) into a dictionary of DataFrames.
            xls = pd.read_excel(file, sheet_name=None)
        except Exception as e:
            return f"Error reading Excel file: {e}", 400

        # Get the first sheet's name and data.
        first_sheet_name = list(xls.keys())[0]
        df = xls[first_sheet_name]

        # === Begin Data Processing (the cleaning steps) ===
        # Filter rows where either 'Unnamed: 2' or 'Unnamed: 5' is not null.
        f_mask = (~df['Unnamed: 2'].isnull()) | (~df['Unnamed: 5'].isnull())
        df = df[f_mask].copy()
        
        # Drop columns where all values are NaN, empty, or whitespace.
        # This version strips and replaces 'nan' with an empty string, then keeps only columns with at least one non-empty value.
        df_cleaned = df.loc[:, df.astype(str).apply(
            lambda col: col.str.strip().replace('nan', '').ne('').any()
        )].copy()
        
        # Fill forward (ffill) to repeat non-empty values downward.
        df_cleaned['Unnamed: 5'] = df_cleaned['Unnamed: 5'].ffill()
        
        # Delete header rows.
        f_mask = (~df_cleaned['Unnamed: 2'].isnull())
        df_cleaned = df_cleaned[f_mask].copy()
        
        # Drop columns again where all values are NaN, empty, or whitespace.
        df_cleaned = df_cleaned.loc[:, df_cleaned.astype(str).apply(
            lambda col: col.str.strip().replace('nan', '').ne('').any()
        )]
        
        # Rename seven columns.
        df_cleaned.columns = ['data', 'plano', 'origem', 'histÛrico', 'valor', 'operaÁ„o', 'usu·rio']
        
        # Convert "valor" column from Brazilian format to numeric floats.
        df_cleaned['valor'] = (
            df_cleaned['valor']
            .astype(str)
            .str.replace('.', '', regex=False)   # remove thousand separator '.'
            .str.replace(',', '.', regex=False)    # replace decimal ',' with '.'
            .replace('', '0')                      # handle empty strings
            .astype(float)
        )
        
        # Convert "data" from text to datetime and format to YYYY-MM-DD.
        df_cleaned['data'] = pd.to_datetime(
            df_cleaned['data'], format='%d/%m/%Y', errors='coerce'
        ).dt.strftime('%Y-%m-%d')
        # === End Data Processing ===

        # Write the original sheets plus the cleaned data into a new Excel file in memory.
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Write each original sheet back into the output workbook.
            for sheet_name, dataframe in xls.items():
                dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
            # Add a new sheet for the cleaned data.
            df_cleaned.to_excel(writer, sheet_name='Cleaned Data', index=False)
            writer.save()
        output.seek(0)

        return send_file(
            output,
            attachment_filename="processed.xlsx",
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
