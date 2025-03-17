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
            # Read the original Excel file (all sheets)
            excel_file = pd.ExcelFile(file)
            # Read the first sheet to process it with your cleaning code
            df = pd.read_excel(excel_file, sheet_name=excel_file.sheet_names[0])
        except Exception as e:
            return f"Error reading Excel file: {e}", 400

        # --- Begin cleaning process ---
        # Filter rows where either column 'Unnamed: 2' or 'Unnamed: 5' is not null
        f = (~df['Unnamed: 2'].isnull()) | (~df['Unnamed: 5'].isnull())
        df = df[f]

        # Drop columns where all values are NaN, empty, or whitespace
        df_cleaned = df.loc[:, ~df.apply(
            lambda col: col.astype(str).str.strip().replace('nan', '').eq('').all()
        )]

        # Fill forward (ffill) for 'Unnamed: 5'
        df_cleaned.loc[:, 'Unnamed: 5'] = df_cleaned['Unnamed: 5'].ffill()

        # Delete header rows based on 'Unnamed: 2'
        f = (~df_cleaned['Unnamed: 2'].isnull())
        df_cleaned = df_cleaned[f]

        # Drop columns again where all values are NaN, empty, or whitespace
        df_cleaned = df_cleaned.loc[:, ~df_cleaned.apply(
            lambda col: col.astype(str).str.strip().replace('nan', '').eq('').all()
        )]

        # Rename columns (assuming there are exactly seven columns)
        df_cleaned.columns = ['data', 'plano', 'origem', 'histÛrico', 'valor', 'operaÁ„o', 'usu·rio']

        # Convert "valor" column from Brazilian format to numeric floats
        df_cleaned['valor'] = (
            df_cleaned['valor']
            .astype(str)
            .str.replace('.', '', regex=False)   # remove thousand separator '.'
            .str.replace(',', '.', regex=False)   # replace decimal ',' with '.'
            .replace('', '0')                     # handle empty strings
            .astype(float)
        )

        # Convert "data" column to datetime and format as YYYY-MM-DD
        df_cleaned['data'] = pd.to_datetime(
            df_cleaned['data'], format='%d/%m/%Y', errors='coerce'
        ).dt.strftime('%Y-%m-%d')
        # --- End cleaning process ---

        # Create an in-memory output file for the new Excel file.
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Write back all the original sheets
            for sheet in excel_file.sheet_names:
                sheet_df = pd.read_excel(excel_file, sheet_name=sheet)
                sheet_df.to_excel(writer, sheet_name=sheet, index=False)
            # Write the cleaned dataframe in a new sheet called 'Cleaned Data'
            df_cleaned.to_excel(writer, sheet_name='Cleaned Data', index=False)
        output.seek(0)

        return send_file(
            output,
            download_name="processed.xlsx",  # using Flask 2.x+ parameter
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
