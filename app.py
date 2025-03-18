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
            # Read the uploaded Excel file (first sheet by default)
            df = pd.read_excel(file)
        except Exception as e:
            return f"Error reading Excel file: {e}", 400
        
        try:
            # --- Begin cleaning logic from Code B ---
            # Filter rows: keep rows where either 'Unnamed: 2' or 'Unnamed: 5' is not null
            filter_condition = (~df['Unnamed: 2'].isnull()) | (~df['Unnamed: 5'].isnull())
            df = df[filter_condition]

            # Drop columns where all values are NaN, empty, or whitespace
            df_cleaned = df.loc[:, ~df.apply(
                lambda col: col.astype(str).str.strip().replace('nan', '').eq('').all()
            )]

            # Forward-fill 'Unnamed: 5'
            df_cleaned.loc[:, 'Unnamed: 5'] = df_cleaned['Unnamed: 5'].ffill()

            # Remove extra header rows by keeping only rows where 'Unnamed: 2' is not null
            df_cleaned = df_cleaned[~df_cleaned['Unnamed: 2'].isnull()]

            # Drop columns where all values are NaN, empty, or whitespace (again)
            df_cleaned = df_cleaned.loc[:, ~df_cleaned.apply(
                lambda col: col.astype(str).str.strip().replace('nan', '').eq('').all()
            )]

            # Rename columns to the desired names
            df_cleaned.columns = ['data', 'plano', 'origem', 'histórico', 'valor', 'operação', 'usuário']

            # Convert "valor" column from Brazilian format to numeric floats
            df_cleaned['valor'] = (
                df_cleaned['valor']
                .astype(str)
                .str.replace('.', '', regex=False)   # remove thousand separator '.'
                .str.replace(',', '.', regex=False)    # replace decimal ',' with '.'
                .replace('', '0')                      # handle empty strings
                .astype(float)
            )

            # Convert "data" column to datetime (format: YYYY-MM-DD)
            df_cleaned['data'] = pd.to_datetime(
                df_cleaned['data'], format='%d/%m/%Y', errors='coerce'
            ).dt.strftime('%Y-%m-%d')
            # --- End cleaning logic ---
        except Exception as e:
            return f"Error processing Excel file: {e}", 400

        output = io.BytesIO()
        try:
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Write the cleaned data to a sheet named "Cleaned Data"
                df_cleaned.to_excel(writer, index=False, sheet_name='Cleaned Data')
            output.seek(0)
        except Exception as e:
            return f"Error writing Excel file: {e}", 500

        return send_file(
            output,
            download_name="cleaned_processed.xlsx",  # Flask 2.x+ parameter
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
