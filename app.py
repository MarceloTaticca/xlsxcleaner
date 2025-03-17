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
            # Read the Excel file into a DataFrame (reads the first sheet by default)
            df = pd.read_excel(file)
        except Exception as e:
            return f"Error reading Excel file: {e}", 400

        # Create an in-memory output file for the new Excel file.
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        output.seek(0)

        return send_file(
            output,
            download_name="processed.xlsx",  # use 'download_name' for Flask 2.x+
            as_attachment=True,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
