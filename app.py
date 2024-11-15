from flask import Flask, request, render_template
import pandas as pd
import matplotlib.pyplot as plt
import os

app = Flask(__name__)

# Ensure the directory for storing charts exists
if not os.path.exists("static/charts"):
    os.makedirs("static/charts")

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        file = request.files["file"]

        try:
            # Read the Excel file
            df = pd.read_excel(file, header=0, index_col=0)

            # Clean up whitespace only for string data
            df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
            df.index = [idx.strip() if isinstance(idx, str) else idx for idx in df.index]

            # Extract and calculate ratios if data exists
            ratios = {}
            missing_data = []

            def calculate_ratio(name, numerator, denominator):
                if numerator in df.index and denominator in df.index:
                    # Perform element-wise division and round to 2 decimal places
                    ratio_series = (df.loc[numerator] / df.loc[denominator]) * 100
                    return ratio_series.round(3)  # Round to 2 decimal places
                else:
                    missing_data.append(name)
                    return pd.Series()  # Return empty Series if data is missing

            # Define and calculate each ratio
            ratios["Return on Equity"] = calculate_ratio("Return on Equity", "Net Income", "Total Shareholder Equity")
            ratios["Return on Assets"] = calculate_ratio("Return on Assets", "Net Income", "Total Assets")
            ratios["Return on Net Assets"] = calculate_ratio("Return on Net Assets", "Net Income", "Property Plant & Equipment")
            ratios["Return on Invested Capital"] = calculate_ratio("Return on Invested Capital", "NOPAT", "Invested Capital (aka Capital Employed)")
            ratios["Gross Margin"] = calculate_ratio("Gross Margin", "Gross Profit", "Revenue")
            ratios["SG&A % of Revenue"] = calculate_ratio("SG&A % of Revenue", "SG&A", "Revenue")
            ratios["Other Operating Expenses % of Revenue"] = calculate_ratio("Other Operating Expenses % of Revenue", "Other", "Revenue")
            ratios["EBITDA Margin"] = calculate_ratio("EBITDA Margin", "EBITDA", "Revenue")

            # Filter out empty Series from ratios
            ratios = {key: value for key, value in ratios.items() if not value.empty}

            # Generate charts for each ratio
            chart_paths = {}
            for ratio_name, ratio_data in ratios.items():
                plt.figure()
                ratio_data.plot(title=ratio_name)
                plt.xlabel("Year")
                plt.ylabel(f"{ratio_name} (%)")

                # Save the chart as a PNG image
                chart_path = f"static/charts/{ratio_name.replace(' ', '_')}.png"
                plt.savefig(chart_path)
                plt.close()

                # Store the path for rendering in HTML
                chart_paths[ratio_name] = chart_path

            return render_template("results.html", results=ratios, chart_paths=chart_paths, missing_data=missing_data)

        except Exception as e:
            return f"An error occurred while processing the file: {str(e)}"

    return '''
    <!doctype html>
    <title>Upload Excel File</title>
    <h1>Upload an Excel file for Financial Analysis</h1>
    <form method=post enctype=multipart/form-data>
      <input type=file name=file>
      <input type=submit value=Upload>
    </form>
    '''

@app.route("/results")
def results():
    return render_template("results.html")

if __name__ == "__main__":
    app.run(debug=True)
