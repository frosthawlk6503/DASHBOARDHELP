from flask import Flask, render_template
import pandas as pd
import datetime

app = Flask(__name__)
def convert_time_to_hours(time_obj):
    # Check if the entry is a valid datetime.time object
    if isinstance(time_obj, pd.Timestamp):
        # If it's already a timestamp, we can handle it
        return time_obj.hour + time_obj.minute / 60 + time_obj.second / 3600
    elif isinstance(time_obj, pd.Timedelta):
        # If it's already a timedelta, just convert it to hours
        return time_obj.total_seconds() / 3600
    elif isinstance(time_obj, datetime.time):
        # For pure time object, manually calculate the hours
        return time_obj.hour + time_obj.minute / 60 + time_obj.second / 3600
    else:
        # Handle other cases if necessary (e.g. strings, NaNs)
        return 0

# Function to load and process Excel data
def load_data():
    file_path = 'JobCardDataAlteredAltered.xlsx'
    
    # Load data from two sheets (representing two models of machines)
    sheet1 = pd.read_excel(file_path, sheet_name='Sheet1')
    sheet2 = pd.read_excel(file_path, sheet_name='Sheet2')

    # Convert 'Start Date' and 'End Date' to datetime for both sheets
    sheet1['START DATE AND TIME'] = pd.to_datetime(sheet1['START DATE AND TIME'], dayfirst=True)
    sheet1['END DATE AND TIME'] = pd.to_datetime(sheet1['END DATE AND TIME '], dayfirst=True)
    sheet2['START DATE AND TIME '] = pd.to_datetime(sheet2['START DATE AND TIME '], dayfirst=True)
    sheet2['END DATE AND TIME'] = pd.to_datetime(sheet2['END DATE AND TIME'], dayfirst=True)


    # Convert 'Man Hours' to hours
    # sheet1['Man Hours'] = pd.to_timedelta(sheet1['Man Hours']).dt.total_seconds() / 3600
    # sheet2['Man Hours'] = pd.to_timedelta(sheet2['Man Hours']).dt.total_seconds() / 3600
    sheet1['Man Hours'] = sheet1['Man Hours'].apply(convert_time_to_hours)
    sheet2['Man Hours'] = sheet2['Man Hours'].apply(convert_time_to_hours)

    return sheet1, sheet2

# Route for the main webpage
@app.route('/')
def index():
    # Load data from Excel
    sheet1, sheet2 = load_data()

    # Prepare data for Gantt and Pie charts (both sheets)
    gantt_data_1 = sheet1[['Sub Assembly', 'START DATE AND TIME', 'END DATE AND TIME ']].to_dict('records')
    gantt_data_2 = sheet2[['Sub Assembly', 'START DATE AND TIME ', 'END DATE AND TIME']].to_dict('records')

    # Prepare Pie chart data: number of attempts per sub-assembly
    attempts_pie_1 = sheet1.groupby('Sub Assembly')['No. of Attempts'].sum().to_dict()
    attempts_pie_2 = sheet2.groupby('Sub Assembly')['No. of Attempts'].sum().to_dict()

    # Render the HTML page with the data passed to it
    return render_template('dashboard.html', 
                           gantt_data_1=gantt_data_1, gantt_data_2=gantt_data_2,
                           attempts_pie_1=attempts_pie_1, attempts_pie_2=attempts_pie_2)

if __name__ == '__main__':
    app.run(debug=True)
