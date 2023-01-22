from flask import Flask,render_template,request;
from flask_wtf import FlaskForm;
from wtforms import FileField,SubmitField;
import pandas as pd
import logging as logger
from werkzeug.utils import secure_filename
import os



app = Flask(__name__)
app.config['SECRET_KEY'] = 'supersecretkey'
app.config['UPLOAD_FOLDER'] = 'static/files'

class UploadFileForm(FlaskForm):
    file = FileField("File")
    submit = SubmitField("Upload File")


@app.route('/',methods = ['GET','POST'])



def index():
        form = UploadFileForm()
      



# @app.route('/upload', methods=['GET', 'POST'])


        
        if form.validate_on_submit():
            

            file = form.file.data
            df = pd.read_excel(file)
            df['Gender'] = df['Gender'].replace({'Female': 0, 'Male': 1})
            df['Gender'] = pd.to_numeric(df['Gender'])
            df['duplicate_location'] = df.duplicated(subset=['Location'], keep=False)
        
            # Group data by 'Surveyor name' and calculate the number of duplicate locations
            result = df.groupby('Surveyor name')['duplicate_location'].sum().reset_index()
            result.columns = ['Surveyor name', 'DUPLICATE LOCATION']
        
            # Merge the result with the original DataFrame
            result = pd.merge(result, df, on='Surveyor name')
        
            # Merge the result with the original DataFrame
        
        
            # Group data by 'Surveyor name' and calculate various statistics
            result =result.groupby('Surveyor name').agg({
                'Audio Duration (in secs)': ['count', 'sum'],
                'duplicate_location':['sum'],
            'Gender': ['count', 'sum'],
                'Timestamp': ['min', 'max'],
            }).reset_index()
        
            result['NO OF SAMPLES'] = result['Audio Duration (in secs)']['count'] - result['duplicate_location']['sum']#ivot_table = result.pivot_table(values='Gender', index='Surveyor name', aggfunc='count')
            # Rename columns
            #result.columns = ['EMPLOYEE NAME', 'NO OF SAMPLES', 'DURATION','DUPLICATE LOCATION', 'MALE', 'FEMALE', 'STARING TIME', 'ENDING TIME']
            result.columns = ['EMPLOYEE NAME', 'NO OF SAMPLES', 'DURATION', 'DUPLICATE LOCATION', 'FEMALE', 'MALE', 'STARING TIME', 'ENDING TIME', 'NEW COLUMN']
        
            # Calculate percentage of Male and Female
        
            result['MALE'] = (result['MALE'] / result['NO OF SAMPLES']) * 100
            result['FEMALE'] = (result['FEMALE'] / result['NO OF SAMPLES']) * 100
        
            result['FEMALE']= result['FEMALE']-result['MALE']
        
        
        
            # Convert 'STARING TIME' and 'ENDING TIME' columns to datetime
            result['STARING TIME'] = pd.to_datetime(result['STARING TIME'])
            result['ENDING TIME'] = pd.to_datetime(result['ENDING TIME'])
        
            # Add new column 'DURATION'
            result['DURATION'] = result['ENDING TIME'] - result['STARING TIME']
            result['DURATION'] = result['DURATION'].apply(lambda x: divmod(x.seconds, 3600))
            # result['DURATION'] = result['DURATION'].apply(lambda x: f"{x[0]:0>2}:{x[1]//60:0>2}:{x[1]%60:0>2}")
            #result['DURATION'] = result['DURATION'].dt.total_seconds().div(60).round(2)
            #result['DURATION'] = result['DURATION'].div(60).round(2)
            # Print the result
            print(result)
            
            result.to_excel("swapnil_New.xlsx", index=False)
            return "Report Generated Successfully"

        return render_template('index.html',form = form)
            
        
      
    
if __name__ == '__main__':
    app.run()