from flask import Flask, render_template
import pandas as pd

app = Flask(__name__)

@app.route('/greet/<name>', methods=['GET'])
def greet(name):
    person_info = get_person_info(name)
    
    if person_info is None:
        greeting = "Person not found"
        person_info = {}
    else:
        greeting = f"Hello, {name}!"

    return render_template('index.html', greeting=greeting, person_info=person_info)

def get_person_info(name):
    data = pd.read_excel('data.xlsx')
    person_data = data[data['Name'] == name]
    
    if not person_data.empty:
        person_info = person_data.iloc[0].to_dict()
        return person_info
    return None

if __name__ == '__main__':
    app.run(debug=True)

