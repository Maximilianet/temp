import pandas as pd
import json
from flask import Flask, request, send_file, render_template, jsonify, make_response
from io import BytesIO

app = Flask(__name__)

# Пример функции для конвертации JSON в Excel
def json_to_excel(json_data):
    rows = []
    look_ids = {}
    look_number = 1
    
    for line in json_data['data']['lines']:
        look_id = line['lookId']
        
        if look_id not in look_ids:
            look_ids[look_id] = look_number
            look_number += 1
        
        for product in line['products']:
            product_data = product['product']
            row = {
                'Номер': line['order'],
                'Бренд': product_data['brand'],
                'colorId': product_data['colorId'],
                'imageUrl': product_data['imageUrl'],
                'itemId': product_data['itemId'],
                'productName': product_data['name'],
                'Номер лука': look_ids[look_id]
            }
            rows.append(row)

    df = pd.DataFrame(rows)
    
    # Убираем колонку looks__lookId и добавляем колонку Картинка с формулой
    df['Картинка'] = df.apply(lambda x: f'=IMAGE(D{df.index.get_loc(x.name) + 2})', axis=1)

    # Сохраняем в Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Basket')
    output.seek(0)
    
    return output

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process_json', methods=['POST'])
def process_json():
    if 'file' not in request.files:
        return jsonify({"error": "Файл не найден"}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({"error": "Файл не выбран"}), 400

    json_data = json.load(file)

    excel_file = json_to_excel(json_data)

    response = make_response(excel_file.read())
    response.headers['Content-Disposition'] = 'attachment; filename=basket_output.xlsx'
    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return response

if __name__ == '__main__':
    app.run(debug=True)
