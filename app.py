import os
from flask import Flask, request, render_template, send_file
import pandas as pd
import json

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process_json', methods=['POST'])
def upload_file():
    file = request.files['file']
    if file:
        # Получаем название входного файла без расширения
        input_filename = os.path.splitext(file.filename)[0]
        
        # Парсинг JSON файла
        data = json.load(file)
        
        # Обработка данных из JSON и создание DataFrame
        lines = data.get('lines', [])
        looks = data.get('looks', [])
        
        # Преобразование данных в формат для таблицы
        rows = []
        look_mapping = {look['lookId']: idx+1 for idx, look in enumerate(looks)}
        
        for line in lines:
            row = {
                "Номер": line.get("order", ""),
                "Бренд": line['products'][0]['product'].get('brand', ''),
                "colorId": line['products'][0]['product'].get('colorId', ''),
                "imageUrl": line['products'][0]['product'].get('imageUrl', ''),
                "itemId": line['products'][0]['product'].get('itemId', ''),
                "productName": line['products'][0]['product'].get('name', ''),
                "Номер лука": look_mapping.get(line.get('lookId', ''), ''),
                "Картинка": f"=IMAGE(D{len(rows) + 2})"  # Добавляем картинку в ячейку
            }
            rows.append(row)
        
        # Создание DataFrame
        df = pd.DataFrame(rows)
        
        # Указываем имя листа
        sheet_name = "Экспорт корзины"
        
        # Путь для сохранения файла Excel
        output_path = f"output/{input_filename}.xlsx"
        
        # Сохраняем в Excel
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            
            # Получаем доступ к рабочей книге и листу
            workbook  = writer.book
            worksheet = writer.sheets[sheet_name]
            
            # Устанавливаем высоту строк с 2 по конец
            for i in range(1, len(df) + 1):
                worksheet.set_row(i, 200)
        
        return send_file(output_path, as_attachment=True, download_name=f"{input_filename}.xlsx")

if __name__ == '__main__':
    if not os.path.exists('output'):
        os.makedirs('output')
    app.run(debug=True)
