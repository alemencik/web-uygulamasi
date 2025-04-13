from flask import Flask, render_template, request

app = Flask(__name__, static_folder='static', template_folder='templates')


VERI_DOSYASI = r'C:\Users\ARYA\Desktop\web-uygulamasi\perakende listesi.xlsm'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    if not os.path.exists(VERI_DOSYASI):
        return "Veri dosyası bulunamadı!"

    try:
        app_excel = xw.App(visible=False)
        wb = app_excel.books.open(VERI_DOSYASI)
        sheet = wb.sheets['Sayfa1']

        data = {
            'B4': request.form.get('input1'),
            'B74': request.form.get('select1'),
            'I44': {"şilte 1": 1, "şilte 2": 2, "şilte 3": 3}.get(request.form.get('select3'), 0),
            'I41': {"süpürgelik 1": 1, "süpürgelik 2": 2, "süpürgelik 3": 3}.get(request.form.get('select2'), 0),
            'J4': request.form.get('input2'),
            'K4': request.form.get('input3'),
            'L4': request.form.get('input4'),
            'N41': 'DOĞRU' if request.form.get('active1') else 'YANLIŞ',
            'O41': 'DOĞRU' if request.form.get('active2') else 'YANLIŞ'
        }

        for hucre, deger in data.items():
            sheet.range(hucre).value = deger

        q40_degeri = sheet.range('Q40').value
        wb.save()
        wb.close()
        app_excel.quit()

        q40_degeri_formatli = f"{round(q40_degeri):,} TL".replace(",", ".")
        return render_template('index.html', sonuc=f"Sonuç: {q40_degeri_formatli}")

    except Exception as e:
        return f"Bir hata oluştu: {str(e)}"

if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=5000)

