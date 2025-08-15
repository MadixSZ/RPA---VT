from flask import Flask, render_template, request, send_file
from datetime import datetime, date
import csv
import os
import zipfile
from io import BytesIO
import calendar

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'recibos_html'

MESES_PT = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
}

def converter_feriados(feriados_str, ano_referencia):
    feriados = []
    for data_str in feriados_str.split(','):
        data_str = data_str.strip()
        if not data_str:
            continue
        try:
            dia, mes = map(int, data_str.split('/'))
            feriados.append(date(ano_referencia, mes, dia))
        except:
            pass
    return feriados

def calcular_data_emissao(data_admissao, mes_referencia, feriados_str):
    ano, mes = map(int, mes_referencia.split('-'))
    feriados = converter_feriados(feriados_str, ano)
    
    primeiro_dia_mes = date(ano, mes, 1)
    
    if data_admissao > primeiro_dia_mes:
        data_emissao = data_admissao
    else:
        dia = 1
        while True:
            data_teste = date(ano, mes, dia)
            condicoes = [
                data_teste.weekday() < 5,
                data_teste not in feriados
            ]
            if all(condicoes):
                data_emissao = data_teste
                break
            dia += 1
    
    return data_emissao

def calcular_dias_uteis(data_admissao, mes_referencia, feriados_str, considerar_sabados):
    ano, mes = map(int, mes_referencia.split('-'))
    data_inicio = max(data_admissao, date(ano, mes, 1))
    ultimo_dia = date(ano, mes, calendar.monthrange(ano, mes)[1])
    
    feriados = converter_feriados(feriados_str, ano)
    sabados_validos = []
    
    for dia in range(1, ultimo_dia.day + 1):
        data = date(ano, mes, dia)
        if data.weekday() == 5:
            sabados_validos.append(data)
    
    sabados_alternados = [s for idx, s in enumerate(sabados_validos) if idx % 2 == 0] if considerar_sabados else []
    
    dias_uteis = 0
    for dia in range(1, ultimo_dia.day + 1):
        data = date(ano, mes, dia)
        condicoes = [
            data >= data_inicio,
            data not in feriados,
            (data.weekday() < 5) or (data in sabados_alternados)
        ]
        if all(condicoes):
            dias_uteis += 1
    
    return dias_uteis

def valor_por_extenso(valor):
    unidades = ["", "um", "dois", "três", "quatro", "cinco", 
               "seis", "sete", "oito", "nove"]
    dez_a_vinte = ["dez", "onze", "doze", "treze", "quatorze", 
                  "quinze", "dezesseis", "dezessete", "dezoito", "dezenove"]
    dezenas = ["", "dez", "vinte", "trinta", "quarenta", "cinquenta", 
              "sessenta", "setenta", "oitenta", "noventa"]
    centenas = ["", "cento", "duzentos", "trezentos", "quatrocentos", 
               "quinhentos", "seiscentos", "setecentos", "oitocentos", "novecentos"]
    
    reais = int(valor)
    centavos = int(round((valor - reais) * 100))
    
    extenso = []
    if reais >= 100:
        extenso.append(centenas[reais // 100])
        reais %= 100
    if reais >= 20:
        extenso.append(dezenas[reais // 10])
        reais %= 10
    if reais >= 10:
        extenso.append(dez_a_vinte[reais - 10])
    elif reais > 0:
        extenso.append(unidades[reais])
    
    texto = " e ".join(extenso)
    texto += " reais" if texto else "zero reais"
    
    if centavos > 0:
        texto += " e "
        if centavos >= 10:
            if centavos < 20:
                texto += dez_a_vinte[centavos - 10]
            else:
                texto += dezenas[centavos // 10]
                centavos %= 10
                if centavos > 0:
                    texto += f" e {unidades[centavos]}"
        else:
            texto += unidades[centavos]
        texto += " centavos"
    
    return texto

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try:
            arquivo_csv = request.files['planilha']
            mes_referencia = request.form['mes_referencia']
            feriados = request.form['feriados']
            considerar_sabados = request.form.get('sabados_alternados') == 'on'
            
            funcionarios = []
            conteudo = arquivo_csv.stream.read().decode('utf-8')
            leitor = csv.DictReader(conteudo.splitlines())
            
            for linha in leitor:
                data_admissao = datetime.strptime(linha['data_admissao'], '%Y-%m-%d').date()
                valor = float(linha['valor_conducao'])
                dias = calcular_dias_uteis(data_admissao, mes_referencia, feriados, considerar_sabados)
                data_emissao = calcular_data_emissao(data_admissao, mes_referencia, feriados)
                ano_ref, mes_ref = map(int, mes_referencia.split('-'))
                
                funcionarios.append({
                    'nome': linha['nome'],
                    'total': dias * 2 * valor,
                    'total_extenso': valor_por_extenso(dias * 2 * valor),
                    'dias': dias,
                    'mes_portugues': MESES_PT[mes_ref],
                    'ano_ref': ano_ref,
                    'data_emissao': data_emissao.strftime('%d de ') + MESES_PT[data_emissao.month] + data_emissao.strftime(' de %Y')
                })
            
            arquivos = []
            for func in funcionarios:
                html = render_template('recibo.html', **func)
                caminho = os.path.join(app.config['UPLOAD_FOLDER'], f"Recibo_{func['nome']}.html")
                with open(caminho, 'w', encoding='utf-8') as f:
                    f.write(html)
                arquivos.append(caminho)
            
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w') as zipf:
                for arquivo in arquivos:
                    zipf.write(arquivo, os.path.basename(arquivo))
            
            zip_buffer.seek(0)
            return send_file(
                zip_buffer,
                mimetype='application/zip',
                as_attachment=True,
                download_name='recibos.zip'
            )
        
        except Exception as e:
            return f"Erro: {str(e)}", 500
    
    return render_template('form.html')

@app.route('/static/<path:filename>')
def static_files(filename):
    return send_from_directory('static', filename)

if __name__ == '__main__':
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    app.run(host='0.0.0.0', port=5000, debug=True)