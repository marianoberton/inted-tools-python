from flask import Flask, request, send_file, render_template, redirect, url_for
from processor import process_excel_file
import os
import tempfile
import zipfile
import uuid


app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return render_template('index.html', error='No se seleccionó ningún archivo')
        
        file = request.files['file']
        if file.filename == '':
            return render_template('index.html', error='No se seleccionó ningún archivo')

        # Verificar extensión del archivo
        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            return render_template('index.html', error='Formato de archivo no válido. Solo se aceptan archivos Excel (.xlsx, .xls)')

        try:
            # Crear directorio temporal único
            process_id = uuid.uuid4().hex
            temp_dir = os.path.join(app.config['UPLOAD_FOLDER'], process_id)
            os.makedirs(temp_dir)

            # Guardar archivo temporal
            input_path = os.path.join(temp_dir, file.filename)
            file.save(input_path)
            
            # Procesar archivo
            result_files = process_excel_file(input_path, temp_dir)
            
            # Crear archivo ZIP con resultados
            zip_filename = f"resultados_{process_id}.zip"
            zip_path = os.path.join(temp_dir, zip_filename)
            
            with zipfile.ZipFile(zip_path, 'w') as zipf:
                for f in result_files:
                    zipf.write(f, os.path.basename(f))
            
            # Redirigir a la descarga
            return redirect(url_for('download', filename=zip_filename))
        
        except ValueError as ve:
            return render_template('index.html', error=f"Error en el formato del archivo: {str(ve)}")
        except Exception as e:
            return render_template('index.html', error=f"Error al procesar el archivo: {str(e)}")
    
    return render_template('index.html')

@app.route('/download/<filename>')
def download(filename):
    try:
        # Extraer el process_id del nombre del archivo
        process_id = filename.split('_')[1].split('.')[0]
        temp_dir = os.path.join(app.config['UPLOAD_FOLDER'], process_id)
        file_path = os.path.join(temp_dir, filename)
        
        if not os.path.exists(file_path):
            raise FileNotFoundError("El archivo solicitado no existe")
            
        return send_file(file_path, as_attachment=True)
    
    except Exception as e:
        return render_template('index.html', error=f"Error al descargar: {str(e)}")

if __name__ == '__main__':
    app.run(port=5000, debug=True)