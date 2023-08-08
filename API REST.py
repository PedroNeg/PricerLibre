from flask import Flask, render_template_string
import subprocess

app = Flask(__name__)

@app.route('/')
def home():
    return render_template_string("""
        <h1>Bienvenido a nuestra página!</h1>
        <p><a href="{{ url_for('run_script') }}">Haz clic aquí</a> para ejecutar el script.</p>
    """)

@app.route('/run_script')
def run_script():
    subprocess.call(['python', 'C:/Users/negri/Desktop/Python/Original Probada.py'])
    return 'Script ejecutado!'

if __name__ == '__main__':
    app.run(debug=True)
