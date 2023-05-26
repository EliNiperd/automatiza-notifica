from flask import Flask, render_template, request
import win32com.client
import time

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/enviar-correos', methods=['POST'])
def enviar_correos():
    destinatarios = request.form['destinatarios']
    asunto = 'Asunto fijo del correo'
    
    for destinatario in destinatarios.split(';'):
        cuerpo = f"Estimado(a) destinatario,\n\n"
        cuerpo += "Este es un ejemplo de correo automatizado.\n"
        cuerpo += "Cuerpo personalizable según el destinatario.\n\n"
        cuerpo += "Saludos,\n"
        cuerpo += "Tu nombre"
        
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)
        mail.To = destinatario.strip()
        mail.Subject = asunto
        mail.Body = cuerpo
        mail.Send()
        
        time.sleep(1800)  # Espera 30 minutos
        
        inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6)
        response_received = False
        for email in inbox.Items:
            if email.Subject == asunto:
                response_received = True
                break
        
        if not response_received:
            # Realiza las acciones necesarias si no se recibió una respuesta
            return 'No se recibió una respuesta al correo enviado.'
    
    return 'Correos enviados exitosamente.'

if __name__ == '__main__':
    app.run(debug=True)
