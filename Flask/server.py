from flask import Flask, send_file, request
import datetime

app = Flask(__name__)

@app.route('/pixel.png')
def pixel():
    # Registrar informações de acesso
    ip = request.remote_addr
    user_agent = request.headers.get('User-Agent')
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # Log das informações
    log_entry = f"Email opened - IP: {ip}, User-Agent: {user_agent}, Time: {timestamp}\n"
    with open('email_log.txt', 'a') as f:
        f.write(log_entry)

    # Servir a imagem do pixel (PIXEL.png)
    return send_file('PIXEL.png', mimetype='image/png')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)