import os
from flask import Flask, jsonify, request
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)

SHAREPOINT_SITE_URL = os.getenv("SHAREPOINT_SITE_URL") # <-- ESTA SERA A URL BASE DO SITE
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

print("--- VARIAVEIS DE AMBIENTE Lidas ---")
print(f"SHAREPOINT_SITE_URL: {SHAREPOINT_SITE_URL}")
print(f"TENANT_ID: {TENANT_ID}")
print(f"CLIENT_ID: {CLIENT_ID}")
print(f"CLIENT_SECRET: {'*' * len(CLIENT_SECRET) if CLIENT_SECRET else 'NÃO DEFINIDO/VAZIO'}")
print("---------------------------------")

@app.route('/')
def home():
    return "API SharePoint está funcionando!"

@app.route('/read-excel', methods=['GET'])
def read_excel_from_sharepoint():
    file_path = request.args.get('file_path')
    list_name = request.args.get('list_name')

    if not file_path or not list_name:
        return jsonify({"error": "Parâmetros 'file_path' e 'list_name' são obrigatórios."}), 400

    try:
        # Volte para esta linha, o Tenant ID é inferido ou não é necessário aqui
        credential = ClientCredential(CLIENT_ID, CLIENT_SECRET)
        ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(credential)

        # ... o restante do seu código ...

    except Exception as e:
        print(f"Erro ao acessar SharePoint: {e}")
        return jsonify({"error": "Ocorreu um erro ao processar sua solicitação."}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=os.getenv('PORT', 5000))
