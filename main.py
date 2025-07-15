import os
from flask import Flask, jsonify, request
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.client_context import ClientContext
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)

SHAREPOINT_SITE_URL = os.getenv("SHAREPOINT_SITE_URL")
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
    print("--- INICIANDO FUNCAO read_excel_from_sharepoint ---")
    file_path = request.args.get('file_path')
    list_name = request.args.get('list_name')

    if not file_path or not list_name:
        print("DEBUG: Retornando 400 - Parametros ausentes.")
        return jsonify({"error": "Parâmetros 'file_path' e 'list_name' são obrigatórios."}), 400

    try:
        print("DEBUG: Tentando autenticar no SharePoint.")
        credential = ClientCredential(CLIENT_ID, CLIENT_SECRET)
        ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(credential)

        print("DEBUG: Tentando obter a lista/biblioteca de documentos.")
        target_list = ctx.web.lists.get_by_title(list_name)
        ctx.load(target_list)
        ctx.execute_query()
        print(f"DEBUG: Biblioteca acessada: {target_list.get_property('Title')}")

        print("DEBUG: Tentando obter o arquivo da planilha.")
        file = target_list.root_folder.files.get_by_url(file_path)
        ctx.load(file)
        ctx.execute_query()
        print(f"DEBUG: Arquivo acessado: {file.name}")

        download_path = f"/tmp/{os.path.basename(file_path)}"
        print(f"DEBUG: Tentando baixar o arquivo para: {download_path}")
        with open(download_path, "wb") as f:
            file.download(f).execute_query()
        print("DEBUG: Arquivo baixado com sucesso.")

        print("DEBUG: Retornando 200 - Sucesso.")
        return jsonify({
            "message": f"Arquivo '{file_path}' acessado com sucesso na biblioteca '{list_name}'.",
            "file_size_bytes": file.length,
            "downloaded_to": download_path,
        }), 200

    except Exception as e:
        print(f"--- ERRO INESPERADO AO ACESSAR SHAREPOINT: {e} ---") # Mensagem de erro mais visível
        # Se a exceção for capturada, esta linha será executada.
        return jsonify({"error": "Ocorreu um erro ao processar sua solicitação."}), 500

    # Adicione este print para pegar casos onde a função termina sem retornar.
    print("--- FIM DA FUNCAO SEM RETORNO VALIDO ---")
    # Pode até retornar um erro aqui para ter certeza que algo é retornado se chegar aqui.
    # return jsonify({"error": "Falha interna na função, sem retorno explícito."}), 500

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=os.getenv('PORT', 5000))
