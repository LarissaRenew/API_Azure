import os
from flask import Flask, jsonify, request
from office365.runtime.auth.client_credential import ClientCredential
#from office365.sharepoint.client_context import ClientContext
from dotenv import load_dotenv # Mantenha para desenvolvimento local, será ignorado no Render

# Carrega variáveis de ambiente do arquivo .env (apenas para desenvolvimento local)
load_dotenv()

app = Flask(__name__)

# --- Configurações do SharePoint (LENDO DE VARIÁVEIS DE AMBIENTE) ---
# NO RENDER, ESTAS VARIÁVEIS SERÃO CONFIGURADAS NO PAINEL DO SERVIÇO
SHAREPOINT_SITE_URL = os.getenv("SHAREPOINT_SITE_URL")
TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

# ADICIONE ESTES PRINTS PARA DEPURAR:
print(f"DEBUG: SHAREPOINT_SITE_URL={SHAREPOINT_SITE_URL}")
print(f"DEBUG: TENANT_ID={TENANT_ID}")
print(f"DEBUG: CLIENT_ID={CLIENT_ID}")
print(f"DEBUG: CLIENT_SECRET={'*' * len(CLIENT_SECRET) if CLIENT_SECRET else 'None'}") # Não printar o segredo real!



# Adicionando uma verificação básica para garantir que as variáveis estão configuradas
# Isso pode ajudar a depurar se algo não estiver definido no Render
if not all([SHAREPOINT_SITE_URL, TENANT_ID, CLIENT_ID, CLIENT_SECRET]):
    print("ATENÇÃO: Uma ou mais variáveis de ambiente do SharePoint não estão definidas. Verifique as configurações no Render ou seu arquivo .env.")
    # Você pode até mesmo lançar um erro ou parar a aplicação aqui em um cenário mais robusto
    # raise ValueError("Credenciais do SharePoint ausentes.")

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
        # A biblioteca Office365-REST-Python-Client usará estas variáveis para autenticação
        #credential = ClientCredential(CLIENT_ID, CLIENT_SECRET)
        #ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(credential)
        credential = ClientCredential(CLIENT_ID, CLIENT_SECRET, tenant_id="491bfc39-ac22-494b-a77b-ffb7fc0073a4") # Passe o tenant_id explicitamente
        ctx = ClientContext(SHAREPOINT_SITE_URL).with_credentials(credential)

        target_list = ctx.web.lists.get_by_title(list_name)
        ctx.load(target_list)
        ctx.execute_query()

        print(f"Acessando a biblioteca: {target_list.get_property('Title')}")

        file = target_list.root_folder.files.get_by_url(file_path)
        ctx.load(file)
        ctx.execute_query()

        download_path = f"/tmp/{os.path.basename(file_path)}"
        with open(download_path, "wb") as f:
            file.download(f).execute_query()

        return jsonify({
            "message": f"Arquivo '{file_path}' acessado com sucesso na biblioteca '{list_name}'.",
            "file_size_bytes": file.length,
            "downloaded_to": download_path,
        }), 200

    except Exception as e:
        print(f"Erro ao acessar SharePoint: {e}")
        # Retorna uma mensagem de erro genérica para o cliente, mas imprime o detalhe no log do servidor
        return jsonify({"error": "Ocorreu um erro ao processar sua solicitação."}), 500

if __name__ == '__main__':
    # O Render define a porta através da variável de ambiente PORT
    app.run(debug=True, host='0.0.0.0', port=os.getenv('PORT', 5000))
