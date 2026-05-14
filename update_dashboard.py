#!/usr/bin/env python3
"""
Recarrega a dashboard com os arquivos mais recentes já baixados.
Chama /api/reload no servidor Flask sem reiniciá-lo.
"""

import requests

RELOAD_URL = "http://localhost:5000/api/reload"


def reload_server():
    print("Recarregando dashboard...")
    try:
        r = requests.get(RELOAD_URL, timeout=10)
        if r.status_code == 200:
            data = r.json()
            print(f"  deb_base : {data.get('deb_base')}")
            print(f"  deb_curr : {data.get('deb_curr')}")
            print(f"  cri_base : {data.get('cri_base')}")
            print(f"  cri_curr : {data.get('cri_curr')}")
            print("  Dashboard atualizada com sucesso.")
        else:
            print(f"  Servidor retornou HTTP {r.status_code}.")
    except requests.ConnectionError:
        print("  Servidor não está rodando — dados serão carregados no próximo start.")


if __name__ == "__main__":
    reload_server()
