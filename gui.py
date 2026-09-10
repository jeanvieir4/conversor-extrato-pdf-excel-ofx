# -*- coding: utf-8 -*-
"""
Interface grafica (janela nativa via pywebview) do Conversor de Extratos.
Reaproveita a mesma logica de conversao de converter.py - so troca a
"casca" do console preto por uma janela normal. Ver CONTEXTO_PROJETO.md,
secao "Ideias futuras", pro motivo dessa interface existir.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from converter import processar_pdf, _pasta_do_programa

import webview


def _caminho_recurso(nome):
    """Localiza um arquivo auxiliar (ex: gui.html) tanto rodando direto
    quanto empacotado como .exe pelo PyInstaller (onefile extrai pra uma
    pasta temporaria apontada por sys._MEIPASS)."""
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, nome)


class Api:
    def __init__(self, arquivos_iniciais=None):
        self.window = None
        self._arquivos_iniciais = arquivos_iniciais or []

    def arquivos_iniciais(self):
        """PDFs que vieram arrastados em cima do icone do .exe (o Windows
        chama o programa com o(s) caminho(s) em sys.argv). Mais confiavel
        que arrastar pra dentro da janela, que depende da versao do
        WebView2 instalada na maquina."""
        return self._arquivos_iniciais

    def escolher_pdfs(self):
        arquivos = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=True,
            file_types=('Arquivos PDF (*.pdf)', 'Todos os arquivos (*.*)'),
        )
        return list(arquivos) if arquivos else []

    def converter(self, caminhos):
        pasta_saida = os.environ.get('PASTA_SAIDA') or _pasta_do_programa()
        resultados = []
        for caminho in caminhos:
            nome_arquivo = os.path.basename(caminho)
            try:
                saida, total, avisos, arquivos_ofx = processar_pdf(caminho, pasta_saida)
                resultados.append({
                    'arquivo': nome_arquivo,
                    'ok': True,
                    'lancamentos': total,
                    'avisos': len(avisos),
                    'ofx': len(arquivos_ofx),
                })
            except Exception as e:
                resultados.append({
                    'arquivo': nome_arquivo,
                    'ok': False,
                    'erro': str(e),
                })
        return resultados


def _mostrar_erro_fatal(mensagem):
    """Sem console (--windowed), um erro na inicializacao faria o programa
    sumir sem explicacao nenhuma. Mostra uma caixa de mensagem nativa do
    Windows como ultimo recurso - nao depende de nenhuma biblioteca alem
    do proprio Windows."""
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            0, mensagem, 'Conversor de Extratos - Erro ao iniciar', 0x10  # MB_ICONERROR
        )
    except Exception:
        pass


def main():
    arquivos_iniciais = [a for a in sys.argv[1:] if a.lower().endswith('.pdf')]
    api = Api(arquivos_iniciais)
    with open(_caminho_recurso('gui.html'), encoding='utf-8') as f:
        html = f.read()
    window = webview.create_window(
        'Conversor de Extratos',
        html=html,
        js_api=api,
        width=760,
        height=680,
        min_size=(560, 480),
        background_color='#EEF2E9',  # mesma cor de fundo do app - sem isso
                                       # a janela fica branco-estourado (e
                                       # parece travada) enquanto o WebView2
                                       # ainda esta carregando o HTML
    )
    api.window = window
    webview.start(gui='edgechromium' if sys.platform == 'win32' else None)


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        import traceback
        _mostrar_erro_fatal(
            f'O programa nao conseguiu abrir.\n\n{e}\n\n'
            f'Detalhes tecnicos:\n{traceback.format_exc()}'
        )
        sys.exit(1)
