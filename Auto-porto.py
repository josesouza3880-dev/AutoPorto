import os
import re
import shutil
import csv
import logging
import logging.handlers
import unicodedata
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from pathlib import Path
from pypdf import PdfReader

# ============================================================
#  SISTEMA MN-AutoPorto v5.0 — Validação e Envio Manual
# ============================================================

DIRETORIO_ATUAL = Path(__file__).parent.resolve() if '__file__' in globals() else Path.cwd()
LOG_CSV   = DIRETORIO_ATUAL / 'relatorio_movimentacao.csv'
LOG_AUDIT = DIRETORIO_ATUAL / 'auditoria.log'

CABECALHO_CSV = [
    'Nome Original', 'Nome Final', 'Tipo Documento', 'Data do Documento',
    'Turno', 'Caminho Destino', 'Tamanho (KB)', 'Data Processamento', 'Status'
]

# ── Funções de Apoio e Processamento de Texto ─────────────────

def configurar_log():
    handler = logging.handlers.RotatingFileHandler(
        LOG_AUDIT, maxBytes=5 * 1024 * 1024, backupCount=3, encoding='utf-8'
    )
    handler.setFormatter(logging.Formatter(
        '%(asctime)s | %(levelname)s | %(message)s', datefmt='%d/%m/%Y %H:%M:%S'
    ))
    logging.getLogger().setLevel(logging.INFO)
    logging.getLogger().addHandler(handler)

    if not LOG_CSV.exists():
        with open(LOG_CSV, 'w', newline='', encoding='utf-8') as f:
            csv.writer(f, delimiter=';').writerow(CABECALHO_CSV)

def extrair_texto_pdf(caminho: Path) -> str:
    try:
        reader = PdfReader(str(caminho))
        texto = ''.join([page.extract_text() or '' for page in reader.pages])
        return re.sub(r'[\x00-\x08\x0b-\x1f\x7f]', ' ', texto)
    except Exception:
        return ''

def detectar_data(texto: str) -> datetime | None:
    if not texto:
        return None
    padrao = r'\b(\d{1,2})[/\-\.](\d{1,2})[/\-\.](\d{4})\b'
    for match in re.finditer(padrao, texto):
        try:
            dia, mes, ano = int(match.group(1)), int(match.group(2)), int(match.group(3))
            if 1 <= dia <= 31 and 1 <= mes <= 12 and 1900 <= ano <= 2100:
                return datetime(ano, mes, dia)
        except ValueError:
            continue
    return None

def detectar_turno(texto: str) -> str:
    if not texto:
        return 'SemTurno'
    match = re.search(r'(\d)[°º]\s*turno\b', texto, re.IGNORECASE) or re.search(r'\bturno\s*[:\-]?\s*(\d)\b', texto, re.IGNORECASE)
    return f"Turno{match.group(1)}" if match else 'SemTurno'

def identificar_tipo(texto: str) -> str:
    texto_u = texto.upper()
    if "MINUTA" in texto_u:
        return "Minuta"
    if "DIARIO" in texto_u or "BORDO" in texto_u or "DB" in texto_u:
        return "DiarioDeBordo"
    return "Documento"

def nome_unico(pasta: Path, novo_nome: str) -> str:
    base, ext = os.path.splitext(novo_nome)
    ext = ext or '.pdf'
    candidato = novo_nome
    contador = 2
    while (pasta / candidato).exists():
        candidato = f"{base}_{contador}{ext}"
        contador += 1
    return candidato

def registrar_csv(nome_orig, nome_final, tipo, data_doc, turno, pasta_dest, tamanho_kb, agora, status):
    try:
        with open(LOG_CSV, 'a', newline='', encoding='utf-8') as f:
            csv.writer(f, delimiter=';').writerow([
                nome_orig, nome_final, tipo, data_doc, turno,
                pasta_dest, tamanho_kb, agora, status
            ])
    except Exception as e:
        logging.error(f"Erro ao salvar CSV: {e}")


# ── Interface de Confirmação e Validação Manual ──────────────

class ValidadorApp:
    def __init__(self, pasta_scan: Path):
        self.pasta_scan = pasta_scan
        self.arquivos = [a for a in os.listdir(pasta_scan) if (pasta_scan / a).is_file() and a.lower().endswith('.pdf')]
        self.indice = 0

        self.root = tk.Tk()
        self.root.title("MN-AutoPorto v5.0 — Validação e Destino Manual")
        self.root.geometry("580x380")
        self.root.resizable(False, False)

        self._construir_interface()
        
        # Validação caso a pasta não tenha PDFs
        if not self.arquivos:
            messagebox.showwarning("Atenção", "Nenhum arquivo PDF foi encontrado na pasta selecionada!")
            self.root.after(100, self.root.destroy)
        else:
            self._carregar_arquivo_atual()

    def _construir_interface(self):
        frame = ttk.Frame(self.root, padding="15")
        frame.pack(fill=tk.BOTH, expand=True)

        self.lbl_contador = ttk.Label(frame, text="", font=("Arial", 10, "bold"))
        self.lbl_contador.pack(anchor=tk.W, pady=(0, 5))

        self.lbl_original = ttk.Label(frame, text="", font=("Arial", 9), foreground="gray")
        self.lbl_original.pack(anchor=tk.W, pady=(0, 15))

        ttk.Label(frame, text="Nome Sugerido (pode editar se quiser):", font=("Arial", 9, "bold")).pack(anchor=tk.W)
        self.ent_novo_nome = ttk.Entry(frame, font=("Arial", 10))
        self.ent_novo_nome.pack(fill=tk.X, pady=(5, 15))

        ttk.Label(frame, text="Selecione a pasta de destino para este arquivo:", font=("Arial", 9, "bold")).pack(anchor=tk.W, pady=(0, 5))

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=5)

        ttk.Button(btn_frame, text="📁 Enviar para DIÁRIO DE BORDO", command=lambda: self._processar_envio("Diario_de_Bordo")).pack(fill=tk.X, pady=3)
        ttk.Button(btn_frame, text="📄 Enviar para MINUTA", command=lambda: self._processar_envio("Minuta")).pack(fill=tk.X, pady=3)
        ttk.Button(btn_frame, text="🔍 Escolher Outra Pasta...", command=lambda: self._processar_envio(None)).pack(fill=tk.X, pady=3)
        
        ttk.Separator(frame, orient='horizontal').pack(fill='x', pady=10)
        
        ttk.Button(frame, text="Pular Este Arquivo", command=self._proximo_arquivo).pack(side=tk.RIGHT)

    def _carregar_arquivo_atual(self):
        if self.indice >= len(self.arquivos):
            messagebox.showinfo("Concluído", "Todos os arquivos da pasta SCAN foram processados!")
            self.root.after(100, self.root.destroy)
            return

        nome_arq = self.arquivos[self.indice]
        caminho_arq = self.pasta_scan / nome_arq

        self.lbl_contador.config(text=f"Arquivo {self.indice + 1} de {len(self.arquivos)}")
        self.lbl_original.config(text=f"Original: {nome_arq}")

        texto = extrair_texto_pdf(caminho_arq)
        tipo = identificar_tipo(texto)
        data_doc = detectar_data(texto)
        turno = detectar_turno(texto)

        data_str = data_doc.strftime('%d-%m-%Y') if data_doc else 'SemData'
        nome_sugerido = f"{tipo}_{data_str}_{turno}.pdf"

        self.ent_novo_nome.delete(0, tk.END)
        self.ent_novo_nome.insert(0, nome_sugerido)
        
        self.data_doc_exib = data_doc.strftime('%d/%m/%Y') if data_doc else '-'
        self.turno_exib = turno
        self.tipo_exib = tipo

    def _processar_envio(self, subpasta_padrao: str | None):
        if self.indice >= len(self.arquivos):
            messagebox.showinfo("Concluído", "Todos os arquivos da pasta SCAN foram processados!")
            self.root.after(100, self.root.destroy)
            return

        nome_original = self.arquivos[self.indice]
        caminho_origem = self.pasta_scan / nome_original
        novo_nome_digitado = self.ent_novo_nome.get().strip()

        if not novo_nome_digitado:
            messagebox.showwarning("Atenção", "O nome do arquivo não pode ficar em branco.")
            return

        if not novo_nome_digitado.lower().endswith('.pdf'):
            novo_nome_digitado += '.pdf'

        if subpasta_padrao:
            pasta_destino = DIRETORIO_ATUAL / subpasta_padrao
            pasta_destino.mkdir(exist_ok=True)
        else:
            caminho_sel = filedialog.askdirectory(title="Selecione a pasta de destino")
            if not caminho_sel:
                return
            pasta_destino = Path(caminho_sel)

        try:
            nome_final = nome_unico(pasta_destino, novo_nome_digitado)
            caminho_final = pasta_destino / nome_final
            
            shutil.move(str(caminho_origem), str(caminho_final))
            
            tamanho_kb = round(caminho_final.stat().st_size / 1024, 2)
            agora_str = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

            registrar_csv(
                nome_original, nome_final, self.tipo_exib, self.data_doc_exib,
                self.turno_exib, str(pasta_destino), tamanho_kb, agora_str, 'MOVIDO'
            )
            
            self._proximo_arquivo()

        except Exception as e:
            messagebox.showerror("Erro ao Mover", f"Não foi possível mover o arquivo: {e}")

    def _proximo_arquivo(self):
        self.indice += 1
        self._carregar_arquivo_atual()

    def iniciar(self):
        self.root.mainloop()


if __name__ == '__main__':
    configurar_log()
    root_init = tk.Tk()
    root_init.withdraw()

    pasta_scan = filedialog.askdirectory(title="Selecione a pasta SCAN (Origem)")
    if pasta_scan:
        app = ValidadorApp(Path(pasta_scan))
        app.iniciar()