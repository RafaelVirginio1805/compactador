import os
import pandas as pd
import shutil
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox

def selecionar_excel():
    caminho = filedialog.askopenfilename(
        title="Selecione a planilha Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if caminho:
        entry_excel.delete(0, END)
        entry_excel.insert(0, caminho)

def selecionar_pastas():
    caminho = filedialog.askdirectory(title="Selecione a pasta onde estão as pastas a compactar")
    if caminho:
        entry_pastas.delete(0, END)
        entry_pastas.insert(0, caminho)

def selecionar_destino():
    caminho = filedialog.askdirectory(title="Selecione o local de destino dos arquivos ZIP")
    if caminho:
        entry_destino.delete(0, END)
        entry_destino.insert(0, caminho)

def iniciar_zip():
    caminho_excel = entry_excel.get().strip()
    caminho_pastas = entry_pastas.get().strip()
    destino_zips = entry_destino.get().strip()

    if not (caminho_excel and caminho_pastas and destino_zips):
        messagebox.showwarning("Campos obrigatórios", "Por favor, selecione todos os caminhos antes de iniciar.")
        return

    try:
        df = pd.read_excel(caminho_excel)
        if "Notificação" not in df.columns:
            messagebox.showerror("Erro", "A planilha precisa ter uma coluna chamada 'Notificação'.")
            return
        notificacoes = df["Notificação"].astype(str).tolist()
    except Exception as e:
        messagebox.showerror("Erro ao ler Excel", str(e))
        return

    total = len(notificacoes)
    progresso['maximum'] = total
    progresso['value'] = 0

    os.makedirs(destino_zips, exist_ok=True)
    encontrados = 0

    for i, notif in enumerate(notificacoes, start=1):
        encontrada = None
        for pasta in os.listdir(caminho_pastas):
            if pasta.startswith(notif):
                encontrada = os.path.join(caminho_pastas, pasta)
                break

        if encontrada and os.path.isdir(encontrada):
            zip_nome = os.path.join(destino_zips, f"{notif}.zip")
            try:
                shutil.make_archive(zip_nome.replace(".zip", ""), 'zip', encontrada)
                encontrados += 1
            except Exception as e:
                print(f"Erro ao compactar {encontrada}: {e}")

        progresso['value'] = i
        root.update_idletasks()

    messagebox.showinfo("Concluído", f" {encontrados} pastas compactadas com sucesso!")

def main():
    global root, entry_excel, entry_pastas, entry_destino, progresso

    root = ttk.Window(themename="darkly")
    root.title("Compactador de Pastas Automático")
    root.geometry("650x400")
    root.resizable(False, False)

    ttk.Label(root, text="Compactador de Pastas", font=("Segoe UI", 16, "bold")).pack(pady=15)

    frame = ttk.Frame(root)
    frame.pack(padx=20, pady=10, fill=X)

    ttk.Label(frame, text="Planilha Excel (coluna: Notificação):").grid(row=0, column=0, sticky=W, pady=5)
    entry_excel = ttk.Entry(frame, width=60)
    entry_excel.grid(row=1, column=0, padx=5)
    ttk.Button(frame, text="Selecionar", command=selecionar_excel, bootstyle=INFO).grid(row=1, column=1, padx=5)

    ttk.Label(frame, text="Caminho das pastas a compactar:").grid(row=2, column=0, sticky=W, pady=5)
    entry_pastas = ttk.Entry(frame, width=60)
    entry_pastas.grid(row=3, column=0, padx=5)
    ttk.Button(frame, text="Selecionar", command=selecionar_pastas, bootstyle=INFO).grid(row=3, column=1, padx=5)

    ttk.Label(frame, text="Destino dos arquivos ZIP:").grid(row=4, column=0, sticky=W, pady=5)
    entry_destino = ttk.Entry(frame, width=60)
    entry_destino.grid(row=5, column=0, padx=5)
    ttk.Button(frame, text="Selecionar", command=selecionar_destino, bootstyle=INFO).grid(row=5, column=1, padx=5)

    progresso = ttk.Progressbar(root, bootstyle=SUCCESS, mode='determinate')
    progresso.pack(padx=20, pady=20, fill=X)

    ttk.Button(root, text="Iniciar Compactação", bootstyle=SUCCESS, command=iniciar_zip).pack(pady=10)
    ttk.Label(root, text="Desenvolvido pelo estagiário de U.M - RAFAEL VIRGINIO", font=("Segoe UI", 9)).pack(side=BOTTOM, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
