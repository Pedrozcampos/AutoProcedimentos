import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import os
from AuditProcess import AuditProcessor
from DashBoard import DashboardWindow

# - Configurações de Design
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("ProcAuditoria")
        self.geometry("500x550")
        self.configure(fg_color="#121212")

        self.main_container = ctk.CTkFrame(self, fg_color="#1e1e1e", corner_radius=25)
        self.main_container.pack(pady=40, padx=40, fill="both", expand=True)

        ctk.CTkLabel(self.main_container, text="Processamento", font=("Roboto", 24, "bold")).pack(pady=20)
        
        self.et_entry = ctk.CTkEntry(self.main_container, width=200, justify="center")
        self.et_entry.insert(0, "10000")
        self.et_entry.pack(pady=10)

        self.btn_run = ctk.CTkButton(self.main_container, text="Selecionar Arquivo", command=self.iniciar_processo)
        self.btn_run.pack(pady=20)

        self.progress_bar = ctk.CTkProgressBar(self.main_container, width=300)
        self.status_label = ctk.CTkLabel(self.main_container, text="")
        
        self.btn_dash = ctk.CTkButton(self.main_container, text="Ver Gráfico", command=self.open_dashboard, fg_color="transparent", border_width=2)
        self.btn_result = ctk.CTkButton(self.main_container, text="Abrir Excel", command=self.open_result, fg_color="transparent", border_width=2)

    def atualizar_interface_progresso(self, valor):
        self.progress_bar.set(valor)
        self.status_label.configure(text=f"Processando: {int(valor*100)}%")

    def iniciar_processo(self):
        path_in = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.csv")])
        if not path_in: return
        path_out = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="Journal Entries - .xlsx")
        if not path_out: return

        self.progress_bar.pack(pady=10); self.status_label.pack()
        self.btn_run.configure(state="disabled")
        threading.Thread(target=self.executar_tarefa, args=(path_in, path_out), daemon=True).start()

    def executar_tarefa(self, path_in, path_out):
        try:
            et = float(self.et_entry.get())
            proc = AuditProcessor(path_in, et, self.atualizar_interface_progresso)
            _, self.stats = proc.process_audit(path_out)
            self.ultimo_resultado = path_out
            self.after(0, self.finalizar_sucesso)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Erro", f"Falha: {str(e)}"))
            self.after(0, lambda: self.btn_run.configure(state="normal"))

    def finalizar_sucesso(self):
        messagebox.showinfo("Sucesso", "Concluído!")
        self.btn_run.configure(state="normal")
        self.btn_dash.pack(pady=5); self.btn_result.pack(pady=5)

    def open_result(self):
        if os.path.exists(self.ultimo_resultado): os.startfile(self.ultimo_resultado)

    def open_dashboard(self):
        if hasattr(self, 'stats'): DashboardWindow(self.stats)

if __name__ == "__main__":
    App().mainloop()