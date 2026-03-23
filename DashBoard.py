import matplotlib.pyplot as plt
import customtkinter as ctk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# -- classe do gráfico --
class DashboardWindow(ctk.CTkToplevel):
    def __init__(self, stats):
        super().__init__()
        self.title("Painel de Análise de Riscos")
        self.geometry("800x600")
        ctk.CTkLabel(self, text="Resumo de Ocorrências", font=("Roboto", 22, "bold")).pack(pady=20)
        
        fig, ax = plt.subplots(figsize=(8, 5))
        fig.patch.set_facecolor('#1e1e1e'); ax.set_facecolor('#2b2b2b')
        
        # - Filtrar apenas chaves numéricas para o gráfico
        plot_data = {k: v for k, v in stats.items() if isinstance(v, (int, float)) and k != 'dif_dc'}
        bars = ax.bar(plot_data.keys(), plot_data.values(), color='#1f538d')
        ax.tick_params(colors='white')
        ax.set_xticklabels(plot_data.keys(), rotation=45, ha='right', color='white')
        
        for bar in bars:
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.1, int(bar.get_height()), ha='center', color='white', fontweight='bold')
        plt.tight_layout()
        FigureCanvasTkAgg(fig, master=self).get_tk_widget().pack(fill="both", expand=True, padx=20, pady=20)
