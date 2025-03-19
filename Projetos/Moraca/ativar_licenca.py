import ttkbootstrap as ttk
from ttkbootstrap.constants import *
import sys
import os
from licenca import verificar_licenca, salvar_licenca, gerar_nova_licenca_demo

class AtivadorLicenca:
    def __init__(self, root=None, callback=None):
        # Se root não for fornecido, criar uma nova janela
        if root is None:
            self.root = ttk.Window(themename="solar")
            self.root.title("Ativação de Licença - MORACA Sistemas")
            self.root.geometry("550x400")
            self.own_root = True
        else:
            self.root = root
            self.own_root = False
        
        # Callback a ser chamado quando a licença for ativada com sucesso
        self.callback = callback
        
        # Frame principal
        self.main_frame = ttk.Frame(self.root, padding=20)
        self.main_frame.pack(fill=BOTH, expand=YES)
        
        # Título
        ttk.Label(
            self.main_frame,
            text="Ativação de Licença - MORACA Sistemas",
            font=("Helvetica", 16, "bold")
        ).pack(pady=(0, 20))
        
        # Imagem (se existir)
        logo_path = os.path.join("assets", "logo.png")
        if os.path.exists(logo_path):
            try:
                from PIL import Image, ImageTk
                img = Image.open(logo_path)
                img = img.resize((200, 60), Image.LANCZOS)
                logo_img = ImageTk.PhotoImage(img)
                
                # Manter referência para evitar garbage collection
                self.logo_img = logo_img
                
                # Exibir logo
                ttk.Label(
                    self.main_frame,
                    image=logo_img
                ).pack(pady=(0, 20))
            except ImportError:
                pass
        
        # Frame para entrada do código
        input_frame = ttk.Frame(self.main_frame)
        input_frame.pack(fill=X, pady=10)
        
        ttk.Label(
            input_frame,
            text="Digite o código de licença:",
            font=("Helvetica", 11)
        ).pack(anchor="w", pady=5)
        
        self.codigo_licenca_entry = ttk.Text(
            input_frame,
            width=40,
            height=6,
            wrap="word"
        )
        self.codigo_licenca_entry.pack(fill=X, pady=5)
        
        # Frame para botões
        btn_frame = ttk.Frame(self.main_frame)
        btn_frame.pack(pady=20)
        
        ttk.Button(
            btn_frame,
            text="Ativar Licença",
            style="primary.TButton",
            command=self.ativar_licenca,
            width=20
        ).pack(side=LEFT, padx=5)
        
        ttk.Button(
            btn_frame,
            text="Gerar Licença Demo",
            style="secondary.TButton",
            command=self.gerar_demo,
            width=20
        ).pack(side=LEFT, padx=5)
        
        # Frame para status
        status_frame = ttk.LabelFrame(self.main_frame, text="Status da Licença", padding=10)
        status_frame.pack(fill=X, pady=10)
        
        self.status_label = ttk.Label(
            status_frame,
            text="Nenhuma licença ativada",
            font=("Helvetica", 10),
            wraplength=500
        )
        self.status_label.pack(fill=X, pady=5)
        
        # Botões de cancelar/sair
        ttk.Button(
            self.main_frame,
            text="Sair",
            style="danger.Outline.TButton",
            command=self.sair,
            width=15
        ).pack(pady=10)
    
    def ativar_licenca(self):
        # Obter código da licença
        codigo = self.codigo_licenca_entry.get("1.0", END).strip()
        
        if not codigo:
            self.status_label.config(text="Digite um código de licença válido.")
            return
        
        # Verificar licença
        resultado = verificar_licenca(codigo)
        
        if resultado["valida"]:
            # Salvar licença válida
            salvar_licenca(codigo)
            
            # Atualizar status
            self.status_label.config(
                text=f"Licença ativada com sucesso para {resultado['cliente']}.\n{resultado['mensagem']}",
                foreground="green"
            )
            
            # Chamar callback se existir
            if self.callback:
                self.callback(resultado)
            
            # Fechar janela se for própria
            if self.own_root:
                self.root.after(2000, self.root.destroy)
        else:
            # Atualizar status com erro
            self.status_label.config(
                text=f"Falha na ativação: {resultado['mensagem']}",
                foreground="red"
            )
    
    def gerar_demo(self):
        # Gerar licença demo
        codigo_demo = gerar_nova_licenca_demo()
        
        # Mostrar no campo de texto
        self.codigo_licenca_entry.delete("1.0", END)
        self.codigo_licenca_entry.insert("1.0", codigo_demo)
        
        # Atualizar status
        self.status_label.config(
            text="Licença de demonstração gerada. Clique em 'Ativar Licença' para ativá-la.",
            foreground="blue"
        )
    
    def sair(self):
        if self.own_root:
            self.root.destroy()
            sys.exit(0)
        else:
            self.main_frame.destroy()

if __name__ == "__main__":
    app = AtivadorLicenca()
    app.root.mainloop() 