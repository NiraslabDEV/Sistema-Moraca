from PIL import Image
import os

# Caminho para o arquivo de logo PNG
logo_png_path = os.path.join("Moraca", "assets", "logo.png")

# Caminho para salvar o arquivo ICO
logo_ico_path = os.path.join("Moraca", "assets", "logo.ico")

# Verificar se o arquivo PNG existe
if os.path.exists(logo_png_path):
    # Abrir a imagem PNG
    img = Image.open(logo_png_path)
    
    # Salvar como ICO
    img.save(logo_ico_path, format='ICO', sizes=[(16,16), (32,32), (48,48), (64,64), (128,128)])
    
    print(f"Ícone criado com sucesso em {logo_ico_path}")
else:
    print(f"Arquivo de logo não encontrado em {logo_png_path}") 