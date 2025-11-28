# Gerador de Calendário com Folgas

Aplicação oficial para geração de calendários de trabalho com sistema de folgas rotativo

## 📥 Download

Vá para [Releases](../../releases) para baixar a versão mais recente para o seu sistema operativo.

## 🚀 Como usar

1. Execute `Gerador_Calendario_Folgas.exe`
2. Escolha a semana de início do ciclo (1-12)
3. Um ficheiro Excel será gerado automaticamente

## 🔧 Desenvolvimento

```bash
# Instalar dependências
pip install -r requirements.txt

# Executar directamente
python gerador_calendario.py

# Criar executável
pyinstaller --onefile --console gerador_calendario.py
