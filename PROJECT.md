# PPTX Merger by Jair Lima

## Descrição
Aplicativo desktop (GUI) e CLI para mesclar múltiplos arquivos PPTX em uma única apresentação. Criado para resolver o limite de 10 slides por geração do Gamma: gera as partes separadas e une tudo num clique.

## Repositório
https://github.com/jairslima/pptx-merger

## Executáveis (em ~/bin, no PATH global)
- `pptx-merger.exe` — GUI com janela (tkinter)
- `pptx-merger-cli.exe` — versão terminal com argparse e barra de progresso

## Stack e Dependências
- **Python 3.10+**
- `python-pptx >= 0.6.21` — manipulação de PPTX
- `tkinter` — GUI nativa (inclusa no Python padrão)
- `lxml` — dependência indireta do python-pptx
- `PyInstaller 6.19` — geração dos executáveis

## Estrutura de Arquivos
```
pptx-merger/
├── main.py          ← GUI tkinter (ponto de entrada gráfico)
├── cli.py           ← CLI com argparse (ponto de entrada terminal)
├── merger.py        ← lógica de mesclagem PPTX (compartilhada)
├── requirements.txt
├── LICENSE          ← MIT
├── .gitignore
└── PROJECT.md
```

## Comandos Essenciais

### Instalar dependências
```bash
pip install -r requirements.txt
```

### Executar GUI
```bash
python main.py
# ou (após build):
pptx-merger
```

### Executar CLI
```bash
# Auto-detecta Parte 1 / Parte 2 na pasta atual
pptx-merger-cli

# Listar arquivos sem mesclar
pptx-merger-cli --list

# Pasta específica
pptx-merger-cli "C:\minha\pasta\"

# Arquivos explícitos em ordem
pptx-merger-cli parte1.pptx parte2.pptx

# Com nome de saída customizado
pptx-merger-cli parte1.pptx parte2.pptx -o final.pptx
```

### Gerar executáveis
```bash
# GUI (sem console)
pyinstaller --onefile --windowed --name "pptx-merger" --clean main.py
cp dist/pptx-merger.exe ~/bin/

# CLI (com console)
pyinstaller --onefile --console --name "pptx-merger-cli" --clean cli.py
cp dist/pptx-merger-cli.exe ~/bin/
```

## Fluxo de Uso (GUI)
1. Abrir o app (`pptx-merger`)
2. Clicar em "Selecionar Pasta" e escolher a pasta com os PPTX
3. O app detecta automaticamente "Parte 1", "Parte 2" etc. e pré-seleciona na ordem correta
4. Ajustar seleção/ordem com checkboxes e botões ▲/▼ se necessário
5. Definir nome do arquivo de saída
6. Clicar em "Mesclar Arquivos"

## Caso de Uso Principal
Apresentações geradas pelo **Gamma** (limite de 10 slides por geração):
- `Apresentacao - Parte 1.pptx` (10 slides) + `Apresentacao - Parte 2.pptx` (8 slides)
- Resultado: `apresentacao_completa.pptx` (18 slides)

## Decisões Arquiteturais
- **Mesclagem via cópia de XML + remapeamento de relacionamentos**: preserva imagens, formas, transições e plano de fundo.
- **Clonagem de partes de mídia com nome único** (`merged_N.ext`): evita colisão de nomes entre apresentações com imagens de mesmo nome interno.
- **Deep copy do `_element` do slide**: mais fiel que recriar formas via API do python-pptx.
- **Thread separada na GUI**: evita travamento durante a operação.
- **Auto-detecção de partes**: regex para "parte N", "part N", "pN", "vol N", "cap N" no nome do arquivo.
- **Layout em branco**: cada slide novo usa o layout em branco da apresentação base; conteúdo original copiado por cima.

## Estado Atual
- Versão 1.0 — funcional e testado com apresentações reais do Gamma
- GUI: seleção por checkbox, reordenação ▲/▼, barra de progresso, auto-detecção
- CLI: auto-detecção, arquivos explícitos, barra de progresso no terminal, --list
- Suporte a imagens, formas, texto, transições e plano de fundo

## Problemas Conhecidos / Limitações
- Slides com **diagramas SmartArt** complexos podem perder formatação (limitação do python-pptx)
- **Animações de objetos** (não transições de slide) não são copiadas
- Se as apresentações usam temas muito diferentes, o fundo vem do primeiro arquivo
- Conteúdo OLE (Excel embedded, etc.) não é suportado

## Próximos Passos
- Arrastar e soltar arquivos na GUI
- Varredura recursiva de subpastas
- Preview miniatura dos slides
