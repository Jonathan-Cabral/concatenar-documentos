# 📄 Concatenador de Documentos

## Descrição
Uma aplicação de desktop em Python para concatenar (mesclar) múltiplos documentos em um único arquivo PDF. Suporta diversos formatos como PDF, Word e arquivos de texto.

## Funcionalidades
- Selecionar múltiplos tipos de documentos:
  - Arquivos PDF (*.pdf)
  - Documentos Microsoft Word (*.docx)
  - Arquivos de Texto (*.txt)
- Reordenar arquivos na lista
- Mover arquivos para cima ou para baixo
- Remover arquivos específicos
- Concatenar documentos em qualquer ordem desejada
- Conversão automática de formatos para PDF
- Barra de progresso detalhada durante o processamento
- Interface gráfica amigável
- Processamento em segundo plano para melhor desempenho

## Requisitos
- Python 3.x
- Bibliotecas:
  - tkinter (geralmente inclusa na instalação padrão)
  - PyPDF2
  - docx2pdf
  - python-docx
  - reportlab
  - Pillow (PIL)

## Instalação de Dependências
```bash
pip install PyPDF2 python-docx docx2pdf reportlab pillow
```

## Como Usar
1. Execute o script
2. Clique em "Selecionar Documentos"
3. Escolha os arquivos desejados (PDF, DOCX, TXT)
4. Reordene os arquivos se necessário
5. Clique em "Concatenar Documentos"
6. Escolha o local para salvar o arquivo final (formato PDF)

## Processo de Concatenação
1. Os arquivos PDF são usados diretamente
2. Os documentos Word (DOCX) são convertidos para PDF
3. Os arquivos de texto (TXT) são formatados e convertidos para PDF
4. Todos os PDFs são mesclados em um único documento final

## Botões e Funções
- **Selecionar Documentos**: Abre diálogo para escolher arquivos
- **Concatenar Documentos**: Processa e mescla os arquivos selecionados
- **Limpar Lista**: Remove todos os arquivos da lista
- **↑ Mover para Cima**: Move arquivo selecionado para cima
- **↓ Mover para Baixo**: Move arquivo selecionado para baixo
- **Remover Selecionado**: Remove arquivos escolhidos da lista

## Características Técnicas
- Suporte completo a documentos TXT de qualquer tamanho
- Paginação automática para arquivos de texto longos
- Preservação da formatação dos documentos originais
- Gerenciamento automático de arquivos temporários
- Tratamento de erros durante a conversão

## Roadmap de Desenvolvimento
🚧 Próximas Implementações:
- Suporte para outros tipos de documentos
  - Apresentações PowerPoint (.pptx)
  - Planilhas Excel (.xlsx)
  - Imagens (.jpg, .png)
- Visualização prévia dos documentos
- Ajuste de margens e orientação de página

## Contribuição
Contribuições são bem-vindas! Sugestões de novos recursos ou melhorias na interface são muito apreciadas.
