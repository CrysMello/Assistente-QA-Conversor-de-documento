🚀 Visão Geral
O Conversor transforma os arquivos gerados por IA Generativa e salvos nos formatos  PDF, Word, TXT em casos de teste prontos para uso no formato exvel. A ferramenta automatiza a extração e estruturação de cenários de teste, economizando horas de trabalho manual.

*Este projeto foi gerado por IA e algumas modificações foram feitas por mim para se adequar as necessidades do meu dia a dia. 

✨ Funcionalidades
🔄 Conversão Inteligente
Multi-formatos: Suporte a PDF, Word, TXT, JSON e XML

Parsing automático: Identifica automaticamente cenários, pré-condições e resultados esperados

Fallback inteligente: Gera casos mesmo em documentos não estruturados

🎨 Templates Personalizáveis
Padrão Gherkin: Dado/Quando/Então

Teste Detalhado: Com pré-condições, passos e prioridades

Simples: Formato básico para documentação rápida

Customizável: Crie seus próprios templates

📊 Análise de Qualidade
Métricas automáticas: Completude, conteúdo e padrões

Score de qualidade: Pontuação de 0-100

Recomendações: Sugestões para melhorar os casos

Relatórios: Análise detalhada com estatísticas

💾 Exportação Avançada
Excel formatado: Estrutura pronta para planilhas

Edição em linha: Clique duplo para editar diretamente na tabela

Persistência: Mantém alterações durante a sessão

📄 Formatos Suportados
Formato	Recursos	Melhor Uso
PDF	Extração de texto	Documentação técnica, requisitos
Word	Parágrafos e estrutura	Especificações funcionais
TXT	Texto puro	User stories, cenários simples
JSON	Estrutura hierárquica	APIs, testes automatizados
XML	Tags e atributos	Configurações, dados estruturados
🔧 Instalação
Pré-requisitos
Python 3.8 ou superior

pip (gerenciador de pacotes Python)

Instalação das Dependências
bash
# Instalar dependências principais
pip install pandas openpyxl pypdf2 python-docx

# Ou usando requirements.txt
pip install -r requirements.txt
Executável (Recomendado para Usuários Finais)
bash
# Gerar executável
python -m PyInstaller --onefile --windowed --name "ConversorDocumentos" conversor_documentos.py

# O executável estará em: dist/ConversorDocumentos.exe
🎯 Como Usar
1. Iniciar a Aplicação
bash
python conversor_documentos.py
Ou execute o arquivo ConversorDocumentos.exe

2. Configuração Inicial
Selecione o template desejado

Anexe seu documento clicando em "📎 Anexar Documento"

3. Conversão
Clique em "🔄 Converter" para processar o documento

Use "👁️ Pré-visualizar" para ver os resultados

4. Edição e Ajustes
Clique duplo em qualquer célula para editar

Ajuste os casos conforme necessário

5. Análise e Exportação
Use "📊 Analisar Qualidade" para métricas

Clique em "💾 Exportar Excel" para salvar

🎨 Templates
📝 Padrão Gherkin (Recomendado)
text
Historia/Requisito | Cenário | Dado | Quando | Então
Ideal para: BDD, testes comportamentais

🔍 Teste Detalhado
text
ID | Requisito | Cenário | Pré-condições | Passos | Resultado Esperado | Prioridade
Ideal para: Documentação formal, processos rigorosos

⚡ Simples
text
Requisito | Descrição Teste | Entrada | Saída Esperada
Ideal para: Prototipagem, projetos ágeis

🛠️ Criando Templates Personalizados
Vá para a aba "🎨 Templates"

Preencha o nome e colunas (separadas por vírgula)

Clique em "➕ Criar Template"

O novo template estará disponível imediatamente

📈 Análise de Qualidade
Métricas Calculadas
Completude: Percentual de campos preenchidos

Conteúdo: Tamanho médio dos textos

Padrões Gherkin: Identificação de keywords

Score Geral: Pontuação consolidada (0-100)

Recomendações Automáticas
🔴 Alta prioridade: Campos vazios, extração problemática

🟡 Média prioridade: Cenários muito curtos

🔵 Sugestões: Melhorias de padrão e linguagem

🏗️ Estrutura do Projeto
text
conversor_documentos/
├── conversor_documentos.py      # Código principal
├── requirements.txt             # Dependências
├── build/                       # Arquivos de build
├── dist/                        # Executável final
└── README.md                    # Esta documentação


Arquitetura da Aplicação
python
DocumentToExcelConverter
├── __init__()                   # Inicialização
├── setup_ui()                   # Interface gráfica
├── extract_content()            # Extração multi-formatos
├── parse_test_cases()           # Análise de conteúdo
├── analyze_quality()            # Métricas de qualidade
└── export_to_excel()            # Exportação

🔧 Desenvolvimento
Estrutura de Classes Principais
python
class DocumentToExcelConverter:
    # Gerenciamento de estado
    - current_file: str
    - extracted_data: List[Dict]
    - preview_data: List[Dict]
    - templates: Dict
    
    # Processamento de documentos
    - extract_from_pdf()
    - extract_from_word()
    - extract_from_json()
    - extract_from_xml()
    
    # Análise e qualidade
    - calculate_metrics()
    - generate_recommendations()
    - calculate_quality_score()
Adicionando Novos Parsers
python
def extract_from_novo_formato(self, file_path):
    # Implementar lógica de extração
    content = self.ler_arquivo(file_path)
    return self.parse_test_cases(content)
🐛 Troubleshooting
Problemas Comuns
❌ Executável não abre

Verifique se todas as dependências estão incluídas

Execute como administrador se necessário

❌ Erro na extração de PDF

Instale pypdf2 ou pypdf: pip install pypdf2

❌ Documento Word não carrega

Verifique se python-docx está instalado: pip install python-docx

❌ Encoding problems em TXT

A aplicação tenta UTF-8 e Latin-1 automaticamente

Logs e Debug
Para debugging, execute via linha de comando:

bash
python conversor_documentos.py
📊 Exemplos de Uso
Caso 1: Documentação de Requisitos
Anexe um PDF com user stories

Use template "Padrão Gherkin"

Converta e edite os cenários

Exporte para Excel para compartilhar com a equipe

Caso 2: Especificação de API
Anexe JSON com endpoints

Use template "Teste Detalhado"

Analise a qualidade

Ajuste baseado nas recomendações

Caso 3: Migração de Testes
Anexe documento Word com casos antigos

Use template personalizado

Converta e refine

Exporte para novo formato


🆕 Changelog
v1.0.0
✅ Conversão multi-formatos

✅ Templates personalizáveis

✅ Análise de qualidade

✅ Exportação para Excel

✅ Interface gráfica intuitiva

