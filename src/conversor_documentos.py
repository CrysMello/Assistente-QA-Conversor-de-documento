import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
import json
import xml.etree.ElementTree as ET
from pathlib import Path
try:
    import PyPDF2
except ImportError:
    try:
        import pypdf as PyPDF2
    except ImportError:
        messagebox.showerror("Erro", "Biblioteca PDF não encontrada. Instale PyPDF2 ou pypdf.")
import docx
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import re
import sys

class DocumentToExcelConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("🧪 Assistente QA - Conversor de Documentos para Casos de Teste")
        self.root.geometry("1200x800")
        self.root.configure(bg='#f0f0f0')
        
        # Centralizar janela
        self.center_window()
        
        self.current_file = None
        self.extracted_data = []
        self.preview_data = []
        self.templates = self.load_templates()
        self.current_template = "padrao_gherkin"
        
        self.setup_ui()
        
    def center_window(self):
        """Centraliza a janela na tela"""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
    def load_templates(self):
        """Carrega os templates de conversão"""
        return {
            "padrao_gherkin": {
                "name": "Padrão Gherkin",
                "columns": ['Historia/Requisito', 'Cenário', 'Dado', 'Quando', 'Então'],
                "mappings": {
                    'historia_requisito': 'Historia/Requisito',
                    'teste': 'Cenário', 
                    'dado': 'Dado',
                    'quando': 'Quando',
                    'entao': 'Então'
                }
            },
            "teste_detalhado": {
                "name": "Teste Detalhado",
                "columns": ['ID', 'Requisito', 'Cenário', 'Pré-condições', 'Passos', 'Resultado Esperado', 'Prioridade'],
                "mappings": {
                    'historia_requisito': 'Requisito',
                    'teste': 'Cenário',
                    'dado': 'Pré-condições',
                    'quando': 'Passos',
                    'entao': 'Resultado Esperado'
                }
            },
            "simple": {
                "name": "Simples",
                "columns": ['Requisito', 'Descrição Teste', 'Entrada', 'Saída Esperada'],
                "mappings": {
                    'historia_requisito': 'Requisito',
                    'teste': 'Descrição Teste',
                    'dado': 'Entrada',
                    'entao': 'Saída Esperada'
                }
            }
        }

    def setup_ui(self):
        # Frame principal com notebook para abas
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Aba principal - Conversão
        self.main_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.main_frame, text="🔄 Conversão Principal")
        
        # Aba Templates
        self.template_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.template_frame, text="🎨 Templates")
        
        # Aba Análise de Qualidade
        self.quality_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.quality_frame, text="📊 Análise de Qualidade")
        
        self.setup_main_tab()
        self.setup_template_tab()
        self.setup_quality_tab()

    def setup_main_tab(self):
        """Configura a aba principal de conversão"""
        # Configurar grid
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(4, weight=1)
        
        # Cabeçalho
        title_label = ttk.Label(self.main_frame, 
                               text="🧪 CONVERSOR DE DOCUMENTOS PARA QA", 
                               font=('Arial', 16, 'bold'),
                               foreground='#2c3e50')
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 15))
        
        # Frame de configurações
        config_frame = ttk.LabelFrame(self.main_frame, text="⚙️ CONFIGURAÇÕES", padding="10")
        config_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        config_frame.columnconfigure(1, weight=1)
        
        # Template selection
        ttk.Label(config_frame, text="Template:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.template_var = tk.StringVar(value=self.current_template)
        template_combo = ttk.Combobox(config_frame, textvariable=self.template_var, 
                                     values=list(self.templates.keys()), state="readonly")
        template_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        template_combo.bind('<<ComboboxSelected>>', self.on_template_change)
        
        # File selection
        file_frame = ttk.Frame(config_frame)
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0))
        file_frame.columnconfigure(1, weight=1)
        
        ttk.Label(file_frame, text="Arquivo:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.file_path_var = tk.StringVar()
        self.file_entry = ttk.Entry(file_frame, textvariable=self.file_path_var, state='readonly')
        self.file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(file_frame, text="📎 Anexar Documento", 
                  command=self.attach_document).grid(row=0, column=2)
        
        # Botões de ação
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=2, column=0, columnspan=3, pady=15)
        
        ttk.Button(button_frame, text="🔄 Converter", 
                  command=self.convert_document).grid(row=0, column=0, padx=5)
        ttk.Button(button_frame, text="👁️ Pré-visualizar", 
                  command=self.preview_conversion).grid(row=0, column=1, padx=5)
        ttk.Button(button_frame, text="💾 Exportar Excel", 
                  command=self.export_to_excel).grid(row=0, column=2, padx=5)
        ttk.Button(button_frame, text="📊 Analisar Qualidade", 
                  command=self.analyze_quality).grid(row=0, column=3, padx=5)
        ttk.Button(button_frame, text="🗑️ Limpar Tudo", 
                  command=self.clear_all).grid(row=0, column=4, padx=5)
        
        # Frame de pré-visualização
        preview_frame = ttk.LabelFrame(self.main_frame, text="👁️ PRÉ-VISUALIZAÇÃO", padding="10")
        preview_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        
        # Treeview
        self.setup_preview_tree(preview_frame)
        
        v_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_tree.yview)
        h_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        self.preview_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))

    def setup_preview_tree(self, parent_frame):
        """Configura a treeview de pré-visualização"""
        current_template = self.templates[self.current_template]
        columns = current_template["columns"]
        self.preview_tree = ttk.Treeview(parent_frame, columns=columns, show='headings', height=15)
        
        for col in columns:
            self.preview_tree.heading(col, text=col)
            self.preview_tree.column(col, width=150)
        
        self.preview_tree.bind('<Double-1>', self.on_double_click)

    def setup_template_tab(self):
        """Configura a aba de templates"""
        # Frame de seleção de template
        select_frame = ttk.LabelFrame(self.template_frame, text="🎨 SELECIONAR TEMPLATE", padding="10")
        select_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(select_frame, text="Template Atual:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.template_display_var = tk.StringVar()
        template_display = ttk.Entry(select_frame, textvariable=self.template_display_var, state='readonly')
        template_display.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(select_frame, text="🔄 Atualizar Visualização", 
                  command=self.update_template_preview).grid(row=0, column=2)
        
        # Visualização do template
        preview_frame = ttk.LabelFrame(self.template_frame, text="👁️ VISUALIZAÇÃO DO TEMPLATE", padding="10")
        preview_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        self.template_text = scrolledtext.ScrolledText(preview_frame, height=10, wrap=tk.WORD)
        self.template_text.pack(fill='both', expand=True)
        
        # Frame de criação de template
        create_frame = ttk.LabelFrame(self.template_frame, text="🛠️ CRIAR/EDITAR TEMPLATE", padding="10")
        create_frame.pack(fill='x')
        
        ttk.Label(create_frame, text="Nome do Template:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.new_template_name = ttk.Entry(create_frame)
        self.new_template_name.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Label(create_frame, text="Colunas (separadas por vírgula):").grid(row=1, column=0, sticky=tk.W, padx=(0, 10))
        self.new_template_columns = ttk.Entry(create_frame)
        self.new_template_columns.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(create_frame, text="➕ Criar Template", 
                  command=self.create_new_template).grid(row=2, column=0, columnspan=2, pady=10)
        
        self.update_template_display()
        
    def setup_quality_tab(self):
        """Configura a aba de análise de qualidade"""
        # Métricas
        metrics_frame = ttk.LabelFrame(self.quality_frame, text="📈 MÉTRICAS DE QUALIDADE", padding="10")
        metrics_frame.pack(fill='x', pady=(0, 10))
        
        self.metrics_text = scrolledtext.ScrolledText(metrics_frame, height=8, wrap=tk.WORD)
        self.metrics_text.pack(fill='both', expand=True)
        
        # Recomendações
        recommendations_frame = ttk.LabelFrame(self.quality_frame, text="💡 RECOMENDAÇÕES", padding="10")
        recommendations_frame.pack(fill='both', expand=True)
        
        self.recommendations_text = scrolledtext.ScrolledText(recommendations_frame, height=10, wrap=tk.WORD)
        self.recommendations_text.pack(fill='both', expand=True)
        
    def on_template_change(self, event=None):
        """Atualiza o template atual"""
        self.current_template = self.template_var.get()
        self.update_template_display()
        # Recriar a treeview com as novas colunas
        for widget in self.main_frame.winfo_children():
            if isinstance(widget, ttk.LabelFrame) and "PRÉ-VISUALIZAÇÃO" in widget.cget('text'):
                for child in widget.winfo_children():
                    if isinstance(child, ttk.Treeview):
                        child.destroy()
                self.setup_preview_tree(widget)
                break
        
    def update_template_display(self):
        """Atualiza a exibição do template"""
        template = self.templates[self.current_template]
        self.template_display_var.set(template["name"])
        
        display_text = f"Template: {template['name']}\n\n"
        display_text += f"Colunas: {', '.join(template['columns'])}\n\n"
        display_text += "Estrutura de Mapeamento:\n"
        for key, value in template["mappings"].items():
            display_text += f"  {key} → {value}\n"
            
        self.template_text.delete(1.0, tk.END)
        self.template_text.insert(1.0, display_text)
        
    def update_template_preview(self):
        """Atualiza a visualização do template"""
        self.update_template_display()
        
    def create_new_template(self):
        """Cria um novo template personalizado"""
        name = self.new_template_name.get().strip()
        columns_text = self.new_template_columns.get().strip()
        
        if not name or not columns_text:
            messagebox.showerror("Erro", "Por favor, preencha todos os campos.")
            return
            
        columns = [col.strip() for col in columns_text.split(',')]
        
        # Criar mapeamento básico
        mappings = {}
        default_fields = ['historia_requisito', 'teste', 'dado', 'quando', 'entao']
        for i, field in enumerate(default_fields):
            if i < len(columns):
                mappings[field] = columns[i]
            else:
                mappings[field] = columns[0]  # Fallback para primeira coluna
                
        self.templates[name] = {
            "name": name,
            "columns": columns,
            "mappings": mappings
        }
        
        # Atualizar combobox
        self.template_var.set(name)
        self.on_template_change()
        
        messagebox.showinfo("Sucesso", f"Template '{name}' criado com sucesso!")
        self.new_template_name.delete(0, tk.END)
        self.new_template_columns.delete(0, tk.END)
        
    def attach_document(self):
        """Anexa um documento para conversão"""
        file_types = [
            ("Todos os Formatos Suportados", "*.pdf *.txt *.docx *.doc *.json *.xml"),
            ("PDF Files", "*.pdf"),
            ("Text Files", "*.txt"),
            ("Word Documents", "*.docx *.doc"),
            ("JSON Files", "*.json"),
            ("XML Files", "*.xml"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.askopenfilename(filetypes=file_types)
        if filename:
            self.current_file = filename
            self.file_path_var.set(filename)
            self.extracted_data = self.extract_content(filename)
            
    def extract_content(self, file_path):
        """Extrai conteúdo baseado no tipo de arquivo"""
        file_extension = Path(file_path).suffix.lower()
        
        try:
            if file_extension == '.pdf':
                return self.extract_from_pdf(file_path)
            elif file_extension == '.txt':
                return self.extract_from_txt(file_path)
            elif file_extension in ['.docx', '.doc']:
                return self.extract_from_word(file_path)
            elif file_extension == '.json':
                return self.extract_from_json(file_path)
            elif file_extension == '.xml':
                return self.extract_from_xml(file_path)
            else:
                messagebox.showerror("Erro", "Formato de arquivo não suportado")
                return []
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao extrair conteúdo: {str(e)}")
            return []
    
    def extract_from_pdf(self, file_path):
        """Extrai texto de PDF"""
        content = ""
        with open(file_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page in pdf_reader.pages:
                content += page.extract_text() + "\n"
        return self.parse_test_cases(content)
    
    def extract_from_txt(self, file_path):
        """Extrai texto de TXT"""
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
        except UnicodeDecodeError:
            with open(file_path, 'r', encoding='latin-1') as file:
                content = file.read()
        return self.parse_test_cases(content)
    
    def extract_from_word(self, file_path):
        """Extrai texto de Word"""
        doc = docx.Document(file_path)
        content = ""
        for paragraph in doc.paragraphs:
            content += paragraph.text + "\n"
        return self.parse_test_cases(content)
    
    def extract_from_json(self, file_path):
        """Extrai dados de JSON"""
        with open(file_path, 'r', encoding='utf-8') as file:
            data = json.load(file)
        return self.parse_json_test_cases(data)
    
    def extract_from_xml(self, file_path):
        """Extrai dados de XML"""
        tree = ET.parse(file_path)
        root = tree.getroot()
        return self.parse_xml_test_cases(root)
    
    def parse_json_test_cases(self, data):
        """Analisa casos de teste de JSON"""
        test_cases = []
        
        if isinstance(data, list):
            for item in data:
                test_cases.extend(self.extract_from_json_object(item))
        elif isinstance(data, dict):
            test_cases.extend(self.extract_from_json_object(data))
            
        return test_cases if test_cases else self.create_fallback_cases(str(data))
    
    def extract_from_json_object(self, obj, path=""):
        """Extrai casos de teste de objeto JSON"""
        test_cases = []
        
        if isinstance(obj, dict):
            # Verificar se é um caso de teste estruturado
            if any(key.lower() in ['test', 'testcase', 'scenario', 'cenario'] for key in obj.keys()):
                test_case = self.create_test_case_from_json(obj)
                if test_case:
                    test_cases.append(test_case)
            else:
                for key, value in obj.items():
                    test_cases.extend(self.extract_from_json_object(value, f"{path}.{key}" if path else key))
                    
        elif isinstance(obj, list):
            for i, item in enumerate(obj):
                test_cases.extend(self.extract_from_json_object(item, f"{path}[{i}]"))
                
        return test_cases
    
    def create_test_case_from_json(self, obj):
        """Cria caso de teste a partir de objeto JSON"""
        test_case = {
            'historia_requisito': obj.get('requirement', obj.get('story', obj.get('feature', ''))),
            'teste': obj.get('test', obj.get('testcase', obj.get('scenario', obj.get('description', '')))),
            'dado': obj.get('given', obj.get('precondition', obj.get('context', ''))),
            'quando': obj.get('when', obj.get('action', obj.get('steps', ''))),
            'entao': obj.get('then', obj.get('expected', obj.get('result', '')))
        }
        
        # Se pelo menos um campo tem conteúdo, retorna o caso
        if any(test_case.values()):
            return test_case
        return None
    
    def parse_xml_test_cases(self, root):
        """Analisa casos de teste de XML"""
        test_cases = []
        
        # Procurar por elementos comuns de teste
        test_elements = root.findall('.//testcase') + root.findall('.//test') + root.findall('.//scenario')
        
        if test_elements:
            for element in test_elements:
                test_case = self.create_test_case_from_xml(element)
                if test_case:
                    test_cases.append(test_case)
        else:
            # Tentar extrair de qualquer estrutura XML
            test_cases.extend(self.extract_from_xml_element(root))
            
        return test_cases if test_cases else self.create_fallback_cases(ET.tostring(root, encoding='unicode'))
    
    def extract_from_xml_element(self, element):
        """Extrai casos de teste de elemento XML"""
        test_cases = []
        
        # Verificar se este elemento parece ser um caso de teste
        test_case = self.create_test_case_from_xml(element)
        if test_case:
            test_cases.append(test_case)
            
        # Recursivamente processar filhos
        for child in element:
            test_cases.extend(self.extract_from_xml_element(child))
            
        return test_cases
    
    def create_test_case_from_xml(self, element):
        """Cria caso de teste a partir de elemento XML"""
        test_case = {
            'historia_requisito': element.get('requirement', element.get('story', '')),
            'teste': element.get('name', element.get('title', element.tag)),
            'dado': '',
            'quando': '',
            'entao': ''
        }
        
        # Extrair texto do elemento e filhos
        text_content = element.text.strip() if element.text else ''
        for child in element:
            if child.text:
                text_content += f" {child.text}"
                
        if text_content:
            # Tentar parsear como Gherkin
            lines = text_content.split('.')
            for line in lines:
                line = line.strip().lower()
                if line.startswith('dado'):
                    test_case['dado'] = line
                elif line.startswith('quando'):
                    test_case['quando'] = line
                elif line.startswith('então') or line.startswith('entao'):
                    test_case['entao'] = line
                    
        return test_case if any(test_case.values()) else None

    # INICIO DA FUNÇÃO
    def parse_test_cases(self, content):
        test_cases = []
        lines = content.split('\n')
        
        current_case = {
            'historia_requisito': '',
            'teste': '',
            'dado': '',
            'quando': '',
            'entao': ''
        }
        # Variável de estado para a palavra-chave Gherkin atual
        current_gherkin_field = None 
        
        # Palavras-chave de detecção para evitar falsos positivos no texto de continuação QACRYS
        case_separators = ['historia', 'requisito', 'user story', 'cenário', 'scenario', 'teste']
        
        def is_new_separator(line):
            """Verifica se a linha é um separador de caso (Requisito/Cenário)"""
            lower_line = line.lower()
            return any(sep in lower_line for sep in case_separators)

        for line in lines:
            line = line.strip()
            
            if not line:
                # Linha vazia: finaliza o bloco de texto atual, mas não reseta o caso
                current_gherkin_field = None
                continue
                
            lower_line = line.lower()
            
            # 1. DETECÇÃO POR SEPARADORES (Requisito/Cenário)
            if 'historia/requisito:' in lower_line or 'história/requisito:' in lower_line or 'feature:' in lower_line or (is_new_separator(line) and 'cenário' not in lower_line and 'teste' not in lower_line):
                # Se encontrar um novo Requisito, salva o caso anterior se tiver conteúdo
                if any(current_case.values()):
                    # Limpa o whitespace final antes de salvar
                    for key in current_case: current_case[key] = current_case[key].strip()
                    test_cases.append(current_case.copy())
                current_case = {key: '' for key in current_case}
                current_case['historia_requisito'] = self.extract_after_colon(line)
                current_gherkin_field = None
                continue
                    
            elif 'cenário:' in lower_line or 'scenario:' in lower_line or 'teste:' in lower_line or ('cenário' in lower_line or 'scenario' in lower_line or 'teste' in lower_line):
                # Se encontrar um novo Cenário, zera os passos Gherkin e define o Cenário
                if current_case['teste'] and any(current_case.values()):
                    # Se já houver um cenário (múltiplos cenários no mesmo Requisito)
                    for key in current_case: current_case[key] = current_case[key].strip()
                    test_cases.append(current_case.copy())
                    current_case = {
                        'historia_requisito': current_case['historia_requisito'], # Mantém o Requisito
                        'teste': '', 'dado': '', 'quando': '', 'entao': ''
                    }
                    
                current_case['teste'] = self.extract_after_colon(line)
                current_gherkin_field = None
                continue
                
            # 2. DETECÇÃO E MUDANÇA DE ESTADO GHERKIN
            elif lower_line.startswith('dado') or lower_line.startswith('given'):
                current_gherkin_field = 'dado'
                current_case['dado'] += self.clean_gherkin_keyword(line) + " "
                continue
                    
            elif lower_line.startswith('quando') or lower_line.startswith('when'):
                current_gherkin_field = 'quando'
                current_case['quando'] += self.clean_gherkin_keyword(line) + " "
                continue
                    
            elif lower_line.startswith('então') or lower_line.startswith('entao') or lower_line.startswith('then'):
                current_gherkin_field = 'entao'
                current_case['entao'] += self.clean_gherkin_keyword(line) + " "
                continue
            
            # 3. CONTINUAÇÃO DE TEXTO
            elif current_gherkin_field:
                # Se estiver em um estado Gherkin e a linha não for um novo passo/separador,
                # adiciona a linha como continuação
                current_case[current_gherkin_field] += line + " "
                
        # Adiciona o último caso (limpa o whitespace final)
        if any(current_case.values()):
            for key in current_case:
                current_case[key] = current_case[key].strip()
            test_cases.append(current_case)
            
        return test_cases if test_cases else self.create_fallback_cases(content)

    def clean_gherkin_keyword(self, text):
        """Remove palavras-chave Gherkin e limpa espaços (v2)"""
        keywords = ['dado que', 'dado', 'given', 'quando', 'when', 'então', 'entao', 'then', 'e', 'and']
        
        # Cria uma lista de padrões de expressão regular para Gherkin seguido por espaço, vírgula ou dois pontos
        regex_patterns = [rf"^{re.escape(kw)}[\s:,]" for kw in keywords]
        
        lower_text = text.lower()
        
        for keyword in keywords:
            # 1. Tenta encontrar a palavra-chave exatamente no início da linha
            if lower_text.startswith(keyword):
                # 2. Tenta remover a palavra-chave seguida por qualquer separador comum (espaço, vírgula, dois pontos)
                # Exemplo: "Dado:..." ou "Quando,..."
                pattern = rf"^{re.escape(keyword)}[\s:,]*"
                match = re.match(pattern, text, re.IGNORECASE)
                
                if match:
                    # Retorna o texto após a correspondência, mantendo a capitalização original do corpo do texto
                    return text[match.end():].strip()
                
        # Fallback: se não começar com palavra-chave reconhecida, retorna o texto original
        return text.strip()
    #FIM DA FUNÇÃO

    def extract_after_colon(self, text):
        """Extrai texto após dois pontos, se existir"""
        if ':' in text:
            return text.split(':', 1)[1].strip()
        return text.strip()
    
    def clean_gherkin_keyword(self, text):
        """Remove palavras-chave Gherkin"""
        keywords = ['dado que', 'dado', 'given', 'quando', 'when', 'então', 'entao', 'then', 'e', 'and']
        cleaned_text = text.lower()
        
        # Tenta remover a palavra-chave no início da linha
        for keyword in keywords:
            if cleaned_text.startswith(keyword):
                # Se a palavra-chave estiver seguida por ":" (como "Quando:"), preserva o texto após ":"
                if ':' in text:
                    return text.split(':', 1)[1].strip()
                # Senão, remove a palavra-chave e capitaliza
                return text[len(keyword):].strip().capitalize()
                
        # Fallback: se não começar com palavra-chave, retorna o texto original
        return text

    def create_fallback_cases(self, content):
        """Cria casos de teste fallback"""
        test_cases = []
        paragraphs = [p for p in content.split('\n\n') if p.strip() and len(p.strip()) > 20]
        
        for i, para in enumerate(paragraphs[:10]):
            test_cases.append({
                'historia_requisito': f"Requisito {i+1}",
                'teste': f"Cenário {i+1}: {para[:80]}..." if len(para) > 80 else para,
                'dado': "Contexto a ser definido",
                'quando': "Ação a ser especificada",
                'entao': "Resultado esperado a ser determinado"
            })
            
        return test_cases
    
    def analyze_quality(self):
        """Analisa a qualidade dos casos de teste"""
        if not self.preview_data:
            messagebox.showwarning("Aviso", "Nenhum dado para analisar. Gere uma pré-visualização primeiro.")
            return
            
        metrics = self.calculate_metrics()
        recommendations = self.generate_recommendations(metrics)
        
        # Atualizar aba de qualidade
        self.metrics_text.delete(1.0, tk.END)
        self.metrics_text.insert(1.0, metrics)
        
        self.recommendations_text.delete(1.0, tk.END)
        self.recommendations_text.insert(1.0, recommendations)
        
        # Mudar para aba de qualidade
        self.notebook.select(2)
        
    def calculate_metrics(self):
        """Calcula métricas de qualidade"""
        total_cases = len(self.preview_data)
        
        # Análise de completude
        complete_cases = 0
        empty_fields = 0
        total_fields = total_cases * 5  # 5 campos por caso
        
        for case in self.preview_data:
            filled_fields = sum(1 for field in case.values() if field and field.strip())
            if filled_fields == 5:
                complete_cases += 1
            empty_fields += (5 - filled_fields)
        
        completeness = ((total_fields - empty_fields) / total_fields * 100) if total_fields > 0 else 0
        
        # Análise de conteúdo
        avg_lengths = {}
        for field in ['historia_requisito', 'teste', 'dado', 'quando', 'entao']:
            lengths = [len(str(case[field])) for case in self.preview_data if case[field]]
            avg_lengths[field] = sum(lengths) / len(lengths) if lengths else 0
        
        # Padrões Gherkin
        gherkin_patterns = 0
        for case in self.preview_data:
            if (self.contains_gherkin_keywords(case['dado']) or 
                self.contains_gherkin_keywords(case['quando']) or 
                self.contains_gherkin_keywords(case['entao'])):
                gherkin_patterns += 1
        
        metrics_text = f"📊 RELATÓRIO DE QUALIDADE - {datetime.now().strftime('%d/%m/%Y %H:%M')}\n\n"
        metrics_text += f"📈 ESTATÍSTICAS GERAIS:\n"
        metrics_text += f"   • Total de Casos de Teste: {total_cases}\n"
        metrics_text += f"   • Casos Completos: {complete_cases} ({complete_cases/total_cases*100:.1f}%)\n"
        metrics_text += f"   • Completude Geral: {completeness:.1f}%\n"
        metrics_text += f"   • Campos Vazios: {empty_fields}\n\n"
        
        metrics_text += f"📝 ANÁLISE DE CONTEÚDO:\n"
        for field, avg_len in avg_lengths.items():
            metrics_text += f"   • {field}: {avg_len:.1f} chars/caso\n"
        metrics_text += f"   • Padrões Gherkin Identificados: {gherkin_patterns} casos\n\n"
        
        metrics_text += f"🎯 SCORE DE QUALIDADE: {self.calculate_quality_score(completeness, avg_lengths, gherkin_patterns, total_cases)}/100"
        
        return metrics_text
    
    def contains_gherkin_keywords(self, text):
        """Verifica se contém palavras-chave Gherkin"""
        if not text:
            return False
        keywords = ['dado', 'given', 'quando', 'when', 'então', 'entao', 'then', 'and', 'e']
        text_lower = text.lower()
        return any(keyword in text_lower for keyword in keywords)
    
    def calculate_quality_score(self, completeness, avg_lengths, gherkin_patterns, total_cases):
        """Calcula score de qualidade"""
        score = 0
        
        # Completude (40%)
        score += completeness * 0.4
        
        # Conteúdo médio (30%)
        content_score = 0
        for field, avg_len in avg_lengths.items():
            if avg_len > 10:  # Mínimo de caracteres
                content_score += min(avg_len / 50, 1)  # Normalizar para máximo 50 chars
        score += (content_score / len(avg_lengths)) * 30
        
        # Padrões Gherkin (30%)
        gherkin_score = (gherkin_patterns / total_cases * 100) if total_cases > 0 else 0
        score += min(gherkin_score, 30)
        
        return min(score, 100)
    
    def generate_recommendations(self, metrics):
        """Gera recomendações de melhoria"""
        recommendations = "💡 RECOMENDAÇÕES PARA MELHORIA:\n\n"
        
        total_cases = len(self.preview_data)
        empty_count = sum(1 for case in self.preview_data for field in case.values() if not field or not field.strip())
        
        if empty_count > total_cases * 0.3:  # Mais de 30% vazios
            recommendations += "🔴 PRIORIDADE ALTA:\n"
            recommendations += "   • Preencha os campos vazios nos casos de teste\n"
            recommendations += "   • Revise a extração automática do documento fonte\n\n"
        
        # Verificar casos muito curtos
        short_cases = 0
        for case in self.preview_data:
            if len(str(case['teste'])) < 10:
                short_cases += 1
                
        if short_cases > total_cases * 0.2:  # Mais de 20% muito curtos
            recommendations += "🟡 PRIORIDADE MÉDIA:\n"
            recommendations += "   • Detalhe melhor os cenários de teste\n"
            recommendations += "   • Adicione mais contexto aos casos curtos\n\n"
        
        # Verificar padrões Gherkin
        gherkin_cases = sum(1 for case in self.preview_data 
                           if self.contains_gherkin_keywords(case['dado']) or 
                           self.contains_gherkin_keywords(case['quando']) or 
                           self.contains_gherkin_keywords(case['entao']))
        
        if gherkin_cases < total_cases * 0.5:  # Menos de 50% com Gherkin
            recommendations += "🔵 SUGESTÕES:\n"
            recommendations += "   • Use padrões Gherkin (Dado/Quando/Então)\n"
            recommendations += "   • Padronize a linguagem dos casos de teste\n\n"
        
        recommendations += "🎯 AÇÕES RECOMENDADAS:\n"
        recommendations += "   • Revise e edite os casos na pré-visualização\n"
        recommendations += "   • Use templates diferentes para diferentes necessidades\n"
        recommendations += "   • Exporte para Excel para análise adicional\n"
        
        return recommendations
    
    def preview_conversion(self):
        """Exibe pré-visualização"""
        if not self.extracted_data:
            messagebox.showwarning("Aviso", "Nenhum dado para pré-visualizar.")
            return
            
        self.preview_data = self.extracted_data.copy()
        self.update_preview_tree()
    
    def update_preview_tree(self):
        """Atualiza a treeview"""
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
            
        template = self.templates[self.current_template]
        
        for case in self.preview_data:
            values = []
            for col in template["columns"]:
                # Encontrar o campo correspondente
                field_value = ""
                for key, mapped_col in template["mappings"].items():
                    if mapped_col == col and key in case:
                        field_value = case[key]
                        break
                values.append(field_value)
            self.preview_tree.insert('', tk.END, values=values)
    
    def on_double_click(self, event):
        """Edição em linha"""
        item = self.preview_tree.selection()[0]
        column = self.preview_tree.identify_column(event.x)
        column_index = int(column[1:]) - 1
        
        current_value = self.preview_tree.item(item, 'values')[column_index]
        
        x, y, width, height = self.preview_tree.bbox(item, column)
        
        entry = ttk.Entry(self.preview_tree)
        entry.place(x=x, y=y, width=width, height=height)
        entry.insert(0, current_value)
        entry.select_range(0, tk.END)
        entry.focus()
        
        def save_edit(event=None):
            new_value = entry.get()
            self.preview_tree.set(item, column=column_index, value=new_value)
            
            # Atualizar dados
            item_index = self.preview_tree.index(item)
            if 0 <= item_index < len(self.preview_data):
                # Encontrar campo correspondente
                template = self.templates[self.current_template]
                col_name = template["columns"][column_index]
                for key, mapped_col in template["mappings"].items():
                    if mapped_col == col_name:
                        self.preview_data[item_index][key] = new_value
                        break
                        
            entry.destroy()
        
        entry.bind('<Return>', save_edit)
        entry.bind('<FocusOut>', lambda e: entry.destroy())
    
    def convert_document(self):
        """Converte documento"""
        if not self.current_file:
            messagebox.showwarning("Aviso", "Anexe um documento primeiro.")
            return
        self.preview_conversion()
    
    def export_to_excel(self):
        """Exporta para Excel"""
        if not self.preview_data:
            messagebox.showwarning("Aviso", "Nenhum dado para exportar.")
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="casos_teste_exportados.xlsx"
        )
        
        if filename:
            try:
                template = self.templates[self.current_template]
                
                # Criar DataFrame
                data_for_export = []
                for case in self.preview_data:
                    row = {}
                    for col in template["columns"]:
                        for key, mapped_col in template["mappings"].items():
                            if mapped_col == col and key in case:
                                row[col] = case[key]
                                break
                        if col not in row:
                            row[col] = ""
                    data_for_export.append(row)
                
                df = pd.DataFrame(data_for_export)
                df.to_excel(filename, index=False)
                
                messagebox.showinfo("Sucesso", f"Excel exportado: {filename}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao exportar: {str(e)}")
    
    def clear_all(self):
        """Limpa tudo"""
        self.current_file = None
        self.file_path_var.set("")
        self.extracted_data = []
        self.preview_data = []
        
        # Limpar treeview
        for item in self.preview_tree.get_children():
            self.preview_tree.delete(item)
            
        # Limpar análises de qualidade
        self.metrics_text.delete(1.0, tk.END)
        self.recommendations_text.delete(1.0, tk.END)
        
        messagebox.showinfo("Limpeza", "Todos os dados foram limpos!")

def main():
    """Função principal para executar a aplicação"""
    try:
        root = tk.Tk()
        app = DocumentToExcelConverter(root)
        root.mainloop()
    except Exception as e:
        print(f"Erro ao iniciar aplicação: {e}")
        messagebox.showerror("Erro", f"Falha ao iniciar aplicação: {e}")

if __name__ == "__main__":
    main()