### **README.md – Sistema de Apuração Domínio**  

---

## **Sobre o Projeto**  
Este sistema automatiza a **apuração de impostos** no sistema **Domínio** utilizando **Python** e a biblioteca **Pywinauto** para automação de tarefas.  

Ele permite:  
1. **Login automático** no sistema Domínio.  
2. **Processamento de apuração** para empresas com base em dados fornecidos via planilha Excel.  
3. **Monitoramento do progresso em tempo real**.  

O projeto utiliza **Flask** para criar uma interface web simples e interativa, facilitando o controle e o acompanhamento do processo.  

---

## **Recursos Principais**  
- **Upload de Planilha**: Permite carregar uma lista de empresas para processamento em lote.  
- **Login Automático**: Efetua login no sistema Domínio sem intervenção manual.  
- **Processamento em Lote**: Processa várias empresas de forma automatizada.  
- **Progresso em Tempo Real**: Exibe mensagens de progresso para o usuário.  
- **Interface Amigável**: Frontend desenvolvido com **HTML5**, **CSS3** e **JavaScript**.  

---

## **Tecnologias Utilizadas**  

### **Backend:**  
- **Python 3.x**  
- **Flask**  
- **Pywinauto** (Automação de GUI)  
- **OpenPyXL** (Manipulação de planilhas Excel)  
- **PyAutoGUI** (Interação com mouse e teclado)  

### **Frontend:**  
- **HTML5**  
- **CSS3**  
- **JavaScript** (com Fetch API)  
- **Bootstrap 5** (Design responsivo)  
- **FontAwesome** (Ícones)  

### **Banco de Dados:**  
- **Arquivos Excel (.xlsx)** para entrada de dados.  

---

## **Instalação e Configuração**  

### **1. Clonar o Repositório**  
```bash
git clone https://github.com/seu-usuario/sistema-apuracao-dominio.git
cd sistema-apuracao-dominio
```

### **2. Criar Ambiente Virtual**  
```bash
python -m venv venv
source venv/bin/activate     # Linux/MacOS
venv\Scripts\activate        # Windows
```

### **3. Instalar Dependências**  
```bash
pip install -r requirements.txt
```

### **4. Configurar Caminhos no Código**  
No arquivo **app.py**, atualize o caminho do executável do sistema Domínio:  

```python
CAMINHO_EXECUTAVEL = r"C:\Caminho\para\dominio.exe"
```

### **5. Inicializar o Servidor Flask**  
```bash
python app.py
```

## **Como Usar**  

### **1. Tela Inicial**  
- Acesse o sistema via navegador.  
- Preencha o formulário com:  
  - Usuário Domínio  
  - Senha Domínio  
  - Planilha Excel com os dados das empresas.  

### **2. Processamento**  
- Clique em **Enviar** para iniciar o processamento.  
- Acompanhe o progresso em tempo real na seção **Progresso**.  

### **3. Planilha Excel (Formato Aceito)**  
A planilha deve conter as colunas:  
- **Código da Empresa**  
- **Mês/Ano (AAAAMM)**  

**Exemplo:**  

| Código Empresa | Mês/Ano   |  
|----------------|-----------|  
| 1234           | 01/2024   |  
| 5678           | 02/2024   |  

---

## **Segurança e Boas Práticas**  
1. **Variáveis Sensíveis**: Nunca compartilhe senhas ou caminhos no código-fonte.  
2. **Controle de Acesso**: Restrinja o acesso ao sistema na rede.  
3. **Backup de Planilhas**: Guarde cópias das planilhas enviadas.  
4. **HTTPS em Produção**: Sempre use HTTPS para proteger os dados.  

---

## **Estrutura do Projeto**  

```
/static/
  css/
    index.css
  js/
    index.js
  images/
    logo.png
/templates/
  index.html
app.py
requirements.txt
README.md
```

---

## **Requisitos e Dependências**  

### **Bibliotecas Utilizadas**  
```plaintext
Flask
pywinauto
pyautogui
openpyxl
```

**Gerar lista automaticamente (caso precise):**  
```bash
pip freeze > requirements.txt
```

---

## **Exemplos de Comandos**  

- **Executar o Servidor Flask:**  
  ```bash
  python app.py
  ```

- **Monitorar Progresso via API:**  
  ```bash

---

## **Problemas Conhecidos**  
- **Timeout na Apuração**: Se o sistema Domínio demorar muito para responder, o código pode gerar timeout.  
- **Erros na Automação**: Alterações na interface do Domínio podem exigir ajustes nos comandos automatizados.  

---

## **Contribuição**  
Contribuições são bem-vindas!  

1. Faça um **fork** do projeto.  
2. Crie um **branch** para sua nova funcionalidade:  
   ```bash
   git checkout -b nova-funcionalidade
   ```
3. Commit suas alterações:  
   ```bash
   git commit -m "Adiciona nova funcionalidade"
   ```
4. Envie o **pull request**.  

---

## **Licença**  
Este projeto é distribuído sob a licença **MIT**. Consulte o arquivo **LICENSE** para mais detalhes.  

---
