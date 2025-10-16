import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import socket
import subprocess
import platform
import re
import csv
import webbrowser
import ipaddress
import threading
import queue
from concurrent.futures import ThreadPoolExecutor
import os

# ========================================================================
# Nome do Sistema: HostFlow
# Projetado por: Carlos Vilela
# Desenvolvimento: 16/10/2025
# Testes realizados com o Sistema:
# Performance de 11 consultas por segundo com 300 threads simultâneas
# Equipamento:
# Intel i5-1345U
# 32GB de RAM
# Windows 11 Enterprise
# ========================================================================

# Defina o número máximo de hosts permitidos
MAX_HOSTS = 3000  # 11 por segundo com 300 threads simultâneas
# Limite de threads em uso
MAX_THREADS = 300  # 11 por segundo com 300 threads simultâneas
# Limites de TTL para determinar o sistema operacional
TTL_MIN_LINUX = 1
TTL_MAX_LINUX = 100
TTL_MIN_WINDOWS = 101
TTL_MAX_WINDOWS = 255

# Variável global para armazenar os hosts
hosts_list = []
# Variável global para armazenar o caminho do arquivo aberto
current_file_path = None

def is_valid_host(host):
    """Verifica se o host é um IP ou um nome de host válido."""
    try:
        ipaddress.ip_address(host)  # Valida se é um IP
        return True
    except ValueError:
        hostname_pattern = re.compile(r'^[a-zA-Z0-9.-]+$')  # Regex para hostname
        return bool(hostname_pattern.match(host))

def ping(host):
    """Realiza um ping no host e retorna o resultado e TTL."""
    param = '-n' if platform.system().lower() == 'windows' else '-c'
    command = ['ping', param, '1', host]
    try:
        output = subprocess.check_output(command, stderr=subprocess.STDOUT, universal_newlines=True)
        ttl_match = re.search(r'TTL=(\d+)', output)
        ttl = ttl_match.group(1) if ttl_match else 'Não encontrado'
        return True, ttl  # Retorna True se o ping for bem-sucedido
    except subprocess.CalledProcessError:
        return False, 'Erro ao pingar'  # Retorna False se o ping falhar

def dns_lookup(host):
    """Realiza a resolução de DNS do host e retorna o IP e o hostname reverso."""
    try:
        ip = socket.gethostbyname(host)
        reverse_host = socket.gethostbyaddr(ip)
        return ip, reverse_host[0]
    except (socket.gaierror, socket.herror):
        return None, None

def check_port(ip, port):
    """Verifica se a porta está aberta no IP fornecido."""
    try:
        with socket.create_connection((ip, port), timeout=2) as sock:
            return True
    except (socket.timeout, ConnectionRefusedError):
        return False

def get_os(ttl):
    """Determina o sistema operacional baseado no TTL."""
    if ttl != 'Não encontrado':
        ttl_value = int(ttl)
        if TTL_MIN_LINUX <= ttl_value <= TTL_MAX_LINUX:
            return 'Linux'
        elif TTL_MIN_WINDOWS <= ttl_value <= TTL_MAX_WINDOWS:
            return 'Windows'
    return 'Desconhecido'

def read_inventory(file_path):
    """Lê o arquivo CSV e retorna um dicionário com as informações."""
    inventory = {}
    try:
        with open(file_path, newline='', encoding='utf-8-sig') as csvfile:
            reader = csv.DictReader(csvfile, delimiter=';')
            for row in reader:
                code = row['Código']  # A chave deve corresponder ao cabeçalho
                inventory[code] = row
    except FileNotFoundError:
        messagebox.showerror("Erro", f"O arquivo {file_path} não foi encontrado.")
    except KeyError as e:
        messagebox.showerror("Erro", f"A chave {e} não foi encontrada no arquivo CSV.")
    return inventory

def analyze_host(host, inventory, result_queue):
    """Analisa um único host e coloca os resultados na fila."""
    original_host = host
    x_host = host + 'x'  # Cria o host com 'x' no final

    # Testa ambos os hosts
    original_ping_result, original_ttl = ping(original_host)
    x_ping_result, x_ttl = ping(x_host)

    # Determina qual host está pingando
    pinging_host = original_host if original_ping_result else x_host if x_ping_result else None

    # Determina informações adicionais
    if pinging_host:
        ip, reverse_host = dns_lookup(pinging_host)
        ssh_open = check_port(ip, 22) if ip else False
        rdp_open = check_port(ip, 3389) if ip else False
        os_name = get_os(original_ttl if pinging_host == original_host else x_ttl)
        localization_info = inventory.get(original_host, {})
    else:
        ip, reverse_host = dns_lookup(original_host)  # Tenta resolver o DNS do original_host
        ssh_open = check_port(ip, 22) if ip else False
        rdp_open = check_port(ip, 3389) if ip else False
        os_name = 'Não encontrado'
        localization_info = inventory.get(original_host, {})

    # Coloca os resultados na fila
    result_queue.put((original_host, pinging_host, reverse_host, 'True' if pinging_host else 'False',
                      ip or 'Não resolvido', original_ttl if pinging_host == original_host else x_ttl,
                      os_name, ssh_open, rdp_open,
                      localization_info.get('Local', 'Não encontrado'),
                      localization_info.get('Prédio', 'Não encontrado'),
                      localization_info.get('Andar', 'Não encontrado'),
                      localization_info.get('Escritório', 'Não encontrado'),
                      localization_info.get('Obsoleto', 'Não encontrado'),
                      localization_info.get('Anotação', 'Não encontrado')))

def analyze_hosts():
    """Analisa os hosts na tabela e preenche os resultados."""
    global hosts_list, current_file_path  # Acessa as variáveis globais

    if not hosts_list:
        messagebox.showwarning("Aviso", "Nenhum host encontrado na tabela!")
        return

    progress['value'] = 0
    progress['maximum'] = len(hosts_list)
    app.update_idletasks()

    inventory = read_inventory('inventário.csv')  # Altere para o caminho correto se necessário
    result_queue = queue.Queue()

    # Usando ThreadPoolExecutor para limitar o número de threads
    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        for host in hosts_list:  # Usa a lista de hosts armazenada
            executor.submit(analyze_host, host, inventory, result_queue)

    # Atualiza a interface enquanto as threads estão rodando
    def update_results():
        while not result_queue.empty():
            result = result_queue.get()

            # Atualiza a linha correspondente ao host
            for item in results_tree.get_children():
                if results_tree.item(item)['values'][0] == result[0]:  # Compara pelo host
                    results_tree.item(item, values=result)

                    # Variáveis para a lógica
                    SO = result[6]
                    RDP_Aberta = result[8]
                    SSH_Aberta = result[7]
                    PING = result[1] == 'True'
                    DNS_Reverso = result[2]
                    HOST_PINGANDO = result[0]

                    # Lógica para determinar se a linha deve ficar vermelha ou verde
                    C_Windows_Acessivel_Remotamente = ((SO == "Windows") and
                                                       (DNS_Reverso.lower().replace('.domain.biz',
                                                                                    '') == HOST_PINGANDO.lower())) and (
                                                              RDP_Aberta is True)
                    C_Windows_Sem_Acesso_Remoto = ((SO == "Windows") and
                                                   (DNS_Reverso.lower().replace('.domain.biz',
                                                                                '') == HOST_PINGANDO.lower())) and (
                                                          RDP_Aberta is False)

                    C_Linux_Acessivel_Remotamente = (SO == "Linux") and (
                            DNS_Reverso.lower().replace('x.domain.biz', '') == HOST_PINGANDO.lower()) and (
                                                            (RDP_Aberta is True) and (SSH_Aberta is True))
                    C_Linux_Sem_Acesso_Remotamente = (SO == "Linux") and (
                            DNS_Reverso.lower().replace('x.domain.biz', '') == HOST_PINGANDO.lower()) and (
                                                             (RDP_Aberta is False) and (SSH_Aberta is False))

                    Erro_DNS = (DNS_Reverso is not None) and not (
                            (DNS_Reverso.lower().replace('.domain.biz', '') == HOST_PINGANDO.lower()) or (
                            DNS_Reverso.lower().replace('x.domain.biz', '') == HOST_PINGANDO.lower()))

                    # Lógica para determinar se a linha deve ficar vermelha ou verde
                    if C_Windows_Acessivel_Remotamente:
                        results_tree.item(item, tags=('green',))  # Adiciona tag verde
                    elif C_Linux_Acessivel_Remotamente:
                        results_tree.item(item, tags=('green',))  # Adiciona tag verde
                    elif C_Windows_Sem_Acesso_Remoto:
                        results_tree.item(item, tags=('yellow',))  # Adiciona tag Amarelo
                    elif C_Linux_Sem_Acesso_Remotamente:
                        results_tree.item(item, tags=('yellow',))  # Adiciona tag Amarelo
                    elif Erro_DNS:
                        results_tree.item(item, tags=('orange',))  # Adiciona tag laranja
                    else:
                        results_tree.item(item, tags=('red',))  # Adiciona tag vermelha

                    # Debugging
                    print(
                        f"Host: {HOST_PINGANDO}, SO: {SO}, PING: {PING}, RDP: {RDP_Aberta}, SSH: {SSH_Aberta}, DNS: {DNS_Reverso}")

                    break

            progress['value'] += 1
            app.update_idletasks()

        if any(thread.is_alive() for thread in threading.enumerate() if thread != threading.main_thread()):
            app.after(100, update_results)  # Continua verificando
        else:
            app.after(100, update_results)  # Finaliza

    app.after(100, update_results)  # Inicia a atualização


def paste_and_analyze(event=None):
    """Pega o texto da área de transferência, limpa a tabela, preenche a tabela e inicia a análise."""
    global hosts_list  # Usar a lista global
    try:
        clipboard_content = app.clipboard_get()
        hosts = clipboard_content.strip().splitlines()

        if len(hosts) > MAX_HOSTS:
            messagebox.showwarning("Aviso", f"O número máximo de hosts permitidos é {MAX_HOSTS}.")
            return

        results_tree.delete(*results_tree.get_children())
        hosts_list = []  # Limpa a lista de hosts

        for host in hosts:
            if is_valid_host(host):
                results_tree.insert('', tk.END, values=(host, "", "", "", "", "", "", "", "", "", "", "", "", "", ""))
                hosts_list.append(host)  # Adiciona o host à lista

        app.after(1000, analyze_hosts)

    except tk.TclError:
        messagebox.showerror("Erro", "Não foi possível acessar a área de transferência.")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

def extract_report():
    """Extrai um relatório em HTML com o resumo dos hosts analisados."""
    hosts_data = []

    for item in results_tree.get_children():
        host_info = results_tree.item(item)['values']
        hosts_data.append({
            'host': host_info[0],
            'ping_result': host_info[3],
            'dns_reverse': host_info[2],
            'os_name': host_info[6],
            'localization': host_info[9],
            'building': host_info[10],
            'floor': host_info[11],
            'office': host_info[12],
            'obsolete': host_info[13],
            'annotation': host_info[14],
            'dns_reverse_issue': host_info[2] != host_info[0],  # Problema de DNS reverso
            'tags': results_tree.item(item)['tags']  # Pega as tags para categorizar
        })

    # Tabelas para cada categoria
    categories = {
        'Windows Acessível Remotamente': [],
        'Linux Acessível Remotamente': [],
        'Windows Sem Acesso Remoto': [],
        'Linux Sem Acesso Remoto': [],
        'Outros Erros': []
    }

    for host in hosts_data:
        if 'green' in host['tags']:
            if host['os_name'] == 'Windows':
                categories['Windows Acessível Remotamente'].append(host)
            elif host['os_name'] == 'Linux':
                categories['Linux Acessível Remotamente'].append(host)
        elif 'yellow' in host['tags']:
            if host['os_name'] == 'Windows':
                categories['Windows Sem Acesso Remoto'].append(host)
            elif host['os_name'] == 'Linux':
                categories['Linux Sem Acesso Remoto'].append(host)
        elif 'orange' in host['tags']:
            categories['Outros Erros'].append(host)
        elif 'red' in host['tags']:
            categories['Outros Erros'].append(host)

    # Criando o conteúdo HTML
    html_content = """
    <html>
    <head>
    <title>Relatório de Análise de Hosts</title>
    <style>
    body { font-family: Arial, sans-serif; }
    h1 { color: #007bff; }
    h2 { color: #343a40; }
    table { width: 100%; border-collapse: collapse; margin: 20px 0; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }
    th { background-color: #f2f2f2; }
    </style>
    </head>
    <body>
    <h1>Relatório de Análise de Hosts</h1>
    """

    for category, hosts in categories.items():
        html_content += f"<h2>{category}</h2>"
        if hosts:
            html_content += "<table><tr><th>Host</th><th>SO</th><th>DNS Reverso</th><th>Local</th><th>Anotação</th></tr>"
            for host in hosts:
                html_content += f"""
                <tr>
                <td>{host['host']}</td>
                <td>{host['os_name']}</td>
                <td>{host['dns_reverse']}</td>
                <td>{host['localization']}</td>
                <td>{host['annotation']}</td>
                </tr>
                """
            html_content += "</table>"
        else:
            html_content += "<p>Nenhum host encontrado nesta categoria.</p>"

    html_content += "</body></html>"

    # Salvando o conteúdo HTML em um arquivo
    report_file_path = "relatorio_hosts.html"
    with open(report_file_path, "w") as report_file:
        report_file.write(html_content)

    # Abrindo o arquivo HTML no navegador padrão
    webbrowser.open(report_file_path)
    messagebox.showinfo("Relatório Gerado", "O relatório foi gerado e aberto com sucesso: relatorio_hosts.html")

def show_credits():
    """Abre uma janela de créditos."""
    credits_window = tk.Toplevel(app)
    credits_window.title("Créditos")
    credits_window.geometry("400x300")
    credits_window.resizable(False, False)

    credits_frame = ttk.Frame(credits_window, padding="10")
    credits_frame.pack(expand=True, fill=tk.BOTH)

    title_label = tk.Label(credits_frame, text="Créditos", font=("Helvetica", 16, "bold"))
    title_label.pack(pady=(0, 10))

    system_name = "Nome do Sistema: HostFlow"
    description = "Descrição: Sistema para análise de conectividade."
    author = "Projetado por: Carlos Vilela"
    creation_date = "Data de Criação: 16/10/2025"
    version = "Versão: 0.0.1 (Versão alpha)"
    license_info = "Licença: MIT"

    tk.Label(credits_frame, text=system_name, font=("Helvetica", 10)).pack(anchor='w', padx=5, pady=2)
    tk.Label(credits_frame, text=description, font=("Helvetica", 10)).pack(anchor='w', padx=5, pady=2)
    tk.Label(credits_frame, text=author, font=("Helvetica", 10)).pack(anchor='w', padx=5, pady=2)
    tk.Label(credits_frame, text=creation_date, font=("Helvetica", 10)).pack(anchor='w', padx=5, pady=2)
    tk.Label(credits_frame, text=version, font=("Helvetica", 10)).pack(anchor='w', padx=5, pady=2)
    tk.Label(credits_frame, text=license_info, font=("Helvetica", 10)).pack(anchor='w', padx=5, pady=2)

    close_button = ttk.Button(credits_frame, text="Fechar", command=credits_window.destroy)
    close_button.pack(pady=(10, 0))

    credits_window.transient(app)
    credits_window.grab_set()
    credits_window.focus_set()
    credits_window.wait_window()

def open_rdp(event, host_pingando):
    """Abre a conexão RDP para o host selecionado."""
    if host_pingando:  # Verifica se há um host pingando
        rdp_command = f'mstsc /v:{host_pingando}'  # Comando para abrir o RDP
        try:
            subprocess.Popen(rdp_command)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a conexão RDP: {e}")
    else:
        messagebox.showwarning("Aviso", "Não há um host pingando disponível para conexão RDP.")

def open_file():
    """Abre um diálogo para selecionar um arquivo e carrega os hosts."""
    global hosts_list, current_file_path  # Usar as variáveis globais

    file_path = filedialog.askopenfilename(title="Abrir Arquivo de Hosts",
                                           filetypes=[("Text Files", "*.txt"), ("CSV Files", "*.csv")])
    if file_path:
        try:
            with open(file_path, 'r') as file:
                hosts = file.readlines()
                hosts = [host.strip() for host in hosts if is_valid_host(host.strip())]  # Filtra hosts válidos

            if len(hosts) > MAX_HOSTS:
                messagebox.showwarning("Aviso", f"O número máximo de hosts permitidos é {MAX_HOSTS}.")
                return

            results_tree.delete(*results_tree.get_children())  # Limpa a tabela
            hosts_list = []  # Limpa a lista de hosts

            for host in hosts:
                results_tree.insert('', tk.END,
                                    values=(host, "", "", "", "", "", "", "", "", "", "", "", "", "", ""))
                hosts_list.append(host)  # Adiciona o host à lista

            # Armazena o caminho do arquivo aberto e atualiza o título da janela
            current_file_path = file_path
            app.title(f"HostFlow - {os.path.splitext(os.path.basename(current_file_path))[0]}")  # Atualiza o título

            app.after(1000, analyze_hosts)  # Inicia a análise após um segundo

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao abrir o arquivo: {str(e)}")

def save_hosts():
    """Salva os hosts atuais no arquivo atual, se houver um arquivo aberto."""
    global current_file_path

    if current_file_path:  # Verifica se há um arquivo atual
        try:
            with open(current_file_path, 'w', newline='', encoding='utf-8') as file:
                writer = csv.writer(file, delimiter=';')
                # Escreve os hosts na tabela
                for item in results_tree.get_children():
                    host_info = results_tree.item(item)['values']
                    writer.writerow([host_info[0]])  # Salva apenas o host

            messagebox.showinfo("Sucesso", f"Hosts salvos com sucesso em: {current_file_path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao salvar os hosts: {str(e)}")
    else:
        save_hosts_as()  # Caso não haja arquivo atual, chama a função para salvar como


def save_hosts_as():
    """Salva os hosts atuais em um arquivo, permitindo ao usuário escolher o nome e o local."""
    global current_file_path  # Usar a variável global para o caminho do arquivo atual

    file_path = filedialog.asksaveasfilename(defaultextension=".csv",
                                             filetypes=[("CSV Files", "*.csv"), ("Text Files", "*.txt")])
    if file_path:  # Verifica se o usuário selecionou um arquivo
        current_file_path = file_path  # Atualiza o caminho do arquivo atual
        save_hosts()  # Chama a função para salvar os hosts no novo arquivo


def open_ssh(host):
    """Abre uma conexão SSH para o host selecionado após solicitar a chave de acesso."""
    ssh_key = simpledialog.askstring("Chave de Acesso", "Digite a chave de acesso:")
    if ssh_key:
        cmd_command = f"start cmd /k ssh {ssh_key}@{host}"
        try:
            subprocess.Popen(cmd_command, shell=True)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir a conexão SSH: {e}")


def edit_host():
    """Edita o host selecionado."""
    selected_items = results_tree.selection()  # Obtém os itens selecionados

    if selected_items:  # Verifica se há itens selecionados
        selected_item = selected_items[0]  # Pega o primeiro item selecionado
        current_host = results_tree.item(selected_item)['values'][0]  # Obtém o host atual

        # Abre uma caixa de diálogo para edição
        new_host = simpledialog.askstring("Modificar Host", "Edite o host:", initialvalue=current_host)

        if new_host and is_valid_host(new_host):  # Verifica se o novo host é válido
            # Atualiza o host na tabela
            new_values = (new_host,) + tuple(
                results_tree.item(selected_item)['values'][1:])  # Converte a lista em tupla
            results_tree.item(selected_item, values=new_values)

            # Atualiza a lista de hosts
            global hosts_list
            hosts_list = [new_host if host == current_host else host for host in hosts_list]
        else:
            messagebox.showerror("Erro", "Host inválido. Por favor, insira um host válido.")


def show_context_menu(event):
    """Mostra o menu de contexto."""
    context_menu = tk.Menu(app, tearoff=0)

    # Obtém a posição do mouse na árvore
    item = results_tree.identify_row(event.y)

    if item:
        # Adiciona a opção de copiar ao menu de contexto
        context_menu.add_command(label="Copiar",
                                 command=lambda: copy_to_clipboard(results_tree.item(item)['values'][0]))

        # Adiciona submenu "Editar"
        edit_menu = tk.Menu(context_menu, tearoff=0)
        edit_menu.add_command(label="Modificar", command=edit_host)  # Adiciona a opção "Modificar"
        context_menu.add_cascade(label="Editar", menu=edit_menu)

        # Adiciona submenu "Conexão"
        connection_menu = tk.Menu(context_menu, tearoff=0)
        connection_menu.add_command(label="Acesso via RDP",
                                    command=lambda: open_rdp(None, results_tree.item(item)['values'][1]))
        connection_menu.add_command(label="Acesso via SSH",
                                    command=lambda: open_ssh(results_tree.item(item)['values'][1]))

        context_menu.add_cascade(label="Conexão", menu=connection_menu)

        # Adiciona opção de remover a linha
        context_menu.add_command(label="Remover", command=lambda: remove_selected_row(None))

        context_menu.post(event.x_root, event.y_root)


def copy_to_clipboard(text):
    """Copia o texto para a área de transferência."""
    app.clipboard_clear()  # Limpa a área de transferência
    app.clipboard_append(text)  # Adiciona o novo texto


def remove_selected_row(event=None):
    """Remove as linhas selecionadas da tabela."""
    selected_items = results_tree.selection()  # Obtém os itens selecionados

    if selected_items:  # Verifica se há itens selecionados
        for selected_item in selected_items:
            item_value = results_tree.item(selected_item)['values'][0] if selected_item else None
            if item_value:  # Verifica se o valor do item existe
                results_tree.delete(selected_item)  # Remove o item da tabela

        # Também remove o host da lista global
        global hosts_list
        hosts_list = [host for host in hosts_list if
                      host not in [results_tree.item(item)['values'][0] for item in selected_items]]


# Variáveis globais para armazenar informações de arrastar e soltar
dragged_item = None
dragged_item_index = None


def on_tree_select(event):
    """Captura o item selecionado."""
    global dragged_item, dragged_item_index
    selected_item = results_tree.selection()
    if selected_item:
        dragged_item = selected_item[0]  # Pega o ID do item arrastado
        dragged_item_index = results_tree.index(dragged_item)

def on_tree_drag(event):
    """Atualiza a posição do item arrastado durante o movimento do mouse."""
    global dragged_item, dragged_item_index
    if dragged_item and dragged_item_index is not None:
        # Obter a posição do mouse
        y = event.y
        # Obter o item que está sob o mouse
        item = results_tree.identify_row(y)
        if item and item != dragged_item:
            # Mover o item para a nova posição
            current_index = results_tree.index(item)
            if current_index != dragged_item_index:
                # Mover o item arrastado para a nova posição
                results_tree.move(dragged_item, '', current_index)
                # Atualizar o índice do item arrastado
                dragged_item_index = current_index


def on_tree_release(event):
    """Solta o item arrastado na nova posição."""
    global dragged_item
    dragged_item = None  # Limpa a referência do item arrastado



def organize_by_color():
    """Organiza os itens da tabela com base nas cores das tags."""
    items = results_tree.get_children()

    colored_items = {'red': [], 'green': [], 'yellow': [], 'orange': []}
    item_values = {}  # Dicionário para armazenar valores do item
    item_tags = {}  # Dicionário para armazenar tags do item

    # Agrupar itens por cor e armazenar os valores e tags
    for item in items:
        tags = results_tree.item(item)['tags']
        item_values[item] = results_tree.item(item)['values']  # Armazena os valores antes da deleção
        item_tags[item] = tags  # Armazena as tags antes da deleção
        if 'red' in tags:
            colored_items['red'].append(item)
        elif 'green' in tags:
            colored_items['green'].append(item)
        elif 'yellow' in tags:
            colored_items['yellow'].append(item)
        elif 'orange' in tags:
            colored_items['orange'].append(item)

    # Limpar a tabela
    results_tree.delete(*items)

    # Reorganizar os itens na ordem desejada
    for color in ['green', 'yellow', 'orange', 'red']:  # Ajuste a ordem conforme necessário
        for item in colored_items[color]:
            # Reinsere os itens usando os valores armazenados no dicionário
            new_item = results_tree.insert('', 'end', values=item_values[item])
            # Reaplica a tag de cor ao item
            for tag in item_tags[item]:
                results_tree.item(new_item, tags=(tag,))

# Lembre-se de adicionar as tags de cor novamente após inserir os itens.



def show_quantitative_report():
    """Mostra o quantitativo de hosts por categoria."""
    categories_count = {
        'Acessível Remotamente': 0,
        'Sem Acesso Remoto': 0,
        'Erro DNS': 0,
        'Outros Erros': 0
    }

    for item in results_tree.get_children():
        tags = results_tree.item(item)['tags']
        if 'green' in tags:
            categories_count['Acessível Remotamente'] += 1
        elif 'yellow' in tags:
            categories_count['Sem Acesso Remoto'] += 1
        elif 'orange' in tags or 'red' in tags:
            categories_count['Outros Erros'] += 1
        else:
            categories_count['Erro DNS'] += 1 # Você pode ajustar essa lógica conforme necessário

    report_message = "\n".join(f"{category}: {count}" for category, count in categories_count.items())
    messagebox.showinfo("Quantitativo de Hosts", report_message)

# Função para chamar o script ping.py
def run_ping_script():
    try:
        subprocess.Popen(['python', 'ping.py'])  # Supondo que ping.py esteja no mesmo diretório
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível executar o teste de ping: {e}")



# Criação da janela principal
app = tk.Tk()
app.title("HostFlow")
app.geometry("1000x600")

# Criação da barra de menu
menu_bar = tk.Menu(app)

# Menu Arquivo
file_menu = tk.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="Abrir", command=open_file)
file_menu.add_command(label="Salvar Hosts", command=save_hosts)  # Adiciona opção para salvar hosts
file_menu.add_command(label="Salvar Como", command=save_hosts_as)  # Adiciona a opção "Salvar Como"
menu_bar.add_cascade(label="Arquivo", menu=file_menu)

# Adicionando o menu "Teste de Rede"
network_test_menu = tk.Menu(menu_bar, tearoff=0)
# Adicionando a opção "Teste de Ping"
network_test_menu.add_command(label="Teste de Ping", command=run_ping_script)
# Adicionando o menu "Teste de Rede" ao menu principal
menu_bar.add_cascade(label="Teste de Rede", menu=network_test_menu)

# Adicionando o menu "Organizar"
organize_menu = tk.Menu(menu_bar, tearoff=0)
organize_menu.add_command(label="Organizar por Cor", command=organize_by_color)
menu_bar.add_cascade(label="Organizar", menu=organize_menu)


# Menu de Análise
analysis_menu = tk.Menu(menu_bar, tearoff=0)
analysis_menu.add_command(label="Quantitativo", command=show_quantitative_report)
menu_bar.add_cascade(label="Análise", menu=analysis_menu)




# Menu de Procedimentos
procedures_menu = tk.Menu(menu_bar, tearoff=0)
problems_menu = tk.Menu(procedures_menu, tearoff=0)

# Submenu "Problemas de IP"
problems_menu.add_command(label="Devolver e Reservar IP", command=lambda: webbrowser.open(
    "https://domain.biz.service-now.com/cs?id=sc_cat_item&sys_id=00000000000000000000000000000000"))
procedures_menu.add_cascade(label="Problemas de IP", menu=problems_menu)
menu_bar.add_cascade(label="Procedimentos", menu=procedures_menu)

# Menu de Relatórios
report_menu = tk.Menu(menu_bar, tearoff=0)
report_menu.add_command(label="Extrair Relatório", command=extract_report)
menu_bar.add_cascade(label="Relatórios", menu=report_menu)

# Adicionando o menu "Ajuda"
help_menu = tk.Menu(menu_bar, tearoff=0)
help_menu.add_command(label="Sobre", command=show_credits)
menu_bar.add_cascade(label="Ajuda", menu=help_menu)

# Configura a barra de menu na janela principal
app.config(menu=menu_bar)

# Frame para a tabela e a barra de rolagem
frame = ttk.Frame(app)
frame.pack(pady=5, fill=tk.BOTH, expand=True)

# Tabela para resultados
columns = ("Host", "Host Pingando", "DNS Reverso", "Ping", "IP", "TTL", "SO", "SSH Aberta", "RDP Aberta",
           "Local", "Prédio", "Andar", "Escritório", "Obsoleto", "Anotação")
results_tree = ttk.Treeview(frame, columns=columns, show='headings')

for col in columns:
    results_tree.heading(col, text=col)

# Define uma largura fixa para todas as colunas
fixed_width = 100  # Ajuste este valor conforme necessário

for col in columns:
    results_tree.column(col, width=fixed_width)

# Barra de rolagem vertical
scrollbar_vertical = ttk.Scrollbar(frame, orient="vertical", command=results_tree.yview)
results_tree.configure(yscrollcommand=scrollbar_vertical.set)
scrollbar_vertical.pack(side='right', fill='y')

# Barra de rolagem horizontal
scrollbar_horizontal = ttk.Scrollbar(frame, orient="horizontal", command=results_tree.xview)
results_tree.configure(xscrollcommand=scrollbar_horizontal.set)
scrollbar_horizontal.pack(side='bottom', fill='x')

results_tree.pack(pady=5, fill=tk.BOTH, expand=True)

# Barra de progresso
progress = ttk.Progressbar(app, orient="horizontal", length=400, mode="determinate")
progress.pack(pady=10)

# Bind dos eventos de arrastar e soltar
results_tree.bind('<ButtonPress-1>', on_tree_select)  # Evento de clique para começar a arrastar
results_tree.bind('<B1-Motion>', on_tree_drag)  # Evento de movimento do mouse para arrastar
results_tree.bind('<ButtonRelease-1>', on_tree_release)  # Evento de soltura

# Adicionando legenda de cores com checkboxes
legend_frame = ttk.Frame(app)
legend_frame.pack(pady=5)

# Dicionário para armazenar variáveis de controle dos checkboxes
checkbox_vars = {
    'green': tk.BooleanVar(value=True),
    'yellow': tk.BooleanVar(value=True),
    'orange': tk.BooleanVar(value=True),
    'red': tk.BooleanVar(value=True)
}



def update_table_visibility():
    """Atualiza a visibilidade dos hosts na tabela com base nas cores selecionadas."""
    for item in results_tree.get_children():
        host_info = results_tree.item(item)['tags']
        # Verifica se o host deve ser exibido com base nas checkboxes
        if any(checkbox_vars[color].get() and color in host_info for color in checkbox_vars):
            results_tree.item(item, open=True)  # Exibe o item
        else:
            results_tree.detach(item)  # Oculta o item


def create_legend_item_with_checkbox(color, text):
    """Cria um item da legenda com um checkbox para controle de visibilidade."""
    color_box = tk.Canvas(legend_frame, width=20, height=20, bg=color)
    color_box.pack(side=tk.LEFT, padx=5)
    checkbox = ttk.Checkbutton(legend_frame, text=text, variable=checkbox_vars[color],
                               command=update_table_visibility)
    checkbox.pack(side=tk.LEFT)


# Adicionando itens à legenda com checkbox
create_legend_item_with_checkbox('green', 'Acessível Remotamente')
create_legend_item_with_checkbox('yellow', 'Sem Acesso Remoto')
create_legend_item_with_checkbox('orange', 'Erro DNS')
create_legend_item_with_checkbox('red', 'Outros Erros')

# Chama para atualizar a visibilidade inicialmente
update_table_visibility()

# Botões
button_frame = ttk.Frame(app)
button_frame.pack(pady=10)

analyze_button = ttk.Button(button_frame, text="Analisar Hosts", command=analyze_hosts)
analyze_button.pack(side=tk.LEFT, padx=5)

# Bind Ctrl+V para colar e analisar
app.bind('<Control-v>', paste_and_analyze)

# Bind do evento de duplo clique na árvore para abrir RDP
results_tree.bind('<Double-1>', lambda event: open_rdp(event, results_tree.item(results_tree.selection())['values'][1]))

# Bind do botão direito do mouse para o menu de contexto
results_tree.bind('<Button-3>', show_context_menu)

# Bind da tecla Delete para remover a linha selecionada
app.bind('<Delete>', remove_selected_row)

# Adiciona as tags de cor
results_tree.tag_configure('red', background='red')
results_tree.tag_configure('green', background='green')
results_tree.tag_configure('yellow', background='yellow')
results_tree.tag_configure('orange', background='orange')

# Inicia a aplicação
app.mainloop()
