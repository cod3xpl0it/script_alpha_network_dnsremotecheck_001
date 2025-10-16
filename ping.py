import tkinter as tk
import subprocess
import time
import threading
import platform


class PingApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Ping Monitor")

        self.host_label = tk.Label(master, text="Host/IP:")
        self.host_label.pack()

        self.host_entry = tk.Entry(master)
        self.host_entry.pack()

        self.start_stop_button = tk.Button(master, text="Iniciar Ping", command=self.start_stop_ping)
        self.start_stop_button.pack()

        self.canvas = tk.Canvas(master, width=200, height=200)
        self.canvas.pack()

        self.log_text = tk.Text(master, height=10, width=30)
        self.log_text.pack()

        self.status_color = "gray"
        self.timeout = 5  # Tempo limite em segundos
        self.ping_thread = None
        self.running = False

        self.update_circle()

    def update_circle(self):
        self.canvas.delete("all")
        self.canvas.create_oval(50, 50, 150, 150, fill=self.status_color)
        self.master.update()

    def start_stop_ping(self):
        if self.running:
            # Parar o ping
            self.running = False
            self.start_stop_button.config(text="Iniciar Ping")
        else:
            # Iniciar o ping
            host = self.host_entry.get()
            if host:
                if self.ping_thread is not None and self.ping_thread.is_alive():
                    self.running = False
                    self.ping_thread.join()
                self.running = True
                self.log_text.delete(1.0, tk.END)  # Limpa o log antes de iniciar
                self.start_stop_button.config(text="Parar Ping")
                self.start_ping(host)

    def start_ping(self, host):
        def ping():
            last_status = True
            while self.running:
                # Comando de ping baseado no sistema operacional
                command = ["ping", "-c", "1", host] if platform.system().lower() != "windows" else ["ping", "-n", "1",
                                                                                                    host]

                response = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                current_status = response.returncode == 0

                # Atualiza a cor do círculo e log com base no status
                if current_status:
                    self.status_color = "green"
                    log_message = f"{host} está respondendo.\n"
                    last_status = True
                else:
                    if last_status:
                        self.status_color = "orange"
                        log_message = f"{host} não está respondendo. Aguardando...\n"
                        time.sleep(self.timeout)
                        if not current_status:
                            self.status_color = "red"
                            log_message = f"{host} não está respondendo. Tempo limite excedido.\n"
                    else:
                        self.status_color = "red"
                        log_message = f"{host} não está respondendo.\n"

                # Atualiza a interface gráfica
                self.update_circle()
                self.log_text.insert(tk.END, log_message)
                self.log_text.see(tk.END)  # Rola para o final do log
                last_status = current_status
                time.sleep(1)  # Intervalo entre pings

        self.ping_thread = threading.Thread(target=ping)
        self.ping_thread.start()

    def on_closing(self):
        self.running = False
        if self.ping_thread is not None:
            self.ping_thread.join()
        self.master.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = PingApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()
