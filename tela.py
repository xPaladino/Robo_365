import tkinter as tk
from tkinter import ttk, filedialog
from tkcalendar import Calendar
from tkinter import messagebox
from PIL import Image, ImageTk, ImageSequence
import threading
from final.back import processamento
from datetime import datetime
from final.back.SQL.mapa_estoque import start_mapa
import random

import time


def browse_folder(entry):
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        entry.delete(0, tk.END)
        entry.insert(tk.END, folder_selected)
    else:
        messagebox.showwarning("Atenção", "Por favor, escolha um diretório para continuar.")


def start_processing(folder_entry, date_entry, date_entry_end):
    folder_path = folder_entry.get()
    search_date = date_entry.get()
    search_date_end = date_entry_end.get()
    escolha_inicio = None
    escolha_fim = None

    if not search_date or not search_date_end:
        messagebox.showwarning("Atenção", "Por favor, escolha um período para continuar.")
        return

    if search_date:
        ajusta = datetime.strptime(search_date, "%d/%m/%Y")
        formata = ajusta.strftime("%Y-%m-%d")
        escolha_inicio = formata
    search_date = escolha_inicio

    if search_date_end:
        ajusta = datetime.strptime(search_date_end, "%d/%m/%Y")
        formata = ajusta.strftime("%Y-%m-%d")
        escolha_fim = formata
    search_date_end = escolha_fim

    if folder_path:
        processing_window = tk.Toplevel()
        processing_window.geometry("300x300")
        processing_window.iconbitmap('sonic-icon.ico')
        processing_window.title("AGUARDE")
        processing_window.grab_set()
        processing_window.protocol("WM_DELETE_WINDOW", lambda: None)
        gif_label = tk.Label(processing_window)
        gif_label.pack(padx=30, pady=30)

        messages = [
            "Aguarde um pouco, pegue um café.\n Ou dois.",
            "Estamos quase lá, mais um instante por favor.\n Ou dois. Ou três.",
            "Só um momento, estamos trabalhando nisso.\n Literalmente.",
            "Aguarde, estamos processando seus dados.\n É verdade!",
            "Estamos a todo vapor! Mais alguns segundos.\n Prometemos.",
            "Obrigado pela paciência,\n só mais um pouquinho.",
            "Quase pronto, estamos finalizando.\n Não, sério!",
            "Um segundo, por favor.\n O hamster está correndo na roda.",
            "Aguarde um pouco.\n A NASA também teve que esperar.",
            "Um momento.\n Prometo.",
            "Estamos no caminho certo.\n Só perdemos o mapa.",
            "Relaxe, e se isso não ajudar,\n um café pode ajudar!",
            "Estamos trabalhando duro.\n Ou fingindo muito bem.",
            "Isso pode demorar um pouco.\n Afinal, boa arte leva tempo.",
            "Quase lá. Por favor,\n mantenha as mãos e pés dentro do veículo.",
            "Seu arquivo está quase pronto.\n É uma promessa de programador!",
            "Enquanto aguarda, que tal um café?\n Pegue um pra mim também!"
        ]

        processing_label = tk.Label(processing_window, text=random.choice(messages))
        processing_label.pack(padx=20, pady=20)
        time_label = tk.Label(processing_window, text="Tempo decorrido: 0.00")
        time_label.pack()
        processing_thread = threading.Thread(target=process_and_show_message,
                                             args=(
                                                 folder_path, search_date, search_date_end, processing_window,
                                                 processing_label, gif_label, time_label, messages))
        processing_thread.start()
    else:
        messagebox.showwarning("Atenção", "Por favor, escolha um diretório antes de continuar.")


def process_and_show_message(folder_path, selected_date, selected_date_end, processing_window, processing_label,
                             gif_label, time_label, messages):

    # Carregar e exibir o GIF animado
    tempo_inicial = time.time()
    success = [False]
    gif = Image.open("sonicgif2.gif")

    gif_frames = [ImageTk.PhotoImage(frame) for frame in ImageSequence.Iterator(gif)]

    def update_gif_frame(frame_num):
        if processing_window.winfo_exists():
            gif_label.config(image=gif_frames[frame_num])
            processing_window.after(100, update_gif_frame, (frame_num + 1) % len(gif_frames))

    update_gif_frame(0)

    def update_timer():
        if processing_window.winfo_exists() and not success[0]:
            tempo = time.time() - tempo_inicial
            min = int(tempo) // 60
            seg = int(tempo) % 60
            time_label.config(text=f"Tempo decorrido: {min:02}:{seg:02}")

            processing_window.after(1000, update_timer)

    update_timer()

    def update_mensagem():
        if processing_window.winfo_exists() and not success[0]:
            processing_label.config(text=random.choice(messages))
            processing_window.after(6000, update_mensagem)

    update_mensagem()

    mensagens = []

    try:
        nf_pdf_count, resultados_pdf_count, nf_zip_count, resultados_zip_count, nf_excel_count, resultados_excel_count \
            = processamento.main(folder_path, selected_date, selected_date_end)

        if nf_pdf_count != 0:
            mensagens.append(f"PDF's lidos: {nf_pdf_count} - Processados: {resultados_pdf_count}")

        if nf_zip_count != 0:
            mensagens.append(f"ZIP's lidos: {nf_zip_count} - Processados: {resultados_zip_count}")

        if nf_excel_count != 0:
            mensagens.append(f"Excel's lidos: {nf_excel_count} - Processados: {resultados_excel_count}")

        if mensagens:
            success[0] = True
            messagebox.showinfo("Concluído", "Processamento concluído com sucesso!\n" + "\n".join(mensagens))

        else:
            success[0] = True
            messagebox.showwarning("Atenção", f"Não encontrei dados entre {selected_date} e {selected_date_end}.\n"
                                              f"Por favor, verifique o periodo e tente novamente.")

        processing_window.destroy()
    except ValueError as e:
        success[0] = True
        messagebox.showwarning("Atenção", "Arquivo excel em aberto.\nPor favor, feche e tente novamente.")
        processing_window.destroy()


def browse_folder_sql(entry):
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        entry.delete(0, tk.END)
        entry.insert(tk.END, folder_selected)
    else:
        messagebox.showwarning("Atenção", "Por favor, escolha um diretório para continuar.")


def get_date_selected_sql(cal, cal_window, date_entry):
    selected_date_str = cal.get_date()

    if selected_date_str:
        selected_date = datetime.strptime(selected_date_str, "%d/%m/%Y")
        formatted_date = selected_date.strftime("%Y-%m-%d")
        date_entry.delete(0, tk.END)
        date_entry.insert(tk.END, formatted_date)
        cal_window.destroy()
    else:
        messagebox.showwarning("Atenção", "Por favor, escolha um período para continuar.")


def open_calendar_sql(sql_date_entry):
    cal_window = tk.Toplevel()
    cal_window.title("Calendário")
    cal_window.iconbitmap('sonic-icon.ico')

    cal = Calendar(cal_window, selectmode="day", date_pattern="dd/MM/yyyy", locale="en_US")
    cal.pack(padx=10, pady=10)

    def confirm_date():
        get_date_selected_sql(cal, cal_window, sql_date_entry)
        cal_window.destroy()

    confirm_button = tk.Button(cal_window, text="Confirmar",
                               command=confirm_date)
    confirm_button.pack(pady=5)
    cal_window.wait_window()


## Mapa de Estoque


def browse_folder_mapa(entry):
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        entry.delete(0, tk.END)
        entry.insert(tk.END, folder_selected)
    else:
        messagebox.showwarning("Atenção", "Nenhum diretório selecionado")


def get_date_selected_mapa(cal, cal_window, date_entry):
    selected_date_str = cal.get_date()

    if selected_date_str:
        selected_date = datetime.strptime(selected_date_str, "%d/%m/%Y")
        formatted_date = selected_date.strftime("%Y-%m-%d 00:00:00")
        date_entry.delete(0, tk.END)
        date_entry.insert(tk.END, formatted_date)
        cal_window.destroy()
    else:
        messagebox.showwarning("Atenção", "Nenhuma data selecionada")


def process_and_show_message_mapa(folder_path, selected_date, selected_date_end, processing_window, gif_label,
                                  time_label):
    tempo_inicial = time.time()
    success = False
    gif = Image.open("sonicgif2.gif")
    gif_frames = [ImageTk.PhotoImage(frame) for frame in ImageSequence.Iterator(gif)]

    def update_gif_frame(frame_num):
        if processing_window.winfo_exists():
            gif_label.config(image=gif_frames[frame_num])
            processing_window.after(100, update_gif_frame, (frame_num + 1) % len(gif_frames))

    update_gif_frame(0)

    def update_timer():
        if processing_window.winfo_exists():
            tempo_decorrido = time.time() - tempo_inicial
            minutos = int(tempo_decorrido) // 60
            segundos = int(tempo_decorrido) % 60
            time_label.config(text=f"Tempo decorrido: {minutos:02}:{segundos:02}")
            if not success:
                processing_window.after(1000, update_timer)  # Chama update_timer novamente após 1 segundo

    update_timer()
    # Chamar a função principal de processamento
    success = start_mapa(folder_path, selected_date, selected_date_end)
    print(start_mapa(folder_path, selected_date, selected_date_end))
    # Exibir mensagem de conclusão do processamento

    if success:
        # messagebox.showinfo("Concluído", "Processamento concluído com sucesso!")
        messagebox.showinfo("Concluído", f"Dados salvos em: {folder_path}\n")

    else:
        messagebox.showerror("Erro", "Ocorreu um erro durante o processamento!")

    processing_window.destroy()
    # Fechar a janela de processamento


def start_mapa_query(folder_entry_mapa, mapa_date_entry, mapa_date_entry_end):
    folder_path = folder_entry_mapa.get()
    search_date = mapa_date_entry.get()
    search_date_end = mapa_date_entry_end.get()
    if not search_date or not search_date_end:
        messagebox.showwarning("Atenção", "Escolha uma data antes de iniciar")
        return

    if not search_date:
        messagebox.showwarning("Atenção", "Escolha uma data antes de iniciar")
        return

    if folder_path:

        processing_window = tk.Toplevel()
        processing_window.iconbitmap('sonic-icon.ico')
        processing_window.title("AGUARDE")
        processing_window.grab_set()
        processing_window.protocol("WM_DELETE_WINDOW", lambda: None)
        gif_label = tk.Label(processing_window)
        gif_label.pack(padx=30, pady=30)
        time_label = tk.Label(processing_window, text="Tempo decorrido: 0.00")
        time_label.pack()
        processing_label = tk.Label(processing_window, text="Processando... Por favor, aguarde.")
        processing_label.pack(padx=20, pady=20)

        processing_thread = threading.Thread(target=process_and_show_message_mapa,
                                             args=(
                                                 folder_path, search_date, search_date_end, processing_window,
                                                 gif_label, time_label))
        processing_thread.start()
    else:
        messagebox.showerror("Erro", "Nenhum diretório selecionado")


def main():
    root = tk.Tk()
    root.title("Nota Complementar - Leitor de E-mails")
    root.iconbitmap('sonic-icon.ico')

    notebook = ttk.Notebook(root)
    notebook.pack(fill='both', expand=True)

    # Aba para processamento de e-mails
    frame_email = tk.Frame(notebook)
    notebook.add(frame_email, text='Processamento de E-mails')

    def handle_keypress(event, entry_widget):
        # responsavel por alterar o campo para dia/mes/ano hora/minuto/segundo
        if event.char.isdigit() or event.keysym == 'BackSpace' or event.keysym == 'Tab':  # or event.char in ["/", ":", " "]:
            resultado = entry_widget.get()
            if len(resultado) == 2 and event.char.isdigit():
                entry_widget.insert(tk.END, '/')
            elif len(resultado) == 5 and event.char.isdigit():
                entry_widget.insert(tk.END, '/')
            elif len(resultado) == 10 and event.char.isdigit():
                entry_widget.insert(tk.END, ' ')
            elif len(resultado) == 13 and event.char.isdigit():
                entry_widget.insert(tk.END, ':')
            elif len(resultado) == 16 and event.char.isdigit():
                entry_widget.insert(tk.END, ':')
            elif len(resultado) >= 18 and event.char.isdigit():
                entry_widget.delete(18, tk.END)
            return True
        return "break"

    def handle_keypress_nohour(event, entry_widget):
        # responsavel por alterar o campo para dia/mes/ano
        if event.char.isdigit() or event.keysym == 'BackSpace' or event.keysym == 'Tab':  # or event.char in ["/", ":", " "]:
            resultado = entry_widget.get()
            if len(resultado) == 2 and event.char.isdigit():
                entry_widget.insert(tk.END, '/')
            elif len(resultado) == 5 and event.char.isdigit():
                entry_widget.insert(tk.END, '/')
            elif len(resultado) >= 9 and event.char.isdigit():
                entry_widget.delete(9, tk.END)
            return True
        return "break"

    def handle_nokeypress(event):
        # responsavel por não permitir digitação
        if event.keysym == 'BackSpace' or event.keysym == 'Tab':
            # resultado = entry_widget.get()
            return True
        return "break"

    data_label = tk.Label(frame_email, text="Periodo de")
    data_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")

    date_entry = tk.Entry(frame_email, width=20)
    date_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    date_entry.config(validate="key")
    date_entry.bind("<Key>", lambda event: handle_keypress_nohour(event, date_entry))

    data_label_end = tk.Label(frame_email, text="até")
    data_label_end.grid(row=1, column=0, padx=5, pady=5, sticky="e")

    date_entry_end = tk.Entry(frame_email)
    date_entry_end.grid(row=1, column=1, padx=5, pady=5, sticky="w")

    date_entry_end.config(validate="key")
    date_entry_end.bind("<Key>", lambda event: handle_keypress_nohour(event, date_entry_end))

    folder_label = tk.Label(frame_email, text="Escolha o diretório:")
    folder_label.grid(row=2, column=0, padx=5, pady=5)

    folder_entry = tk.Entry(frame_email, width=50)
    folder_entry.grid(row=2, column=1, padx=5, pady=5)

    folder_entry.config(validate="key")
    folder_entry.bind("<Key>", lambda event: handle_nokeypress(event))

    browse_button = tk.Button(frame_email, text="Procurar", command=lambda: browse_folder(folder_entry))
    browse_button.grid(row=2, column=2, padx=5, pady=5)

    start_button = tk.Button(frame_email, text="Começar o processamento",
                             command=lambda: start_processing(folder_entry, date_entry, date_entry_end))
    start_button.grid(row=3, columnspan=3, padx=5, pady=5)

    # MAPA DE ESTOQUE
    frame_mapa = tk.Frame(notebook)
    notebook.add(frame_mapa, text='Mapa de Estoque')

    mapa_date_label = tk.Label(frame_mapa, text="Periodo de")
    mapa_date_label.grid(row=0, column=0, padx=5, pady=5, sticky="e")

    mapa_date_entry = tk.Entry(frame_mapa, width=20)
    mapa_date_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    mapa_date_entry.config(validate="key")
    mapa_date_entry.bind("<Key>", lambda event: handle_keypress(event, mapa_date_entry))

    mapa_data_label_end = tk.Label(frame_mapa, text="até")
    mapa_data_label_end.grid(row=1, column=0, padx=5, pady=5, sticky="e")

    mapa_date_entry_end = tk.Entry(frame_mapa)
    mapa_date_entry_end.grid(row=1, column=1, padx=5, pady=5, sticky="w")

    mapa_date_entry_end.config(validate="key")
    mapa_date_entry_end.bind("<Key>", lambda event: handle_keypress(event, mapa_date_entry_end))

    folder_label = tk.Label(frame_mapa, text="Escolha o diretório:")
    folder_label.grid(row=2, column=0, padx=5, pady=5)

    folder_entry_mapa = tk.Entry(frame_mapa, width=50)
    folder_entry_mapa.grid(row=2, column=1, padx=5, pady=5)

    browse_button = tk.Button(frame_mapa, text="Procurar", command=lambda: browse_folder_mapa(folder_entry_mapa))
    browse_button.grid(row=2, column=2, padx=5, pady=5)

    mapa_start_button = tk.Button(frame_mapa, text="Busca Mapa de Estoque",
                                  command=lambda: start_mapa_query(folder_entry_mapa, mapa_date_entry,
                                                                   mapa_date_entry_end))
    mapa_start_button.grid(row=3, columnspan=3, padx=5, pady=5)

    root.protocol("WM_DELETE_WINDOW", root.quit)
    root.mainloop()


if __name__ == "__main__":
    main()
