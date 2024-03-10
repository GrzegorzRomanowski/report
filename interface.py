import sys
import tkinter as tk


def report_welcome_message_box() -> dict:
    """ Tkinter welcome window that asks for the date and number of the production line to be reported.
    :return: Dictionary with data taken from user.
    """
    ask_window = tk.Tk()
    ask_window.title("inBet  Report")
    ask_window.geometry("250x195")

    ask_label = tk.Label(ask_window, text="Za jaki dzień zrobić raport?\nZ której linii produkcyjnej?")
    ask_label.pack(pady=5, padx=5)

    ask_frame = tk.Frame(ask_window)
    ask_frame.pack(pady=5, padx=5)

    ask_day = tk.Entry(ask_frame, width=6)
    ask_day.grid(pady=2, padx=2, row=0, column=0)
    ask_day.insert(0, "dd")

    ask_month = tk.Entry(ask_frame, width=6)
    ask_month.grid(pady=2, padx=2, row=0, column=1)
    ask_month.insert(0, "mm")

    ask_year = tk.Entry(ask_frame, width=10)
    ask_year.grid(pady=2, padx=2, row=0, column=2)
    ask_year.insert(0, "yyyy")

    ebawe_label = tk.Label(ask_frame, text="EBAWE: ")
    ebawe_label.grid(pady=18, padx=2, row=1, columnspan=2, column=0)

    ask_ebawe = tk.Entry(ask_frame, width=10)
    ask_ebawe.grid(pady=18, padx=2, row=1, column=2)
    ask_ebawe.insert(0, "1")

    output_dict = {
        'dd': str(),
        'mm': str(),
        'yyyy': str(),
        'ebawe': str()
    }

    def confirm_button(out_dict: dict = output_dict):
        dd = ask_day.get()
        mm = ask_month.get()
        yyyy = ask_year.get()
        ebawe = ask_ebawe.get()
        out_dict['dd'] = dd
        out_dict['mm'] = mm
        out_dict['yyyy'] = yyyy
        out_dict['ebawe'] = ebawe
        if dd == "" or mm == "" or yyyy == "" or ebawe == "":
            warning_msg1 = "Nie podano wszystkich danych w oknie startowym!"
            warning_msg_box(warning_msg1, end=False)
        elif not dd.isnumeric() or not mm.isnumeric() or not yyyy.isnumeric():
            warning_msg2 = "Wszystkie dane w oknie startowym muszą być numeryczne!"
            warning_msg_box(warning_msg2, end=False)
        else:
            ask_window.destroy()

    ask_button = tk.Button(ask_window, text="Zatwierdź", command=confirm_button, width=25, height=2)
    ask_button.pack(pady=5, padx=5)

    ask_window.mainloop()
    return output_dict


# Warning message box
def warning_msg_box(msg: str, end: bool = True):
    """ Message box with error.
    :param msg: Warning message to be shown
    :param end: Should kill script execution?
    :return:
    """
    warning_window = tk.Tk()
    warning_window.title("Warning !")

    warning_label = tk.Label(warning_window, text=msg)
    warning_label.pack(pady=20, padx=25)
    print(msg)

    def warning_button(end_button: bool = end):
        warning_window.destroy()
        if end_button:
            sys.exit()
    warning_button = tk.Button(warning_window, text="OK", command=warning_button, width=25, height=2)
    warning_button.pack(pady=10, padx=15)

    warning_window.mainloop()
    if end:
        sys.exit()
