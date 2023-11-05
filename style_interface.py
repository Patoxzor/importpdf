from tkinter import ttk

def configure_treeview_style():
    style = ttk.Style()
    style.configure("Treeview",
                    background="white",
                    foreground="black",
                    rowheight=25,
                    fieldbackground="white")
    style.configure("Treeview.Heading",
                    font=('Arial', 12, 'bold'),
                    foreground='black')
    style.map("Treeview",
              background=[('selected', 'blue')],
              foreground=[('selected', 'white')])
    return style

def button_style():
    return {
        "font": ("Arial", 12, "bold"),
        "bg": "#ffffff",  # background color
        "fg": "#0000ff",  # foreground color
        "activebackground": "#e6e6e6",  # background color when button is pressed
        "activeforeground": "#0000ff",  # foreground color when button is pressed
        "borderwidth": 1,
        "relief": "raised",
        "padx": 10,
        "pady": 5
    }

def label_style():
    return {
        "font": ("Arial", 10, "bold"),
        "bg": "#ffffff",  # Cor de fundo padrão para Labels
        "fg": "#000000",  # Cor do texto padrão para Labels
    }