import tkinter as tk
from tkinter import PhotoImage
import os
import sys

class ImageLoader:
    def __init__(self):
        self.images = {}

    def load(self, name, path):
        """Carrega uma imagem do disco."""
        self.images[name] = PhotoImage(file=path)

    def get(self, name):
        """Recupera uma imagem carregada pelo nome."""
        return self.images.get(name)
    
