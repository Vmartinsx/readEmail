from base64 import encode
from importlib.resources import path
from logging import root
import win32com.client as win32
from pathlib import Path
import os
import lerEmail
import pandas as pd


destino_Arquivo = lerEmail.lendoEmails()
#erro na criação de pasta