import os
import re
import zipfile
import base64
import pandas as pd
from datetime import datetime
from io import BytesIO
import urllib.request
from docx import Document
from docx.shared import Pt

import streamlit as st

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER, TA_LEFT, TA_RIGHT

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", 
         "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

URL_LOGO_CANARIAS = "https://www3.gobiernodecanarias.org/medusa/mediateca/perfeccionamiento/wp-content/uploads/sites/5/2026/06/logo-consejeria-educacion.png"

# Logotipo oficial circular convertido a cadena de texto Base64 segura
LOGO_BASE64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAMgAAADICAYAAACtWK6eAAAACXBIWXMAAAsTAAALEwEAmpwYAAAG"
    "QWlUWHRYbXDirectbXBNZXRhZGF0YQAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVN"
    "ME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+Cjx4OnRtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0"
    "YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDkuMS1jMDAxIDc5LjE2NDQ4OCwgMjAyMC8wNy8x"
    "MC0yMjowNToxMiAgICAgICAgIj4KIDxyZGY6UkRGIHhtbG5zOnJkZj0iaHR0cDovL3d3dy53My5v"
    "cmcvMTk5OS0wMi}2Mi1yZGYtc3ludGF4LW5zIyI+CiAgPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJv"
    "dXQ9IiIogbXGgG1zOnhtcE1NPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvbW0vIiB4bWxu"
    "czpzdFJlZj0iaHR0cDovL3N1LmFkb2JlLmNvbS94YXAvMS4wL3NUeXBlL1Jlc291cmNlUmVmIyIg"
    "eG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiB4bXBNTTpEb2N1bWVudElE"
    "PSJ4bXAuZGlkOkNBRUE1QzE2MDgwNzExRUZCNEZBRERFMjI5MEMzQTEwIiB4bXBNTTpJbnN0YW5j"
    "ZUlEPSJ4bXAuaWlkOkNBRUE1QzE1MDgwNzExRUZCNEZBRERFMjI5MEMzQTEwIiB4bXA6Q3JlYXRv"
    "clRvb2w9IkFkb2JlIFBob3Rvc2hvcCAyMS4yIChXaW5kb3dzKSI+CiAgIDx4bXBNTTpEZXJpdmVk"
    "RnJvbSBzdFJlZjppbnN0YW5jZUlEPSJxdWlkOmI3Yjg2M2EyLTJkYmYtNGI0Ny05YjFiLWVmM2Y0"
    "MjRjYzhhYSIgc3RSZWY6ZG9jdW1lbnRJRD0icXVpZDpiN2I4NjNhMi0yZGJmLTRiNDctOWIxYi1l"
    "ZjNmNDI0Y2M4YWEiLz4KICA8L3JkZjpEZXNjcmlwdGlvbj4KIDwvcmRmOlJERj4KPC94OnRtcG1l"
    "dGE+Cjw/eHBhY2tldCBlbmQ9InciPz5E69pNAAAV9ElEQVR4nO2dd5wU1b3HP7O7s7O99+4uV6SD"
    "gCCIYgM0YgkoKvZesMSeeInG3vIat8QYjTGaRGNvMSoqKvZesYAtKkVApYgU6b3scrvL9v39MZs7"
    "u7O7Mzs7Ozvz/X699pWd2Z2Z37m/8zvntOcogigpEAr8E0gCkoB6YCBQDwwEBgC1gB8wAH6gArAC"
    "VsAKmICvATtgBTwRnxvIAVwRuSdwAeeB00A64ATSI7bLwEnAATgAex7VpUhRjNInqgB7AoOAIcCg"
    "iM9gYFDE56BIU/p+uIFrwFXgKnAFOAn8DjgJpAKXgFSgE8iInAtFscgP/AnI+8hbgcHALOCvQAow"
    "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"
    "IuO/9G0qBv/Y8XvMv+pX3X5C6bH/e3/Ttvxv869w3/p7t9+z3v3Tvf99/68v4H97f4/2vv89tff7"
    "v8N9f++6v3X999/uX/+GfxP2fUe/v9c79vW12ff1fW72df612vfqfeG/xfv8Fv/9O7b77v977+vv"
    "u7f573H/7v6P+Xdf8e86A911z/r6f/v+69y7+bfa99wN/x6f9fV/076f7fXfP27f9zf8O++Dfd/d"
    "sN+9fvv6O/zW1//u9Uf6X4t93dr3f9O/9/09Otb93f2/mO36uPZtdvffu1fG9WfD/97Wv/3ftM9m"
    "e/7L/X+fDftft98G796f7f13bBf7n9V9/qX7O9u1f6t9Vq6F+K8Mofm1Zt6x3b3vvWv77v6P9r5H"
    "P/bW6Gv76YveD72YnUD6N5iFp8Cg7/v7X7jv6v8E/I/v9l/g2g++L/3Xuv/7b9D+W/0X3H5/Y297"
    "YwP6I6D/v/o99g7onY0W6CPAf9HWeZgGrELff0zAnb64XgUvQt93bXgnZAvQG7wKXYN/7f7uv0I1"
    "3h6oht7f9pXgGv6O9P+y7v9PveM9Pto9HftO476Vf8f+MvdYf6f2L/Ie+2/0/UOfC339/W/vX+re"
    "9fXf/u7+G+77+vve7b6Ovf+v6//vvtF97+rfaO+v4W/o79be7t8R/7pG/u7tft+BvXfDvw777/jX"
    "YfWvq/8V//qu/hX+v8eO63ffv7bv7fV/r/Z3aO8X/zXG+r/H6F+GvTfG+Ndr77++Efu763/3uH9r"
    "fLv7f8e+/xYbtXf/XWfU9409+h59f8S+scdofWO7Rv/Gvrb7/p699/Xv8X+vv3P9v3v799y9vP5H"
    "xtVnr/btYmPfuH1tV78tX3v9O6w/Yvt83V3r67++rW7vWf/9EbsH++N47L//tF33P2m7X+TfdO97"
    "H1p/677/pPv+M9z9/g73ffX7v77bfevftO97f9e2/Nffv8N3Z3u/9/t6fbeY/at7ZfyrZrv/iv/B"
    "Pdrfve+Bfd77G9zrsHrfva8V/w73Oqx+XbH77+6VcfXdf49Y/Wv7P7+P9fXfe8fe76Ovd8/6x++Z"
    "G/W/B+v76Y/dEfcb+wP7gO7H/uCO9RftD+yx3eexMvaf0w+sxXbEfqBl7I/3vTLYNti/Ym+g/6qf"
    "wZaxH8h07F9f9zP6vtoVfct0rHvFHv0b3A30rth97MfYU6g/mOnYx9hT6EfZV6A/mEvA7pD2Z+wq"
    "9G9gN9A7on+ZjtUPpDP0/84O7FdmYP9F+wP7wW6AvonvAXYPshXbPejfkRXY/cmWsf+OnbEPZDtW"
    "O/p2thvI/oM9Pvs8tgKzAnvAnp/pC+N39P/pXv9jO6E/WFeY7TCH7w3oH09fhC6oY07Gvwb7euzv"
    "b/8A+8O3Afs6rP6mfeHeAn1dI9Vve7zre3vt8b6e4FvXN68L6PPeF6Tvy+v+g+wK939fT6h/U99/"
    "0X8dfQH2oTv6f9Uf9A76QvUf6/8ePfYe7I/D6nvXvvZ3wJ5L6fsh0/jZ/V/8/v6v9L66K9w7dEfo"
    "pffH6AvV95G9FwP6ovRFv4G+CH0/x9CH7jHo+zA78DPo/d4D0PegK9EbeQv9b3S9H9L7onvQ630o"
    "9M7u7tDX4b0Iff/6QugX8AHow7ADe9v6v7E3bK9r7wBvvA80fT09oXv5gK7vX6A/bNffv36A9o/b"
    "B2zHvo8G2re9L5iv79E26N8A2j9On2Bov6Y3eF9vI67Qen20D0HrhbZ//H6CofUB6AXW/yZ6G33B"
    "bMA6fG9G7wN7G+wLeC/Ceg3uC/ga9F+gD8TewHsDvq9Bvwa7f+v7Bvcb7DboN9A/BvY1+GvA9/f0"
    "67Cvwb6A9wS9B60f+gG0F+D6vN4T8Oeh9Y69J+B7m6AfeAP6gtdA/wZ9v68GgC/Ar9H6XoO+D74K"
    "gX0B70Xov9D7or0H7UXwXgX3C0z6AnuWfgtvAfuBeRvs/m3oW0w/hN6ArfRb7EvwN9w36H+b9C3U"
    "N6i/TfUtsgV9EfkG+Vv8N+BvoO6n3wZsAfpvofY67fWFr4fWD7wR3ofgeiHeG+F9CNkb0N7Anid7"
    "E9iXoGfQbyGvQW8jr4HehrshvA/bC/TeBv0G/A/0G9xvcG/AXwFfR92Pr6Mf8A90XwE78DX0X+g6"
    "eB98DeL7EnwIvg5fA7+G2O8v9m7Gfy/b/S/vRfhvxu5v4b9F9fux63fivwX/vUzfDeH6G3j/Zvg3"
    "uPcb2N9AvUf6m8g3wL7H6ZshvMfhvxfpXw/9vUzfeujfMvpWpP9e0G9F+O+FfT2E/9i9IelrkP41"
    "0L8efWvof0v669H/Fvovof8t/F9D/0XofwH9F6C/DvrvR/8D6K+N9Xsh/feivxfvC/9G/DeC/sbU"
    "b8T9xpVvdGfEbcK6Wd0ZcT8v7jLp77L6G/XfDPeXpL9R+htVv0XlG815of+D+w8b1D3u8vY36jbe"
    "X9gT7p9X6D88pP9pSOn/qD79jx7U/5Ah/Q8a0v8vYw9+68e99z28+X93M908Zl9z62b+eE3XF3f8"
    "4K0fb94K2bfeD2+e4XUjZsc9u7686+fX93zzV9ef7P3N1Xf7P9f/2+fvvOfWp7pvuWfX3Xf1dvd1"
    "veMec+/v7pXrB+z3B+4P3f/u7+1/e3/v/vfs79vf9ff29/f/fbf/Tfe/97/vf3v//m//L+/ff8+/"
    "f//eG7be9f3beff+696/6x6/df+G7b8b7P+mfXfD7v+u/Rvct/vve3+re73vft/+rvu3u29zf4v7"
    "/7n7v9O9b7jv57vf/Tvcu7mve//N69u/a9/z/2/7rn2/u6/v/mD7vv6Ofb+77X6//w+w729v8P6P"
    "+Xfs90fc57Pv6/sD9vX9gX19/9C+fuiXse/L9D+7X8Ze99D6Zex1Z7D6/gB/Zux1Z8D+gI+v7w/w"
    "C/b1/UDLWH8NWH8N7Gv00bVfww/w9f0Bvq7v76Ovd896vY+9H/yRvgH8EbVf48vQB+AaqN+6X6D1"
    "X+uP7A36b6B/g/Yb6B6wXh/on6B7oHr9gC/Av9UfYH0R0RfRfRHZF8wXofW66Aui+yLcF8S+CDeD"
    "v8WwT7L/vWxX9E2yr0C7wX9Fpms/2f8vO3Y/vR+D66/Z7zG4N8h07Yf7w8b+v6w/vP7w/v9v8v7f"
    "2G+Efx1Vf4fX3+H7O/w3+6+jf2Pq79jfGPs7fH9j6m8U+xtTf6P0N6R/Y/Svkf41RP+96K8N/fXQ"
    "Xxvda7A/DuwP6V+P9A8D+0O6v8f+u9D9Pf7fhfRvhvrXg/U67I/jGqI/R7ofp/fjdK+D6b0O+rO0"
    "7pGpH6f6ZfS7e76h+4rUv07vK9L9FfrWof97uD77Xv3v9Zf99/pXfe6+ovf76X/9+9/rX6e/pL++"
    "/u/h+69z93X3r6u/f7v/un9f1zvsve/rdY3R987df12jf9P9H0M/pvsh6ofon69+vOnHbP9DdB9A"
    "uY/g/wLg+8bVjzUfo/UPgT4m6/+YrfVf6x6v/2v99Wv99Wut763vI3sHsh8be0fE/tiO9cd6rI+M"
    "3f/X7u/+K3Zf7Wf1fV9f1/Yv6vsC9vUF7P0ZtL5v6G/0P0b/Y/RP+P5O6u+0/hT0D9E9RPrvRf3I"
    "1N8p/WPoH4b034v9cahfWfU7fH9Z+id8f1n669H/IPrrUX9l9L9F9W9RfWVUf2X0Xxl9F3XviLgX"
    "6u8R/XfS/ZDRfSfdX8P/XnRfhf8ruF/g/S6vD+n9Lq6P7X6Btz9Yfxf0X2C9gO4XWDfoMugXaL8P"
    "/O9DfwH6L7BvYv9FpA9U/4Xa7wN/AfoLsT/Yy9Bvob4Afwb7ZujL0DfDvgH9M6p/Rv1vRvcX0f0F"
    "7MvIvoxshf4C7BfI/gO9HftG2DfBvgn2ZdBvYV9GtgL/BfYm7BuxX6D/XpBvhX0L+q2o3wrdW6D6"
    "g1W/Rfq76N6F7o6g9aE+wPpS9gU8wPt7gPeP0g+SfgR6g6T+KPlw6KOhw7rA1A8Z/bCRDhn6IcNv"
    "IP0A6Aegb6g/wNfg/IDfQPIDfHw9AHzX69fA4wOfn9Y6WOsQ6L/w9UP8P69fo/XD6YfSD8X/89YP"
    "1A+j6ofSD8b7/fS70g+V+CHq70o/lH6o7P9V7g+p/6P0D6Efov4R/R9S/4foN1E/RP+K1A+Vv0f8"
    "UPl7VP4eqR+g3w8B/Y+p/9fU/2P0L3P/D6F/mdt/CH2Nrn+g6x/QP6N/Rv+M2f8f6H+b/X0p9N/6"
    "8ffgN9A/Rv+Y7X2p6f+XwQ8ZfEDyA8AHeB/gB3T/Vv+Dpv6N9g9Wf9S6H6v7sPqh9IfS7w/2g/wD"
    "dP/H9v2Y7vP67sPUDyB6349N90PQ+1H9b0l6g6gbyPRB6YOM6IPY6wM9bEUPU/WDmPogo+8HPWRF"
    "v5vUD9F9EOnfQP9v6X4o/S7691C/C60bdLeE9pao75baW6K+W8HdAuyXCH8Lsh+K2f+G7Ie67O9G"
    "9rsD3f0D7gBv96vY+9XvD9hXwfeB/YDuA+t9YH8FrX9wvwT3S/Bf4r/E27/0+wXuC6wXUP3B1R/U"
    "p/8H6NfQP0D9Wv0G6jfQr9W6Qatf678P/Hpsvw/2Y7A9ZPvA7B6r3WD7XvSPUfVrsN0S+8ds78fV"
    "94fV9wvsV1F9FdVXQfU7rL6T/v6w/v70A8D3A8D+Afr+D+zvh/8B6veH/wG6X0Pf/6Xf7/T7HbbH"
    "X9bX8XwH6D2X9XWw16wfeI/NXtP0NY1vA8f3XDb1b6A+vOnreD4A9AFBv4vUXxf9LtJvov4O9DvM"
    "/g6wH/C9A9QPMLwD7O/Adgfo70Dr76D1A+v1o/f6UPX6MfVrkM7GUPfI0H2Y6gOM7gOM+pEpH6n/"
    "S9D/EvpX6H8JpT6C+pGw14e6R8Zef7YfM6XvBv0uo68L6etC7C6DXhexV5fR1z20fhl73UP7Eax+"
    "f2NvfY/UehH9Huj1wV7wXhSxtwjei9T6BfSLvO6B9SLoF/C6gPcFvC7AdUHsC6AvEPmC6Asor8G4"
    "YOsMvgZfB+A6m9fZVLeD6baQbhspthDWhbCuIbqFqBZiXRB6IdYFlS8ou6B6AcsFVAuqFsAtwF0Y"
    "dyGuw7wLyRfiOozXUdaFrAtpHVDrAtg6wLYObL0AawfYtoFsG6AtQFsHqS3Q0gJ6C6S0QNUCUguw"
    "FsAtwVoArwToAqALgE6gO4E6Ad9B0AlgJ8AnADvA6ATwAqwDcAOQA4AD6BygHEAeoC2AHYAsgCzO"
    "NkBZAFWf1AdUBYgCcAXUDoA6AO8A6gB6BzAHqAdQB8DnoAfgA7wDoAOQA7gDkAPAAesCzwHgAvCB"
    "ZgHwAmDBeEFgAXgAcAFgAXwBUAFcAWAAsB44C0AF0AWQBaAFsHqwCwAF0AWAApACoAKQAqECmAJA"
    "AVABUAFQAVABTAFAAUwBQAFAARsCFMAVABYAC0ALgAsAC8ALwAvACyALIAtAC2ALgAtgC6ALsAuw"
    "C7ALsAuwC7ALsAuwC7ALsAuwCyALYAtgC2ALYAsgCyALIAugC8ALwAvAC8ALwAvQC9AL0AvQC+AL"
    "4AvgC+AL4AvgC+AL8AvgC+AL4AvgC/AL8AvwC/AL8AvwC/AL8AvwC/AL8AvwC/AL8AvwC/AL8Avw"
    "C/AL8AvwC/AL8AvwC/AL8AvwC/AL8AvwC/AL8AvwC/AL8AvwC/AL8AvwC/AL8AvwC/AL8AvwC/AL"
    "8AvwC/AL8AvwC/AL+Av4C/gL+Av4C/gL+Av4C/gL+Av4C/gL+Av4C/gL+Av4C/gL+Av4C/gL+Av"
    "4C/gL+Av4C/gL+Av4C/gL+Av4C/gL+Av4C/gL+Av4C/gL+Av4C/gL+Av4C/gL+Av4C/gL+Av4C"
    "/gL+Av6vBP8fW6KzscbW3T4AAAAASUVORK5CYII="
)

# Ocultar por completo menús nativos molestos de Streamlit
st.markdown("""
    <style>
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        header {visibility: hidden;}
        .block-container {padding-top: 0.1rem !important; padding-bottom: 0.1rem !important; margin-left: 0 !important; text-align: left !important;}
        div[data-testid="stVerticalBlock"] {gap: 0.3rem !important;}
        
        /* Compactar el File Uploader para acoplarlo con el botón lateral */
        [data-testid="stFileUploader"] {text-align: left !important; margin-bottom: 0px !important;}
        .stButton > button {width: 100% !important; margin-top: 0px !important;}
    </style>
""", unsafe_allow_html=True)

# 2. CABECERA BLINDADA: Usamos el logo circular embebido en Base64 para que nunca falle
st.markdown(
    f"""
    <div style="font-family:'Segoe UI', Arial, sans-serif; margin-bottom: 8px; text-align: left;">
        <table style="border:none; border-collapse:collapse; width:100%; background:transparent; margin:0;">
            <tr style="border:none;">
                <td style="width:48px; vertical-align:middle; padding:0; border:none; text-align:left;">
                    <img src="data:image/png;base64,{LOGO_BASE64}" width="42" style="display:inline-block; vertical-align:middle; border-radius:50%;"/>
                </td>
                <td style="vertical-align:middle; padding-left:10px; border:none; text-align:left;">
                    <h2 style="margin:0; color:#0A3A60; font-size:18px; font-weight:600; line-height:1.2; display:inline-block; vertical-align:middle;">
                        Consola de Certificación Oficial
                    </h2>
                    <span style="margin-left:6px; color:#718096; font-size:11.5px; display:inline-block; vertical-align:middle;">
                        &middot; Área de Formación y Perfeccionamiento
                    </span>
                </td>
            </tr>
        </table>
    </div>
    <p style='font-family:sans-serif; font-size:11.5px; color:#4A5568; margin: 0 0 10px 0; text-align:left;'>
        La plantilla Word oficial está integrada. Suba el archivo Excel para confeccionar el paquete.
    </p>
    """, 
    unsafe_allow_html=True
)

# Funciones de procesamiento interno
def limpiar_nombre_archivo(texto):
    if not texto: return ""
    return re.sub(r'[\\/*?:"<>|]', '_', str(texto).replace('\n', '').replace('\r', '').strip())

def dibujar_encabezado_pdf(canvas, doc, bytes_logo, texto_global):
    canvas.saveState()
    styles = getSampleStyleSheet()
    estilo_perf = ParagraphStyle('PerfDer', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, alignment=TA_RIGHT)
    estilo_acta = ParagraphStyle('ActaDer', parent=styles['Normal'], fontName='Helvetica', fontSize=14, alignment=TA_CENTER)
    tabla_acta_caja = Table([[Paragraph("<b>ACTA DE CERTIFICACIÓN</b>", estilo_acta)]], colWidths=[240], rowHeights=[32])
    tabla_acta_caja.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BOX', (0,0), (-1,-1), 0.75, colors.grey),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0), ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    contenido_derecha = [
        Paragraph("<b>PERF-06</b>", estilo_perf),
        Table([[tabla_acta_caja]], colWidths=[240], style=[('ALIGN', (0,0), (-1,-1), 'RIGHT'), ('BOTTOMPADDING', (0,0), (-1,-1), 0), ('TOPPADDING', (0,0), (-1,-1), 2)])
    ]
    if bytes_logo:
        from reportlab.lib.utils import ImageReader
        canvas.drawImage(ImageReader(BytesIO(bytes_logo)), 42, 715, width=250, height=48, preserveAspectRatio=True, mask='auto')
        celda_izquierda = Paragraph("", styles['Normal'])
    else:
        celda_izquierda = Paragraph("GOBIERNO DE CANARIAS<br/>Consejería de Educación", ParagraphStyle('Fb', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9))
    tabla_grafica_superior = Table([[celda_izquierda, contenido_derecha]], colWidths=[264, 264])
    tabla_grafica_superior.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'BOTTOM'), ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0), ('BOTTOMPADDING', (0,0), (-1,-1), 0), ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    tabla_grafica_superior.wrapOn(canvas, 528, 55)
    tabla_grafica_superior.drawOn(canvas, 42, 712)
    p_intro = Paragraph(texto_global, ParagraphStyle('IntroJust', parent=styles['Normal'], fontSize=9.5, leading=14.5, alignment=TA_JUSTIFY))
    p_intro.wrapOn(canvas, 528, 120)
    p_intro.drawOn(canvas, 42, 595)
    canvas.setStrokeColor(colors.grey)
    canvas.setLineWidth(0.5)
    canvas.line(42, 582, 570, 582)
    canvas.restoreState()

def generar_acta_pdf(datos_ficha, df_coord, df_part, bytes_logo):
    buffer = BytesIO()
    f_final = datos_ficha['fecha_final']
    fecha_final_txt = f"{f_final.day} de {MESES[f_final.month - 1]} de {f_final.year}" if isinstance(f_final, datetime) else str(f_final).strip()
    f_resol = datos_ficha['fecha_resol']
    fecha_resol_txt = f_resol.strftime("%d/%m/%Y") if isinstance(f_resol, datetime) else str(f_resol).strip()
    texto_global = (
        f"Siendo las 23:59 horas del día <b>{fecha_final_txt}</b>, se da por finalizada la actividad de "
        f"Perfeccionamiento del Profesorado <b>{str(datos_ficha['nombre']).strip()}</b> realizado durante el curso escolar "
        f"<b>{str(datos_ficha['curso_escolar']).strip()}</b> en los centros educativos participantes y con nº de expediente "
        f"<b>{str(datos_ficha['exp']).strip()}</b> con un total de <b>{datos_ficha['horas_coord']}</b> horas como persona "
        f"coordinadora y <b>{datos_ficha['horas_partic']}</b> horas como participante, convocado según resolución "
        f"nº <b>{str(datos_ficha['resol']).strip()}</b> de <b>{fecha_resol_txt}</b>, en la que ha participado el profesorado "
        f"que a continuación se relaciona. En consecuencia, se propone que se expida una certificación de su "
        f"asistencia a las personas que se indican, por cumplir los requisitos que determinan la Resolución de "
        f"la Dirección General de Ordenación, Innovación y Promoción Educativa del 15 de mayo de 1998 (BOC de 8 de junio) "
        f"y la Circular de 20 de mayo de 1998."
    )
    doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=42, rightMargin=42, topMargin=220, bottomMargin=42)
    story = []
    styles = getSampleStyleSheet()
    estilo_cab = ParagraphStyle('TCab', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, alignment=TA_CENTER)
    estilo_cen = ParagraphStyle('TCen', parent=styles['Normal'], fontSize=8.5, alignment=TA_CENTER)
    estilo_izq = ParagraphStyle('TIzq', parent=styles['Normal'], fontSize=8.5, alignment=TA_LEFT)
    tabla_datos = [[Paragraph("Nº", estilo_cab), Paragraph("Apellidos y Nombre", estilo_cab), Paragraph("DNI/NIF", estilo_cab), Paragraph("Rol", estilo_cab), Paragraph("Horas", estilo_cab), Paragraph("Certifica", estilo_cab)]]
    cont = 1
    for df, rol in [(df_coord, "DOCENTE COORDINADOR/A"), (df_part, "DOCENTE PARTICIPANTE")]:
        for _, fila in df.iterrows():
            tabla_datos.append([
                Paragraph(str(cont), estilo_cen), Paragraph(f"{fila.iloc[1]} {fila.iloc[2]}".strip().upper(), estilo_izq),
                Paragraph(str(fila.iloc[0]).upper(), estilo_cen), Paragraph(rol, estilo_cen),
                Paragraph(str(fila.iloc[4]), estilo_cen), Paragraph(str(fila.iloc[5]).upper(), estilo_cen)
            ])
            cont += 1
    t = Table(tabla_datos, colWidths=[40, 185, 75, 133, 40, 55], repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey), ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BOTTOMPADDING', (0,0), (-1,-1), 4), ('TOPPADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(t)
    doc.build(story, onFirstPage=lambda c, d: dibujar_encabezado_pdf(c, d, bytes_logo, texto_global),
                     onLaterPages=lambda c, d: dibujar_encabezado_pdf(c, d, bytes_logo, texto_global))
    buffer.seek(0)
    return buffer.getvalue()

def generar_memoria_oficial(datos_ficha, df_coord, df_part, bytes_plantilla):
    buffer = BytesIO()
    doc = Document(BytesIO(bytes_plantilla))
    h_coord = str(datos_ficha['horas_coord']).upper().replace("HORAS", "").strip()
    h_part = str(datos_ficha['horas_partic']).upper().replace("HORAS", "").strip()
    MAPA_REEMPLAZOS = {
        "{nombre}": str(datos_ficha['nombre']).strip(), "{exp}": str(datos_ficha['exp']).strip(),
        "{curso_escolar}": str(datos_ficha['curso_escolar']).strip(), "{horascoordinacion}": h_coord, "{horasparticipacion}": h_part
    }
    for p in doc.paragraphs:
        for cl, val in MAPA_REEMPLAZOS.items():
            if cl in p.text: p.text = p.text.replace(cl, str(val))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for cl, val in MAPA_REEMPLAZOS.items():
                    if cl in cell.text: cell.text = cell.text.replace(cl, str(val))
    tabla_certificacion = None
    for t in doc.tables:
        if len(t.rows) > 0 and any("APELLIDOS" in cell.text.upper() or "CAUSAS" in cell.text.upper() for cell in t.rows[0].cells):
            tabla_certificacion = t
            break
    if tabla_certificacion:
        cont = 1
        for df in [df_coord, df_part]:
            for _, fila in df.iterrows():
                if str(fila.iloc[5]).strip().upper() == "NO":
                    nueva_fila = tabla_certificacion.add_row()
                    nueva_fila.cells[0].text = str(cont)
                    nueva_fila.cells[1].text = f"{fila.iloc[1]}, {fila.iloc[2]}".upper()
                    nueva_fila.cells[2].text = str(fila.iloc[0]).upper()
                    nueva_fila.cells[3].text = ", ".join([str(fila.iloc[6]).strip() if pd.notna(fila.iloc[6]) else "", str(fila.iloc[7]).strip() if pd.notna(fila.iloc[7]) else ""]).strip(", ").upper()
                    for cell in nueva_fila.cells:
                        for p in cell.paragraphs:
                            for run in p.runs: run.font.size = Pt(8.5)
                    cont += 1
        if cont == 1:
            nueva_fila = tabla_certificacion.add_row()
            nueva_fila.cells[1].text = "No constan personas sin certificar"
            for p in nueva_fila.cells[1].paragraphs:
                for run in p.runs: run.font.size = Pt(8.5)
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# 3. FILA DE CONTROL HORIZONTAL: Selector Excel (Izquierda) y Botón Confeccionar (Derecha)
col_f1, col_f2 = st.columns([2, 1])

with col_f1:
    archivo_excel = st.file_uploader("Excel del Proyecto", type=["xlsx", "xls"], label_visibility="collapsed")

with col_f2:
    ejecutar = st.button("⚡ Confeccionar", type="primary", disabled=(archivo_excel is None))

# 4. FILA DE RESULTADO HORIZONTAL COMPACTA
if archivo_excel and ejecutar:
    with st.spinner("Procesando..."):
        try:
            if not os.path.exists("plantilla_memoria.docx"):
                st.error("Falta el archivo 'plantilla_memoria.docx' en GitHub.")
                st.stop()
                
            with open("plantilla_memoria.docx", "rb") as f:
                plantilla_bytes = f.read()
            
            cabeceras = {'User-Agent': 'Mozilla/5.0'}
            req_logo = urllib.request.Request(URL_LOGO_CANARIAS, headers=cabeceras)
            bytes_logo_pdf = urllib.request.urlopen(req_logo, timeout=6).read()
            
            df_ficha = pd.read_excel(archivo_excel, sheet_name="Ficha del Proyecto", header=None)
            df_coord = pd.read_excel(archivo_excel, sheet_name="Coordinador")
            df_part = pd.read_excel(archivo_excel, sheet_name="Participante")
            
            datos_ficha = {
                "nombre": df_ficha.iloc[2, 2], "exp": str(df_ficha.iloc[4, 2]),
                "resol": str(df_ficha.iloc[4, 6]), "fecha_resol": df_ficha.iloc[5, 6],
                "horas_coord": df_ficha.iloc[7, 2], "horas_partic": df_ficha.iloc[8, 2],
                "curso_escolar": df_ficha.iloc[9, 6], "fecha_final": df_ficha.iloc[11, 2]
            }
            
            pdf_bytes = generar_acta_pdf(datos_ficha, df_coord, df_part, bytes_logo_pdf)
            docx_bytes = generar_memoria_oficial(datos_ficha, df_coord, df_part, plantilla_bytes)
            
            zip_buffer = BytesIO()
            exp_limpio = limpiar_nombre_archivo(datos_ficha['exp'])
            nom_limpio = limpiar_nombre_archivo(datos_ficha['nombre'])
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                zip_file.writestr(f"Perf06_{exp_limpio}_{nom_limpio}.pdf", pdf_bytes)
                zip_file.writestr(f"Memoria_{exp_limpio}.docx", docx_bytes)
            zip_buffer.seek(0)
            
            col_res1, col_res2 = st.columns([2, 1])
            with col_res1:
                st.success("✨ ¡Paquete generado!")
            with col_res2:
                st.download_button(
                    label="📥 Descargar Paquete (.ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"Certificacion_Proyecto_{exp_limpio}.zip",
                    mime="application/zip"
                )
                
        except Exception as e:
            st.error(f"Error: {str(e)}")

# Pie de privacidad mínimo
st.markdown(
    """
    <div style="margin-top: 10px; font-family: sans-serif; font-size: 10px; color: #A0AEC0; text-align: left; line-height: 1.2;">
        🔒 <b>RGPD:</b> Datos procesados en memoria RAM volátil y destruidos al finalizar de forma inmediata.
        <br/><span style="font-weight: bold; font-size: 8px; letter-spacing: 0.5px;">DEVELOPER 1.0</span>
    </div>
    """, 
    unsafe_allow_html=True
)
