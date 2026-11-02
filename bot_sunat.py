import os
import re
import time

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

ARCHIVO_ENTRADA = "rucs.xlsx"
ARCHIVO_SALIDA = "resultados_ruc.xlsx"
URL_SUNAT = "https://e-consultaruc.sunat.gob.pe/cl-ti-itmrconsruc/jcrS00Alias"
PERFIL_CHROME = os.path.abspath("perfil_chrome")  # guarda cookies entre ejecuciones
CARPETA_DEBUG = "debug"

COLUMNAS = ["Razón Social", "Dirección Fiscal", "Condición", "Fecha Inicio"]


def valor_de(texto, etiqueta):
    """Busca 'etiqueta' en el texto de la página y devuelve su valor
    (lo que sigue en la misma línea o, si no hay nada, la línea siguiente)."""
    lineas = [l.strip() for l in texto.splitlines()]
    for i, linea in enumerate(lineas):
        if linea.startswith(etiqueta):
            resto = linea[len(etiqueta):].strip(" :\t")
            if resto:
                return resto
            for sig in lineas[i + 1:]:
                if sig:
                    return sig
    return ""


def extraer_datos(page):
    # Se toma todo el texto visible del resultado y se parsea por etiquetas.
    # Es mucho más robusto que navegar el DOM con locators relativos.
    texto = page.locator("body").inner_text()

    num_ruc = valor_de(texto, "Número de RUC")          # "2060... - RAZON SOCIAL SAC"
    razon = num_ruc.split(" - ", 1)[1].strip() if " - " in num_ruc else num_ruc
    condicion = valor_de(texto, "Condición del Contribuyente")
    fecha = valor_de(texto, "Fecha de Inicio de Actividades")
    direccion = valor_de(texto, "Domicilio Fiscal")
    direccion = re.sub(r"\s+", " ", direccion)

    return {
        "Razón Social": razon,
        "Dirección Fiscal": direccion,
        "Condición": condicion,
        "Fecha Inicio": fecha,
    }


def ejecutar_bot_rpa():
    print(f"Leyendo RUCs desde {ARCHIVO_ENTRADA}...")
    # dtype=str evita que el RUC se lea como número
    df = pd.read_excel(ARCHIVO_ENTRADA, dtype={"RUC": str})
    for col in COLUMNAS:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype("object")

    os.makedirs("videos_evidencia", exist_ok=True)
    os.makedirs(CARPETA_DEBUG, exist_ok=True)
    total = len(df)

    with sync_playwright() as p:
        # Se usa el Chrome REAL instalado (no el Chromium de Playwright) con un
        # perfil persistente. SUNAT bloquea el Chromium automatizado: por eso el
        # captcha devolvía 401 (Unauthorized).
        context = p.chromium.launch_persistent_context(
            user_data_dir=PERFIL_CHROME,
            channel="chrome",
            headless=False,
            slow_mo=300,
            no_viewport=True,
            record_video_dir="videos_evidencia/",
            ignore_default_args=["--enable-automation"],
            args=["--disable-blink-features=AutomationControlled"],
        )
        page = context.pages[0] if context.pages else context.new_page()

        for index, row in df.iterrows():
            ruc = str(row["RUC"]).strip()
            print(f"[{index + 1}/{total}] Consultando RUC: {ruc}...")

            try:
                # Cada consulta necesita un captcha nuevo; recargar la página
                # lo regenera.
                page.goto(URL_SUNAT, wait_until="domcontentloaded")
                page.wait_for_selector("#txtRuc", timeout=15000)
                page.locator("#txtRuc").fill(ruc)

                print("   >>> Escribe el captcha en el navegador y pulsa 'Buscar' "
                      "(tienes 2 minutos)...")

                # Espera a que aparezca el resultado (lo dispara el usuario al pulsar Buscar)
                page.wait_for_selector("text=Condición del Contribuyente", timeout=120000)

                datos = extraer_datos(page)
                for col, val in datos.items():
                    df.at[index, col] = val

                print(f"   -> Éxito: {datos['Razón Social']} | {datos['Condición']}")

            except PWTimeout:
                print(f"   -> Tiempo agotado / no se encontró resultado para {ruc}.")
                df.at[index, "Condición"] = "ERROR DE LECTURA"
                page.screenshot(path=f"{CARPETA_DEBUG}/{ruc}.png", full_page=True)
                with open(f"{CARPETA_DEBUG}/{ruc}.html", "w", encoding="utf-8") as f:
                    f.write(page.content())
            except Exception as e:
                print(f"   -> Error en {ruc}: {e}")
                df.at[index, "Condición"] = "ERROR DE LECTURA"
                page.screenshot(path=f"{CARPETA_DEBUG}/{ruc}.png", full_page=True)

            # Guardado incremental: si algo falla no pierdes lo ya consultado
            df.to_excel(ARCHIVO_SALIDA, index=False)
            time.sleep(2)

        context.close()

    df.to_excel(ARCHIVO_SALIDA, index=False)
    print(f"\nProceso finalizado. Datos guardados en '{ARCHIVO_SALIDA}'.")


if __name__ == "__main__":
    ejecutar_bot_rpa()