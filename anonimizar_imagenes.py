"""Anonimiza imágenes (JPG) y videos (MP4) de ecocardiografía Mindray.

TODO SE EJECUTA LOCALMENTE (OCR con el motor Vision de macOS, sin internet).
El script NUNCA imprime el nombre ni el RUT: los escribe directo al Excel.

Para cada archivo:
  1. Lee (OCR) el RUT y el nombre de la franja superior (1er frame en videos).
  2. Agrupa por RUT -> un ID aleatorio por paciente.
  3. Tapa la franja con negro y estampa encima "ID  IMG01" / "ID  CLIP01".
  4. Guarda como <ID>_IMG01.jpg / <ID>_CLIP01.mp4 (sin metadatos).
  5. Escribe <salida>_clave_sensible.xlsx con RUT, Nombre, ID (+ archivos).

Uso:
    .venv/bin/python anonimizar_imagenes.py prueba_entrada prueba_salida
"""
import argparse
import re
import secrets
import sys
from collections import Counter, defaultdict
from pathlib import Path

import cv2
import Vision
from Foundation import NSData
from openpyxl import Workbook
from openpyxl.styles import Font

ALTO_ENCABEZADO = 90                 # px a tapar desde arriba
ZONA_OCR = (slice(0, 88), slice(495, 900))   # (filas, columnas): nombre + RUT
EXT_IMG = {".jpg", ".jpeg"}
EXT_VID = {".mp4"}


# ---------- OCR local ----------
def ocr_lineas(bgr):
    bgr = cv2.resize(bgr, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC)
    ok, buf = cv2.imencode(".png", bgr)
    data = NSData.dataWithBytes_length_(buf.tobytes(), len(buf))
    handler = Vision.VNImageRequestHandler.alloc().initWithData_options_(data, None)
    req = Vision.VNRecognizeTextRequest.alloc().init()
    req.setRecognitionLevel_(Vision.VNRequestTextRecognitionLevelAccurate)
    req.setUsesLanguageCorrection_(False)
    handler.performRequests_error_([req], None)
    res = [(o.boundingBox().origin.y, o.topCandidates_(1)[0].string())
           for o in (req.results() or [])]
    return [t for _, t in sorted(res, reverse=True)]   # de arriba hacia abajo


def dv_rut(cuerpo: str) -> str:
    s, m = 0, 2
    for d in reversed(cuerpo):
        s += int(d) * m
        m = m + 1 if m < 7 else 2
    r = 11 - s % 11
    return "0" if r == 11 else "K" if r == 10 else str(r)


def parsear(lineas):
    """-> (rut 'NNNNNNNN-D' o '', nombre, rut_valido)"""
    rut, nombre = "", ""
    for t in lineas:
        sol = re.sub(r"[^0-9kK]", "", t)
        if not rut and sum(c.isdigit() for c in sol) >= 7 and 8 <= len(sol) <= 9:
            cuerpo, dv = sol[:-1], sol[-1].upper()
            if cuerpo.isdigit():
                rut = f"{cuerpo}-{dv}"
        elif not nombre and not re.search(r"\d", t) and len(t.strip()) > 3:
            nombre = t.strip()
    valido = bool(rut) and dv_rut(rut.split("-")[0]) == rut.split("-")[1]
    return rut, nombre, valido


def primer_frame(p: Path):
    if p.suffix.lower() in EXT_IMG:
        return cv2.imread(str(p))
    cap = cv2.VideoCapture(str(p))
    ok, fr = cap.read()
    cap.release()
    return fr if ok else None


# ---------- anonimización ----------
def tapar(frame, texto):
    frame[:ALTO_ENCABEZADO, :] = 0
    cv2.putText(frame, texto, (20, 58), cv2.FONT_HERSHEY_SIMPLEX, 1.4,
                (255, 255, 255), 3, cv2.LINE_AA)
    return frame


def procesar_imagen(origen, destino, texto):
    img = cv2.imread(str(origen))
    if img is None:
        raise RuntimeError("no se pudo leer la imagen")
    cv2.imwrite(str(destino), tapar(img, texto), [cv2.IMWRITE_JPEG_QUALITY, 95])


def procesar_video(origen, destino, texto):
    cap = cv2.VideoCapture(str(origen))
    if not cap.isOpened():
        raise RuntimeError("no se pudo abrir el video")
    fps = cap.get(cv2.CAP_PROP_FPS)
    w, h = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH)), int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    out = cv2.VideoWriter(str(destino), cv2.VideoWriter_fourcc(*"mp4v"), fps, (w, h))
    while True:
        ok, frame = cap.read()
        if not ok:
            break
        out.write(tapar(frame, texto))
    cap.release()
    out.release()


def nuevo_id(usados):
    while True:
        i = secrets.token_hex(4).upper()
        if i not in usados:
            usados.add(i)
            return i


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("entrada", type=Path)
    ap.add_argument("salida", type=Path)
    args = ap.parse_args()
    args.salida.mkdir(parents=True, exist_ok=True)
    clave = args.salida.parent / f"{args.salida.name}_clave_sensible.xlsx"

    archivos = sorted(p for p in args.entrada.iterdir()
                      if p.suffix.lower() in EXT_IMG | EXT_VID and not p.name.startswith("."))

    # 1) OCR de cada archivo -> agrupar por RUT (no se imprime nada del OCR)
    info = {}                      # archivo -> (rut, nombre, valido)
    for p in archivos:
        fr = primer_frame(p)
        info[p] = parsear(ocr_lineas(fr[ZONA_OCR])) if fr is not None else ("", "", False)

    grupos = defaultdict(list)
    for p in archivos:
        rut = info[p][0]
        grupos[rut if rut else f"__sin_rut_{p.name}"].append(p)

    usados, pacientes, filas, sin_leer, invalidos = set(), [], [], 0, 0
    for clave_g, lista in grupos.items():
        pid = nuevo_id(usados)
        rut = info[lista[0]][0]
        nombre = Counter(info[p][1] for p in lista if info[p][1]).most_common(1)
        nombre = nombre[0][0] if nombre else ""
        valido = bool(rut) and all(info[p][2] for p in lista)
        obs = "" if valido else "REVISAR (RUT no leído o dígito verificador inválido)"
        sin_leer += not rut
        invalidos += bool(rut) and not valido
        pacientes.append((rut, nombre, pid, obs))
        n_img = n_vid = 0
        for p in lista:                                  # orden cronológico por nombre
            if p.suffix.lower() in EXT_IMG:
                n_img += 1
                etiqueta, ext, f = f"IMG{n_img:02d}", ".jpg", procesar_imagen
            else:
                n_vid += 1
                etiqueta, ext, f = f"CLIP{n_vid:02d}", ".mp4", procesar_video
            destino = args.salida / f"{pid}_{etiqueta}{ext}"
            try:
                f(p, destino, f"{pid}  {etiqueta}")
                filas.append((rut, nombre, pid, destino.name, p.name, obs))
            except Exception as e:
                print(f"ERROR {p.name}: {e}", file=sys.stderr)

    # 2) Excel con la clave (columna 1 = RUT, 2 = Nombre, 3 = ID)
    wb = Workbook()
    ws = wb.active
    ws.title = "Pacientes"
    ws.append(["RUT", "Nombre", "ID", "Observación"])
    for r in pacientes:
        ws.append(list(r))
    ws2 = wb.create_sheet("Archivos")
    ws2.append(["RUT", "Nombre", "ID", "Archivo anónimo", "Archivo original", "Observación"])
    for r in filas:
        ws2.append(list(r))
    for hoja in (ws, ws2):
        for c in hoja[1]:
            c.font = Font(bold=True)
        for col in hoja.columns:
            hoja.column_dimensions[col[0].column_letter].width = 28
    wb.save(clave)

    print(f"{len(filas)} archivos procesados de {len(archivos)}.")
    print(f"{len(pacientes)} paciente(s) distintos detectados.")
    print(f"RUT no leído: {sin_leer} | RUT con dígito verificador inválido: {invalidos}")
    print(f"Clave: {clave}")


if __name__ == "__main__":
    main()
