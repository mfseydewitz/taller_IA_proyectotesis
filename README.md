# Anonimización de imágenes de ecocardiografía

Script local (macOS) que, para imágenes JPG y videos MP4 de ecografía Mindray:

1. Lee por OCR (motor Vision de macOS, sin internet) el RUT y nombre de la franja superior.
2. Asigna un ID aleatorio por paciente (agrupando por RUT).
3. Tapa la franja con negro y estampa `ID  IMG01` / `ID  CLIP01`.
4. Guarda los archivos como `<ID>_IMG01.jpg` / `<ID>_CLIP01.mp4`, sin metadatos.
5. Genera `<salida>_clave_sensible.xlsx` (RUT, Nombre, ID) que **no debe compartirse ni subirse**.

## Uso
```
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python anonimizar_imagenes.py carpeta_entrada carpeta_salida
```

## Privacidad
Este repositorio contiene solo código. `.gitignore` excluye imágenes, videos, planillas y claves.
