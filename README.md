# Copiado selectivo de fotos por matrícula

Programa que lee un Excel con matrículas (números enteros, columna A, sin
cabecera) y **copia** de una carpeta de origen solo las fotos cuyo nombre
coincide con esas matrículas, dejándolas en una carpeta nueva
`fotos-campeonato`.

- **Solo lee y copia.** Nunca borra ni mueve nada del origen.
- **No importa el formato del nombre:** reconoce `7799`, `7.799`, `07799`,
  `foto_7799`, `7799_1`, `7.799 (2)`… todas como la matrícula `7799`.
- **No importa la extensión:** si el nombre coincide, copia el archivo sea
  `.jpg`, `.png`, `.heic`, sin extensión, etc.
- Una matrícula puede tener **varias fotos**: las copia todas.
- Genera reportes: `no_encontradas.txt` y `reporte.txt`.

---

## Uso en la PC que tiene las fotos (cero instalación)

`copiar_fotos.exe` **lleva Python y todo lo necesario dentro del archivo**. En
esa computadora **no se instala nada**: ni Python, ni paquetes, ni PATH, ni
internet. Solo:

1. Pon **`copiar_fotos.exe`** y **`matriculas.xlsx`** en una misma carpeta
   (donde quieras; el programa se ubica solo).
2. Doble clic a `copiar_fotos.exe`.
3. Al terminar, en esa misma carpeta quedan:
   - `fotos-campeonato/` con las fotos copiadas,
   - `no_encontradas.txt` con las matrículas sin foto,
   - `reporte.txt` con el detalle completo.

La ventana espera un **Enter** al final para que se pueda leer el resumen.

> La carpeta de origen está fijada dentro del programa:
> `C:\Users\ROXANA\Desktop\CREDENCIALES\FOTOS TOTAL CREDENCIAL`.
> Si cambia, hay que editar la variable `ORIGEN` en `copiar_fotos.py` y
> recompilar.

### Posible aviso de Windows (una sola vez, no es instalar nada)

Como el `.exe` no está firmado digitalmente, la **primera** vez Windows puede
mostrar *"Windows protegió su PC"* (SmartScreen). Es un clic, no una
instalación:

> **Más información → Ejecutar de todas formas.**

Después ya no vuelve a aparecer. (Eliminar ese aviso por completo requeriría un
certificado de firma de código de pago; no es necesario para que funcione.)

---

## Cómo se genera el .exe (GitHub Actions)

El `.exe` es de Windows y aquí trabajamos en Mac, así que se compila en la nube:

1. Sube esta carpeta a un repo de GitHub.
2. Pestaña **Actions** → workflow **"Compilar EXE (Windows)"** → **Run
   workflow** (o se lanza solo en cada push a `main`).
3. Cuando termine, entra al run y descarga el artefacto **`copiar_fotos-exe`**:
   dentro está `copiar_fotos.exe`.

El workflow está en `.github/workflows/build.yml` (runner `windows-latest` +
PyInstaller `--onefile`, así que sale un único `.exe`).

### Compilarlo en una PC con Windows (alternativa)

```bat
pip install pyinstaller
pyinstaller --onefile --console --name copiar_fotos copiar_fotos.py
```

El `.exe` queda en `dist\copiar_fotos.exe`.

---

## Configuración (parte superior de `copiar_fotos.py`)

| Variable      | Qué es                                                    |
|---------------|-----------------------------------------------------------|
| `ORIGEN`      | Carpeta con todas las fotos (se lee, nunca se modifica).  |
| `EXCEL`       | Nombre del Excel; por defecto `matriculas.xlsx` al lado.  |
| `DESTINO`     | Carpeta donde se copian; por defecto `fotos-campeonato`.  |
| `EXTENSIONES` | `None` = cualquier extensión. O un set para limitar.      |
| `DRY_RUN`     | `True` = solo reporta lo que haría, sin copiar nada.      |

El Excel se lee con la **librería estándar** de Python (no necesita `openpyxl`).

---

## Cómo empareja los nombres (a prueba de formatos)

Para cada nombre de archivo y cada matrícula:

1. Quita separadores de miles (puntos, comas, espacios, apóstrofes).
2. Toma la secuencia de dígitos más larga (así `7799_1` → `7799`).
3. Quita ceros a la izquierda (`07799` → `7799`).

Si un nombre tiene dos números del mismo largo (p. ej. `7799-8596`), se copia
igual pero queda **marcado como ambiguo** en `reporte.txt` para revisarlo a
mano.
