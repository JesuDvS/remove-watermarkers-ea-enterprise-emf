"""
Script para eliminar marca de agua de archivo EMF,
corregir codificación de caracteres (mojibake UTF-8/Latin-1),
aplicar efecto espejo vertical a todo el texto,
ajustar posición Y del texto,
rotar el diagrama 180° y aplicar flip horizontal,
y re-exportar el resultado como EMF limpio usando pywin32.
"""
import sys
import os

try:
    import win32com.client
except ImportError:
    print("❌ Error: Falta la librería pywin32.")
    print("Por favor, instálala ejecutando: pip install pywin32")
    sys.exit(1)

# =========================================================
# CONFIGURACIÓN DE DESPLAZAMIENTO DE TEXTO
# =========================================================
# Valor que se sumará/restará a la posición Y de los textos.
# 0 = Sin cambios. Valores negativos suben, positivos bajan.
TEXT_Y_OFFSET = -10
# =========================================================

MSO_PICTURE         = 13
MSO_GROUP           = 6
MSO_FLIP_VERTICAL   = 1
MSO_FLIP_HORIZONTAL = 0


def remove_watermark_from_emf(input_file, output_file=None):
    """Elimina la marca de agua de un EMF buscando y reemplazando bytes."""
    if output_file is None:
        base, ext = os.path.splitext(input_file)
        output_file = f"{base}_limpio{ext}"

    print(f"🔍 Leyendo: {input_file}")
    with open(input_file, 'rb') as f:
        data = bytearray(f.read())

    print(f"📊 Tamaño original: {len(data):,} bytes")

    watermark_text  = "EA 17.1 Unregistered Trial Version   "
    watermark_bytes = watermark_text.encode('utf-16le')

    positions = []
    pos = 0
    while True:
        pos = data.find(watermark_bytes, pos)
        if pos == -1:
            break
        positions.append(pos)
        pos += 1

    print(f"🎯 Marca de agua encontrada en {len(positions)} posiciones")

    if not positions:
        print("✅ ¡El archivo ya parece estar limpio!")
        return input_file

    replacement_bytes = (" " * len(watermark_text)).encode('utf-16le')
    count = 0
    for pos in positions:
        if data[pos:pos + len(watermark_bytes)] == watermark_bytes:
            data[pos:pos + len(watermark_bytes)] = replacement_bytes
            count += 1

    print(f"🧹 Reemplazadas {count} ocurrencias")

    with open(output_file, 'wb') as f:
        f.write(data)

    remaining = data.count(watermark_bytes)
    if remaining == 0:
        print(f"✅ ¡Éxito! Guardado en: {output_file}")
    else:
        print(f"⚠️ Aún quedan {remaining} marcas de agua.")

    return output_file


def fix_encoding(text):
    """
    Corrige mojibake: texto UTF-8 que fue leído como Latin-1.
    Ejemplo: 'Ã±' → 'ñ',  'Â«' → '«',  'Â»' → '»'
    """
    try:
        return text.encode('latin-1').decode('utf-8')
    except (UnicodeEncodeError, UnicodeDecodeError):
        return text  # si no aplica, devolver el texto original


def fix_text_encoding(slide):
    """
    Recorre todas las formas con texto y corrige la codificación
    de aquellas que presenten caracteres mojibake.
    """
    fixed = 0
    for i in range(1, slide.Shapes.Count + 1):
        s = slide.Shapes(i)
        try:
            if not (s.HasTextFrame and s.TextFrame.HasText):
                continue

            tr = s.TextFrame.TextRange
            original  = tr.Text
            corrected = fix_encoding(original)

            if corrected != original:
                tr.Text = corrected
                fixed += 1
                print(f"  ✏️  '{original.strip()}' → '{corrected.strip()}'")

        except Exception:
            pass

    print(f"🔤 Codificación corregida en {fixed} forma(s) de texto")
    return fixed


def full_ungroup(slide):
    """
    Desagrupa completamente todas las formas de la diapositiva.
    Primera pasada: convierte el EMF (msoPicture=13) en objetos dibujables.
    Siguientes pasadas: deshace grupos anidados (msoGroup=6).
    """
    rondas = 0
    while True:
        hubo_desagrupable = False
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            try:
                if shape.Type in (MSO_PICTURE, MSO_GROUP):
                    shape.Ungroup()
                    hubo_desagrupable = True
                    break
            except Exception:
                pass
        if not hubo_desagrupable:
            break
        rondas += 1

    print(f"📦 Desagrupado completo en {rondas} ronda(s). Formas resultantes: {slide.Shapes.Count}")


def flip_all_text(slide):
    """Aplica flip vertical a todas las formas con texto."""
    flipped = 0
    for i in range(1, slide.Shapes.Count + 1):
        s = slide.Shapes(i)
        try:
            if s.HasTextFrame and s.TextFrame.HasText:
                s.Flip(MSO_FLIP_VERTICAL)
                flipped += 1
        except Exception:
            pass
    print(f"🪞 Espejo vertical en texto: {flipped} elemento(s) invertido(s)")
    return flipped


def adjust_text_y_position(slide, offset):
    """Ajusta la posición Y (Top) de todas las formas con texto."""
    if offset == 0:
        return 0
        
    adjusted = 0
    for i in range(1, slide.Shapes.Count + 1):
        s = slide.Shapes(i)
        try:
            if s.HasTextFrame and s.TextFrame.HasText:
                s.Top += offset
                adjusted += 1
        except Exception:
            pass
    print(f"↕️  Desplazamiento Y ({offset}) aplicado a {adjusted} texto(s)")
    return adjusted


def regroup_all(slide):
    """Agrupa todas las formas sueltas en una sola."""
    count = slide.Shapes.Count
    if count == 0:
        raise RuntimeError("No hay formas en la diapositiva.")
    if count == 1:
        return slide.Shapes(1)

    names = [slide.Shapes(i).Name for i in range(1, count + 1)]
    return slide.Shapes.Range(names).Group()


def recrop_emf(input_emf, output_emf=None):
    """
    Pipeline completo:
      1. Desagrupar completamente (EMF → objetos dibujables)
      2. Corregir codificación de caracteres (mojibake)
      3. Flip vertical en todo el texto
      4. Ajustar desplazamiento Y de todos los textos
      5. Reagrupar
      6. Rotar el diagrama completo 180°
      7. Flip horizontal del diagrama completo
      8. Recortar márgenes y exportar como EMF y PNG
    """
    if output_emf is None:
        base, ext = os.path.splitext(input_emf)
        output_emf = f"{base}_recortado{ext}"

    abs_emf_in  = os.path.abspath(input_emf)
    abs_emf_out = os.path.abspath(output_emf)

    print(f"\n🔄 Procesando EMF vía API de Windows...")

    ppt = None
    presentation = None
    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        ppt.DisplayAlerts = 1

        presentation = ppt.Presentations.Add(WithWindow=False)
        slide = presentation.Slides.Add(1, 12)

        slide.Shapes.AddPicture(
            FileName=abs_emf_in,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=0, Top=0
        )

        # ── 1. Desagrupar completamente ───────────────────────────────────
        full_ungroup(slide)

        # ── 2. Corregir mojibake en todo el texto ─────────────────────────
        fix_text_encoding(slide)

        # ── 3. Flip vertical en todo el texto ─────────────────────────────
        flip_all_text(slide)

        # ── 4. Ajustar posición Y de los textos ───────────────────────────
        adjust_text_y_position(slide, TEXT_Y_OFFSET)

        # ── 5. Reagrupar todo en una sola forma ───────────────────────────
        shape = regroup_all(slide)

        # ── 6. Rotar el diagrama completo 180° ────────────────────────────
        shape.Rotation = 180
        print("🔄 Diagrama rotado 180°")

        # ── 7. Flip horizontal del diagrama completo ──────────────────────
        shape.Flip(MSO_FLIP_HORIZONTAL)
        print("↔️  Flip horizontal aplicado al diagrama")

        # ── 8. Recortar márgenes y exportar ───────────────────────────────
        shape.Left = 0
        shape.Top  = 0
        presentation.PageSetup.SlideWidth  = shape.Width
        presentation.PageSetup.SlideHeight = shape.Height

        shape.Export(abs_emf_out, 5)  # 5 = ppShapeFormatEMF
        
        # --- NUEVO: Exportar PNG para vista previa ---
        abs_png_out = os.path.splitext(abs_emf_out)[0] + ".png"
        shape.Export(abs_png_out, 2)  # 2 = ppShapeFormatPNG
        # ---------------------------------------------

        size = os.path.getsize(abs_emf_out)
        print(f"📁 EMF exportado exitosamente: {output_emf} ({size:,} bytes)")
        return output_emf

    except Exception as e:
        print(f"❌ Error durante la re-exportación: {e}")
        return None

    finally:
        if presentation is not None:
            presentation.Close()
        if ppt is not None:
            if ppt.Presentations.Count == 0:
                ppt.Quit()


def main():
    if len(sys.argv) < 2:
        print("Uso: python remove_watermark.py <archivo_emf> [archivo_salida]")
        sys.exit(1)

    input_file  = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None

    if not os.path.exists(input_file):
        print(f"❌ Error: El archivo '{input_file}' no existe.")
        sys.exit(1)

    clean_emf = remove_watermark_from_emf(input_file, output_file)

    if clean_emf:
        recrop_emf(clean_emf)


if __name__ == "__main__":
    main()