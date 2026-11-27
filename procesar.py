#!/usr/bin/env python3
"""Convierte los Word (.docx) y los textos (.md) de la carpeta entrada/ en notas del sitio.

Lee cada .docx de entrada/, arma la nota (título, bajada, fotos con epígrafe
y crédito, subtítulos, texto y fuentes), guarda las fotos en img/, actualiza
propias.json y mueve el Word a procesadas/. Si un Word no se puede leer, lo
mueve a errores/ con un .txt que explica el problema y termina con error.
"""
import datetime, glob, json, os, re, shutil, subprocess, sys, tempfile, unicodedata, urllib.request
from PIL import Image

RAIZ = os.path.dirname(os.path.abspath(__file__))
NOTAS = os.path.join(RAIZ, 'propias.json')

PAISES = {  # nombre en mayúsculas sin tildes -> nombre del sitio
    'ARGENTINA': 'Argentina', 'BOLIVIA': 'Bolivia', 'BRASIL': 'Brasil', 'CHILE': 'Chile',
    'COLOMBIA': 'Colombia', 'CUBA': 'Cuba', 'ECUADOR': 'Ecuador', 'EL SALVADOR': 'El Salvador',
    'GUATEMALA': 'Guatemala', 'MEXICO': 'México', 'PARAGUAY': 'Paraguay', 'PERU': 'Perú',
    'URUGUAY': 'Uruguay', 'VENEZUELA': 'Venezuela', 'ESTADOS UNIDOS': 'Estados Unidos',
    'EE.UU.': 'Estados Unidos', 'EEUU': 'Estados Unidos', 'ESPANA': 'España', 'ALEMANIA': 'Alemania',
    'ISRAEL': 'Israel', 'REINO UNIDO': 'Reino Unido', 'INGLATERRA': 'Reino Unido',
}


def sin_tildes(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')


def slug(s):
    return re.sub(r'[^a-z0-9]+', '-', sin_tildes(s).lower()).strip('-')[:60]


def limpiar(s):
    s = re.sub(r'\[([^\]]*)\]\([^)]*\)', r'\1', s)          # enlaces -> texto
    s = re.sub(r'\{[^}]*\}', '', s)                          # atributos {...}
    s = s.replace('**', '').replace('__', '')
    s = re.sub(r'(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?![\w*])', r'\1', s)  # *itálica*
    s = re.sub(r'\\(.)', r'\1', s)                           # escapes de pandoc
    s = re.sub(r'\s*\n\s*', ' ', s)
    return s.strip()


def nombre_pais(s):
    k = sin_tildes(s).upper().strip()
    return PAISES.get(k, s.strip().title())


def seccion(pais, volanta):
    resto = [x.strip() for x in re.split(r'[·|]', volanta)[1:] if x.strip()]
    chicas = {'a', 'al', 'con', 'de', 'del', 'el', 'en', 'la', 'las', 'los', 'o', 'para', 'por', 'sin', 'y', 'e', 'u'}
    def frase(x):
        if not x.isupper(): return x
        ps = x.lower().split()
        return ' '.join(p if (k and p in chicas) else p.capitalize() for k, p in enumerate(ps))
    resto = [frase(x) for x in resto]
    return ' · '.join([pais] + resto)


def es_negrita(b):
    return bool(re.fullmatch(r'\*\*[^\n]+\*\*', b.strip()))


def imagen(b):
    m = re.match(r'!\[[^\]]*\]\(([^)\s]+)', b.strip()) or re.search(r'<img[^>]+src="([^"]+)"', b)
    return m.group(1) if m else None


def procesar(docx, previas):
    base = os.path.splitext(os.path.basename(docx))[0]
    nid = slug(base) or 'nota'
    tmp = tempfile.mkdtemp()
    fmt = ['-f', 'markdown'] if docx.lower().endswith('.md') else []
    md = subprocess.run(['pandoc', docx] + fmt + ['-t', 'markdown-smart', '--wrap=none', '--extract-media=' + tmp],
                        capture_output=True, text=True, check=True).stdout
    bloques = [b.strip() for b in re.split(r'\n\s*\n', md) if b.strip()]

    # Encabezado: volanta (PAÍS · SECCIÓN), título y bajada.
    i = 0
    volanta = limpiar(bloques[i]); i += 1
    if '·' not in volanta and '|' not in volanta:
        raise ValueError('El primer renglón tiene que ser la volanta: "PAÍS · SECCIÓN" (por ejemplo "REINO UNIDO · RELACIÓN CON EUROPA").')
    pais = nombre_pais(re.split(r'[·|]', volanta)[0])
    if i >= len(bloques):
        raise ValueError('Falta el título.')
    titulo = limpiar(re.sub(r'^#+\s*', '', bloques[i])); i += 1
    bajada = ''
    if i < len(bloques) and not imagen(bloques[i]) and not bloques[i].startswith('#'):
        bajada = limpiar(bloques[i]); i += 1

    fotos, partes, fuentes = [], [], []
    en_fuentes = False
    while i < len(bloques):
        b = bloques[i]; i += 1
        t = limpiar(re.sub(r'^#+\s*', '', b))
        if re.fullmatch(r'(fuentes( consultadas)?|fuentes y referencias)\s*:?', sin_tildes(t).lower()):
            en_fuentes = True; continue
        if en_fuentes:
            for linea in b.split('\n'):
                linea = limpiar(re.sub(r'^\s*([•\-*]|\d+[.)])\s*', '', linea))
                if linea: fuentes.append(linea)
            continue
        src = imagen(b)
        if src:
            epi = cred = ''
            if i < len(bloques) and not imagen(bloques[i]) and not bloques[i].startswith('#'):
                epi = limpiar(bloques[i]); i += 1
            if i < len(bloques) and re.match(r'(foto|imagen|crédito|credito)\b', sin_tildes(limpiar(bloques[i])).lower()):
                cred = limpiar(bloques[i]); i += 1
            if not cred and re.search(r'\b(foto|wikimedia|licencia|CC BY)\b', epi, re.I) and re.match(r'(foto|imagen)', epi, re.I):
                epi, cred = '', epi
            if not cred:
                raise ValueError(f'La foto {len(fotos)+1} no tiene crédito. Debajo del epígrafe tiene que ir un renglón que empiece con "Foto:" (autor, sitio y licencia).')
            if re.match(r'https?://', src):
                req = urllib.request.Request(src, headers={'User-Agent': 'PoliticaResumida/1.0 (https://politicaresumida.com.ar)'})
                archivo = os.path.join(tmp, f'web-{len(fotos)}')
                with urllib.request.urlopen(req, timeout=60) as r, open(archivo, 'wb') as out:
                    out.write(r.read())
                src = archivo
            im = Image.open(src if os.path.isabs(src) else os.path.join(tmp, src)).convert('RGB')
            w, h = im.size
            im.thumbnail((1400, 1400))
            nombre = f'img/{nid}-{len(fotos)+1}.jpg'
            im.save(os.path.join(RAIZ, nombre), quality=80, optimize=True, progressive=True)
            fotos.append({'src': nombre, 'epi': epi, 'cred': cred, 'vertical': h > w * 1.15})
            partes.append(f'[[foto:{len(fotos)-1}]]')
            continue
        if b.startswith('#') or (es_negrita(b) and len(t) < 90):
            partes.append('## ' + t); continue
        if t: partes.append(t)

    texto = [p for p in partes if not p.startswith('[[')]
    if not titulo or len(' '.join(texto).split()) < 80:
        raise ValueError('La nota quedó sin título o casi sin texto. Revisá que el Word tenga volanta, título, bajada y párrafos.')

    previa = next((n for n in previas if n['id'] == nid), None)
    hoy = (datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(hours=3)).strftime('%Y-%m-%d')
    return {'id': nid, 'p': pais, 'f': previa['f'] if previa else hoy, 'seccion': seccion(pais, volanta),
            'titulo': titulo, 'bajada': bajada, 'fotos': fotos, 'texto': '\n\n'.join(partes), 'fuentes': fuentes}


def main():
    notas = json.load(open(NOTAS, encoding='utf-8')) if os.path.exists(NOTAS) else []
    os.makedirs(os.path.join(RAIZ, 'img'), exist_ok=True)
    errores = 0
    archivos = glob.glob(os.path.join(RAIZ, 'entrada', '*.docx')) + glob.glob(os.path.join(RAIZ, 'entrada', '*.md'))
    for docx in sorted(archivos):
        nombre = os.path.basename(docx)
        try:
            n = procesar(docx, notas)
        except Exception as e:
            errores += 1
            os.makedirs(os.path.join(RAIZ, 'errores'), exist_ok=True)
            shutil.move(docx, os.path.join(RAIZ, 'errores', nombre))
            with open(os.path.join(RAIZ, 'errores', nombre + '.txt'), 'w', encoding='utf-8') as f:
                f.write(f'No se pudo publicar {nombre}.\n\nMotivo: {e}\n\nCorregí el Word y volvé a subirlo a la carpeta entrada.\n')
            print(f'ERROR {nombre}: {e}')
            continue
        notas = [x for x in notas if x['id'] != n['id']]
        notas.insert(0, n)
        os.makedirs(os.path.join(RAIZ, 'procesadas'), exist_ok=True)
        shutil.move(docx, os.path.join(RAIZ, 'procesadas', nombre))
        for viejo in glob.glob(os.path.join(RAIZ, 'errores', nombre + '*')):
            os.remove(viejo)
        print(f'OK {nombre} -> {n["p"]}: {n["titulo"]} ({len(n["fotos"])} fotos)')
    notas.sort(key=lambda x: x['f'], reverse=True)
    with open(NOTAS, 'w', encoding='utf-8') as f:
        json.dump(notas, f, ensure_ascii=False, indent=1)
    sys.exit(1 if errores else 0)


if __name__ == '__main__':
    main()
