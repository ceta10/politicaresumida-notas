# Notas de Política Resumida

Acá viven las notas propias que se publican en [politicaresumida.com.ar](https://politicaresumida.com.ar).

## Cómo publicar una nota

1. Abrí la carpeta **`entrada`**.
2. Tocá **Add file → Upload files** y arrastrá el Word (`.docx`) de la nota.
3. Tocá **Commit changes**.

En uno o dos minutos la nota aparece en la sección del país (el sitio puede tardar hasta 5 minutos más en mostrarla). El Word se mueve solo a `procesadas/`.

- **Corregir una nota:** subí el Word corregido **con el mismo nombre de archivo**. Reemplaza la anterior.
- **Si algo falla:** el Word pasa a `errores/` junto con un `.txt` que explica qué corregir, y GitHub te manda un mail.

## Cómo tiene que ser el Word

1. **Volanta** (primer renglón, en negrita): `PAÍS · SECCIÓN`, por ejemplo `REINO UNIDO · RELACIÓN CON EUROPA`. El país define en qué sección del sitio aparece.
2. **Título** (segundo renglón, en negrita).
3. **Bajada**: un párrafo que resume la nota.
4. **Texto**, con subtítulos en negrita o con estilo de título.
5. **Fotos** pegadas dentro del Word. Debajo de cada una, el **epígrafe** y en el renglón siguiente el **crédito**, empezando con `Foto:` (autor, sitio y licencia). Solo fotos con licencia libre (por ejemplo, de Wikimedia Commons).
6. Al final, opcional: un renglón **Fuentes consultadas** y debajo la lista.

## Archivos

- `propias.json`: las notas publicadas (lo lee el sitio). No hace falta tocarlo.
- `img/`: las fotos de las notas.
- `procesar.py`: el programa que convierte los Word.
