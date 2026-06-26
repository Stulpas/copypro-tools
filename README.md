# CopyPro Tools

Šis projektas paruoštas taip, kad kasdieniai duomenų pakeitimai nereikalautų
keisti Python kodo.

## Svarbiausi failai

- `data/copypro_kodai.csv` – visi PLU kodai, originalūs lietuviški pavadinimai,
  kainos, kategorijos ir plataus formato pasiūlymo duomenys.
- `data/copypro_popieriaus_dydziai.csv` – numatytieji ir naudotojo sukurti
  popieriaus dydžiai.
- `copypro_tools.py` – pagrindinė programa.
- `copypro_updater.py` – atskiras atnaujintojas.
- `copypro_update_support.py` – atnaujinimų ir pirmojo paleidimo logika.
- `.github/workflows/build-release.yml` – automatinis Windows programos,
  ZIP ir diegiklio kūrimas.
- `installer/CopyProTools.iss` – pirmojo diegimo programos nustatymai.

## Duomenų saugojimas darbo kompiuteryje

Programa ir atnaujinami vykdomieji failai:

`%LOCALAPPDATA%\CopyPro Tools\`

Nustatymai, kainos, PLU kodai, popieriaus dydžiai ir naudotojo išsaugoti maketai:

`%APPDATA%\CopyPro\`

Šie aplankai nėra Desktop ar Downloads aplankuose.

## CSV stulpeliai

### copypro_kodai.csv

- `category` – kategorija lietuviškai.
- `code` – PLU kodas.
- `name_lt` – originalus lietuviškas pavadinimas.
- `price_eur` – vieneto kaina, naudojamas taškas, pvz. `2.50`.
- `search_aliases_lt` – papildomi lietuviški paieškos žodžiai, atskirti kableliais.
- `wide_format_coverage` – `dalinis`, `pilnas`, `brėžinys`, `fiksuotas` arba tuščia.
- `wide_format_label_lt` – trumpas lietuviškas pavadinimas plataus formato pasiūlyme.
- `active` – `taip` arba `ne`.
- `notes_lt` – pastabos lietuviškai.

### copypro_popieriaus_dydziai.csv

- `group` – dydžių grupė.
- `name` – dydžio pavadinimas.
- `width_mm` – plotis milimetrais.
- `height_mm` – aukštis milimetrais.
- `active` – `taip` arba `ne`.
- `is_custom` – `taip`, jei dydį sukūrė naudotojas.
- `notes_lt` – pastabos lietuviškai.

## Prekių ženklai ir trumpiniai

Prekių ženklai ir techniniai trumpiniai palikti originalūs, pavyzdžiui:
`CopyPro`, `Pioneer Navigator`, `Curious Collection`, `INEO`, `USB`, `CD`,
`DVD`, `PVC`, `SRA3`, `A4`.


## Sticker Outline

Skirtukas `Sticker Outline` aptinka ne baltą objekto kraštą ir eksportuoja tikrą
vektorinę raudoną liniją. Galima keisti balto fono toleranciją, kontūro
gludinimą, mažiausią objekto plotą, poslinkį milimetrais ir vidinių baltų
sričių įtraukimą.

Eksportas:

- tik kontūras – vienas PDF arba SVG;
- vaizdas ir kontūras – vienas PDF arba SVG su atskirais `Artwork` ir
  `Cut line` sluoksniais / grupėmis.

PDF šaltinio vaizdas PDF eksporte išlieka vektorinis. Įterpiami rastriniai
vaizdai nemažinami ir papildomai nuostolingai nesuspaudžiami. Kontūro spalva –
RGB 255, 0, 0.
