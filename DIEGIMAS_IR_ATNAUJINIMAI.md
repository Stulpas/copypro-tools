# GitHub, atnaujinimų ir diegimo instrukcija

## 1. Sukurkite GitHub saugyklą

Rekomenduojamas pavadinimas:

`copypro-tools`

Paprasčiausia naudoti viešą saugyklą, nes programai nereikės saugoti GitHub
prieigos rakto.

## 2. Nurodykite GitHub paskyrą

Atidarykite `update_config.json` ir pakeiskite:

```json
"github_owner": "CHANGE_ME"
```

į savo GitHub naudotojo vardą. Jei saugyklos pavadinimas kitoks, pakeiskite ir
`github_repository`.

## 3. Įkelkite projektą

Atidarykite Command Prompt projekto aplanke:

```bat
git init
git add .
git commit -m "Pirmoji CopyPro Tools versija"
git branch -M main
git remote add origin https://github.com/JUSU_VARDAS/copypro-tools.git
git push -u origin main
```

## 4. Sukurkite pirmą versiją

```bat
git tag v1.0.0
git push origin v1.0.0
```

GitHub Actions automatiškai sukurs:

- `CopyPro-Tools-Setup.exe` – rekomenduojamas pirmam diegimui;
- `CopyPro-Tools-Windows.zip` – nešiojama versija;
- `CopyPro Tools.exe`;
- `CopyPro Updater.exe`.

Failai bus matomi GitHub saugyklos skiltyje `Releases`.

## 5. Pirmas diegimas darbo kompiuteryje

Rekomenduojama naudoti diegiklį:

1. Atsisiųskite `CopyPro-Tools-Setup.exe`.
2. Paleiskite failą.
3. Diegiklis nereikalauja administratoriaus teisių.
4. Programa įdiegiama į `%LOCALAPPDATA%\CopyPro Tools\`.
5. Sukuriama darbalaukio ir Start meniu nuoroda.
6. Paleidus pirmą kartą, redaguojami duomenys nukopijuojami į
   `%APPDATA%\CopyPro\`.

## 6. Naujo atnaujinimo paskelbimas

Pakeitus kodą arba numatytuosius duomenis:

```bat
git add .
git commit -m "Trumpas pakeitimų aprašymas"
git push
git tag v1.1.0
git push origin v1.1.0
```

Kiekvieną kartą naudokite didesnį versijos numerį:

- `v1.0.1` – mažas klaidos pataisymas;
- `v1.1.0` – nauja funkcija;
- `v2.0.0` – didelis pakeitimas.

Programa paleidimo metu patikrina naujausią GitHub Release. Radusi naujesnę
versiją, pasiūlo ją įdiegti. Atnaujintojas uždaro programą, pakeičia failus ir
paleidžia naują versiją. `%APPDATA%\CopyPro\` duomenys nepašalinami.

## 7. Kainų ir kodų keitimas be programos atnaujinimo

Darbo kompiuteryje atidarykite:

`%APPDATA%\CopyPro\copypro_kodai.csv`

Pakeiskite kainą ar pavadinimą ir išsaugokite CSV UTF-8 formatu. Programos
nustatymuose galima pasirinkti ir kitą CSV failo vietą.

Svarbu: numatytųjų CSV pakeitimas GitHub projekte nepakeičia darbuotojų jau
redaguotų AppData kopijų. Jei norite centralizuotai pakeisti aktyvius kainynus,
reikia pakeisti AppData failą arba naudoti bendrą tinklo CSV kelią programos
nustatymuose.
