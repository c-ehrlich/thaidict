# Volubilis Thai-English for Yomitan

An importable Yomitan dictionary built from the Volubilis Multilingual Thai
Dictionary and Database, edition 26.2 (July 2026).

The current build contains:

- 100,827 Thai-English dictionary entries
- 111,407 English senses
- 4,801 indexed alternate Thai headwords
- 65 part-of-speech tags
- 11 Yomitan term banks

The dictionary includes pronunciations, parts of speech, usage notes,
classifiers, synonyms, domains, scientific names, etymologies,
syllabification, and notes when the source provides them.

Import [`dist/volubilis-th-en-26.2-yomitan.zip`](dist/volubilis-th-en-26.2-yomitan.zip)
into Yomitan.

## Build and test

The builder uses the Python standard library. Run these commands from the
repository root:

```sh
python3 build.py
python3 -m unittest discover -s . -p 'test_*.py'
```

The first build downloads the source workbook and verifies its SHA-256
checksum. To use a local workbook, run:

```sh
python3 build.py --source /path/to/VOLUBILIS\ Database.xlsx
```

## License and attribution

Volubilis is by Francis Bastien (Belisan) and is licensed under
[Creative Commons Attribution-ShareAlike 4.0 International](https://creativecommons.org/licenses/by-sa/4.0/).
See the [publisher's project page](https://belisan-volubilis.blogspot.com/2007/)
for the source data.

The generated dictionary is adapted material under the same license. Preserve
the attribution and license metadata in `index.json` when redistributing it.

