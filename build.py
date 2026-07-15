#!/usr/bin/env python3
"""Build a clean Thai-English Yomitan dictionary from the Volubilis XLSX database."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import sys
import unicodedata
import urllib.request
import zipfile
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Iterator, Mapping, Sequence
from xml.etree import ElementTree


TOOL_DIR = Path(__file__).resolve().parent
DEFAULT_CACHE_DIR = TOOL_DIR / ".cache"
DEFAULT_DIST_DIR = TOOL_DIR / "dist"
SOURCE_FILENAME = "volubilis-database-26.2.xlsx"
SOURCE_URL = "https://downloads.sourceforge.net/project/belisan/VOLUBILIS%20Database.xlsx"
SOURCE_SHA256 = "b9ab74187a1c369d03bf1a0b94cdc0523edb77a4da72759ee85d81626a20fc0c"
SOURCE_PAGE_URL = "https://belisan-volubilis.blogspot.com/2007/"
LICENSE_URL = "https://creativecommons.org/licenses/by-sa/4.0/"
DICTIONARY_TITLE = "Volubilis Thai-English"
CONVERTER_REVISION = 1
BANK_SIZE = 10_000
THAI_RE = re.compile(r"[\u0e00-\u0e7f]")
CELL_REFERENCE_RE = re.compile(r"[A-Z]+")
XLSX_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
ALT_DELIMITERS = frozenset({";", "="})
ZIP_TIMESTAMP = (1980, 1, 1, 0, 0, 0)


@dataclass(frozen=True)
class SourceMetadata:
    name: str
    revision: str
    edition_date: str
    advertised_entries: str


@dataclass(frozen=True)
class Sense:
    english: str
    part_of_speech: str = ""
    usage: str = ""
    scientific: str = ""
    domain: str = ""
    classifier: str = ""
    synonyms: str = ""
    etymology: str = ""
    syllables: str = ""
    note: str = ""
    other_forms: tuple[str, ...] = ()


@dataclass
class BuildReport:
    source_url: str
    source_sha256: str
    source_revision: str
    source_edition_date: str
    source_rows: int = 0
    bilingual_rows: int = 0
    skipped_without_english: int = 0
    skipped_non_thai_headword: int = 0
    emitted_entries: int = 0
    emitted_senses: int = 0
    alternate_headwords_added: int = 0
    exact_duplicate_senses_removed: int = 0
    term_banks: int = 0
    part_of_speech_tags: int = 0
    archive_sha256: str = ""


def normalize_text(value: str) -> str:
    """Normalize source strings while preserving meaningful internal whitespace."""
    value = value.replace("\u00a0", " ").replace("\r\n", "\n").replace("\r", "\n")
    value = "\n".join(part.strip() for part in value.split("\n"))
    value = re.sub(r"[ \t]+", " ", value).strip()
    return unicodedata.normalize("NFC", value)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def column_index(cell_reference: str) -> int:
    match = CELL_REFERENCE_RE.match(cell_reference)
    if match is None:
        raise ValueError(f"Invalid XLSX cell reference: {cell_reference!r}")
    result = 0
    for character in match.group(0):
        result = result * 26 + ord(character) - ord("A") + 1
    return result - 1


def read_shared_strings(workbook: zipfile.ZipFile) -> list[str]:
    strings: list[str] = []
    with workbook.open("xl/sharedStrings.xml") as source:
        for _, element in ElementTree.iterparse(source, events=("end",)):
            if element.tag != XLSX_NS + "si":
                continue
            strings.append("".join(node.text or "" for node in element.iter(XLSX_NS + "t")))
            element.clear()
    return strings


def first_worksheet_path(workbook: zipfile.ZipFile) -> str:
    candidates = sorted(
        name
        for name in workbook.namelist()
        if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")
    )
    if not candidates:
        raise ValueError("The XLSX file does not contain a worksheet")
    return candidates[0]


def iter_xlsx_rows(path: Path) -> Iterator[tuple[int, dict[int, str]]]:
    """Yield sparse, zero-indexed cells from the first XLSX worksheet."""
    with zipfile.ZipFile(path) as workbook:
        shared_strings = read_shared_strings(workbook)
        worksheet_path = first_worksheet_path(workbook)
        with workbook.open(worksheet_path) as worksheet:
            for _, row in ElementTree.iterparse(worksheet, events=("end",)):
                if row.tag != XLSX_NS + "row":
                    continue
                row_number = int(row.attrib["r"])
                values: dict[int, str] = {}
                for cell in row.findall(XLSX_NS + "c"):
                    reference = cell.attrib.get("r")
                    if reference is None:
                        continue
                    value_node = cell.find(XLSX_NS + "v")
                    value = "" if value_node is None else value_node.text or ""
                    cell_type = cell.attrib.get("t")
                    if cell_type == "s" and value:
                        value = shared_strings[int(value)]
                    elif cell_type == "inlineStr":
                        value = "".join(node.text or "" for node in cell.iter(XLSX_NS + "t"))
                    values[column_index(reference)] = normalize_text(value)
                yield row_number, values
                row.clear()


def canonical_header(value: str) -> str | None:
    header = value.split("\n", 1)[0].strip().upper()
    matchers = (
        ("THAIROM", "thai_romanized"),
        ("EASYTHAI", "easy_thai"),
        ("THAIPHON", "thai_phonetic"),
        ("ETYMO", "etymology"),
        ("THA (", "thai"),
        ("ENG (", "english"),
        ("TYPE", "part_of_speech"),
        ("USAGE", "usage"),
        ("SCIENT", "scientific"),
        ("DOM", "domain"),
        ("CLASSIF", "classifier"),
        ("SYLLAB", "syllables"),
        ("NOTE", "note"),
        ("SYN", "synonyms"),
        ("LEVEL", "level"),
    )
    for prefix, key in matchers:
        if header.startswith(prefix):
            return key
    return None


def parse_source_metadata(row: Mapping[int, str]) -> SourceMetadata:
    values = [value for _, value in sorted(row.items()) if value]
    joined = " | ".join(values)
    revision_match = re.search(r"v\.\s*([0-9.]+)\s*\(([^)]+)\)", joined, flags=re.IGNORECASE)
    if revision_match is None:
        revision_match = re.search(r"v\.\s*([0-9.]+)", joined, flags=re.IGNORECASE)
    revision = revision_match.group(1) if revision_match else "unknown"
    edition_date = revision_match.group(2) if revision_match and revision_match.lastindex == 2 else "unknown"
    advertised = next((value for value in values if "entr" in value.lower()), "unknown")
    name = values[0] if values else "Volubilis"
    return SourceMetadata(name=name, revision=revision, edition_date=edition_date, advertised_entries=advertised)


def split_top_level_alternatives(value: str) -> list[str]:
    """Split semicolon/equal alternatives, but not delimiters nested in brackets."""
    value = normalize_text(value)
    if not value:
        return []
    parts: list[str] = []
    start = 0
    depth = 0
    pairs = {"(": ")", "[": "]", "{": "}"}
    closing = set(pairs.values())
    for index, character in enumerate(value):
        if character in pairs:
            depth += 1
        elif character in closing and depth:
            depth -= 1
        elif character in ALT_DELIMITERS and depth == 0:
            part = normalize_text(value[start:index])
            if part:
                parts.append(part)
            start = index + 1
    final = normalize_text(value[start:])
    if final:
        parts.append(final)
    return parts or [value]


def split_thai_headwords(value: str) -> list[str]:
    candidates = split_top_level_alternatives(value)
    if len(candidates) > 1 and all(THAI_RE.search(candidate) for candidate in candidates):
        return list(dict.fromkeys(candidates))
    return [normalize_text(value)]


def tone_marked_romanization(value: str) -> str:
    """Convert Volubilis THAIPHON tone prefixes to readable Unicode accents."""
    value = normalize_text(value).lower()
    if not value:
        return ""

    # Preserve Volubilis's two open-o values before expanding long vowels.
    value = value.replace("ǿ", "\u0000").replace("ø̅", "ɔ̄").replace("ø", "ɔ").replace("\u0000", "ø")
    for source, target in {
        "ē": "ee",
        "ā": "aa",
        "ī": "ii",
        "ū": "uu",
        "ō": "oo",
        "ɔ̄": "ɔɔ",
    }.items():
        value = value.replace(source, target)

    accents = {"¯": "\u0301", "\\": "\u0302", "/": "\u030c", "_": "\u0300", "-": ""}
    tone_pattern = re.compile(r"([¯\\/_-])([bcdfghjklmnpqrstvwxyz]*)([aeiouɔø])", flags=re.IGNORECASE)

    def apply_tone(match: re.Match[str]) -> str:
        marker, consonants, vowel = match.groups()
        return consonants + vowel + accents[marker]

    value = tone_pattern.sub(apply_tone, value)
    value = re.sub(r"\s+", " ", value).strip()
    return unicodedata.normalize("NFC", value)


def pair_headwords_and_readings(thai: str, pronunciation: str) -> list[tuple[str, str, tuple[str, ...]]]:
    headwords = split_thai_headwords(thai)
    raw_readings = split_top_level_alternatives(pronunciation)
    readings = [tone_marked_romanization(value) for value in raw_readings]
    if not readings:
        readings = [""]
    if len(readings) == len(headwords):
        paired_readings = readings
    elif len(readings) == 1:
        paired_readings = readings * len(headwords)
    else:
        paired_readings = [tone_marked_romanization(pronunciation)] * len(headwords)
    result = []
    for index, headword in enumerate(headwords):
        other_forms = tuple(candidate for candidate in headwords if candidate != headword)
        result.append((headword, paired_readings[index], other_forms))
    return result


def tag_slug(part_of_speech: str) -> str:
    value = part_of_speech.lower().strip()
    value = re.sub(r"[^a-z0-9]+", "-", value).strip("-")
    return f"pos-{value or 'unspecified'}"


def span(content: str, *, bold: bool = False, italic: bool = False, lang: str | None = None) -> dict[str, Any]:
    node: dict[str, Any] = {"tag": "span", "content": content}
    style: dict[str, str] = {}
    if bold:
        style["fontWeight"] = "bold"
    if italic:
        style["fontStyle"] = "italic"
    if style:
        node["style"] = style
    if lang:
        node["lang"] = lang
    return node


def labelled_line(label: str, value: str, *, lang: str | None = None, italic: bool = False) -> dict[str, Any]:
    return {
        "tag": "div",
        "content": [span(label + ": ", bold=True), span(value, italic=italic, lang=lang)],
        "style": {"fontSize": "0.9em", "marginTop": "0.15em"},
    }


def sense_content(sense: Sense) -> list[Any]:
    content: list[Any] = [
        {"tag": "div", "content": span(sense.english, bold=True, lang="en"), "style": {"marginBottom": "0.2em"}}
    ]
    fields = (
        ("Part of speech", sense.part_of_speech, None, True),
        ("Usage", sense.usage, None, False),
        ("Classifier", sense.classifier, "th", False),
        ("Synonyms", sense.synonyms, "th", False),
        ("Other forms", "; ".join(sense.other_forms), "th", False),
        ("Domain", sense.domain, None, False),
        ("Scientific name", sense.scientific, None, True),
        ("Etymology", sense.etymology, None, False),
        ("Syllables", sense.syllables, "th", False),
        ("Note", sense.note, None, False),
    )
    for label, value, lang, italic in fields:
        if value:
            content.append(labelled_line(label, value, lang=lang, italic=italic))
    return content


def structured_glossary(senses: Sequence[Sense]) -> dict[str, Any]:
    if len(senses) == 1:
        content: Any = {"tag": "div", "content": sense_content(senses[0])}
    else:
        content = {
            "tag": "ol",
            "content": [
                {"tag": "li", "content": sense_content(sense), "style": {"marginBottom": "0.65em"}}
                for sense in senses
            ],
            "style": {"marginTop": 0, "marginBottom": 0, "paddingLeft": "1.5em"},
        }
    return {"type": "structured-content", "content": content}


def source_record(values: Mapping[str, str], other_forms: tuple[str, ...]) -> Sense:
    return Sense(
        english=values["english"],
        part_of_speech=values.get("part_of_speech", ""),
        usage=values.get("usage", ""),
        scientific=values.get("scientific", ""),
        domain=values.get("domain", ""),
        classifier=values.get("classifier", ""),
        synonyms=values.get("synonyms", ""),
        etymology=values.get("etymology", ""),
        syllables=values.get("syllables", ""),
        note=values.get("note", ""),
        other_forms=other_forms,
    )


def load_source(path: Path) -> tuple[SourceMetadata, dict[tuple[str, str], list[Sense]], Counter[str], BuildReport]:
    rows = iter_xlsx_rows(path)
    try:
        first_number, first_row = next(rows)
        second_number, second_row = next(rows)
    except StopIteration as error:
        raise ValueError("The source workbook does not contain metadata and header rows") from error
    if (first_number, second_number) != (1, 2):
        raise ValueError("Expected source metadata on row 1 and column headers on row 2")

    metadata = parse_source_metadata(first_row)
    columns = {
        canonical: index
        for index, value in second_row.items()
        if (canonical := canonical_header(value)) is not None
    }
    missing_columns = {"thai", "english"} - columns.keys()
    if missing_columns:
        raise ValueError(f"Missing required source columns: {', '.join(sorted(missing_columns))}")

    report = BuildReport(
        source_url=SOURCE_URL,
        source_sha256=sha256_file(path),
        source_revision=metadata.revision,
        source_edition_date=metadata.edition_date,
    )
    grouped: dict[tuple[str, str], list[Sense]] = defaultdict(list)
    seen: dict[tuple[str, str], set[Sense]] = defaultdict(set)
    part_of_speech_counts: Counter[str] = Counter()

    for _, sparse_row in rows:
        report.source_rows += 1
        values = {
            name: sparse_row.get(index, "")
            for name, index in columns.items()
        }
        thai = values.get("thai", "")
        english = values.get("english", "")
        if not thai or not english:
            report.skipped_without_english += 1
            continue
        if THAI_RE.search(thai) is None:
            report.skipped_non_thai_headword += 1
            continue
        report.bilingual_rows += 1
        pairs = pair_headwords_and_readings(thai, values.get("thai_phonetic", ""))
        report.alternate_headwords_added += max(0, len(pairs) - 1)
        for expression, reading, other_forms in pairs:
            sense = source_record(values, other_forms)
            key = (expression, reading)
            if sense in seen[key]:
                report.exact_duplicate_senses_removed += 1
                continue
            seen[key].add(sense)
            grouped[key].append(sense)
            part_of_speech_counts[sense.part_of_speech] += 1

    return metadata, grouped, part_of_speech_counts, report


def make_index(metadata: SourceMetadata) -> dict[str, Any]:
    revision = f"{metadata.revision}-yomitan.{CONVERTER_REVISION}"
    attribution = (
        "Volubilis Multilingual Thai Dictionary & Database by Francis Bastien (Belisan), "
        f"licensed under CC BY-SA 4.0 ({LICENSE_URL}). Source: {SOURCE_PAGE_URL} "
        "This adapted Yomitan edition normalizes Unicode and whitespace, selects Thai-English records, "
        "indexes alternate Thai spellings, converts THAIPHON tone notation to Unicode-accented romanization, "
        "groups duplicate headwords by pronunciation, removes exact duplicate senses, and uses structured content."
    )
    return {
        "title": DICTIONARY_TITLE,
        "revision": revision,
        "format": 3,
        "sequenced": False,
        "author": "Francis Bastien (Belisan); Yomitan conversion by the Axiom project",
        "url": SOURCE_PAGE_URL,
        "description": (
            f"Thai-English edition derived from {metadata.name} v{metadata.revision} "
            f"({metadata.edition_date}; {metadata.advertised_entries})."
        ),
        "attribution": attribution,
        "sourceLanguage": "th",
        "targetLanguage": "en",
    }


def make_tag_bank(part_of_speech_counts: Counter[str]) -> list[list[Any]]:
    grouped: dict[str, list[str]] = defaultdict(list)
    totals: Counter[str] = Counter()
    for value, count in part_of_speech_counts.items():
        slug = tag_slug(value)
        grouped[slug].append(value or "Unspecified part of speech")
        totals[slug] += count
    slugs = sorted(grouped, key=lambda slug: (-totals[slug], slug))
    return [
        [slug, "partOfSpeech", order, " / ".join(sorted(grouped[slug], key=str.lower)), 0]
        for order, slug in enumerate(slugs)
    ]


def make_term_entries(grouped: Mapping[tuple[str, str], Sequence[Sense]]) -> list[list[Any]]:
    entries: list[list[Any]] = []
    for (expression, reading), senses in sorted(grouped.items(), key=lambda item: item[0]):
        definition_tags = " ".join(sorted({tag_slug(sense.part_of_speech) for sense in senses}))
        entries.append(
            [
                expression,
                reading,
                definition_tags,
                "",
                0,
                [structured_glossary(senses)],
                0,
                "",
            ]
        )
    return entries


def json_bytes(value: Any) -> bytes:
    return (json.dumps(value, ensure_ascii=False, separators=(",", ":")) + "\n").encode("utf-8")


def write_deterministic_zip(path: Path, files: Mapping[str, bytes]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for filename in sorted(files):
            info = zipfile.ZipInfo(filename=filename, date_time=ZIP_TIMESTAMP)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o100644 << 16
            info.create_system = 3
            archive.writestr(info, files[filename], compress_type=zipfile.ZIP_DEFLATED, compresslevel=9)


def validate_structured_content(content: Any) -> None:
    if isinstance(content, str):
        return
    if isinstance(content, list):
        for child in content:
            validate_structured_content(child)
        return
    if not isinstance(content, dict) or "tag" not in content:
        raise ValueError(f"Invalid structured content node: {content!r}")
    allowed_tags = {"br", "span", "div", "ol", "ul", "li", "details", "summary"}
    if content["tag"] not in allowed_tags:
        raise ValueError(f"Unsupported structured content tag: {content['tag']!r}")
    if "content" in content:
        validate_structured_content(content["content"])


def validate_package(path: Path) -> dict[str, int]:
    with zipfile.ZipFile(path) as archive:
        names = sorted(archive.namelist())
        if "index.json" not in names:
            raise ValueError("Dictionary archive is missing index.json")
        index = json.loads(archive.read("index.json"))
        for field in ("title", "revision", "attribution", "sourceLanguage", "targetLanguage"):
            if not isinstance(index.get(field), str) or not index[field]:
                raise ValueError(f"index.json has an invalid {field!r} field")
        if index.get("format") != 3 or index.get("sourceLanguage") != "th" or index.get("targetLanguage") != "en":
            raise ValueError("index.json has incompatible format or language metadata")

        tag_files = [name for name in names if re.fullmatch(r"tag_bank_\d+\.json", name)]
        term_files = [name for name in names if re.fullmatch(r"term_bank_\d+\.json", name)]
        if not tag_files or not term_files:
            raise ValueError("Dictionary archive is missing tag or term banks")

        defined_tags: set[str] = set()
        for filename in tag_files:
            bank = json.loads(archive.read(filename))
            for row in bank:
                if not isinstance(row, list) or len(row) != 5 or not isinstance(row[0], str):
                    raise ValueError(f"Invalid tag row in {filename}: {row!r}")
                if row[0] in defined_tags:
                    raise ValueError(f"Duplicate tag name: {row[0]}")
                defined_tags.add(row[0])

        entry_count = 0
        expression_reading_pairs: set[tuple[str, str]] = set()
        for filename in term_files:
            bank = json.loads(archive.read(filename))
            if len(bank) > BANK_SIZE:
                raise ValueError(f"{filename} exceeds the configured bank size")
            for row in bank:
                if not isinstance(row, list) or len(row) != 8:
                    raise ValueError(f"Invalid term row in {filename}: {row!r}")
                expression, reading, definition_tags, rules, score, glossary, sequence, term_tags = row
                if not isinstance(expression, str) or not expression or THAI_RE.search(expression) is None:
                    raise ValueError(f"Invalid Thai expression in {filename}: {expression!r}")
                if not isinstance(reading, str) or not isinstance(definition_tags, str):
                    raise ValueError(f"Invalid reading or tags for {expression!r}")
                if (expression, reading) in expression_reading_pairs:
                    raise ValueError(f"Duplicate expression/reading pair: {(expression, reading)!r}")
                expression_reading_pairs.add((expression, reading))
                if not set(definition_tags.split()).issubset(defined_tags):
                    raise ValueError(f"Undefined tag on {expression!r}: {definition_tags!r}")
                if rules != "" or score != 0 or sequence != 0 or term_tags != "":
                    raise ValueError(f"Unexpected term metadata on {expression!r}")
                if not isinstance(glossary, list) or not glossary:
                    raise ValueError(f"Missing glossary for {expression!r}")
                for definition in glossary:
                    if definition.get("type") != "structured-content":
                        raise ValueError(f"Non-structured definition for {expression!r}")
                    validate_structured_content(definition["content"])
                entry_count += 1

    return {"entries": entry_count, "tags": len(defined_tags), "term_banks": len(term_files)}


def download_source(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary_path = path.with_suffix(path.suffix + ".part")
    request = urllib.request.Request(SOURCE_URL, headers={"User-Agent": "volubilis-yomitan-builder/1"})
    print(f"Downloading {SOURCE_URL}")
    try:
        with urllib.request.urlopen(request) as response, temporary_path.open("wb") as destination:
            shutil.copyfileobj(response, destination)
        temporary_path.replace(path)
    finally:
        temporary_path.unlink(missing_ok=True)


def resolve_source(source: Path | None, cache_dir: Path) -> Path:
    if source is not None:
        if not source.is_file():
            raise FileNotFoundError(source)
        return source
    cached_source = cache_dir / SOURCE_FILENAME
    if not cached_source.exists():
        download_source(cached_source)
    actual_hash = sha256_file(cached_source)
    if actual_hash != SOURCE_SHA256:
        raise ValueError(
            "The publisher download does not match the reviewed Volubilis 26.2 source. "
            f"Expected {SOURCE_SHA256}, received {actual_hash}. Remove {cached_source} and retry, "
            "or inspect the publisher's update before changing SOURCE_SHA256."
        )
    return cached_source


def build(source: Path, dist_dir: Path) -> tuple[Path, BuildReport]:
    metadata, grouped, part_of_speech_counts, report = load_source(source)
    index = make_index(metadata)
    tag_bank = make_tag_bank(part_of_speech_counts)
    terms = make_term_entries(grouped)

    files: dict[str, bytes] = {
        "index.json": json_bytes(index),
        "tag_bank_1.json": json_bytes(tag_bank),
    }
    for bank_number, start in enumerate(range(0, len(terms), BANK_SIZE), start=1):
        files[f"term_bank_{bank_number}.json"] = json_bytes(terms[start : start + BANK_SIZE])

    dist_dir.mkdir(parents=True, exist_ok=True)
    output_path = dist_dir / f"volubilis-th-en-{metadata.revision}-yomitan.zip"
    write_deterministic_zip(output_path, files)
    validation = validate_package(output_path)

    report.emitted_entries = validation["entries"]
    report.emitted_senses = sum(len(senses) for senses in grouped.values())
    report.term_banks = validation["term_banks"]
    report.part_of_speech_tags = validation["tags"]
    report.archive_sha256 = sha256_file(output_path)
    report_path = dist_dir / "build-report.json"
    report_path.write_text(json.dumps(asdict(report), ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return output_path, report


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path, help="Use a local Volubilis XLSX instead of the pinned publisher download")
    parser.add_argument("--cache-dir", type=Path, default=DEFAULT_CACHE_DIR)
    parser.add_argument("--dist-dir", type=Path, default=DEFAULT_DIST_DIR)
    parser.add_argument("--validate", type=Path, help="Validate an existing Yomitan ZIP and exit")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    try:
        if args.validate:
            summary = validate_package(args.validate)
            print(json.dumps(summary, indent=2))
            return 0
        source = resolve_source(args.source, args.cache_dir)
        output, report = build(source, args.dist_dir)
        print(json.dumps(asdict(report), ensure_ascii=False, indent=2))
        print(f"Created {output}")
        return 0
    except (OSError, ValueError, zipfile.BadZipFile) as error:
        print(f"error: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
