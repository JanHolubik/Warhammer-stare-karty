from __future__ import annotations

from pathlib import Path
from io import BytesIO
from typing import Dict
import re

import pandas as pd
from docx import Document

VAT_RATE = 1.21


BASE_DIR = Path(__file__).resolve().parent
PROMPT_TEMPLATE_DIR = BASE_DIR / "prompt_templates"
TEMPLATE_DIR = BASE_DIR / "sablony"


def make_docx_bytes(text: str) -> bytes:
    doc = Document()
    for line in text.splitlines():
        doc.add_paragraph(line)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def pick_existing(template_dir: Path, candidates: list[str]) -> Path:
    for name in candidates:
        path = template_dir / name
        if path.exists():
            return path
    raise FileNotFoundError(
        f"Nenalezena žádná z těchto šablon ve složce {template_dir}: {candidates}"
    )


def get_template_paths(template_dir: Path, template_type: str) -> Dict[str, Path]:
    template_type = template_type.strip().lower()

    short_map = {
        "miniatures": ["kratky.docx"],
        "books": ["kratky_univ.docx"],
        "warscroll": ["kratky_univ.docx"],
        "dice": ["kratky.docx"],
        "upgrades": ["kratky.docx"],
        "accessories": ["kratky_univ.docx"],
    }

    detail_map = {
        "miniatures": ["detailni.docx"],
        "books": ["detailni_kniha.docx"],
        "warscroll": ["detailni_warscoll.docx", "detailni_warscroll.docx"],
        "dice": ["detailni_kostky.docx"],
        "upgrades": ["detailni_upgrades.docx"],
        "accessories": ["stetce_pris.docx"],
    }

    short_map_en = {
        "miniatures": ["kratky_en.docx"],
        "books": ["kratky_univ_en.docx"],
        "warscroll": ["kratky_univ_en.docx"],
        "dice": ["kratky_en.docx"],
        "upgrades": ["kratky_en.docx"],
        "accessories": ["kratky_univ_en.docx"],
    }

    detail_map_en = {
        "miniatures": ["detailni_en.docx"],
        "books": ["detailni_kniha_en.docx"],
        "warscroll": ["detailni_warscoll_en.docx", "detailni_warscroll_en.docx"],
        "dice": ["detailni_kostky_en.docx"],
        "upgrades": ["detailni_upgrades_en.docx"],
        "accessories": ["stetce_pris en.docx"],
    }

    short_map_sk = {
        "miniatures": ["kratky_sk.docx"],
        "books": ["kratky_univ_sk.docx"],
        "warscroll": ["kratky_univ_sk.docx"],
        "dice": ["kratky_sk.docx"],
        "upgrades": ["kratky_sk.docx"],
        "accessories": ["kratky_univ_sk.docx"],
    }

    detail_map_sk = {
        "miniatures": ["detailni_sk.docx"],
        "books": ["detailni_kniha_sk.docx"],
        "warscroll": ["detailni_warscoll_sk.docx", "detailni_warscroll_sk.docx"],
        "dice": ["detailni_kostky_sk.docx"],
        "upgrades": ["detailni_upgrades_sk.docx"],
        "accessories": ["stetce_pris sk.docx"],
    }

    if template_type not in short_map:
        raise ValueError(f"Neznámý TEMPLATE_TYPE: {template_type}")

    return {
        "short_cs": pick_existing(template_dir, short_map[template_type]),
        "detail_cs": pick_existing(template_dir, detail_map[template_type]),
        "short_en": pick_existing(template_dir, short_map_en[template_type]),
        "detail_en": pick_existing(template_dir, detail_map_en[template_type]),
        "short_sk": pick_existing(template_dir, short_map_sk[template_type]),
        "detail_sk": pick_existing(template_dir, detail_map_sk[template_type]),
    }


def create_kartovani_card_row(
    name: str,
    code: str,
    external_code: str,
    ean: str,
    price: float,
    product_type: str,
    description: str = "",
    image_urls: list[str] | None = None,
) -> pd.DataFrame:
    image_urls = image_urls or []
    price_without_vat = round(float(price) / VAT_RATE, 2) if price else 0.0

    row = {
        "code": code,
        "pairCode": "",
        "externalCode": external_code.strip() or code,
        "name": name,
        "ean": ean,
        "price": price,
        "priceWithoutVat": price_without_vat,
        "description": description,
        "manufacturer": "",
        "availabilityInStock": "Skladem",
        "availabilityOutOfStock": "Na dotaz",
        "itemType": product_type,
        "productVisibility": "visible",
    }

    for idx in range(1, 11):
        col = "image" if idx == 1 else f"image{idx}"
        row[col] = image_urls[idx - 1] if idx <= len(image_urls) else ""

    return pd.DataFrame([row])


def create_kartovani_source_row(
    name: str,
    code: str,
    external_code: str,
    ean: str,
    price: float,
    intro_image_src: str = "",
    image_urls: list[str] | None = None,
    video_url: str = "",
    template_kind: str = "miniatures",
    prompt_type: str = "miniatures",
) -> pd.DataFrame:
    image_urls = image_urls or []
    price_without_vat = round(float(price) / VAT_RATE, 2) if price else 0.0

    row = {
        "code": code,
        "pairCode": "",
        "externalCode": external_code.strip() or code,
        "product_name": name,
        "name": name,
        "ean": ean,
        "price": price,
        "priceWithoutVat": price_without_vat,
        "description": "",
        "manufacturer": "",
        "availabilityInStock": "Skladem",
        "availabilityOutOfStock": "Na dotaz",
        "itemType": "product",
        "intro_image_src": intro_image_src,
        "productVisibility": "visible",
        "video_url": video_url,
        "template_kind": template_kind,
        "prompt_type": prompt_type,
    }

    for idx in range(1, 11):
        row[f"img{idx}_src"] = image_urls[idx - 1] if idx <= len(image_urls) else ""
        row[f"image_alt_{idx}"] = ""

        target_col = "image" if idx == 1 else f"image{idx}"
        row[target_col] = image_urls[idx - 1] if idx <= len(image_urls) else ""

    for lang in ["cs", "en", "sk"]:
        row[f"name:{lang}"] = ""
        row[f"shortDescription:{lang}"] = ""
        row[f"description:{lang}"] = ""
        row[f"seoTitle:{lang}"] = ""
        row[f"metaDescription:{lang}"] = ""
        row[f"xmlFeedName:{lang}"] = ""

    return pd.DataFrame([row])


def build_kartovani_prompt(
    prompt_type: str,
    product_name: str,
    product_ean: str,
    product_code: str = "",
) -> str:
    template_path = PROMPT_TEMPLATE_DIR / f"{prompt_type}.txt"
    if not template_path.exists():
        raise FileNotFoundError(f"Šablona promptu nenalezena: {template_path}")

    template_text = template_path.read_text(encoding="utf-8")

    return f"""{template_text}

--------------------------------------------------
PRODUKT
{product_name}

EAN
{product_ean}

CODE
{product_code}
--------------------------------------------------
"""


def parse_ai_output_to_lang_blocks(text: str) -> dict:
    pattern = r"\[LANG=(cs|en|sk)\]\s*(.*?)(?=\[LANG=cs\]|\[LANG=en\]|\[LANG=sk\]|\Z)"
    matches = re.findall(pattern, text, flags=re.DOTALL | re.IGNORECASE)

    out = {"cs": "", "en": "", "sk": ""}
    for lang, content in matches:
        out[lang.lower()] = content.strip()
    return out


def parse_key_value_block(block_text: str) -> dict:
    result = {}
    current_key = None
    current_value_lines = []

    for line in block_text.splitlines():
        if ":" in line:
            maybe_key, maybe_value = line.split(":", 1)
            key = maybe_key.strip()
            if key:
                if current_key is not None:
                    result[current_key] = "\n".join(current_value_lines).strip()
                current_key = key
                current_value_lines = [maybe_value.strip()]
                continue

        if current_key is not None:
            current_value_lines.append(line.strip())

    if current_key is not None:
        result[current_key] = "\n".join(current_value_lines).strip()

    return result


def replace_placeholders_in_docx(template_path: Path, values: dict) -> str:
    doc = Document(template_path)

    for paragraph in doc.paragraphs:
        full_text = "".join(run.text for run in paragraph.runs)
        new_text = full_text
        for key, value in values.items():
            new_text = new_text.replace("{" + key + "}", value or "")
        if new_text != full_text:
            for i in range(len(paragraph.runs) - 1, -1, -1):
                paragraph._element.remove(paragraph.runs[i]._element)
            paragraph.add_run(new_text)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    full_text = "".join(run.text for run in paragraph.runs)
                    new_text = full_text
                    for key, value in values.items():
                        new_text = new_text.replace("{" + key + "}", value or "")
                    if new_text != full_text:
                        for i in range(len(paragraph.runs) - 1, -1, -1):
                            paragraph._element.remove(paragraph.runs[i]._element)
                        paragraph.add_run(new_text)

    html_lines = []

    for paragraph in doc.paragraphs:
        txt = paragraph.text.strip()
        if txt:
            html_lines.append(txt)

    for table in doc.tables:
        html_lines.append("<table>")
        for row in table.rows:
            html_lines.append("<tr>")
            for cell in row.cells:
                html_lines.append(f"<td>{cell.text.strip()}</td>")
            html_lines.append("</tr>")
        html_lines.append("</table>")

    return "\n".join(html_lines)


def normalize_miniature_values(values: dict) -> dict:
    normalized = dict(values)

    if "nazev_produktu" in normalized:
        normalized.setdefault("název_sady", normalized["nazev_produktu"])
        normalized.setdefault("nazev_sady", normalized["nazev_produktu"])

    if "pro_koho" in normalized:
        normalized.setdefault("doporučení_pro_koho", normalized["pro_koho"])
        normalized.setdefault("doporuceni_pro_koho", normalized["pro_koho"])

    if "potrebne_vybaveni" in normalized:
        normalized.setdefault("potřebné_vybavení", normalized["potrebne_vybaveni"])

    if "doporuceni_pro_koho" in normalized:
        normalized.setdefault("doporučení_pro_koho", normalized["doporuceni_pro_koho"])

    return normalized


def build_kartovani_html(ai_output: str, template_kind: str, extra_values: dict | None = None) -> dict:
    lang_blocks = parse_ai_output_to_lang_blocks(ai_output)

    template_paths = get_template_paths(TEMPLATE_DIR, template_kind)

    short_files = {
        "cs": template_paths["short_cs"],
        "en": template_paths["short_en"],
        "sk": template_paths["short_sk"],
    }
    long_files = {
        "cs": template_paths["detail_cs"],
        "en": template_paths["detail_en"],
        "sk": template_paths["detail_sk"],
    }

    out = {}

    for lang in ["cs", "en", "sk"]:
        values = parse_key_value_block(lang_blocks.get(lang, ""))

        if extra_values:
            values.update(extra_values)

        if template_kind == "miniatures":
            values = normalize_miniature_values(values)

        out[f"shortDescription:{lang}"] = replace_placeholders_in_docx(short_files[lang], values)
        out[f"description:{lang}"] = replace_placeholders_in_docx(long_files[lang], values)

        product_name = (
            values.get("nazev_produktu", "").strip()
            or values.get("název_sady", "").strip()
            or values.get("nazev_sady", "").strip()
        )
        short_desc = values.get("kratky_popis", "").strip()

        final_short = f"{product_name} {short_desc}".strip()
        final_short = final_short.replace("\n", " ").replace("\r", " ")

        out[f"name:{lang}"] = product_name
        out[f"seoTitle:{lang}"] = product_name
        out[f"xmlFeedName:{lang}"] = product_name
        out[f"metaDescription:{lang}"] = final_short[:155] if final_short else ""

    return out


def apply_kartovani_output_to_csv(
    df: pd.DataFrame,
    row_index: int,
    ai_output: str,
    template_kind: str,
    extra_values: dict | None = None,
) -> pd.DataFrame:
    html_map = build_kartovani_html(ai_output, template_kind, extra_values=extra_values)
    df_out = df.copy()

    for col, value in html_map.items():
        if col not in df_out.columns:
            df_out[col] = ""
        df_out.at[row_index, col] = value

    if extra_values:
        for idx in range(1, 11):
            source_key = f"img{idx}_src"
            target_col = "image" if idx == 1 else f"image{idx}"
            if source_key in extra_values:
                if target_col not in df_out.columns:
                    df_out[target_col] = ""
                df_out.at[row_index, target_col] = extra_values[source_key]

        if "video_url" in extra_values:
            if "video_url" not in df_out.columns:
                df_out["video_url"] = ""
            df_out.at[row_index, "video_url"] = extra_values["video_url"]

    df_out = df_out.drop(
        columns=[c for c in ["product_name", "name", "description", "shortDescription"] if c in df_out.columns],
        errors="ignore"
    )

    preferred_order = [
        "code",
        "pairCode",
        "externalCode",
        "name:cs",
        "name:en",
        "name:sk",
        "shortDescription:cs",
        "shortDescription:en",
        "shortDescription:sk",
        "description:cs",
        "description:en",
        "description:sk",
        "price",
        "priceWithoutVat",
        "manufacturer",
        "itemType",
        "image",
        "image2",
        "image3",
        "image4",
        "image5",
        "image6",
        "image7",
        "image8",
        "image9",
        "image10",
        "video_url",
        "availabilityInStock",
        "availabilityOutOfStock",
        "ean",
        "productVisibility",
        "xmlFeedName:cs",
        "xmlFeedName:en",
        "xmlFeedName:sk",
        "seoTitle:cs",
        "seoTitle:en",
        "seoTitle:sk",
        "metaDescription:cs",
        "metaDescription:en",
        "metaDescription:sk",
        "template_kind",
        "prompt_type",
    ]

    existing_preferred = [col for col in preferred_order if col in df_out.columns]
    remaining_cols = [col for col in df_out.columns if col not in existing_preferred]

    df_out = df_out[existing_preferred + remaining_cols]
    df_out = df_out.fillna("")

    return df_out