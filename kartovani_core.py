from __future__ import annotations

from pathlib import Path
from io import BytesIO
from typing import Dict
import re

import pandas as pd
from docx import Document
from bs4 import BeautifulSoup

VAT_RATE = 1.21


BASE_DIR = Path(__file__).resolve().parent
PROMPT_TEMPLATE_DIR = BASE_DIR / "prompt_templates"
TEMPLATE_DIR = BASE_DIR / "sablony"

INVALID_IMAGE_VALUES = {
    "", "/", "placeholder",
    "intro_image_src", "/intro_image_src",
    "img1_src", "/img1_src",
    "img2_src", "/img2_src",
    "img3_src", "/img3_src",
    "img4_src", "/img4_src",
    "img5_src", "/img5_src",
    "img6_src", "/img6_src",
    "img7_src", "/img7_src",
    "img8_src", "/img8_src",
    "img9_src", "/img9_src",
    "img10_src", "/img10_src",
}

INVALID_VIDEO_VALUES = {
    "", "/", "https://www.youtube.com/embed/", "https://www.youtube.com/embed"
}


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
        "albi": ["kratky_albi.docx"],
    }

    detail_map = {
        "miniatures": ["detailni.docx"],
        "books": ["detailni_kniha.docx"],
        "warscroll": ["detailni_warscoll.docx", "detailni_warscroll.docx"],
        "dice": ["detailni_kostky.docx"],
        "upgrades": ["detailni_upgrades.docx"],
        "accessories": ["stetce_pris.docx"],
        "albi": ["detailni_albi.docx"],
    }

    short_map_en = {
        "miniatures": ["kratky_en.docx"],
        "books": ["kratky_univ_en.docx"],
        "warscroll": ["kratky_univ_en.docx"],
        "dice": ["kratky_en.docx"],
        "upgrades": ["kratky_en.docx"],
        "accessories": ["kratky_univ_en.docx"],
        "albi": ["kratky_albi_en.docx"],
    }

    detail_map_en = {
        "miniatures": ["detailni_en.docx"],
        "books": ["detailni_kniha_en.docx"],
        "warscroll": ["detailni_warscoll_en.docx", "detailni_warscroll_en.docx"],
        "dice": ["detailni_kostky_en.docx"],
        "upgrades": ["detailni_upgrades_en.docx"],
        "accessories": ["stetce_pris en.docx"],
        "albi": ["detailni_albi_en.docx"],
    }

    short_map_sk = {
        "miniatures": ["kratky_sk.docx"],
        "books": ["kratky_univ_sk.docx"],
        "warscroll": ["kratky_univ_sk.docx"],
        "dice": ["kratky_sk.docx"],
        "upgrades": ["kratky_sk.docx"],
        "accessories": ["kratky_univ_sk.docx"],
        "albi": ["kratky_albi_sk.docx"],
    }

    detail_map_sk = {
        "miniatures": ["detailni_sk.docx"],
        "books": ["detailni_kniha_sk.docx"],
        "warscroll": ["detailni_warscoll_sk.docx", "detailni_warscroll_sk.docx"],
        "dice": ["detailni_kostky_sk.docx"],
        "upgrades": ["detailni_upgrades_sk.docx"],
        "accessories": ["stetce_pris sk.docx"],
        "albi": ["detailni_albi_sk.docx"],
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


def is_valid_image(value: str | None) -> bool:
    if value is None:
        return False
    return str(value).strip() not in INVALID_IMAGE_VALUES


def is_valid_video(value: str | None) -> bool:
    if value is None:
        return False
    return str(value).strip() not in INVALID_VIDEO_VALUES


def add_class(tag, class_name: str) -> None:
    classes = tag.get("class", [])
    if class_name not in classes:
        classes.append(class_name)
        tag["class"] = classes


def make_images_clickable(soup: BeautifulSoup) -> None:
    for img in soup.find_all("img"):
        src = (img.get("src") or "").strip()
        if not is_valid_image(src):
            continue

        parent = img.parent
        if parent and parent.name == "a" and "image-link" in (parent.get("class") or []):
            continue

        a = soup.new_tag(
            "a",
            href=src,
            target="_blank",
            rel="noopener noreferrer"
        )
        a["class"] = ["image-link"]
        img.wrap(a)


def cleanup_rendered_html(html: str, values: dict) -> str:
    soup = BeautifulSoup(html, "html.parser")

    intro_valid = is_valid_image(values.get("intro_image_src", ""))
    img1_valid = is_valid_image(values.get("img1_src", ""))
    img2_valid = is_valid_image(values.get("img2_src", ""))
    img3_valid = is_valid_image(values.get("img3_src", ""))
    img4_valid = is_valid_image(values.get("img4_src", ""))
    video_valid = is_valid_video(values.get("video_url", ""))

    # Hero image
    hero_image = soup.select_one(".hero-image")
    hero = soup.select_one(".hero")
    if hero_image and not intro_valid:
        hero_image.decompose()
        if hero:
            add_class(hero, "no-image")

    # Story image + story note
    story_media = soup.select_one(".story-media")
    story_split = soup.select_one(".story-split")
    if story_media and not img1_valid:
        story_media.decompose()
        if story_split:
            add_class(story_split, "no-image")

    # Full bleed image + divider
    full_bleed = soup.select_one(".full-bleed-image")
    if full_bleed and not img2_valid:
        prev = full_bleed.find_previous_sibling()
        if prev and "soft-divider" in (prev.get("class") or []):
            prev.decompose()
        full_bleed.decompose()

    # Duo gallery cards
    duo_gallery = soup.select_one(".duo-gallery")
    if duo_gallery:
        duo_cards = duo_gallery.select(":scope > .duo-card")
        if len(duo_cards) >= 1 and not img3_valid:
            duo_cards[0].decompose()

        duo_cards = duo_gallery.select(":scope > .duo-card")
        if len(duo_cards) >= 2 and not img4_valid:
            duo_cards[1].decompose()

        duo_cards = duo_gallery.select(":scope > .duo-card")
        if len(duo_cards) == 1:
            add_class(duo_gallery, "one-item")
        elif len(duo_cards) == 0:
            duo_gallery.decompose()

    # Video section
    if not video_valid:
        for section in soup.select("section.section-card"):
            iframe = section.select_one("iframe")
            if iframe:
                section.decompose()

    # Remove any invalid images that still remain
    for img in soup.find_all("img"):
        src = (img.get("src") or "").strip()
        if not is_valid_image(src):
            parent = img.parent
            if parent and parent.name == "a" and "image-link" in (parent.get("class") or []):
                parent.decompose()
            else:
                img.decompose()

    # Remove invalid iframes
    for iframe in soup.find_all("iframe"):
        src = (iframe.get("src") or "").strip()
        if not is_valid_video(src):
            section = iframe.find_parent("section")
            if section:
                section.decompose()
            else:
                iframe.decompose()

    make_images_clickable(soup)

    return str(soup)


def validate_final_html(html: str) -> list[str]:
    forbidden = [
        'src="intro_image_src"',
        'src="/intro_image_src"',
        'src="img1_src"',
        'src="/img1_src"',
        'src="img2_src"',
        'src="/img2_src"',
        'src="img3_src"',
        'src="/img3_src"',
        'src="img4_src"',
        'src="/img4_src"',
        'src="img5_src"',
        'src="/img5_src"',
        'src="img6_src"',
        'src="/img6_src"',
        'src="img7_src"',
        'src="/img7_src"',
        'src="img8_src"',
        'src="/img8_src"',
        'src="img9_src"',
        'src="/img9_src"',
        'src="img10_src"',
        'src="/img10_src"',
        'src=""',
        'src="/"',
        'https://www.youtube.com/embed/"',
    ]
    return [x for x in forbidden if x in html]


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

        if not values.get("intro_image_src", "").strip():
            if values.get("img1_src", "").strip():
                values["intro_image_src"] = values["img1_src"]

        if template_kind == "miniatures":
            values = normalize_miniature_values(values)

        short_html = replace_placeholders_in_docx(short_files[lang], values)
        long_html = replace_placeholders_in_docx(long_files[lang], values)

        # SHORT nech bez cleanupu, ať se nesahá na <style>
        long_html = cleanup_rendered_html(long_html, values)

        short_errors = validate_final_html(short_html)
        long_errors = validate_final_html(long_html)

        if short_errors:
            raise ValueError(f"Neplatné placeholdery v shortDescription:{lang}: {short_errors}")
        if long_errors:
            raise ValueError(f"Neplatné placeholdery v description:{lang}: {long_errors}")

        out[f"shortDescription:{lang}"] = short_html
        out[f"description:{lang}"] = long_html

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