import json
import pandas as pd
import streamlit as st

from kartovani_core import (
    create_kartovani_card_row,
    create_kartovani_source_row,
    build_kartovani_prompt,
    make_docx_bytes,
    apply_kartovani_output_to_csv,
    build_kartovani_html,
    is_valid_image,
    is_valid_video,
)


def get_albi_filter_config() -> dict[str, list[str]]:
    return {
        "filteringProperty:Typ hry": [
            "Strategická hra",
            "Rodinná hra",
            "Karetní hra",
            "Párty hra",
            "Dětská hra",
            "Hlavolam",
            "Venkovní hra",
            "Magnetická hra",
        ],
        "filteringProperty:Počet hráčů": [
            "Pro 1 hráče",
            "Pro 2 hráče",
            "Pro 3-4 hráče",
            "Více než 4",
        ],
        "filteringProperty:Herní čas": [
            "Do 15 minut",
            "15-30 minut",
            "31-45 minut",
            "do 60 minut",
            "61-90 minut",
            "Více než 90 minut",
        ],
        "filteringProperty:Věk": [
            "Do 8 let",
            "Věk 8+",
            "Věk 12+",
            "Věk 14+",
            "Věk 18+",
        ],
        "filteringProperty:Jazyk": [
            "Čeština",
            "Angličtina",
            "Slovenština",
        ],
    }


def build_albi_filters_prompt_text(product_name: str, product_ean: str, product_code: str) -> str:
    config = get_albi_filter_config()

    lines = []
    lines.append("[FILTERS]")
    lines.append("")
    lines.append("Použij pouze přesné hodnoty z povoleného seznamu.")
    lines.append("Vyplň všechny filtry, které lze z produktu bezpečně určit.")
    lines.append("Pokud si nejsi jistý jen u konkrétního pole, nech prázdné pouze to pole.")
    lines.append("Nikdy nevymýšlej vlastní variantu.")
    lines.append("")
    lines.append("Pokud má filtr více hodnot, odděl je středníkem bez mezer navíc.")
    lines.append("Příklad:")
    lines.append("Rodinná hra;Karetní hra")
    lines.append("")

    for key, values in config.items():
        lines.append(key)
        for value in values:
            lines.append(f"- {value}")
        lines.append("")

    lines.append("VRAŤ POUZE TENTO BLOK VE FORMÁTU key=value:")
    lines.append("")

    for key in config.keys():
        lines.append(f"{key}=")

    lines.append("")
    lines.append("--------------------------------------------------")
    lines.append("PRODUKT")
    lines.append(product_name or "")
    lines.append("")
    lines.append("EAN")
    lines.append(product_ean or "")
    lines.append("")
    lines.append("CODE")
    lines.append(product_code or "")
    lines.append("--------------------------------------------------")

    return "\n".join(lines)


def parse_filters_from_text(text: str) -> dict[str, str]:
    parsed: dict[str, str] = {}

    for raw_line in str(text or "").splitlines():
        line = raw_line.strip()
        if not line.startswith("filteringProperty:"):
            continue
        if "=" not in line:
            continue

        key, value = line.split("=", 1)
        parsed[key.strip()] = value.strip()

    return parsed


def validate_and_normalize_albi_filters(parsed_filters: dict[str, str]) -> dict[str, str]:
    config = get_albi_filter_config()
    result = {key: "" for key in config.keys()}

    for key, value in parsed_filters.items():
        if key not in config:
            continue

        if not value:
            result[key] = ""
            continue

        values = [v.strip() for v in value.split(";") if v.strip()]
        valid_values = [v for v in values if v in config[key]]
        result[key] = ";".join(valid_values) if valid_values else ""

    return result


def enrich_albi_csv_with_filters(df: pd.DataFrame, filters_text: str, row_index: int) -> pd.DataFrame:
    df_out = df.copy()

    parsed_filters = parse_filters_from_text(filters_text)
    validated_filters = validate_and_normalize_albi_filters(parsed_filters)

    for key, value in validated_filters.items():
        if key not in df_out.columns:
            df_out[key] = ""
        df_out.at[row_index, key] = value

    return df_out


def build_albi_lang_prompt(product_name: str, product_ean: str, product_code: str, lang_mode: str) -> str:
    return build_kartovani_prompt(
        prompt_type="albi",
        product_name=product_name,
        product_ean=product_ean,
        product_code=product_code,
        lang=lang_mode,
    )


def split_related_video(value: str) -> tuple[str, str]:
    raw = str(value or "").strip()
    if ";" not in raw:
        return "", ""
    title, url = raw.split(";", 1)
    return title.strip(), url.strip()


def render_albi_page():
    st.title("ALBI")
    st.caption("Samostatná aplikace pro ALBI produkty")

    if "albi_prompt_cs" not in st.session_state:
        st.session_state["albi_prompt_cs"] = ""

    if "albi_prompt_en" not in st.session_state:
        st.session_state["albi_prompt_en"] = ""

    if "albi_prompt_sk" not in st.session_state:
        st.session_state["albi_prompt_sk"] = ""

    if "albi_ai_output_cs" not in st.session_state:
        st.session_state["albi_ai_output_cs"] = ""

    if "albi_ai_output_en" not in st.session_state:
        st.session_state["albi_ai_output_en"] = ""

    if "albi_ai_output_sk" not in st.session_state:
        st.session_state["albi_ai_output_sk"] = ""

    if "albi_export_csv_bytes" not in st.session_state:
        st.session_state["albi_export_csv_bytes"] = None

    if "albi_preview_html_map" not in st.session_state:
        st.session_state["albi_preview_html_map"] = None

    if "albi_ai_output_filters" not in st.session_state:
        st.session_state["albi_ai_output_filters"] = ""

    if "albi_create_csv_bytes" not in st.session_state:
        st.session_state["albi_create_csv_bytes"] = None

    if "albi_source_csv_bytes" not in st.session_state:
        st.session_state["albi_source_csv_bytes"] = None

    if "albi_create_filename" not in st.session_state:
        st.session_state["albi_create_filename"] = "albi_CREATE.csv"

    if "albi_source_filename" not in st.session_state:
        st.session_state["albi_source_filename"] = "albi_SOURCE.csv"

    subtab1, subtab2 = st.tabs(["Nová karta", "Prompt + Fill"])

    with subtab1:
        st.subheader("Vytvoření nové ALBI karty")

        name = st.text_input("Název produktu", key="albi_name")
        code = st.text_input("Kód produktu - nutné a povinné!", key="albi_code")
        external_code = st.text_input("External code / skladový kód", key="albi_external_code")
        ean = st.text_input("EAN kód", key="albi_ean")
        price = st.number_input(
            "Prodejní cena s DPH",
            min_value=0.0,
            step=1.0,
            key="albi_price"
        )
        description = st.text_area("Základní popis (nemusí být nic)", key="albi_desc")

        if st.button("Vytvořit CREATE + SOURCE CSV", key="albi_create_btn"):
            if not name or not code:
                st.warning("Vyplň alespoň název produktu a kód produktu.")
            else:
                create_df = create_kartovani_card_row(
                    name=name,
                    code=code,
                    external_code=external_code,
                    ean=ean,
                    price=price,
                    product_type="product",
                    description=description,
                    image_urls=[],
                )

                source_df = create_kartovani_source_row(
                    name=name,
                    code=code,
                    external_code=external_code,
                    ean=ean,
                    price=price,
                    intro_image_src="",
                    image_urls=[],
                    video_url="",
                    template_kind="albi",
                    prompt_type="albi",
                )

                st.session_state["albi_create_csv_bytes"] = create_df.to_csv(
                    index=False,
                    sep=";"
                ).encode("utf-8-sig")

                st.session_state["albi_source_csv_bytes"] = source_df.to_csv(
                    index=False,
                    sep=";"
                ).encode("utf-8-sig")

                st.session_state["albi_create_filename"] = f"{code}_CREATE.csv"
                st.session_state["albi_source_filename"] = f"{code}_SOURCE.csv"

                st.success("ALBI CREATE i SOURCE jsou připravené ke stažení.")

        if st.session_state["albi_create_csv_bytes"] is not None:
            st.download_button(
                "Stáhnout CREATE CSV",
                data=st.session_state["albi_create_csv_bytes"],
                file_name=st.session_state["albi_create_filename"],
                mime="text/csv",
                key="albi_download_create",
            )

        if st.session_state["albi_source_csv_bytes"] is not None:
            st.download_button(
                "Stáhnout SOURCE CSV",
                data=st.session_state["albi_source_csv_bytes"],
                file_name=st.session_state["albi_source_filename"],
                mime="text/csv",
                key="albi_download_source",
            )

    with subtab2:
        st.subheader("Prompt + Fill")

        uploaded_source_csv = st.file_uploader(
            "Nahraj SOURCE CSV",
            type=["csv"],
            key="albi_uploaded_source_csv",
        )

        df = None
        row_index = None
        product_name = ""
        product_ean = ""
        product_code = ""

        if uploaded_source_csv is not None:
            try:
                df = pd.read_csv(uploaded_source_csv, sep=";", dtype=str).fillna("")

                product_options = [
                    f"{i} | {row.get('product_name', '')}"
                    for i, row in df.iterrows()
                ]

                if product_options:
                    selected = st.selectbox(
                        "Vyber produkt",
                        product_options,
                        key="albi_select_product"
                    )
                    row_index = int(selected.split("|")[0].strip())

                    product_name = df.iloc[row_index].get("product_name", "")
                    product_ean = df.iloc[row_index].get("ean", "")
                    product_code = df.iloc[row_index].get("code", "")

                    st.info(product_name)
                else:
                    st.warning("SOURCE CSV neobsahuje žádné produkty.")

            except Exception as e:
                st.error(f"Nepodařilo se načíst SOURCE CSV: {e}")

        if st.button("Generovat ALBI prompty", key="albi_generate_prompt"):
            if not product_name:
                st.warning("Nejdřív nahraj SOURCE CSV a vyber produkt.")
            else:
                st.session_state["albi_prompt_cs"] = ""
                st.session_state["albi_prompt_en"] = ""
                st.session_state["albi_prompt_sk"] = ""

                missing_langs = []

                for lang in ["cs", "en", "sk"]:
                    try:
                        st.session_state[f"albi_prompt_{lang}"] = build_albi_lang_prompt(
                            product_name=product_name,
                            product_ean=product_ean,
                            product_code=product_code,
                            lang_mode=lang,
                        )
                    except FileNotFoundError:
                        missing_langs.append(lang)
                        st.session_state[f"albi_prompt_{lang}"] = ""
                    except Exception as e:
                        st.error(f"Chyba při načítání promptu pro jazyk {lang}: {e}")
                        st.session_state[f"albi_prompt_{lang}"] = ""

                if missing_langs:
                    st.warning(f"Chybí prompt šablony pro jazyky: {', '.join(missing_langs)}")
                else:
                    st.success("ALBI prompty byly vygenerovány.")

        generated_prompt_cs = st.session_state.get("albi_prompt_cs", "").strip()
        generated_prompt_en = st.session_state.get("albi_prompt_en", "").strip()
        generated_prompt_sk = st.session_state.get("albi_prompt_sk", "").strip()

        if generated_prompt_cs or generated_prompt_en or generated_prompt_sk:
            st.subheader("Vygenerované ALBI prompty")

            col_cs, col_en, col_sk = st.columns(3)

            with col_cs:
                if generated_prompt_cs:
                    st.text_area(
                        "Prompt CS",
                        value=generated_prompt_cs,
                        height=420,
                        key="albi_prompt_preview_cs",
                    )
                    st.components.v1.html(
                        f"""
                        <button onclick='navigator.clipboard.writeText({json.dumps(generated_prompt_cs)})'>
                        📋 Kopírovat CS prompt
                        </button>
                        """,
                        height=40,
                    )

            with col_en:
                if generated_prompt_en:
                    st.text_area(
                        "Prompt EN",
                        value=generated_prompt_en,
                        height=420,
                        key="albi_prompt_preview_en",
                    )
                    st.components.v1.html(
                        f"""
                        <button onclick='navigator.clipboard.writeText({json.dumps(generated_prompt_en)})'>
                        📋 Kopírovat EN prompt
                        </button>
                        """,
                        height=40,
                    )

            with col_sk:
                if generated_prompt_sk:
                    st.text_area(
                        "Prompt SK",
                        value=generated_prompt_sk,
                        height=420,
                        key="albi_prompt_preview_sk",
                    )
                    st.components.v1.html(
                        f"""
                        <button onclick='navigator.clipboard.writeText({json.dumps(generated_prompt_sk)})'>
                        📋 Kopírovat SK prompt
                        </button>
                        """,
                        height=40,
                    )

        st.subheader("⬇ VLOŽ SEM AI OUTPUTY")

        output_tab_cs, output_tab_en, output_tab_sk = st.tabs(["🇨🇿 Output CS", "🇬🇧 Output EN", "🇸🇰 Output SK"])

        with output_tab_cs:
            st.text_area(
                "AI Output CS",
                height=420,
                key="albi_ai_output_cs",
                placeholder="""
[LANG=cs]
...
""",
            )

        with output_tab_en:
            st.text_area(
                "AI Output EN",
                height=420,
                key="albi_ai_output_en",
                placeholder="""
[LANG=en]
...
""",
            )

        with output_tab_sk:
            st.text_area(
                "AI Output SK",
                height=420,
                key="albi_ai_output_sk",
                placeholder="""
[LANG=sk]
...
""",
            )

        ai_output_cs = (st.session_state.get("albi_ai_output_cs") or "").strip()
        ai_output_en = (st.session_state.get("albi_ai_output_en") or "").strip()
        ai_output_sk = (st.session_state.get("albi_ai_output_sk") or "").strip()

        ai_output = "\n\n".join([x for x in [ai_output_cs, ai_output_en, ai_output_sk] if x])

        st.markdown("### Prompt pro filtry")

        filters_prompt_text = build_albi_filters_prompt_text(
            product_name=product_name,
            product_ean=product_ean,
            product_code=product_code,
        )

        st.text_area(
            "Filtry - prompt",
            value=filters_prompt_text,
            height=260,
            key="albi_filters_prompt_preview"
        )

        st.components.v1.html(
            f"""
            <button onclick='navigator.clipboard.writeText({json.dumps(filters_prompt_text)})'>
            📋 Kopírovat prompt pro filtry
            </button>
            """,
            height=40,
        )

        st.markdown("### AI Output Filters")

        ai_output_filters = st.text_area(
            "AI Output Filters",
            height=220,
            key="albi_ai_output_filters",
            placeholder="""[FILTERS]
filteringProperty:Typ hry=Rodinná hra
filteringProperty:Počet hráčů=Pro 2 hráče;Pro 3-4 hráče
filteringProperty:Herní čas=15-30 minut
filteringProperty:Věk=Věk 8+
filteringProperty:Jazyk=Čeština
""",
        )

        st.markdown("### Odkazy na média")

        def source_value(col_name: str) -> str:
            if df is not None and row_index is not None and row_index < len(df):
                return str(df.iloc[row_index].get(col_name, "") or "")
            return ""

        intro_image_src = st.text_input(
            "Odkaz na úvodní obrázek (první velký obrázek)",
            value=source_value("intro_image_src"),
            key="albi_intro_image_src"
        )
        img1_src = st.text_input(
            "Odkaz na obrázek 1",
            value=source_value("img1_src"),
            key="albi_img1_src"
        )
        img2_src = st.text_input(
            "Odkaz na obrázek 2 (velký roztažený obrazek, dát něco kde nejsou dulezite věci na krajích, ale uprostřed)",
            value=source_value("img2_src"),
            key="albi_img2_src"
        )
        img3_src = st.text_input(
            "Odkaz na obrázek 3",
            value=source_value("img3_src"),
            key="albi_img3_src"
        )
        img4_src = st.text_input(
            "Odkaz na obrázek 4",
            value=source_value("img4_src"),
            key="albi_img4_src"
        )
        video_url = st.text_input(
            "Video URL -.youtube.com/embed/Wo5rp6CEUmw.  - PŘEDĚLAT NA EMBED",
            value=source_value("video_url"),
            key="albi_video_url"
        )

        related_video_title_default, related_video_url_default = split_related_video(source_value("relatedVideo"))

        related_video_title = st.text_input(
            "Název souvisejícího videa",
            value=related_video_title_default,
            key="albi_related_video_title"
        )

        related_video_url = st.text_input(
            "URL souvisejícího videa – vlož ve stejném formátu, jak zkopíruješ z YouTube",
            value=related_video_url_default,
            key="albi_related_video_url"
        )

        extra_values = {
            "intro_image_src": intro_image_src.strip(),
            "img1_src": img1_src.strip(),
            "img2_src": img2_src.strip(),
            "img3_src": img3_src.strip(),
            "img4_src": img4_src.strip(),
            "video_url": video_url.strip(),
        }

        related_video_value = ""
        if related_video_title.strip() and related_video_url.strip():
            related_video_value = f"{related_video_title.strip()};{related_video_url.strip()}"

        st.markdown("### Stav médií")

        media_rows = [
            ("intro_image_src", extra_values["intro_image_src"], is_valid_image(extra_values["intro_image_src"])),
            ("img1_src", extra_values["img1_src"], is_valid_image(extra_values["img1_src"])),
            ("img2_src", extra_values["img2_src"], is_valid_image(extra_values["img2_src"])),
            ("img3_src", extra_values["img3_src"], is_valid_image(extra_values["img3_src"])),
            ("img4_src", extra_values["img4_src"], is_valid_image(extra_values["img4_src"])),
            ("video_url", extra_values["video_url"], is_valid_video(extra_values["video_url"])),
        ]

        media_df = pd.DataFrame(
            [
                {
                    "pole": name,
                    "vyplněno": "ano" if value else "ne",
                    "platné": "ano" if valid else "ne",
                    "hodnota": value,
                }
                for name, value, valid in media_rows
            ]
        )
        st.dataframe(media_df, use_container_width=True)

        if ai_output_cs and ai_output_en and ai_output_sk:
            st.download_button(
                "Stáhnout DOCX",
                data=make_docx_bytes(ai_output),
                file_name="albi_prompt.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="albi_download_docx",
            )
        else:
            st.caption("Tlačítko ke stažení se zobrazí až po vyplnění CS, EN i SK outputu.")

        col1, col2 = st.columns(2)

        with col1:
            if st.button("Náhled HTML", key="albi_preview_btn"):
                if not ai_output.strip():
                    st.warning("Vlož AI output.")
                else:
                    try:
                        cleaned_extra_values = {k: v for k, v in extra_values.items() if v}
                        html_map = build_kartovani_html(
                            ai_output=ai_output,
                            template_kind="albi",
                            extra_values=cleaned_extra_values,
                        )
                        st.session_state["albi_preview_html_map"] = html_map
                        st.success("Náhled HTML připraven.")
                    except Exception as e:
                        st.error(f"Chyba při generování náhledu: {e}")

        with col2:
            if st.button("Zpracovat do CSV", key="albi_fill_btn"):
                if df is None or row_index is None:
                    st.warning("Nejdřív nahraj SOURCE CSV a vyber produkt.")
                elif not ai_output.strip():
                    st.warning("Vlož AI output.")
                else:
                    try:
                        cleaned_extra_values = {k: v for k, v in extra_values.items() if v}

                        out_df = apply_kartovani_output_to_csv(
                            df=df,
                            row_index=row_index,
                            ai_output=ai_output,
                            template_kind="albi",
                            extra_values=cleaned_extra_values,
                        )

                        if ai_output_filters.strip():
                            out_df = enrich_albi_csv_with_filters(
                                df=out_df,
                                filters_text=ai_output_filters,
                                row_index=row_index,
                            )

                        if "relatedVideo" not in out_df.columns:
                            out_df["relatedVideo"] = ""

                        out_df.at[row_index, "relatedVideo"] = related_video_value

                        st.session_state["albi_export_csv_bytes"] = out_df.to_csv(
                            index=False,
                            sep=";"
                        ).encode("utf-8-sig")

                        st.success("ALBI CSV je připravené ke stažení.")
                    except Exception as e:
                        st.error(f"Chyba při zpracování: {e}")

        preview_map = st.session_state.get("albi_preview_html_map")
        if preview_map:
            st.markdown("## Náhled výsledku")

            preview_lang = st.selectbox(
                "Jazyk náhledu",
                ["cs", "en", "sk"],
                key="albi_preview_lang"
            )

            short_key = f"shortDescription:{preview_lang}"
            long_key = f"description:{preview_lang}"

            tab1, tab2 = st.tabs(["Krátký popis", "Detailní popis"])

            with tab1:
                short_html = preview_map.get(short_key, "")
                st.code(short_html, language="html")
                st.components.v1.html(short_html, height=800, scrolling=True)

            with tab2:
                long_html = preview_map.get(long_key, "")
                st.code(long_html, language="html")
                st.components.v1.html(long_html, height=1000, scrolling=True)

        if st.session_state["albi_export_csv_bytes"] is not None:
            st.download_button(
                "Stáhnout CSV",
                data=st.session_state["albi_export_csv_bytes"],
                file_name="ALBI_FILLED.csv",
                mime="text/csv",
                key="albi_download_csv",
            )