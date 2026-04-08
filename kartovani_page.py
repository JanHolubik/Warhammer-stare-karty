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


def render_kartovani_page():
    st.title("Kartování")
    st.caption("Samostatná aplikace pro CREATE CSV, SOURCE CSV, prompt a FILLED CSV.")

    st.markdown("""
    <style>
    button[data-baseweb="tab"] > div[data-testid="stMarkdownContainer"] p {
        font-size: 22px !important;
        font-weight: 700 !important;
    }

    button[data-baseweb="tab"] {
        padding-top: 14px !important;
        padding-bottom: 14px !important;
    }
    </style>
    """, unsafe_allow_html=True)

    if "kartovani_generated_prompt_text" not in st.session_state:
        st.session_state["kartovani_generated_prompt_text"] = ""

    if "kartovani_create_csv_bytes" not in st.session_state:
        st.session_state["kartovani_create_csv_bytes"] = None

    if "kartovani_source_csv_bytes" not in st.session_state:
        st.session_state["kartovani_source_csv_bytes"] = None

    if "kartovani_create_filename" not in st.session_state:
        st.session_state["kartovani_create_filename"] = "kartovani_CREATE.csv"

    if "kartovani_source_filename" not in st.session_state:
        st.session_state["kartovani_source_filename"] = "kartovani_SOURCE.csv"

    if "kartovani_selected_template_kind" not in st.session_state:
        st.session_state["kartovani_selected_template_kind"] = "miniatures"

    if "generated_prompt_text" not in st.session_state:
        st.session_state["generated_prompt_text"] = ""

    if "generated_prompt_type" not in st.session_state:
        st.session_state["generated_prompt_type"] = ""

    if "kartovani_preview_html_map" not in st.session_state:
        st.session_state["kartovani_preview_html_map"] = None

    state_key = "kartovani_export_csv_bytes"
    if state_key not in st.session_state:
        st.session_state[state_key] = None

    subtab1, subtab2 = st.tabs(["Nová karta", "Prompt + Fill"])

    with subtab1:
        st.subheader("Vytvoření nové karty")

        name = st.text_input("Název produktu", key="kartovani_name")
        code = st.text_input("Kód produktu - nutné a povinné!", key="kartovani_code")
        external_code = st.text_input("External code SKLAD - Kód značky", key="kartovani_external_code")
        ean = st.text_input("EAN kód - dáváme stejné jako CODE", key="kartovani_ean")
        price = st.number_input(
            "Prodejní cena s DPH",
            min_value=0.0,
            step=1.0,
            key="kartovani_price"
        )
        description = st.text_area("Základní popis (nemusí být nic)", key="kartovani_desc")

        if st.button("Vytvořit CREATE + SOURCE CSV", key="kartovani_create_btn"):
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
                    template_kind="miniatures",
                    prompt_type="miniatures",
                )

                st.session_state["kartovani_create_csv_bytes"] = create_df.to_csv(
                    index=False,
                    sep=";"
                ).encode("utf-8-sig")

                st.session_state["kartovani_source_csv_bytes"] = source_df.to_csv(
                    index=False,
                    sep=";"
                ).encode("utf-8-sig")

                st.session_state["kartovani_create_filename"] = f"{code}_CREATE.csv"
                st.session_state["kartovani_source_filename"] = f"{code}_SOURCE.csv"

                st.success("CREATE i SOURCE jsou připravené ke stažení.")

        if st.session_state["kartovani_create_csv_bytes"] is not None:
            st.download_button(
                "Stáhnout CREATE CSV",
                data=st.session_state["kartovani_create_csv_bytes"],
                file_name=st.session_state["kartovani_create_filename"],
                mime="text/csv",
                key="kartovani_download_create",
            )

        if st.session_state["kartovani_source_csv_bytes"] is not None:
            st.download_button(
                "Stáhnout SOURCE CSV",
                data=st.session_state["kartovani_source_csv_bytes"],
                file_name=st.session_state["kartovani_source_filename"],
                mime="text/csv",
                key="kartovani_download_source",
            )

    with subtab2:
        st.subheader("Prompt + Fill")

        uploaded_source_csv = st.file_uploader(
            "Nahraj SOURCE CSV",
            type=["csv"],
            key="kartovani_uploaded_source_csv",
        )

        df = None
        row_index = None
        product_name = ""
        product_ean = ""
        product_code = ""

        if uploaded_source_csv is not None:
            try:
                df = pd.read_csv(uploaded_source_csv, sep=";", dtype=str).fillna("")

                name_col = None
                if "product_name" in df.columns:
                    name_col = "product_name"
                elif "name" in df.columns:
                    name_col = "name"
                elif "name:cs" in df.columns:
                    name_col = "name:cs"

                if name_col:
                    product_options = [f"{i} | {row.get(name_col, '')}" for i, row in df.iterrows()]
                    selected = st.selectbox(
                        "Vyber produkt",
                        product_options,
                        key="kartovani_select_product"
                    )

                    row_index = int(selected.split("|")[0].strip())
                    product_name = df.iloc[row_index].get(name_col, "")
                    product_ean = df.iloc[row_index].get("ean", "")
                    product_code = df.iloc[row_index].get("code", "")

                    template_kind_from_source = (
                        df.iloc[row_index].get("template_kind", "").strip().lower() or "miniatures"
                    )
                    st.session_state["kartovani_selected_template_kind"] = template_kind_from_source

                    st.info(f"Produkt: {product_name}")
                    st.write(f"EAN: {product_ean}")
                    st.write(f"CODE: {product_code}")
                    st.write(f"Výchozí template_kind ze SOURCE: {template_kind_from_source}")
                else:
                    st.error(
                        "SOURCE CSV neobsahuje sloupec 'product_name', 'name' ani 'name:cs'. "
                        f"Nalezené sloupce: {list(df.columns)}"
                    )

            except Exception as e:
                st.error(f"Nepodařilo se načíst SOURCE CSV: {e}")

        st.markdown("### Vyber prompt")

        col1, col2, col3, col4, col5, col6 = st.columns(6)

        def generate_prompt(prompt_type: str):
            if not product_name:
                st.warning("Nejdřív nahraj SOURCE CSV a vyber produkt.")
                return

            try:
                prompt_text = build_kartovani_prompt(
                    prompt_type=prompt_type,
                    product_name=product_name,
                    product_ean=product_ean,
                    product_code=product_code,
                )

                st.session_state["generated_prompt_text"] = prompt_text
                st.session_state["generated_prompt_type"] = prompt_type
                st.session_state["kartovani_selected_template_kind"] = prompt_type

            except Exception as e:
                st.error(f"Chyba při načítání promptu: {e}")

        with col1:
            if st.button("Miniatures", key="prompt_btn_miniatures"):
                generate_prompt("miniatures")
        with col2:
            if st.button("Books", key="prompt_btn_books"):
                generate_prompt("books")
        with col3:
            if st.button("Dice", key="prompt_btn_dice"):
                generate_prompt("dice")
        with col4:
            if st.button("Warscroll", key="prompt_btn_warscroll"):
                generate_prompt("warscroll")
        with col5:
            if st.button("Upgrades", key="prompt_btn_upgrades"):
                generate_prompt("upgrades")
        with col6:
            if st.button("Accessories", key="prompt_btn_accessories"):
                generate_prompt("accessories")

        st.info(f"Aktuálně vybraná šablona: {st.session_state['kartovani_selected_template_kind']}")

        if st.session_state["generated_prompt_text"]:
            prompt_text = st.session_state["generated_prompt_text"]

            st.text_area(
                f"Vygenerovaný prompt ({st.session_state['generated_prompt_type']})",
                value=prompt_text,
                height=350,
                key="generated_prompt_preview",
            )

            copy_text = json.dumps(prompt_text)

            st.components.v1.html(
                f"""
                <button onclick='navigator.clipboard.writeText({copy_text})'
                style="
                    background-color:#1f77b4;
                    color:white;
                    padding:8px 16px;
                    border:none;
                    border-radius:6px;
                    cursor:pointer;
                    font-size:14px;
                    margin-top:8px;
                ">
                📋 Kopírovat prompt
                </button>
                """,
                height=50,
            )

        ai_output = st.text_area(
            "AI Output",
            height=420,
            key="kartovani_ai_output",
            placeholder="""[LANG=cs]
nazev_produktu:
...

[LANG=en]
...

[LANG=sk]
...""",
        )

        st.markdown("### Odkazy na média")

        def source_value(col_name: str) -> str:
            if df is not None and row_index is not None and row_index < len(df):
                return str(df.iloc[row_index].get(col_name, "") or "")
            return ""

        intro_image_src = st.text_input(
            "Odkaz na úvodní obrázek (intro image)",
            value=source_value("intro_image_src"),
            key="kartovani_intro_image_src"
        )
        img1_src = st.text_input(
            "Odkaz na obrázek 1",
            value=source_value("img1_src"),
            key="kartovani_img1_src"
        )
        img2_src = st.text_input(
            "Odkaz na obrázek 2",
            value=source_value("img2_src"),
            key="kartovani_img2_src"
        )
        img3_src = st.text_input(
            "Odkaz na obrázek 3",
            value=source_value("img3_src"),
            key="kartovani_img3_src"
        )
        img4_src = st.text_input(
            "Odkaz na obrázek 4",
            value=source_value("img4_src"),
            key="kartovani_img4_src"
        )
        img5_src = st.text_input(
            "Odkaz na obrázek 5",
            value=source_value("img5_src"),
            key="kartovani_img5_src"
        )
        video_url = st.text_input(
            "Video URL",
            value=source_value("video_url"),
            key="kartovani_fill_video_url"
        )

        extra_values = {
            "intro_image_src": intro_image_src.strip(),
            "img1_src": img1_src.strip(),
            "img2_src": img2_src.strip(),
            "img3_src": img3_src.strip(),
            "img4_src": img4_src.strip(),
            "img5_src": img5_src.strip(),
            "video_url": video_url.strip(),
        }

        st.markdown("### Stav médií")

        media_rows = [
            ("intro_image_src", extra_values["intro_image_src"], is_valid_image(extra_values["intro_image_src"])),
            ("img1_src", extra_values["img1_src"], is_valid_image(extra_values["img1_src"])),
            ("img2_src", extra_values["img2_src"], is_valid_image(extra_values["img2_src"])),
            ("img3_src", extra_values["img3_src"], is_valid_image(extra_values["img3_src"])),
            ("img4_src", extra_values["img4_src"], is_valid_image(extra_values["img4_src"])),
            ("img5_src", extra_values["img5_src"], is_valid_image(extra_values["img5_src"])),
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

        if ai_output.strip():
            prompt_docx_bytes = make_docx_bytes(ai_output)
            st.download_button(
                "Stáhnout vystup_prompt.docx",
                data=prompt_docx_bytes,
                file_name="vystup_prompt.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="kartovani_download_docx",
            )

        action_col1, action_col2 = st.columns(2)

        with action_col1:
            if st.button("Náhled HTML", key="kartovani_preview_btn"):
                if not ai_output.strip():
                    st.warning("Vlož AI output.")
                else:
                    try:
                        cleaned_extra_values = {k: v for k, v in extra_values.items() if v}

                        html_map = build_kartovani_html(
                            ai_output=ai_output,
                            template_kind=st.session_state["kartovani_selected_template_kind"],
                            extra_values=cleaned_extra_values,
                        )

                        st.session_state["kartovani_preview_html_map"] = html_map
                        st.success("Náhled HTML připraven.")
                    except Exception as e:
                        st.error(f"Chyba při generování náhledu: {e}")

        with action_col2:
            if st.button("Zpracovat do CSV", key="kartovani_fill_btn"):
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
                            template_kind=st.session_state["kartovani_selected_template_kind"],
                            extra_values=cleaned_extra_values,
                        )

                        st.session_state[state_key] = out_df.to_csv(
                            index=False,
                            sep=";"
                        ).encode("utf-8-sig")

                        st.success("CSV připraveno ke stažení.")
                    except Exception as e:
                        st.error(f"Chyba při zpracování: {e}")

        preview_map = st.session_state.get("kartovani_preview_html_map")
        if preview_map:
            st.markdown("## Náhled výsledku")

            preview_lang = st.selectbox(
                "Jazyk náhledu",
                ["cs", "en", "sk"],
                key="kartovani_preview_lang"
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

        if st.session_state[state_key] is not None:
            st.download_button(
                "Stáhnout FILLED CSV",
                data=st.session_state[state_key],
                file_name="kartovani_FILLED.csv",
                mime="text/csv",
                key="kartovani_download_filled",
            )