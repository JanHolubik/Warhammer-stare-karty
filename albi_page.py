import json
import pandas as pd
import streamlit as st

from kartovani_core import (
    build_kartovani_prompt,
    make_docx_bytes,
    apply_kartovani_output_to_csv,
    build_kartovani_html,
    is_valid_image,
    is_valid_video,
)


def render_albi_page():
    st.title("ALBI – Prompt + Fill")
    st.caption("Jednoduchý režim pro ALBI produkty")

    if "albi_prompt" not in st.session_state:
        st.session_state["albi_prompt"] = ""

    if "albi_export_csv_bytes" not in st.session_state:
        st.session_state["albi_export_csv_bytes"] = None

    if "albi_preview_html_map" not in st.session_state:
        st.session_state["albi_preview_html_map"] = None

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

    if st.button("Generovat ALBI prompt", key="albi_generate_prompt"):
        if not product_name:
            st.warning("Nejdřív nahraj SOURCE CSV a vyber produkt.")
        else:
            try:
                prompt_text = build_kartovani_prompt(
                    prompt_type="albi",
                    product_name=product_name,
                    product_ean=product_ean,
                    product_code=product_code,
                )
                st.session_state["albi_prompt"] = prompt_text
            except Exception as e:
                st.error(f"Chyba při načítání promptu: {e}")

    if st.session_state["albi_prompt"]:
        prompt_text = st.session_state["albi_prompt"]

        st.text_area(
            "Prompt",
            value=prompt_text,
            height=300,
            key="albi_prompt_preview"
        )

        st.components.v1.html(
            f"""
            <button onclick='navigator.clipboard.writeText({json.dumps(prompt_text)})'>
            📋 Kopírovat prompt
            </button>
            """,
            height=40,
        )

    ai_output = st.text_area(
        "AI Output",
        height=400,
        key="albi_ai_output"
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
        "Video URL -.youtube.com/embed/Wo5rp6CEUmw/",
        value=source_value("video_url"),
        key="albi_video_url"
    )

    extra_values = {
        "intro_image_src": intro_image_src.strip(),
        "img1_src": img1_src.strip(),
        "img2_src": img2_src.strip(),
        "img3_src": img3_src.strip(),
        "img4_src": img4_src.strip(),
        "video_url": video_url.strip(),
    }

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

    if ai_output.strip():
        st.download_button(
            "Stáhnout DOCX",
            data=make_docx_bytes(ai_output),
            file_name="albi_prompt.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="albi_download_docx",
        )

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