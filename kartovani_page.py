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


ALLOWED_TEMPLATE_TYPES = [
    "miniatures",
    "books",
    "dice",
    "warscroll",
    "upgrades",
    "accessories",
    "albi",
]


def build_kartovani_lang_prompt(
    product_name: str,
    product_ean: str,
    product_code: str,
    prompt_type: str,
    lang_mode: str,
) -> str:
    return build_kartovani_prompt(
        prompt_type=prompt_type,
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


def get_filter_config() -> dict[str, list[str]]:
    return {
        "filteringProperty:Box": [
            "Combat Patrol",
            "Starter Set",
            "Vanguard / Spearhead",
            "Battleforce",
            "Army Set",
            "Expert Kit",
        ],
        "filteringProperty:herni system": [
            "Warhammer 40K",
            "Age of Sigmar",
            "The Horus Heresy",
            "Warhammer: The Old World",
            "Necromunda",
            "Kill Team",
            "Warhammer Underworlds",
            "Blood Bowl",
            "Warcry",
            "Warhammer Quest",
            "Middle-earth Strategy Battle Game",
        ],
        "filteringProperty:skupina armad W40K": [
            "Space Marines",
            "Imperium of Man / Impérium lidstva",
            "Chaos",
            "Xenos",
        ],
        "filteringProperty:skupina armad AOS": [
            "Grand Alliance Order / Řád",
            "Grand Alliance Chaos / Chaos",
            "Grand Alliance Death / Smrt",
            "Grand Alliance Destruction / Destrukce",
        ],
        "filteringProperty:skupina armad HH": [
            "Loyalist Legions / Loajální legie",
            "Traitor Legions / Zrádné legie",
            "Mechanicum",
            "Solar Auxilia",
        ],
        "filteringProperty:skupina armad OLDWORLD": [
            "Forces of Fantasy / Řád",
            "Ravening Hordes / Destrukce",
        ],
        "filteringProperty:Frakce W40K": [
            "Adepta Sororitas",
            "Adeptus Custodes",
            "Adeptus Mechanicus",
            "Aeldari",
            "Astra Militarum",
            "Chaos Daemons",
            "Chaos Knights",
            "Chaos Space Marines",
            "Death Guard",
            "Drukhari",
            "Genestealer Cults",
            "Grey Knights",
            "Imperial Agents",
            "Imperial Knights",
            "Leagues of Votann",
            "Necrons",
            "Orks",
            "Space Marines",
            "T’au Empire",
            "Thousand Sons",
            "Tyranids",
            "World Eaters",
        ],
        "filteringProperty:Chaos bůh": [
            "Nurgle",
            "Khorne",
            "Tzeentch",
            "Slaanesh",
            "Chaos Undivided",
        ],
        "filteringProperty:subFRAKCE Space Marines": [
            "Black Templars",
            "Blood Angels",
            "Dark Angels",
            "Space Wolves",
            "Ultramarines",
            "Imperial Fists",
            "Iron Hands",
            "Raven Guard",
            "Salamanders",
            "White Scars",
        ],
        "filteringProperty:subFRAKCE Chaos Space Marines": [
            "Emperor’s Children",
        ],
        "filteringProperty:Frakce AOS": [
            "Blades of Khorne",
            "Cities of Sigmar",
            "Daughters of Khaine",
            "Disciples of Tzeentch",
            "Flesh-eater Courts",
            "Fyreslayers",
            "Gloomspite Gitz",
            "Hedonites of Slaanesh",
            "Idoneth Deepkin",
            "Kharadron Overlords",
            "Lumineth Realm-lords",
            "Maggotkin of Nurgle",
            "Nighthaunt",
            "Ogor Mawtribes",
            "Orruk Warclans",
            "Ossiarch Bonereapers",
            "Seraphon",
            "Skaven",
            "Slaves to Darkness",
            "Sons of Behemat",
            "Soulblight Gravelords",
            "Stormcast Eternals",
            "Sylvaneth",
        ],
        "filteringProperty:Frakce HH": [
            "Legiones Astartes",
            "Mechanicum",
            "Solar Auxilia",
            "Talons of the Emperor",
        ],
        "filteringProperty:subFRAKCE Legiones Astartes": [
            "Ultramarines",
            "Imperial Fists",
            "Blood Angels",
            "Dark Angels",
            "White Scars",
            "Space Wolves",
            "Iron Hands",
            "Raven Guard",
            "Salamanders",
            "Sons of Horus",
            "World Eaters",
            "Death Guard",
            "Emperor’s Children",
            "Thousand Sons",
            "Word Bearers",
            "Night Lords",
            "Iron Warriors",
            "Alpha Legion",
        ],
        "filteringProperty:subFRAKCE Talons of the Emperor": [
            "Adeptus Custodes",
            "Sisters of Silence",
        ],
        "filteringProperty:Typ jednotky W4OK": [
            "Infantry / Pěchota",
            "Character / Postavy",
            "Vehicle / Vozidla",
            "Monster / Monstra",
            "Mounted / Jízda",
            "Fly / Létající jednotky",
            "Transport / Transporty",
            "Walker / Kráčející stroj",
            "Terminator / Terminátoři",
            "Battleline / Základní jednotky",
            "Epic Hero / Epický hrdina",
            "Psyker / Psyker",
            "Beast / Zvíře",
            "Fortifiacation / Opevnění",
            "Swarm / Roj",
        ],
        "filteringProperty:Typ jednotky AOS": [
            "Infantry / Pěchota",
            "Hero / Hrdina",
            "Monster / Monstrum",
            "Cavalry / Kavalerie",
            "Beast / Zvíře",
            "Fly / Létající jednotka",
            "Wizard / Kouzelník",
            "Warmaster / Vrchní velitel",
            "Battleline / Základní jednotky",
            "Artillery / Dělostřelectvo",
        ],
        "filteringProperty:Typ jednotky HH": [
            "Infantry / Pěchota",
            "Character / Postava",
            "Vehicle / Vozidlo",
            "Dreadnought / Dreadnought",
            "Fly / Létající jednotka",
            "Transport / Transport",
            "Heavy Support / Těžká podpora",
            "Fast Attack / Rychlý útok",
            "HQ / Velitelství",
        ],
        "filteringProperty:typ modelu": [
            "Upgrades",
            "Terrain / Scenery",
            "Base",
            "Basing / dekorace",
            "Accessories / Příslušenství",
            "Rules & Books",
        ],
        "filteringProperty:druh knihy": [
            "Codex",
            "Core Rulebook",
            "Warscroll Cards",
            "Lore Book",
            "Battletome",
            "Datacards",
            "Novel (Black Library)",
            "Campaign Book",
        ],
    }


def build_filters_template_text(product_name: str, product_ean: str, product_gw_url: str) -> str:
    config = get_filter_config()

    lines = []
    lines.append("[FILTERS]")
    lines.append("")
    lines.append("Použij pouze přesné hodnoty z povoleného seznamu.")
    lines.append("Vyplň všechny filtry, které lze z produktu bezpečně určit.")
    lines.append("Pokud jde o Age of Sigmar Spearhead box, vrať přesně hodnotu: Vanguard / Spearhead")
    lines.append("Pokud si nejsi jistý jen u konkrétního pole, nech prázdné pouze to pole.")
    lines.append("Nikdy nevymýšlej vlastní variantu.")
    lines.append("")
    lines.append("Pokud má filtr více hodnot, odděl je středníkem bez mezer navíc:")
    lines.append("Příklad:")
    lines.append("Vehicle / Vozidla;Walker / Kráčející stroj")
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
    lines.append("GW URL")
    lines.append(product_gw_url or "")
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


def validate_and_normalize_filters(parsed_filters: dict[str, str]) -> dict[str, str]:
    config = get_filter_config()
    result = {key: "" for key in config.keys()}

    for key, value in parsed_filters.items():
        if key not in config:
            continue
        if not value:
            result[key] = ""
            continue

        values = [v.strip() for v in value.split(";") if v.strip()]
        valid_values = []

        for v in values:
            if v in config[key]:
                valid_values.append(v)

        if valid_values:
            result[key] = ";".join(valid_values)
        else:
            result[key] = ""

    system_value = result.get("filteringProperty:herni system", "")

    if system_value != "Warhammer 40K":
        result["filteringProperty:skupina armad W40K"] = ""
        result["filteringProperty:Frakce W40K"] = ""
        result["filteringProperty:subFRAKCE Space Marines"] = ""
        result["filteringProperty:subFRAKCE Chaos Space Marines"] = ""
        result["filteringProperty:Typ jednotky W4OK"] = ""

    if system_value != "Age of Sigmar":
        result["filteringProperty:skupina armad AOS"] = ""
        result["filteringProperty:Frakce AOS"] = ""
        result["filteringProperty:Typ jednotky AOS"] = ""

    if system_value != "The Horus Heresy":
        result["filteringProperty:skupina armad HH"] = ""
        result["filteringProperty:Frakce HH"] = ""
        result["filteringProperty:subFRAKCE Legiones Astartes"] = ""
        result["filteringProperty:subFRAKCE Talons of the Emperor"] = ""
        result["filteringProperty:Typ jednotky HH"] = ""

    if system_value != "Warhammer: The Old World":
        result["filteringProperty:skupina armad OLDWORLD"] = ""

    if result.get("filteringProperty:Frakce W40K") != "Space Marines":
        result["filteringProperty:subFRAKCE Space Marines"] = ""

    if result.get("filteringProperty:Frakce W40K") != "Chaos Space Marines":
        result["filteringProperty:subFRAKCE Chaos Space Marines"] = ""

    if result.get("filteringProperty:Frakce HH") != "Legiones Astartes":
        result["filteringProperty:subFRAKCE Legiones Astartes"] = ""

    if result.get("filteringProperty:Frakce HH") != "Talons of the Emperor":
        result["filteringProperty:subFRAKCE Talons of the Emperor"] = ""

    box_value = result.get("filteringProperty:Box", "")
    if box_value:
        result["filteringProperty:Typ jednotky W4OK"] = ""
        result["filteringProperty:Typ jednotky AOS"] = ""
        result["filteringProperty:Typ jednotky HH"] = ""

    valid_chaos_factions_w40k = {
        "Chaos Daemons",
        "Chaos Space Marines",
        "Death Guard",
        "Thousand Sons",
        "World Eaters",
    }

    valid_chaos_factions_aos = {
        "Blades of Khorne",
        "Disciples of Tzeentch",
        "Hedonites of Slaanesh",
        "Maggotkin of Nurgle",
        "Slaves to Darkness",
        "Skaven",
    }

    if system_value not in ["Warhammer 40K", "Age of Sigmar"]:
        result["filteringProperty:Chaos bůh"] = ""
    elif (
        result.get("filteringProperty:Frakce W40K") not in valid_chaos_factions_w40k
        and result.get("filteringProperty:Frakce AOS") not in valid_chaos_factions_aos
    ):
        result["filteringProperty:Chaos bůh"] = ""

    return result


def enrich_filled_csv_with_filters(df: pd.DataFrame, filters_text: str, row_index: int) -> pd.DataFrame:
    df_out = df.copy()

    parsed_filters = parse_filters_from_text(filters_text)
    validated_filters = validate_and_normalize_filters(parsed_filters)

    for key, value in validated_filters.items():
        if key not in df_out.columns:
            df_out[key] = ""
        df_out.at[row_index, key] = value

    return df_out


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

    if "kartovani_prompt_cs" not in st.session_state:
        st.session_state["kartovani_prompt_cs"] = ""

    if "kartovani_prompt_en" not in st.session_state:
        st.session_state["kartovani_prompt_en"] = ""

    if "kartovani_prompt_sk" not in st.session_state:
        st.session_state["kartovani_prompt_sk"] = ""

    if "kartovani_ai_output_cs" not in st.session_state:
        st.session_state["kartovani_ai_output_cs"] = ""

    if "kartovani_ai_output_en" not in st.session_state:
        st.session_state["kartovani_ai_output_en"] = ""

    if "kartovani_ai_output_sk" not in st.session_state:
        st.session_state["kartovani_ai_output_sk"] = ""

    if "kartovani_ai_output_filters" not in st.session_state:
        st.session_state["kartovani_ai_output_filters"] = ""

    if "kartovani_preview_html_map" not in st.session_state:
        st.session_state["kartovani_preview_html_map"] = None

    state_key = "kartovani_export_csv_bytes"
    if state_key not in st.session_state:
        st.session_state[state_key] = None

    subtab1, subtab2 = st.tabs(["Nová karta", "Prompt + Fill"])

    with subtab1:
        st.subheader("Vytvoření nové karty")

        name = st.text_input("Název produktu", key="kartovani_name")
        code = st.text_input("Kód produktu - DŮLEŽITÉ ! - NĚKDY JE STENÉ JAKO EAN", key="kartovani_code")
        external_code = st.text_input("External code SKLAD - Kód značky", key="kartovani_external_code")
        ean = st.text_input("EAN kód", key="kartovani_ean")
        price = st.number_input(
            "Prodejní cena s DPH",
            min_value=0.0,
            step=1.0,
            key="kartovani_price"
        )
        description = st.text_area("Základní popis (nemusí být nic)", key="kartovani_desc")

        template_kind_new = st.selectbox(
            "Typ šablony pro produkt",
            ALLOWED_TEMPLATE_TYPES,
            key="kartovani_new_template_kind",
        )

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
                    template_kind=template_kind_new,
                    prompt_type=template_kind_new,
                )

                st.session_state["kartovani_selected_template_kind"] = template_kind_new

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

                    template_kind_from_source = str(
                        df.iloc[row_index].get("template_kind", "")
                    ).strip().lower()

                    if template_kind_from_source in ALLOWED_TEMPLATE_TYPES:
                        st.session_state["kartovani_selected_template_kind"] = template_kind_from_source
                    else:
                        template_kind_from_source = ""

                    st.info(f"Produkt: {product_name}")
                    st.write(f"EAN: {product_ean}")
                    st.write(f"CODE: {product_code}")
                    st.write(
                        f"Výchozí template_kind ze SOURCE: "
                        f"{template_kind_from_source or st.session_state['kartovani_selected_template_kind']}"
                    )
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
                st.session_state["kartovani_prompt_cs"] = build_kartovani_lang_prompt(
                    product_name=product_name,
                    product_ean=product_ean,
                    product_code=product_code,
                    prompt_type=prompt_type,
                    lang_mode="cs",
                )
                st.session_state["kartovani_prompt_en"] = build_kartovani_lang_prompt(
                    product_name=product_name,
                    product_ean=product_ean,
                    product_code=product_code,
                    prompt_type=prompt_type,
                    lang_mode="en",
                )
                st.session_state["kartovani_prompt_sk"] = build_kartovani_lang_prompt(
                    product_name=product_name,
                    product_ean=product_ean,
                    product_code=product_code,
                    prompt_type=prompt_type,
                    lang_mode="sk",
                )

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

        generated_prompt_cs = st.session_state.get("kartovani_prompt_cs", "").strip()
        generated_prompt_en = st.session_state.get("kartovani_prompt_en", "").strip()
        generated_prompt_sk = st.session_state.get("kartovani_prompt_sk", "").strip()

        if generated_prompt_cs or generated_prompt_en or generated_prompt_sk:
            st.subheader("Vygenerované prompty")

            col_cs, col_en, col_sk = st.columns(3)

            with col_cs:
                if generated_prompt_cs:
                    st.text_area(
                        "Prompt CS",
                        value=generated_prompt_cs,
                        height=420,
                        key="kartovani_prompt_preview_cs",
                    )

                    st.components.v1.html(
                        f"""
                        <button onclick='navigator.clipboard.writeText({json.dumps(generated_prompt_cs)})'
                        style="
                            background-color:#1f77b4;
                            color:white;
                            padding:8px 16px;
                            border:none;
                            border-radius:6px;
                            cursor:pointer;
                            font-size:14px;
                            margin-top:8px;
                            width:100%;
                        ">
                        📋 Kopírovat CS prompt
                        </button>
                        """,
                        height=50,
                    )

            with col_en:
                if generated_prompt_en:
                    st.text_area(
                        "Prompt EN",
                        value=generated_prompt_en,
                        height=420,
                        key="kartovani_prompt_preview_en",
                    )

                    st.components.v1.html(
                        f"""
                        <button onclick='navigator.clipboard.writeText({json.dumps(generated_prompt_en)})'
                        style="
                            background-color:#1f77b4;
                            color:white;
                            padding:8px 16px;
                            border:none;
                            border-radius:6px;
                            cursor:pointer;
                            font-size:14px;
                            margin-top:8px;
                            width:100%;
                        ">
                        📋 Kopírovat EN prompt
                        </button>
                        """,
                        height=50,
                    )

            with col_sk:
                if generated_prompt_sk:
                    st.text_area(
                        "Prompt SK",
                        value=generated_prompt_sk,
                        height=420,
                        key="kartovani_prompt_preview_sk",
                    )

                    st.components.v1.html(
                        f"""
                        <button onclick='navigator.clipboard.writeText({json.dumps(generated_prompt_sk)})'
                        style="
                            background-color:#1f77b4;
                            color:white;
                            padding:8px 16px;
                            border:none;
                            border-radius:6px;
                            cursor:pointer;
                            font-size:14px;
                            margin-top:8px;
                            width:100%;
                        ">
                        📋 Kopírovat SK prompt
                        </button>
                        """,
                        height=50,
                    )

        st.subheader("⬇ VLOŽ SEM AI OUTPUTY")

        output_tab_cs, output_tab_en, output_tab_sk = st.tabs(["🇨🇿 Output CS", "🇬🇧 Output EN", "🇸🇰 Output SK"])

        with output_tab_cs:
            st.text_area(
                "AI Output CS",
                height=420,
                key="kartovani_ai_output_cs",
                placeholder="""
[LANG=cs]
...
""",
            )

        with output_tab_en:
            st.text_area(
                "AI Output EN",
                height=420,
                key="kartovani_ai_output_en",
                placeholder="""
[LANG=en]
...
""",
            )

        with output_tab_sk:
            st.text_area(
                "AI Output SK",
                height=420,
                key="kartovani_ai_output_sk",
                placeholder="""
[LANG=sk]
...
""",
            )

        ai_output_cs = (st.session_state.get("kartovani_ai_output_cs") or "").strip()
        ai_output_en = (st.session_state.get("kartovani_ai_output_en") or "").strip()
        ai_output_sk = (st.session_state.get("kartovani_ai_output_sk") or "").strip()

        ai_output = "\n\n".join([x for x in [ai_output_cs, ai_output_en, ai_output_sk] if x])

        st.subheader("Prompt pro filtry")

        filters_prompt_text = build_filters_template_text(
            product_name=product_name,
            product_ean=product_ean,
            product_gw_url="",
        )

        st.text_area(
            "Filtry - prompt",
            value=filters_prompt_text,
            height=250,
            key="kartovani_filters_prompt_preview",
        )

        st.components.v1.html(
            f"""
            <button onclick='navigator.clipboard.writeText({json.dumps(filters_prompt_text)})'
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
            📋 Kopírovat prompt pro filtry
            </button>
            """,
            height=50,
        )

        st.subheader("⬇ VLOŽ SEM AI OUTPUT FILTRŮ")

        ai_output_filters = st.text_area(
            "AI Output Filters",
            height=320,
            key="kartovani_ai_output_filters",
            placeholder="""
[FILTERS]
filteringProperty:herni system=Age of Sigmar
filteringProperty:skupina armad AOS=Grand Alliance Chaos / Chaos
filteringProperty:Frakce AOS=Skaven
filteringProperty:Typ jednotky AOS=Infantry / Pěchota
""",
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
        video_url = st.text_input(
            "Video URL",
            value=source_value("video_url"),
            key="kartovani_fill_video_url"
        )

        related_video_title_default, related_video_url_default = split_related_video(
            source_value("relatedVideo")
        )

        related_video_title = st.text_input(
            "Název souvisejícího videa",
            value=related_video_title_default,
            key="kartovani_related_video_title"
        )

        related_video_url = st.text_input(
            "URL videa – vlož ve stejném formátu, jak zkopíruješ z YouTube",
            value=related_video_url_default,
            key="kartovani_related_video_url"
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
            prompt_docx_bytes = make_docx_bytes(ai_output)
            st.download_button(
                "Stáhnout vystup_prompt.docx",
                data=prompt_docx_bytes,
                file_name="vystup_prompt.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="kartovani_download_docx",
            )
        else:
            st.caption("Tlačítko ke stažení se zobrazí až po vyplnění CS, EN i SK outputu.")

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

                        if ai_output_filters.strip():
                            out_df = enrich_filled_csv_with_filters(
                                df=out_df,
                                filters_text=ai_output_filters,
                                row_index=row_index,
                            )

                        if "relatedVideo" not in out_df.columns:
                            out_df["relatedVideo"] = ""

                        out_df.at[row_index, "relatedVideo"] = related_video_value

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