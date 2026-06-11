"""
Streamlit Web-UI
----------------
Bossinfo Detailkonzept-Generator: Holt Work Items eines Epics aus Azure DevOps
und erzeugt ein Word-Dokument im bossinfo-Style.

Start lokal:   streamlit run app.py
Oder bequemer: Doppelklick auf start.command (macOS) / start.bat (Windows)
"""
from __future__ import annotations

import re
import traceback
from pathlib import Path

import streamlit as st

from devops_client import AzureDevOpsClient, WorkItem
from docx_builder import DetailkonzeptBuilder


APP_DIR = Path(__file__).parent
TEMPLATE_PATH = APP_DIR / "template_detailkonzept.docx"
CONFIG_PATH = APP_DIR / "config.txt"


# ----- Konfiguration laden/speichern (kein PAT im Code!) -----

def load_config() -> dict:
    if not CONFIG_PATH.exists():
        return {}
    cfg = {}
    for line in CONFIG_PATH.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        cfg[k.strip()] = v.strip()
    return cfg


def save_config(cfg: dict) -> None:
    lines = [
        "# Bossinfo Detailkonzept Generator - Konfiguration",
        "# Diese Datei NICHT in Git/SharePoint einchecken (enthält PAT).",
        "",
    ]
    for k, v in cfg.items():
        lines.append(f"{k}={v}")
    CONFIG_PATH.write_text("\n".join(lines), encoding="utf-8")


def sanitize_filename(name: str) -> str:
    """Erzeugt einen dateisystem-tauglichen Dateinamen aus beliebigem Text."""
    name = name.strip()
    name = re.sub(r"[^\w\-\. ]+", "_", name, flags=re.UNICODE)
    name = re.sub(r"_+", "_", name).strip("_ ")
    return name or "Detailkonzept"


def format_azure_devops_error(exc: Exception) -> str:
    """Macht 401/403-Fehler für den Nutzer verständlich."""
    status = getattr(getattr(exc, "response", None), "status_code", None)
    if status in (401, 403):
        return (
            "Azure DevOps hat den PAT abgelehnt (ungültig/abgelaufen oder Rechte fehlen). "
            "Bitte ein neues PAT mit 'Project and Team (Read)' und 'Work Items (Read)' erzeugen "
            "und erneut speichern."
        )
    return str(exc)


def normalize_project_name(name: str) -> str:
    """Normalisiert Projektnamen fuer robuste Vergleiche."""
    return re.sub(r"\s+", " ", (name or "")).strip().casefold()


# ----- UI -----

st.set_page_config(
    page_title="bossinfo • Detailkonzept-Generator",
    page_icon="📄",
    layout="centered",
)

# Blaues Farbschema fuer Auswahl-Chips (Multiselect) und Primary-Buttons,
# da aeltere Streamlit-Versionen die Theme-primaryColor dort nicht anwenden.
st.markdown("""
<style>
span[data-baseweb="tag"] {
    background-color: #2563EB !important;
    color: #FFFFFF !important;
}
button[kind="primary"], button[kind="primaryFormSubmit"] {
    background-color: #2563EB !important;
    border-color: #2563EB !important;
    color: #FFFFFF !important;
}
button[kind="primary"]:hover, button[kind="primaryFormSubmit"]:hover {
    background-color: #1D4ED8 !important;
    border-color: #1D4ED8 !important;
}
/* Checkbox: angekreuzte Box */
label[data-baseweb="checkbox"] input:checked + span {
    background-color: #2563EB !important;
    border-color: #2563EB !important;
}
/* Toggle: aktivierter Schalter (Track) */
label[data-baseweb="checkbox"] input:checked + div {
    background-color: #2563EB !important;
    border-color: #2563EB !important;
}
/* Radio-Buttons: Ring + Punkt der Auswahl */
label[data-baseweb="radio"] input:checked + div {
    border-color: #2563EB !important;
}
label[data-baseweb="radio"] input:checked + div > div {
    background-color: #2563EB !important;
}
</style>
""", unsafe_allow_html=True)

st.title("📄 Detailkonzept-Generator")
st.caption("Exportiert Azure DevOps Work Items eines Epics in ein Word-Dokument im bossinfo-Style.")

if not TEMPLATE_PATH.exists():
    st.error(
        f"Vorlage `{TEMPLATE_PATH.name}` nicht gefunden. Bitte die Datei in den "
        "Tool-Ordner kopieren und das Tool neu starten."
    )
    st.stop()

cfg = load_config()


def get_project_id(name: str) -> str:
    target = normalize_project_name(name)
    for item in st.session_state.projects:
        if normalize_project_name(item.get("name", "")) == target:
            return item.get("id", "")
    return ""

# Session-State initialisieren
for key, default in [("projects", []), ("epics", []),
                     ("epics_for_project", None), ("selected_project", cfg.get("PROJECT", "")),
                     ("tags", []), ("queries", [])]:
    if key not in st.session_state:
        st.session_state[key] = default


# ===== 1. Schritt: Zugang =====

with st.expander("🔐 1. Azure DevOps Zugang", expanded=not bool(cfg.get("PAT"))):
    org_url = st.text_input(
        "Organisations-URL",
        value=cfg.get("ORG_URL", "https://dev.azure.com/bossinfo/"),
        help="z.B. https://dev.azure.com/bossinfo/",
    )
    pat = st.text_input(
        "Personal Access Token (PAT)",
        value=cfg.get("PAT", ""),
        type="password",
        help="DevOps → Profilbild → Personal Access Tokens. Für Projekte/Tags wird mindestens 'Project and Team (Read)' plus 'Work Items (Read)' benötigt.",
    )

    col_a, col_b = st.columns([1, 1])
    with col_a:
        if st.button("🔄 Projekte laden", disabled=not pat, use_container_width=True):
            try:
                client = AzureDevOpsClient(org_url, "", pat, "")
                st.session_state.projects = client.list_projects()
                if st.session_state.projects:
                    saved_project = cfg.get("PROJECT", "").strip()
                    if saved_project and get_project_id(saved_project):
                        st.session_state.selected_project = saved_project
                    else:
                        st.session_state.selected_project = st.session_state.projects[0].get("name", "")
                # Epics-/Tag-Cache invalidieren beim neuen Laden
                st.session_state.epics = []
                st.session_state.epics_for_project = None
                st.session_state.tags = []
                st.success(f"{len(st.session_state.projects)} Projekte gefunden.")
            except Exception as e:
                st.error(f"Projekte konnten nicht geladen werden: {format_azure_devops_error(e)}")
    with col_b:
        if st.button("💾 Zugang speichern", use_container_width=True):
            save_config({
                "ORG_URL": org_url,
                "PROJECT": st.session_state.selected_project,
                "PAT": pat,
            })
            st.success("Gespeichert in `config.txt` (lokal).")


# ===== 2. Schritt: Projekt =====

st.subheader("📂 2. Projekt auswählen")

if st.session_state.projects:
    project_names = [p["name"] for p in st.session_state.projects]
    default_idx = 0
    if st.session_state.selected_project:
        matches = [i for i, name in enumerate(project_names) if normalize_project_name(name) == normalize_project_name(st.session_state.selected_project)]
        if matches:
            default_idx = matches[0]
    project = st.selectbox(
        "Projekt",
        project_names,
        index=default_idx,
        label_visibility="collapsed",
    )
    # Wenn Projekt gewechselt wurde -> Epic-Cache leeren
    if project != st.session_state.selected_project:
        st.session_state.selected_project = project
        st.session_state.epics = []
        st.session_state.epics_for_project = None
else:
    st.info("Bitte oben zuerst Zugang eingeben und Projekte laden anklicken.")
    project = st.text_input(
        "Projekt (manuell)",
        value=st.session_state.selected_project,
        help="Falls der automatische Abruf nicht klappt, hier den Projektnamen exakt eintragen.",
    )
    st.session_state.selected_project = project


# ===== 3. Schritt: Epic =====

st.subheader("🎯 3. Quelle auswählen")

source_mode = st.radio(
    "Quelle der Work Items",
    ["Epics", "Query"],
    index=1 if cfg.get("SOURCE", "Epics") == "Query" else 0,
    horizontal=True,
    help="Epics: Epics auswählen wie bisher. Query: eine in Azure DevOps "
         "gespeicherte Query (Shared/My Queries) als Quelle verwenden – "
         "exportiert werden die Resultate der Query.",
)

selected_epic_ids: list[int] = []
selected_query: dict | None = None

if source_mode == "Epics":
    if project and pat:
        col_x, col_y = st.columns([2, 1])
        with col_x:
            only_active = st.checkbox(
                "Nur aktive Epics anzeigen (ohne 'Closed' / 'Removed')",
                value=True,
            )
        with col_y:
            if st.button("🔄 Epics laden", use_container_width=True):
                try:
                    client = AzureDevOpsClient(
                        org_url,
                        project,
                        pat,
                        get_project_id(project) or project,
                    )
                    with st.spinner(f"Lade Epics aus '{project}'..."):
                        st.session_state.epics = client.list_epics(only_active=only_active)
                    st.session_state.epics_for_project = (project, only_active)
                    st.success(f"{len(st.session_state.epics)} Epics gefunden.")
                except Exception as e:
                    st.error(f"Epics konnten nicht geladen werden: {format_azure_devops_error(e)}")
                    with st.expander("Details"):
                        st.code(traceback.format_exc())

            if st.button("🏷️ Tags aus Projekt laden", use_container_width=True):
                try:
                    client = AzureDevOpsClient(
                        org_url,
                        project,
                        pat,
                        get_project_id(project) or project,
                    )
                    with st.spinner(f"Lade Tags aus '{project}'..."):
                        st.session_state.tags = client.list_tags()
                    if st.session_state.tags:
                        st.success(f"{len(st.session_state.tags)} Tags gefunden.")
                    else:
                        st.info("Keine Tags gefunden – oder der PAT hat keine ausreichenden Rechte für die Tag-Abfrage.")
                except Exception as e:
                    st.error(f"Tags konnten nicht geladen werden: {format_azure_devops_error(e)}")
                    with st.expander("Details"):
                        st.code(traceback.format_exc())

        # Hinweis falls Cache zu Projekt-/Filter-Mix passt
        cache_state = st.session_state.epics_for_project
        if st.session_state.epics and cache_state and cache_state != (project, only_active):
            st.warning("Projekt- oder Filter-Auswahl wurde geändert. Bitte „Epics laden\" erneut klicken.")

        if st.session_state.epics:
            filtered = st.session_state.epics

            if filtered:
                def fmt(e: dict) -> str:
                    state_tag = f" [{e['state']}]" if e.get("state") else ""
                    return f"#{e['id']} – {e['title']}{state_tag}"

                selected = st.multiselect(
                    "Epics",
                    filtered,
                    default=[],
                    format_func=fmt,
                    help="Mehrere Epics können gleichzeitig ausgewählt werden.",
                )
                if selected:
                    selected_epic_ids = [item["id"] for item in selected]
                    meta_bits = []
                    for item in selected[:3]:
                        if item.get("area_path"):
                            meta_bits.append(f"{item['title']} ({item['area_path']})")
                    if meta_bits:
                        st.caption("Ausgewählt: " + " · ".join(meta_bits))
            else:
                st.warning("Keine Epics passen zum Filter.")
        else:
            st.info("Klicke auf 🔄 Epics laden, um die Epics aus dem Projekt zu holen.")
    else:
        st.info("Zugang eingeben und Projekt wählen, dann können die Epics geladen werden.")

else:
    if project and pat:
        if st.button("🔄 Queries laden", use_container_width=True):
            try:
                client = AzureDevOpsClient(
                    org_url,
                    project,
                    pat,
                    get_project_id(project) or project,
                )
                with st.spinner(f"Lade Queries aus '{project}'..."):
                    st.session_state.queries = client.list_queries()
                if st.session_state.queries:
                    st.success(f"{len(st.session_state.queries)} Queries gefunden.")
                else:
                    st.info("Keine gespeicherten Queries im Projekt gefunden.")
            except Exception as e:
                st.error(f"Queries konnten nicht geladen werden: {format_azure_devops_error(e)}")
                with st.expander("Details"):
                    st.code(traceback.format_exc())

        if st.session_state.queries:
            _QUERY_TYPE_LABEL = {"flat": "Liste", "tree": "Hierarchie", "oneHop": "Direkte Links"}

            def fmt_query(q: dict) -> str:
                t = _QUERY_TYPE_LABEL.get(q.get("query_type", ""), "")
                return f"{q['path']}  [{t}]" if t else q["path"]

            query_ids = [q["id"] for q in st.session_state.queries]
            default_q_idx = 0
            if cfg.get("QUERY_ID", "") in query_ids:
                default_q_idx = query_ids.index(cfg.get("QUERY_ID", ""))
            selected_query = st.selectbox(
                "Query",
                st.session_state.queries,
                index=default_q_idx,
                format_func=fmt_query,
                help="Gespeicherte Query aus Azure DevOps. Bei Hierarchie-Queries wird "
                     "die Hierarchie übernommen, bei Listen-Queries aus den "
                     "Parent-Links rekonstruiert.",
            )
        else:
            st.info("Klicke auf 🔄 Queries laden, um die gespeicherten Queries aus dem Projekt zu holen.")
    else:
        st.info("Zugang eingeben und Projekt wählen, dann können die Queries geladen werden.")


# ===== 4. Schritt: Output =====

st.subheader("⚙️ 4. Output")

customer = st.text_input(
    "Kundenname (ersetzt [Kunde] im Dokument)",
    value=cfg.get("CUSTOMER", "LEP AG"),
)

default_filename = f"Detailkonzept_{sanitize_filename(customer)}" if customer else "Detailkonzept"
filename_hint = st.text_input(
    "Dateiname (ohne Endung)",
    value=default_filename,
)

# ----- Work-Item-Typen Filter -----
ALL_WORK_ITEM_TYPES = [
    "Epic",
    "Feature",
    "User Story",
    "Task",
]

# Vorbelegung aus config.txt (wenn vorhanden), sonst alle Typen aktiv
saved_types_str = cfg.get("TYPES", "")
if saved_types_str:
    default_types = [t.strip() for t in saved_types_str.split(",") if t.strip()]
else:
    default_types = ALL_WORK_ITEM_TYPES.copy()

selected_types = st.multiselect(
    "Work-Item-Typen, die exportiert werden sollen",
    options=ALL_WORK_ITEM_TYPES,
    default=default_types,
    help=(
        "Standardmässig alle. Deaktivierst du z.B. 'Task', werden Tasks nicht "
        "ins Dokument aufgenommen. Hinweis: wird ein Eltern-Typ deaktiviert, "
        "werden auch dessen Kinder übersprungen."
    ),
)
if not selected_types:
    st.warning("Mindestens ein Typ muss gewählt sein – sonst kommt nichts ins Dokument.")

saved_tags_str = cfg.get("TAGS", "")
selected_tags = [t.strip() for t in saved_tags_str.split(",") if t.strip()]

if st.session_state.tags:
    selected_tags = st.multiselect(
        "Tags filtern (aus dem Projekt)",
        options=st.session_state.tags,
        default=selected_tags,
        help="Wähle einen oder mehrere Tags aus der Projekt-Liste aus Azure DevOps.",
    )
else:
    selected_tags_input = st.text_input(
        "Tags filtern (1 oder mehrere, getrennt durch Komma oder Semikolon)",
        value=",".join(selected_tags),
        placeholder="z.B. backend, wichtig",
        help="Bitte zuerst auf 'Tags aus Projekt laden' klicken, damit eine Auswahl verfügbar ist.",
    )

    def parse_tags(value: str) -> list[str]:
        parts = re.split(r"[,;]\s*", value.strip())
        return [p.strip() for p in parts if p and p.strip()]

    selected_tags = parse_tags(selected_tags_input)

if selected_tags:
    st.caption(f"Tag-Filter aktiv: {', '.join(selected_tags)}")
else:
    st.caption("Kein Tag-Filter aktiv – alle Work Items mit den gewählten Typen werden aufgenommen.")

# ----- Work-Item-Nummer neben Titel anzeigen? -----
show_wi_id_default = cfg.get("SHOW_WI_ID", "true").lower() in ("1", "true", "yes", "ja", "y")
show_wi_id = st.checkbox(
    "Work-Item-Nr. neben Titel anzeigen (z.B. Stammdaten (#175910))",
    value=show_wi_id_default,
    help="Wenn aktiviert, erscheint hinter jedem Titel die DevOps-ID als klickbarer Hyperlink.",
)

# ----- Aufwandschätzungs-Abschnitt anzeigen? -----
show_estimates_default = cfg.get("SHOW_ESTIMATES", "true").lower() in ("1", "true", "yes", "ja", "y")
show_estimates = st.checkbox(
    "Abschnitt 'Aufwandschätzung Entwicklungen' (Tabelle vor 'Freigabe') anzeigen",
    value=show_estimates_default,
    help="Wenn aktiviert, wird vor dem 'Freigabe'-Feature automatisch eine Tabelle "
         "mit allen Work Items eingefügt, die Original Estimate / Original Estimate PL gesetzt haben.",
)


# ===== 5. Generieren =====

st.divider()

source_ok = bool(selected_query) if source_mode == "Query" else bool(selected_epic_ids)
generate_disabled = not (pat and project and source_ok and selected_types)
help_msg = None
if not pat:
    help_msg = "PAT fehlt"
elif not project:
    help_msg = "Projekt fehlt"
elif source_mode == "Epics" and not selected_epic_ids:
    help_msg = "Mindestens ein Epic muss ausgewählt sein"
elif source_mode == "Query" and not selected_query:
    help_msg = "Eine Query muss ausgewählt sein"
elif not selected_types:
    help_msg = "Mindestens ein Work-Item-Typ muss gewählt sein"

if st.button(
    "📥 Detailkonzept generieren",
    type="primary",
    disabled=generate_disabled,
    help=help_msg,
    use_container_width=True,
):
    try:
        with st.status("Verbinde zu Azure DevOps...", expanded=True) as status:
            client = AzureDevOpsClient(
                org_url,
                project,
                pat,
                get_project_id(project) or project,
            )
            client.test_connection()

            trees = []
            total_items = 0
            if source_mode == "Query":
                status.update(label=f"Führe Query '{selected_query['name']}' aus...")
                roots = client.run_query(selected_query["id"])
                if not roots:
                    status.update(label="Query lieferte keine Ergebnisse", state="error")
                    st.error(f"Die Query '{selected_query['name']}' lieferte keine Work Items.")
                    st.stop()

                def count_nodes(n: WorkItem) -> int:
                    return 1 + sum(count_nodes(c) for c in n.children)

                total_items = sum(count_nodes(r) for r in roots)
                # Pseudo-Epic als Wurzel: so bleibt die Dokumentstruktur erhalten
                # (Titelblatt mit Query-Name, Resultate als Kapitel-Hierarchie).
                pseudo = WorkItem(id=0, work_item_type="Epic", title=selected_query["name"])
                pseudo.children = roots
                trees = [pseudo]
                st.write(f"📦 {total_items} Work Items aus Query '{selected_query['name']}'")
            else:
                status.update(label="Lade Work-Item-Hierarchie...")
                for epic_id in selected_epic_ids:
                    items = client.get_descendants(epic_id)
                    total_items += len(items)
                    tree = client.build_tree(items, epic_id)
                    if tree is None:
                        status.update(label="Epic nicht gefunden", state="error")
                        st.error(f"Epic mit ID {epic_id} nicht im Projekt '{project}' gefunden.")
                        st.stop()
                    trees.append(tree)

                st.write(f"📦 {total_items} Work Items gefunden über {len(trees)} Epic(s)")
                n_features = sum(sum(1 for c in tree.children if c.work_item_type == "Feature") for tree in trees)
                n_stories = sum(
                    sum(1 for f in tree.children for s in f.children
                        if s.work_item_type in ("User Story", "Product Backlog Item", "Issue", "Bug"))
                    for tree in trees
                )
                st.write(f"🧩 Davon Features: {n_features}, Stories: {n_stories}")

            status.update(label="Erstelle Word-Dokument...")
            out_path = APP_DIR / "output" / f"{sanitize_filename(filename_hint)}.docx"
            out_path.parent.mkdir(exist_ok=True)
            builder = DetailkonzeptBuilder(
                str(TEMPLATE_PATH),
                devops_client=client,
                allowed_types=set(selected_types) if selected_types else None,
                allowed_tags=set(selected_tags) if selected_tags else None,
                show_work_item_id=show_wi_id,
                show_estimates_section=show_estimates,
            )
            builder.build(trees, customer, str(out_path))

            status.update(label="Fertig!", state="complete")

        # Konfiguration persistieren (Komfort fuer naechsten Start)
        save_config({
            "ORG_URL": org_url,
            "PROJECT": project,
            "PAT": pat,
            "CUSTOMER": customer,
            "TYPES": ",".join(selected_types),
            "SHOW_WI_ID": "true" if show_wi_id else "false",
            "SHOW_ESTIMATES": "true" if show_estimates else "false",
            "TAGS": ",".join(selected_tags),
            "SOURCE": source_mode,
            "QUERY_ID": (selected_query or {}).get("id", ""),
        })

        st.success(f"Dokument erstellt: {out_path.name}")
        with open(out_path, "rb") as f:
            st.download_button(
                "⬇️ Word-Dokument herunterladen",
                data=f.read(),
                file_name=out_path.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

    except Exception as e:
        st.error(f"Fehler: {format_azure_devops_error(e)}")
        with st.expander("Technische Details"):
            st.code(traceback.format_exc())


st.divider()
st.caption(
    "💡 Hinweis: Beim ersten Öffnen des Dokuments in Word einmal mit **F9** "
    "das Inhaltsverzeichnis aktualisieren."
)
