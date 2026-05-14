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

from devops_client import AzureDevOpsClient
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


# ----- UI -----

st.set_page_config(
    page_title="bossinfo • Detailkonzept-Generator",
    page_icon="📄",
    layout="centered",
)

st.title("📄 Detailkonzept-Generator")
st.caption("Exportiert Azure DevOps Work Items eines Epics in ein Word-Dokument im bossinfo-Style.")

if not TEMPLATE_PATH.exists():
    st.error(
        f"Vorlage `{TEMPLATE_PATH.name}` nicht gefunden. Bitte die Datei in den "
        "Tool-Ordner kopieren und das Tool neu starten."
    )
    st.stop()

cfg = load_config()

# Session-State initialisieren
for key, default in [("projects", []), ("epics", []),
                     ("epics_for_project", None), ("selected_project", cfg.get("PROJECT", ""))]:
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
        help="DevOps → Profilbild → Personal Access Tokens. Berechtigung 'Work Items: Read' reicht.",
    )

    col_a, col_b = st.columns([1, 1])
    with col_a:
        if st.button("🔄 Projekte laden", disabled=not pat, use_container_width=True):
            try:
                client = AzureDevOpsClient(org_url, "", pat)
                st.session_state.projects = client.list_projects()
                # Epics-Cache invalidieren beim neuen Laden
                st.session_state.epics = []
                st.session_state.epics_for_project = None
                st.success(f"{len(st.session_state.projects)} Projekte gefunden.")
            except Exception as e:
                st.error(f"Projekte konnten nicht geladen werden: {e}")
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
    if st.session_state.selected_project in project_names:
        default_idx = project_names.index(st.session_state.selected_project)
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

st.subheader("🎯 3. Epic auswählen")

epic_id: int | None = None

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
                client = AzureDevOpsClient(org_url, project, pat)
                with st.spinner(f"Lade Epics aus '{project}'..."):
                    st.session_state.epics = client.list_epics(only_active=only_active)
                st.session_state.epics_for_project = (project, only_active)
                st.success(f"{len(st.session_state.epics)} Epics gefunden.")
            except Exception as e:
                st.error(f"Epics konnten nicht geladen werden: {e}")
                with st.expander("Details"):
                    st.code(traceback.format_exc())

    # Hinweis falls Cache zu Projekt-/Filter-Mix passt
    cache_state = st.session_state.epics_for_project
    if st.session_state.epics and cache_state and cache_state != (project, only_active):
        st.warning("Projekt- oder Filter-Auswahl wurde geändert. Bitte „Epics laden\" erneut klicken.")

    if st.session_state.epics:
        # Suchfeld für lange Listen
        search = st.text_input(
            "🔎 Suche im Epic-Titel",
            value="",
            placeholder="z.B. Detailkonzept, DYCE, ..."
        )
        filtered = st.session_state.epics
        if search.strip():
            q = search.strip().lower()
            filtered = [e for e in st.session_state.epics if q in (e["title"] or "").lower()]
        st.caption(f"{len(filtered)} von {len(st.session_state.epics)} Epics passen zum Filter.")

        if filtered:
            def fmt(e: dict) -> str:
                state_tag = f" [{e['state']}]" if e.get("state") else ""
                return f"#{e['id']} – {e['title']}{state_tag}"

            selected = st.selectbox(
                "Epic",
                filtered,
                format_func=fmt,
                label_visibility="collapsed",
            )
            if selected:
                epic_id = selected["id"]
                # Zusätzliche Info zum gewählten Epic
                meta_bits = []
                if selected.get("area_path"):
                    meta_bits.append(f"Bereich: {selected['area_path']}")
                if selected.get("iteration_path"):
                    meta_bits.append(f"Iteration: {selected['iteration_path']}")
                if meta_bits:
                    st.caption(" · ".join(meta_bits))
        else:
            st.warning("Keine Epics passen zum Filter.")
    else:
        st.info("Klicke auf 🔄 Epics laden, um die Epics aus dem Projekt zu holen.")
else:
    st.info("Zugang eingeben und Projekt wählen, dann können die Epics geladen werden.")

# Optional: manuelle Eingabe als Fallback (falls Epic in einem anderen Projekt liegt oder
# noch nicht über den Filter sichtbar ist)
with st.expander("➕ Epic manuell per ID/URL angeben", expanded=False):
    manual = st.text_input(
        "Epic-ID oder Browser-URL",
        value="",
        placeholder="175907 oder https://dev.azure.com/.../_workitems/edit/175907",
    )
    if manual.strip():
        m = re.search(r"(\d+)", manual)
        if m:
            epic_id = int(m.group(1))
            st.success(f"Manuelle ID: #{epic_id}")


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

generate_disabled = not (pat and project and epic_id and selected_types)
help_msg = None
if not pat:
    help_msg = "PAT fehlt"
elif not project:
    help_msg = "Projekt fehlt"
elif not epic_id:
    help_msg = "Epic noch nicht ausgewählt"
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
            client = AzureDevOpsClient(org_url, project, pat)
            client.test_connection()

            status.update(label="Lade Work-Item-Hierarchie...")
            items = client.get_descendants(epic_id)
            st.write(f"📦 {len(items)} Work Items gefunden")

            tree = client.build_tree(items, epic_id)
            if tree is None:
                status.update(label="Epic nicht gefunden", state="error")
                st.error(f"Epic mit ID {epic_id} nicht im Projekt '{project}' gefunden.")
                st.stop()

            n_features = sum(1 for c in tree.children if c.work_item_type == "Feature")
            n_stories = sum(
                1 for f in tree.children for s in f.children
                if s.work_item_type in ("User Story", "Product Backlog Item", "Issue", "Bug")
            )
            st.write(f"🧩 Davon Features: {n_features}, Stories: {n_stories}")

            status.update(label="Erstelle Word-Dokument...")
            out_path = APP_DIR / "output" / f"{sanitize_filename(filename_hint)}.docx"
            out_path.parent.mkdir(exist_ok=True)
            builder = DetailkonzeptBuilder(
                str(TEMPLATE_PATH),
                devops_client=client,
                allowed_types=set(selected_types) if selected_types else None,
                show_work_item_id=show_wi_id,
                show_estimates_section=show_estimates,
            )
            builder.build(tree, customer, str(out_path))

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
        st.error(f"Fehler: {e}")
        with st.expander("Technische Details"):
            st.code(traceback.format_exc())


st.divider()
st.caption(
    "💡 Hinweis: Beim ersten Öffnen des Dokuments in Word einmal mit **F9** "
    "das Inhaltsverzeichnis aktualisieren."
)
