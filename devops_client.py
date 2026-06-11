"""
Azure DevOps Client
-------------------
Holt Work Items eines Epics samt Hierarchie (Feature -> User Story -> Task)
sowie Anhänge (Bilder) über die REST API.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Optional
from urllib.parse import quote

import requests


# ---------- Datenmodell ----------

@dataclass
class WorkItem:
    id: int
    work_item_type: str   # "Epic" | "Feature" | "User Story" | "Task" | ...
    title: str
    description_html: str = ""
    acceptance_html: str = ""
    state: str = ""
    assigned_to: str = ""
    tags: str = ""
    url: str = ""
    parent_id: Optional[int] = None
    # Manuelle Reihenfolge aus dem DevOps-Backlog (drag&drop pflegt das Feld).
    # Kleinere Werte = weiter oben im Backlog.
    stack_rank: Optional[float] = None
    children: list["WorkItem"] = field(default_factory=list)
    raw_fields: dict = field(default_factory=dict)


# ---------- Client ----------

class AzureDevOpsClient:
    """
    Dünner Wrapper um die Azure DevOps REST API für Work Items.
    Authentifizierung via Personal Access Token (PAT).
    """

    def __init__(
        self,
        organization_url: str,
        project: str = "",
        pat: str = "",
        project_id: str = "",
        *extra_args,
    ):
        # Organisations-URL normalisieren (Trailing-Slash etc.)
        # Einige Hot-Reload-/Legacy-Pfade übergeben einen zusätzlichen Positionsparameter;
        # wir akzeptieren ihn, ohne die App daran zu scheitern.
        self.organization_url = organization_url.rstrip("/") if organization_url else ""
        self.project = project
        self.project_id = project_id or project
        self.pat = (pat or "").strip()
        self.session = requests.Session()
        self.session.auth = requests.auth.HTTPBasicAuth("", self.pat)
        self.session.headers.update({
            "Accept": "application/json",
            "User-Agent": "DevOpsExport/1.0",
        })
        self.api_version = "7.1"

    # ----- Low-level Helfer -----

    def _get(self, url: str, **params) -> dict:
        params.setdefault("api-version", self.api_version)
        r = self.session.get(url, params=params, timeout=60)
        r.raise_for_status()
        return r.json()

    def _post(self, url: str, body: dict, **params) -> dict:
        params.setdefault("api-version", self.api_version)
        r = self.session.post(url, json=body, params=params, timeout=60)
        r.raise_for_status()
        return r.json()

    # ----- High-level Methoden -----

    def test_connection(self) -> str:
        """Prüft den Zugang. Gibt den Project-Namen zurück."""
        project_path = quote(self.project_id, safe="")
        url = f"{self.organization_url}/_apis/projects/{project_path}"
        data = self._get(url)
        return data.get("name", "")

    def list_epics(self, only_active: bool = False) -> list[dict]:
        """
        Listet alle Epics im konfigurierten Projekt.
        Gibt eine Liste von Dicts: [{id, title, state, area_path, iteration_path}, ...].
        Sortiert nach Titel.
        """
        wiql = {
            "query": (
                "SELECT [System.Id], [System.Title], [System.State], "
                "[System.AreaPath], [System.IterationPath] "
                "FROM WorkItems "
                f"WHERE [System.TeamProject] = '{self.project}' "
                "AND [System.WorkItemType] = 'Epic' "
                + ("AND [System.State] <> 'Closed' AND [System.State] <> 'Removed' " if only_active else "")
                + "ORDER BY [System.Title] ASC"
            )
        }
        project_path = quote(self.project_id, safe="")
        url = f"{self.organization_url}/{project_path}/_apis/wit/wiql"
        result = self._post(url, wiql)
        ids = [r["id"] for r in result.get("workItems", [])]
        if not ids:
            return []

        epics: list[dict] = []
        for i in range(0, len(ids), 200):
            chunk = ids[i:i + 200]
            url = f"{self.organization_url}/_apis/wit/workitemsbatch"
            body = {
                "ids": chunk,
                "fields": [
                    "System.Id", "System.Title", "System.State",
                    "System.AreaPath", "System.IterationPath",
                ],
            }
            data = self._post(url, body)
            for raw in data.get("value", []):
                f = raw.get("fields", {})
                epics.append({
                    "id": raw.get("id"),
                    "title": f.get("System.Title", ""),
                    "state": f.get("System.State", ""),
                    "area_path": f.get("System.AreaPath", ""),
                    "iteration_path": f.get("System.IterationPath", ""),
                })
        epics.sort(key=lambda e: (e["title"] or "").lower())
        return epics

    def list_projects(self) -> list[dict]:
        """Listet alle Projekte der Organisation. Gibt [{name, id, description}, ...] zurück."""
        url = f"{self.organization_url}/_apis/projects"
        projects = []
        continuation_token = None
        while True:
            params = {"$top": 200, "stateFilter": "wellFormed"}
            if continuation_token:
                params["continuationToken"] = continuation_token
            r = self.session.get(url, params={**params, "api-version": self.api_version}, timeout=60)
            r.raise_for_status()
            data = r.json()
            for p in data.get("value", []):
                projects.append({
                    "name": p.get("name", ""),
                    "id": p.get("id", ""),
                    "description": p.get("description", ""),
                })
            continuation_token = r.headers.get("X-MS-ContinuationToken")
            if not continuation_token:
                break
        projects.sort(key=lambda x: x["name"].lower())
        return projects

    def list_tags(self) -> list[str]:
        """Liest alle eindeutigen Tags aus dem konfigurierten Projekt aus Azure DevOps."""
        project_path = quote(self.project_id, safe="")

        # Primaer: offizieller Tags-Endpunkt (GET .../{project}/_apis/wit/tags)
        try:
            url = f"{self.organization_url}/{project_path}/_apis/wit/tags"
            r = self.session.get(url, params={"api-version": "7.1-preview.1"}, timeout=60)
            r.raise_for_status()
            data = r.json()
            tags = []
            for entry in data.get("value", []) or []:
                name = entry.get("name") or ""
                if isinstance(name, str) and name.strip():
                    tags.append(name.strip())
            if tags:
                return sorted({t.lower() for t in tags})
        except Exception:
            pass

        # Fallback: Tags direkt aus Work Items im aktuellen Projekt lesen.
        # Hinweis: [System.Tags] unterstuetzt in WIQL nur Contains/Not Contains,
        # daher alle Work Items des Projekts holen und Tags clientseitig filtern.
        url = f"{self.organization_url}/{project_path}/_apis/wit/wiql"
        try:
            result = self._post(url, {
                "query": (
                    "SELECT [System.Id] "
                    "FROM WorkItems "
                    "WHERE [System.TeamProject] = @project "
                    "ORDER BY [System.ChangedDate] DESC"
                )
            }, **{"$top": 5000})
        except Exception:
            return []

        ids = [wi["id"] for wi in result.get("workItems", [])]
        if not ids:
            return []

        tags_set = set()
        for i in range(0, len(ids), 200):
            chunk = ids[i:i + 200]
            batch_url = f"{self.organization_url}/_apis/wit/workitemsbatch"
            body = {"ids": chunk, "fields": ["System.Tags"]}
            try:
                data = self._post(batch_url, body)
            except Exception:
                continue
            for raw in data.get("value", []):
                field_tags = raw.get("fields", {}).get("System.Tags", "") or ""
                for tag in re.split(r"[;,]+", field_tags):
                    clean = tag.strip().lower()
                    if clean:
                        tags_set.add(clean)

        return sorted(tags_set)

    def list_queries(self) -> list[dict]:
        """
        Listet alle gespeicherten Queries des Projekts (Shared Queries + My Queries).
        Gibt [{id, name, path, query_type}, ...] zurueck (nur Queries, keine Ordner).
        query_type: "flat" | "tree" | "oneHop"
        """
        project_path = quote(self.project_id, safe="")
        base = f"{self.organization_url}/{project_path}/_apis/wit/queries"
        queries: list[dict] = []

        def collect(node: dict):
            if node.get("isFolder"):
                children = node.get("children")
                # Tiefere Ordner liefern Kinder nicht mit -> gezielt nachladen
                if children is None and node.get("hasChildren"):
                    try:
                        data = self._get(f"{base}/{node['id']}", **{"$depth": 2, "$expand": "minimal"})
                        children = data.get("children", [])
                    except Exception:
                        children = []
                for ch in children or []:
                    collect(ch)
            else:
                queries.append({
                    "id": node.get("id", ""),
                    "name": node.get("name", ""),
                    "path": node.get("path", node.get("name", "")),
                    "query_type": node.get("queryType", ""),
                })

        data = self._get(base, **{"$depth": 2, "$expand": "minimal"})
        for root in data.get("value", []) or []:
            collect(root)
        queries.sort(key=lambda q: (q["path"] or "").lower())
        return queries

    def run_query(self, query_id: str) -> list[WorkItem]:
        """
        Fuehrt eine gespeicherte Query aus und liefert die Resultate als Liste von
        Wurzel-WorkItems (mit children-Hierarchie).
          - Tree-/OneHop-Queries: Hierarchie kommt direkt aus den Query-Relationen.
          - Flat-Queries: Hierarchie wird aus den Parent-Links innerhalb des
            Resultats rekonstruiert; Items ohne Parent im Resultat werden Wurzeln.
        Die Reihenfolge der Query-Resultate bleibt erhalten.
        """
        ids, parent_of = self._run_query_ids(query_id)
        if not ids:
            return []

        by_id = {wi.id: wi for wi in self._get_work_items_batch(ids)}
        for wi in by_id.values():
            wi.children = []

        roots: list[WorkItem] = []
        for wid in ids:
            wi = by_id.get(wid)
            if wi is None:
                continue
            # Parent aus Query-Relation, sonst aus dem Work-Item-Link
            pid = parent_of.get(wid, wi.parent_id)
            if pid is not None and pid != wid and pid in by_id:
                by_id[pid].children.append(wi)
            else:
                roots.append(wi)
        return roots

    def run_query_flat(self, query_id: str) -> list[WorkItem]:
        """
        Fuehrt eine gespeicherte Query aus und liefert die Resultate als flache
        Liste EXAKT in der Reihenfolge der Query (keine Umgruppierung).
        """
        ids, _ = self._run_query_ids(query_id)
        if not ids:
            return []
        by_id = {wi.id: wi for wi in self._get_work_items_batch(ids)}
        items = [by_id[i] for i in ids if i in by_id]
        for wi in items:
            wi.children = []
        return items

    def _run_query_ids(self, query_id: str) -> tuple[list[int], dict[int, int]]:
        """Query ausfuehren; liefert (IDs in Query-Reihenfolge, Parent-Zuordnung aus Relationen)."""
        project_path = quote(self.project_id, safe="")
        url = f"{self.organization_url}/{project_path}/_apis/wit/wiql/{query_id}"
        result = self._get(url, **{"$top": 5000})

        if result.get("queryResultType") == "workItemLink" or result.get("workItemRelations"):
            # Tree-/OneHop-Query: Relationen auswerten
            order: list[int] = []
            parent_of: dict[int, int] = {}
            for rel in result.get("workItemRelations", []) or []:
                tgt = (rel.get("target") or {}).get("id")
                src = (rel.get("source") or {}).get("id")
                if tgt is None:
                    continue
                if tgt not in order:
                    order.append(tgt)
                if src is not None:
                    parent_of.setdefault(tgt, src)
            ids = order + [t for t in parent_of if t not in order]
        else:
            # Flat-Query
            parent_of = {}
            ids = [w["id"] for w in result.get("workItems", []) or []]

        return list(dict.fromkeys(ids)), parent_of

    def get_work_item(self, work_item_id: int) -> WorkItem:
        """Einzelnes Work Item mit allen Feldern und Relationen abrufen."""
        project_path = quote(self.project_id, safe="")
        url = f"{self.organization_url}/{project_path}/_apis/wit/workitems/{work_item_id}"
        data = self._get(url, **{"$expand": "all"})
        return self._parse_work_item(data)

    def get_descendants(self, epic_id: int) -> list[WorkItem]:
        """
        Alle Nachkommen eines Epics (Features, Stories, Tasks, Bugs, ...) via WIQL.
        Liefert eine flache Liste; Hierarchie wird separat aufgebaut.
        """
        wiql = {
            "query": (
                "SELECT [System.Id] "
                "FROM WorkItemLinks "
                f"WHERE [Source].[System.Id] = {epic_id} "
                "AND [System.Links.LinkType] = 'System.LinkTypes.Hierarchy-Forward' "
                "MODE (Recursive)"
            )
        }
        project_path = quote(self.project_id, safe="")
        url = f"{self.organization_url}/{project_path}/_apis/wit/wiql"
        result = self._post(url, wiql)
        rels = result.get("workItemRelations", [])
        ids = []
        for rel in rels:
            target = rel.get("target")
            if target and "id" in target:
                ids.append(target["id"])
            source = rel.get("source")
            if source and "id" in source and source["id"] not in ids:
                ids.append(source["id"])

        # Stellt sicher, dass das Epic selbst dabei ist
        if epic_id not in ids:
            ids.insert(0, epic_id)

        return self._get_work_items_batch(ids)

    def _get_work_items_batch(self, ids: list[int]) -> list[WorkItem]:
        """Batch-Abruf (max. 200 Items pro Request)."""
        items: list[WorkItem] = []
        for i in range(0, len(ids), 200):
            chunk = ids[i:i + 200]
            url = f"{self.organization_url}/_apis/wit/workitemsbatch"
            body = {"ids": chunk, "$expand": "all"}
            data = self._post(url, body)
            for raw in data.get("value", []):
                items.append(self._parse_work_item(raw))
        return items

    def _parse_work_item(self, raw: dict) -> WorkItem:
        f = raw.get("fields", {})
        parent_id = None
        for rel in raw.get("relations", []) or []:
            if rel.get("rel") == "System.LinkTypes.Hierarchy-Reverse":
                m = re.search(r"/workItems/(\d+)$", rel.get("url", ""))
                if m:
                    parent_id = int(m.group(1))
                    break

        assigned = ""
        assigned_field = f.get("System.AssignedTo")
        if isinstance(assigned_field, dict):
            assigned = assigned_field.get("displayName", "")
        elif isinstance(assigned_field, str):
            assigned = assigned_field

        # Manuelle Backlog-Reihenfolge: je nach Prozess-Template heisst das Feld
        # "Microsoft.VSTS.Common.StackRank" (Agile, CMMI) oder
        # "Microsoft.VSTS.Common.BacklogPriority" (Scrum).
        stack_rank: Optional[float] = None
        for candidate in ("Microsoft.VSTS.Common.StackRank",
                          "Microsoft.VSTS.Common.BacklogPriority"):
            val = f.get(candidate)
            if val is not None:
                try:
                    stack_rank = float(val)
                    break
                except (TypeError, ValueError):
                    continue

        wi = WorkItem(
            id=raw.get("id"),
            work_item_type=f.get("System.WorkItemType", ""),
            title=f.get("System.Title", ""),
            description_html=f.get("System.Description", "") or "",
            acceptance_html=f.get("Microsoft.VSTS.Common.AcceptanceCriteria", "") or "",
            state=f.get("System.State", ""),
            assigned_to=assigned,
            tags=f.get("System.Tags", "") or "",
            url=raw.get("_links", {}).get("html", {}).get("href", ""),
            parent_id=parent_id,
            stack_rank=stack_rank,
            raw_fields=f,
        )
        return wi

    def build_tree(self, items: list[WorkItem], root_id: int) -> Optional[WorkItem]:
        """Baut aus flacher Liste den Hierarchie-Baum unter dem Root (Epic)."""
        by_id = {it.id: it for it in items}
        if root_id not in by_id:
            return None
        for it in items:
            it.children = []  # reset
        for it in items:
            if it.parent_id and it.parent_id in by_id and it.id != root_id:
                by_id[it.parent_id].children.append(it)

        # Sortierung:
        #  1. Typ-Hierarchie (Feature vor Story vor Task)
        #  2. Manuelle Backlog-Reihenfolge (StackRank / BacklogPriority)
        #     - Kleinere Werte = weiter oben im Backlog
        #     - Items ohne StackRank kommen ans Ende
        #  3. Als Tie-Breaker: ID aufsteigend
        type_order = {"Epic": 0, "Feature": 1, "User Story": 2, "Product Backlog Item": 2,
                      "Issue": 2, "Bug": 3, "Task": 4}

        def sort_key(c: WorkItem):
            rank_is_none = c.stack_rank is None
            return (
                type_order.get(c.work_item_type, 99),
                rank_is_none,                          # False (=0) vor True (=1)
                c.stack_rank if not rank_is_none else 0.0,
                c.id,
            )

        def sort_recursive(node: WorkItem):
            node.children.sort(key=sort_key)
            for ch in node.children:
                sort_recursive(ch)

        root = by_id[root_id]
        sort_recursive(root)
        return root

    # ----- Attachments / Bilder -----

    def download_attachment(self, attachment_url: str) -> bytes:
        """Lädt einen Anhang (z.B. eingebettetes Bild) als Bytes."""
        # Direkt mit Auth-Headern fetchen (DevOps-Auth via Basic-PAT)
        r = self.session.get(attachment_url, timeout=120, allow_redirects=True)
        r.raise_for_status()
        return r.content

    def fetch_image(self, src: str) -> Optional[bytes]:
        """
        Holt ein Bild aus einer beliebigen Quelle:
          - data: URI (base64 inline)
          - Absolute DevOps-Attachment-URL (mit Auth)
          - Absolute andere URL (ohne Auth)
          - Relative URL (gegen organization_url aufgeloest)
        Gibt None zurueck, wenn das Bild nicht geladen werden kann.
        """
        if not src:
            return None
        src = src.strip()

        # 1) data:image/...;base64,...
        if src.startswith("data:"):
            import base64
            try:
                _, payload = src.split(",", 1)
                if ";base64" in src.split(",", 1)[0]:
                    return base64.b64decode(payload)
                return payload.encode("utf-8")
            except Exception:
                return None

        # 2) Relative URL -> auf Org-URL ergaenzen
        if src.startswith("/"):
            full_url = self.organization_url + src
        elif src.startswith("http://") or src.startswith("https://"):
            full_url = src
        else:
            # Unbekanntes Schema (z.B. cid:, mailto:) - ueberspringen
            return None

        # 3) Holen - mit DevOps-Auth wenn es eine DevOps-URL ist, sonst ohne
        try:
            if "dev.azure.com" in full_url or "visualstudio.com" in full_url:
                r = self.session.get(full_url, timeout=120, allow_redirects=True)
            else:
                # Externe URL ohne Auth (PAT wuerde da auch nicht passen)
                r = requests.get(full_url, timeout=120, allow_redirects=True,
                                 headers={"User-Agent": "DetailkonzeptGenerator/1.0"})
            r.raise_for_status()
            return r.content
        except Exception:
            return None
