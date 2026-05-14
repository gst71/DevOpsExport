"""
Azure DevOps Client
-------------------
Holt Work Items eines Epics samt Hierarchie (Feature -> User Story -> Task)
sowie Anhänge (Bilder) über die REST API.
"""
from __future__ import annotations

import base64
import re
from dataclasses import dataclass, field
from typing import Optional

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

    def __init__(self, organization_url: str, project: str, pat: str):
        # Organisations-URL normalisieren (Trailing-Slash etc.)
        self.organization_url = organization_url.rstrip("/")
        self.project = project
        self.pat = pat
        # PAT in HTTP Basic Auth (leerer Username)
        token = base64.b64encode(f":{pat}".encode("utf-8")).decode("ascii")
        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": f"Basic {token}",
            "Accept": "application/json",
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
        url = f"{self.organization_url}/_apis/projects/{self.project}"
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
        url = f"{self.organization_url}/{self.project}/_apis/wit/wiql"
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

    def get_work_item(self, work_item_id: int) -> WorkItem:
        """Einzelnes Work Item mit allen Feldern und Relationen abrufen."""
        url = f"{self.organization_url}/{self.project}/_apis/wit/workitems/{work_item_id}"
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
        url = f"{self.organization_url}/{self.project}/_apis/wit/wiql"
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
