"""Structured LLM search layer for Generic Local File Search v2.3.0."""
from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any


def install(appmod: Any) -> None:
    appmod.VERSION = "2.3.0"

    def file_type_description(extension: str) -> str:
        ext = extension.casefold().lstrip(".")
        labels = {
            "xlsx": "excel spreadsheet workbook", "xlsm": "excel spreadsheet macro workbook",
            "xls": "excel spreadsheet workbook", "xlsb": "excel binary spreadsheet workbook",
            "ods": "open document spreadsheet", "csv": "csv spreadsheet table data",
            "tsv": "tsv spreadsheet table data", "docx": "word document office document",
            "doc": "word document office document", "pptx": "powerpoint presentation slides",
            "ppt": "powerpoint presentation slides", "pdf": "pdf document",
            "txt": "text document", "md": "markdown text document", "markdown": "markdown text document",
            "json": "json data configuration", "xml": "xml data configuration",
            "yaml": "yaml configuration", "yml": "yaml configuration", "py": "python source code",
            "js": "javascript source code", "ts": "typescript source code", "sql": "sql source code query",
        }
        return labels.get(ext, f"{ext} file")

    def file_metadata(path: str, extension: str, size: int, modified: float) -> dict[str, Any]:
        p = Path(path)
        try:
            modified_text = datetime.fromtimestamp(modified).isoformat(timespec="seconds")
        except Exception:
            modified_text = str(modified)
        return {
            "file_name": p.name,
            "file_name_without_extension": p.stem,
            "extension": extension,
            "file_type": file_type_description(extension),
            "parent_folder": p.parent.name,
            "folder_path": str(p.parent),
            "full_path": path,
            "size_bytes": size,
            "modified": modified_text,
        }

    def deterministic_score(query: str, path: str, extension: str, content: str) -> tuple[float, str]:
        q = query.casefold().strip()
        content_low = content.casefold()
        metadata_low = f"{path} {file_type_description(extension)} {extension}".casefold()
        terms = appmod.query_terms(query)
        raw = 0.0
        if q in metadata_low:
            raw += 90
        if q in content_low:
            raw += 75
        matched = 0
        for term in terms:
            mc = metadata_low.count(term)
            cc = content_low.count(term)
            if mc or cc:
                matched += 1
            raw += min(mc, 3) * 22 + min(cc, 8) * 4
            if cc:
                raw += 7
        if terms:
            raw += (matched / len(terms)) * 55
            if matched == len(terms):
                raw += 25
        return round(min(100.0, raw / 2.5), 2), appmod.find_best_snippet(content, terms)

    def ollama_json(url: str, model: str, system: str, prompt: str, temperature: float) -> dict[str, Any]:
        if appmod.httpx is None:
            raise RuntimeError("httpx is not installed")
        payload = {
            "model": model, "stream": False, "format": "json", "options": {"temperature": temperature},
            "messages": [{"role": "system", "content": system}, {"role": "user", "content": prompt}],
        }
        with appmod.httpx.Client(base_url=url.rstrip("/"), timeout=appmod.httpx.Timeout(180, connect=5)) as client:
            response = client.post("/api/chat", json=payload)
            response.raise_for_status()
            message = (response.json().get("message") or {}).get("content", "").strip()
        try:
            return json.loads(message)
        except json.JSONDecodeError:
            match = re.search(r"\{.*\}", message, flags=re.S)
            if not match:
                raise RuntimeError("Ollama returned invalid JSON")
            return json.loads(match.group(0))

    def ollama_interpret_query(url: str, model: str, query: str) -> dict[str, Any]:
        prompt = f"""Interpret this as SEARCH INSTRUCTIONS for a local file index. Do not rank a file yet.

USER REQUEST:
{query}

Return JSON only:
{{
  "intent": "what a match should represent",
  "file_name_hints": [],
  "folder_hints": [],
  "file_type_hints": [],
  "content_hints": [],
  "exclude_hints": [],
  "must_conditions": [],
  "preference_conditions": []
}}

Treat filename, folder/full path, file type/extension, modified date/size, and extracted file content as separate evidence sources. Never turn a metadata instruction into a content requirement."""
        return ollama_json(
            url, model,
            "Convert natural-language local-file searches into structured instructions. Return JSON only.",
            prompt, 0.0,
        )

    def ollama_rank(url: str, model: str, query: str, instructions: dict[str, Any], result: Any, content: str) -> tuple[int, str]:
        metadata = file_metadata(result.path, result.extension, result.size, result.modified)
        prompt = f"""Apply these SEARCH INSTRUCTIONS to this ONE candidate file.

ORIGINAL REQUEST:
{query}

INTERPRETED INSTRUCTIONS:
{json.dumps(instructions, indent=2, ensure_ascii=False)}

FILE METADATA:
{json.dumps(metadata, indent=2, ensure_ascii=False)}

EXTRACTED FILE CONTENT:
{content[:22000]}

Evaluate independently: file name, folder/full path, extension/file type, date/size if requested, and document content. Respect must-conditions and exclusions. A metadata match can be decisive even when the words are absent from the body; a content match can be decisive with a generic filename.

Return JSON only:
{{"relevance": 0-100, "reason": "concise explanation identifying metadata and/or content evidence"}}"""
        data = ollama_json(
            url, model,
            "Judge relevance by applying structured search instructions to file metadata and content separately. Return JSON only.",
            prompt, 0.1,
        )
        try:
            score = max(0, min(100, int(float(data.get("relevance", 0)))))
        except (TypeError, ValueError):
            score = 0
        return score, re.sub(r"\s+", " ", str(data.get("reason", ""))).strip()

    def search_worker(self: Any, root: Path, query: str, options: dict[str, Any]) -> None:
        conn = appmod.db_connect()
        try:
            sql = "SELECT path, extension, size, modified, status, content FROM files WHERE root = ?"
            params: list[Any] = [str(root)]
            if options["start"] is not None:
                sql += " AND modified >= ?"; params.append(options["start"])
            if options["end"] is not None:
                sql += " AND modified <= ?"; params.append(options["end"])

            selected_type = options.get("file_type", "All supported files")
            selected_exts = appmod.FILE_TYPE_PRESETS.get(selected_type)
            if selected_exts:
                marks = ",".join("?" for _ in selected_exts)
                sql += f" AND extension IN ({marks})"; params.extend(sorted(selected_exts))

            instructions: dict[str, Any] = {"intent": query}
            ollama_available = False
            if options["use_ollama"] and options["ollama_top"] > 0:
                ok, _models, error = appmod.ollama_status(options["ollama_url"])
                ollama_available = ok
                if ok:
                    try:
                        self.events.put(("status", "Interpreting search instructions with Ollama…"))
                        instructions = ollama_interpret_query(options["ollama_url"], options["model"], query)
                    except Exception as exc:
                        self.events.put(("status", f"Instruction interpretation failed; using raw prompt: {exc}"))
                else:
                    self.events.put(("status", f"Ollama unavailable; local ranking only: {error}"))

            if not selected_exts:
                hints = " ".join(str(x) for x in instructions.get("file_type_hints", [])).casefold()
                inferred: set[str] = set()
                if any(x in hints for x in ("spreadsheet", "excel", "workbook", "xlsx", "xlsm", "xlsb")):
                    inferred |= set(appmod.FILE_TYPE_PRESETS.get("Spreadsheets") or [])
                if any(x in hints for x in ("word", "docx", "doc document")):
                    inferred |= set(appmod.FILE_TYPE_PRESETS.get("Word documents") or [])
                if any(x in hints for x in ("powerpoint", "presentation", "pptx", "slides")):
                    inferred |= set(appmod.FILE_TYPE_PRESETS.get("PowerPoint") or [])
                if "pdf" in hints:
                    inferred.add(".pdf")
                if inferred:
                    marks = ",".join("?" for _ in inferred)
                    sql += f" AND extension IN ({marks})"; params.extend(sorted(inferred))

            metadata_words = {
                "spreadsheet", "spreadsheets", "excel", "workbook", "workbooks", "document", "documents",
                "word", "powerpoint", "presentation", "presentations", "slides", "pdf", "folder", "folders",
                "type", "xlsx", "xlsm", "xls", "xlsb", "docx", "doc", "pptx", "ppt",
            }
            searchable_terms = [t for t in appmod.query_terms(query) if t not in metadata_words][:10]
            if searchable_terms:
                clauses = []
                for term in searchable_terms:
                    clauses.append("(path LIKE ? OR content LIKE ?)")
                    wildcard = f"%{term}%"; params.extend([wildcard, wildcard])
                sql += " AND (" + " OR ".join(clauses) + ")"

            rows = conn.execute(sql, params).fetchall()
            if not rows:
                indexed = conn.execute("SELECT COUNT(*) FROM files WHERE root = ?", (str(root),)).fetchone()[0]
                msg = "No indexed files found for this folder. Build / Update Index first." if indexed == 0 else "No indexed files matched this search."
                self.events.put(("done", msg)); return

            results, contents = [], {}
            for path, ext, size, modified, status, content in rows:
                self.check_stop()
                score, snippet = deterministic_score(query, path, ext, content)
                if score <= 0:
                    continue
                result = appmod.SearchResult(path, ext, modified, size, score, score, snippet, status)
                results.append(result); contents[path] = content

            results.sort(key=lambda x: x.local_score, reverse=True)
            max_results, ollama_top = options["max_results"], options["ollama_top"]
            working = results[:max(max_results, ollama_top)]
            if options["use_ollama"] and working and ollama_top > 0 and ollama_available:
                top_n = min(ollama_top, len(working))
                self.events.put(("progress_mode", ("determinate", top_n)))
                for idx, result in enumerate(working[:top_n], start=1):
                    self.check_stop()
                    self.events.put(("progress", (idx, f"Ollama {idx}/{top_n}: {Path(result.path).name}")))
                    try:
                        score, reason = ollama_rank(options["ollama_url"], options["model"], query, instructions, result, contents[result.path])
                        result.ollama_score, result.ollama_reason = score, reason
                        result.final_score = round(result.local_score * 0.45 + score * 0.55, 2)
                    except Exception as exc:
                        result.ollama_reason = f"Ollama check failed: {exc}"

            working.sort(key=lambda x: (x.final_score, x.local_score), reverse=True)
            final = working[:max_results]
            self.events.put(("results", final))
            self.events.put(("done", f"Search complete: {len(final):,} ranked results. Structured LLM instructions: {'on' if ollama_available else 'off'}."))
        except appmod.StopRequested:
            self.events.put(("done", "Search stopped."))
        except Exception as exc:
            self.events.put(("error", f"Search failed: {exc}"))
        finally:
            conn.close()

    appmod.file_type_description = file_type_description
    appmod.file_metadata = file_metadata
    appmod.deterministic_score = deterministic_score
    appmod.ollama_interpret_query = ollama_interpret_query
    appmod.ollama_rank = ollama_rank
    appmod.GenericSearchApp._search_worker = search_worker
