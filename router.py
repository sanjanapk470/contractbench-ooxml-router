import difflib
import json
import time
from datetime import datetime, timezone
from typing import Literal

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from dotenv import load_dotenv
from google import genai
from google.genai import types
from pydantic import BaseModel, Field

# Loads GEMINI_API_KEY from your local .env file
load_dotenv()
client = genai.Client()


class ClauseDecision(BaseModel):
    action: Literal["EDIT", "PRESERVE", "ESCALATE"] = Field(
        description="EDIT for surgical redline, PRESERVE if within acceptable fallback, ESCALATE if human approval is required."
    )
    rule_id: str | None = Field(
        default=None, description="Matching playbook rule ID, if applicable."
    )
    rationale: str = Field(
        description="Concise legal/commercial reasoning for the routing decision."
    )
    revised_clause: str | None = Field(
        default=None,
        description="Required ONLY when action == 'EDIT'. Must preserve all non-violating words verbatim."
    )


class ClauseBatchItem(BaseModel):
    index: int
    decision: ClauseDecision


class ClauseBatchDecision(BaseModel):
    items: list[ClauseBatchItem]


def create_sample_msa(path: str):
    """Generates a 5-clause SaaS MSA testing PRESERVE, EDIT, and ESCALATE paths."""
    doc = Document()
    doc.add_heading("MASTER SUBSCRIPTION AGREEMENT", level=1)
    doc.add_paragraph(
        "1. Services. Provider shall make the AI Contract Intelligence Platform available to Customer during the Subscription Term."
    )
    # Tests PRESERVE (Net 45 is within Enterprise fallback)
    doc.add_paragraph(
        "2. Payment Terms. Customer shall pay all undisputed invoices within Net 45 days of invoice date."
    )
    # Tests EDIT (Nominal $100 cap violates playbook -> surgical inline edit)
    doc.add_paragraph(
        "3. Limitation of Liability. Provider's total aggregate liability under this Agreement shall not exceed $100 USD."
    )
    # Tests ESCALATE (Uncapped IP indemnity -> flag for human counsel instead of silent rewrite)
    doc.add_paragraph(
        "4. Indemnification. Provider shall defend and indemnify Customer against all third-party claims with unlimited liability."
    )
    # Tests EDIT (Texas venue -> surgical inline edit to California)
    doc.add_paragraph(
        "5. Governing Law. This Agreement shall be governed by the laws of the State of Texas, with exclusive venue in Austin, TX."
    )
    doc.save(path)


def route_clauses_batch(clauses: list[tuple[int, str]], playbook_data: dict) -> dict[int, ClauseDecision]:
    """Dynamically discovers available Gemini Flash models on your API key and runs batch routing."""
    formatted_clauses = "\n".join([f"[{idx}] {text}" for idx, text in clauses])
    prompt = f"""
    You are an expert commercial contracts agent. Evaluate each numbered contract clause below against the deal context and playbook.

    Deal Context & Playbook:
    {json.dumps(playbook_data, indent=2)}

    Contract Clauses:
    {formatted_clauses}

    Routing Rules:
    1. PRESERVE: Complies with standard position, falls within acceptable fallback, or is low-risk boilerplate. No stylistic edits.
    2. EDIT: Violates a rule that authorizes an automated fallback/standard fix. Provide `revised_clause` starting with the exact clause number and heading, making only surgical word-level edits to the violating terms.    
    3. ESCALATE: Triggers an escalation threshold (e.g., unlimited liability, Net 90+) requiring human approval.
    """

    # Discover all models enabled for your API key that support generateContent
    available = [
        m.name.replace("models/", "")
        for m in client.models.list()
        if "generateContent" in (m.supported_actions or [])
        and "gemini" in m.name.lower()
        and ("flash" in m.name.lower() or "lite" in m.name.lower())
    ]
    # Prefer lite/flash variants first
    candidates = sorted(available, key=lambda n: ("lite" not in n, "flash" not in n)) or ["gemini-3.8-flash"]

    for model_name in candidates:
        for attempt in range(3):
            try:
                print(f"[*] Calling {model_name} (attempt {attempt + 1}, batch of {len(clauses)} clauses)...")
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=ClauseBatchDecision,
                        temperature=0.0,
                        http_options={"timeout": 20000},
                    ),
                )
                batch = ClauseBatchDecision.model_validate_json(response.text)
                return {item.index: item.decision for item in batch.items}
            except Exception as e:
                print(f"[!] {model_name} failed ({type(e).__name__}), retrying...")
                time.sleep(2)

    raise RuntimeError(f"All discovered models failed: {candidates}")


def _build_run(text: str, is_deletion: bool = False) -> OxmlElement:
    run = OxmlElement("w:r")
    node = OxmlElement("w:delText" if is_deletion else "w:t")
    node.set(qn("xml:space"), "preserve")
    node.text = text
    run.append(node)
    return run


def apply_surgical_ooxml_redline(paragraph, old_text: str, new_text: str, rev_id: int, author="ContractRouter AI") -> int:
    """Injects word-level OOXML <w:del> and <w:ins> tags while keeping unchanged tokens clean."""
    timestamp = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    p_el = paragraph._p

    for child in list(p_el):
        if child.tag != qn("w:pPr"):
            p_el.remove(child)

    old_tokens = old_text.split(" ")
    new_tokens = new_text.split(" ")
    matcher = difflib.SequenceMatcher(None, old_tokens, new_tokens)

    for op, i1, i2, j1, j2 in matcher.get_opcodes():
        if op == "equal":
            p_el.append(_build_run(" ".join(old_tokens[i1:i2]) + " ", is_deletion=False))
        elif op in ("replace", "delete"):
            del_el = OxmlElement("w:del")
            del_el.set(qn("w:id"), str(rev_id))
            del_el.set(qn("w:author"), author)
            del_el.set(qn("w:date"), timestamp)
            del_el.append(_build_run(" ".join(old_tokens[i1:i2]) + " ", is_deletion=True))
            p_el.append(del_el)
            rev_id += 1
        if op in ("replace", "insert"):
            ins_el = OxmlElement("w:ins")
            ins_el.set(qn("w:id"), str(rev_id))
            ins_el.set(qn("w:author"), author)
            ins_el.set(qn("w:date"), timestamp)
            ins_el.append(_build_run(" ".join(new_tokens[j1:j2]) + " ", is_deletion=False))
            p_el.append(ins_el)
            rev_id += 1

    return rev_id


def attach_escalation_marker(paragraph, rule_id: str | None, rationale: str):
    """Appends a visible inline review flag for escalated clauses requiring human sign-off."""
    run = paragraph.add_run(f"  [ESCALATE ({rule_id or 'HIGH_RISK'}): {rationale}]")
    run.bold = True


def process_contract(input_docx: str, output_docx: str, playbook_path: str):
    with open(playbook_path, "r") as f:
        playbook_data = json.load(f)

    doc = Document(input_docx)
    rev_id = 1
    metrics = {"EDIT": 0, "PRESERVE": 0, "ESCALATE": 0}

    clauses = [
        (i, p.text.strip())
        for i, p in enumerate(doc.paragraphs)
        if p.text.strip() and not p.style.name.startswith("Heading")
    ]

    print("\n--- Running Clause Judgment & OOXML Routing Pipeline ---")
    decisions_by_idx = route_clauses_batch(clauses, playbook_data)

    for idx, text in clauses:
        decision = decisions_by_idx[idx]
        metrics[decision.action] += 1
        print(f"[{decision.action:8}] {text[:52]}... -> {decision.rationale}")

        para = doc.paragraphs[idx]
        if decision.action == "EDIT" and decision.revised_clause:
            rev_id = apply_surgical_ooxml_redline(
                para, old_text=text, new_text=decision.revised_clause, rev_id=rev_id
            )
        elif decision.action == "ESCALATE":
            attach_escalation_marker(para, decision.rule_id, decision.rationale)

    doc.save(output_docx)
    print("\n--- Run Summary ---")
    print(f"Decisions: {metrics}")
    print(f"Saved native Track Changes output -> {output_docx}\n")


if __name__ == "__main__":
    create_sample_msa("sample_msa.docx")
    process_contract("sample_msa.docx", "redlined_msa.docx", "playbook.json")