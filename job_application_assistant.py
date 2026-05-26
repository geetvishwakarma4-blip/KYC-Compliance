#!/usr/bin/env python3
"""Job Application Assistant for legal/compliance/risk profiles.

Features:
1) Accept LinkedIn job links with pasted descriptions or direct pasted JDs.
2) Compare each job with a fixed candidate profile.
3) Output match score /100.
4) List missing keywords.
5) Generate customized resume summary + key skills.
6) Draft short cover message.
7) Export full report to .docx (Word).
"""
from __future__ import annotations
import argparse
import json
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Iterable
from zipfile import ZipFile, ZIP_DEFLATED
from xml.sax.saxutils import escape

PROFILE = {
    "years_experience": 16,
    "domains": ["legal", "compliance", "risk", "shipping"],
    "capabilities": [
        "aml", "sanctions", "trade compliance", "abac", "contract management",
        "india dpdp act", "corporate governance", "regulatory advisory"
    ],
    "target_roles": [
        "compliance head", "legal counsel", "risk", "sanctions compliance", "shipping compliance"
    ],
}

KEYWORD_BANK = {
    "core": [
        "aml", "anti money laundering", "kyc", "sanctions", "ofac", "trade compliance", "export control",
        "abac", "anti bribery", "corruption", "contract management", "regulatory", "governance", "risk",
        "shipping", "maritime", "legal", "compliance", "due diligence", "investigations", "policy",
        "monitoring", "internal controls", "india dpdp act", "data protection", "privacy"
    ],
    "leadership": ["head", "lead", "manager", "director", "stakeholder", "board", "cross functional"],
    "tools": ["excel", "power bi", "sap", "oracle", "erp", "grc"],
}

@dataclass
class JobInput:
    title: str
    description: str
    link: str = ""


def normalize(text: str) -> str:
    return re.sub(r"\s+", " ", text.lower()).strip()


def present_keywords(text: str, keywords: Iterable[str]) -> set[str]:
    t = normalize(text)
    return {k for k in keywords if re.search(rf"\b{re.escape(k)}\b", t)}


def score_job(job: JobInput) -> dict:
    jd = normalize(job.description)
    core = present_keywords(jd, KEYWORD_BANK["core"])
    leadership = present_keywords(jd, KEYWORD_BANK["leadership"])
    tools = present_keywords(jd, KEYWORD_BANK["tools"])

    # Weighted rubric
    score = 0.0
    score += min(55, (len(core) / max(len(KEYWORD_BANK["core"]), 1)) * 55)
    score += min(15, (len(leadership) / max(len(KEYWORD_BANK["leadership"]), 1)) * 15)
    score += min(5, (len(tools) / max(len(KEYWORD_BANK["tools"]), 1)) * 5)

    # Role fit
    role_hits = present_keywords(job.title + " " + jd, PROFILE["target_roles"])
    score += 15 if role_hits else 7

    # Experience fit signal
    if re.search(r"(1[0-9]|[2-9][0-9])\+?\s*(years|yrs)", jd):
        score += 10
    elif re.search(r"[5-9]\+?\s*(years|yrs)", jd):
        score += 7
    else:
        score += 5

    final_score = int(round(min(score, 100)))
    missing = [k for k in KEYWORD_BANK["core"] if k not in core][:15]

    resume_summary = (
        f"Compliance and legal leader with {PROFILE['years_experience']}+ years across shipping, risk, "
        "and regulatory frameworks. Proven delivery in AML/KYC, sanctions screening, trade compliance, "
        "ABAC programs, contract governance, and advisory support to senior stakeholders. "
        "Brings strong command of India DPDP Act, corporate governance, and cross-border compliance controls, "
        "aligned to the requirements of this role."
    )

    key_skills = sorted(set([
        "AML/KYC", "Sanctions Compliance", "Trade Compliance", "ABAC", "Contract Management",
        "Regulatory Advisory", "Corporate Governance", "India DPDP Act", "Risk Assessment",
        "Shipping/Maritime Compliance", "Policy & Controls", "Stakeholder Management"
    ] + [kw.title() for kw in list(core)[:6]]))

    cover = (
        f"Subject: Application for {job.title}\n\n"
        f"Hi Hiring Team,\n\n"
        f"I’m interested in the {job.title} opportunity. I bring {PROFILE['years_experience']} years of experience "
        "across legal, compliance, and risk in the shipping industry, with deep exposure to AML, sanctions, "
        "trade compliance, ABAC, and contract management. I have also advised on India DPDP Act and corporate "
        "governance matters, and I’m confident I can add immediate value to your compliance framework.\n\n"
        "I would welcome the chance to discuss how my background maps to your priorities.\n\n"
        "Best regards"
    )

    return {
        "title": job.title,
        "link": job.link,
        "score": final_score,
        "matched_keywords": sorted(core),
        "missing_keywords": missing,
        "resume_summary": resume_summary,
        "key_skills": key_skills,
        "cover_message": cover,
    }


def _paragraph_xml(text: str) -> str:
    safe = escape(text)
    return f"<w:p><w:r><w:t xml:space='preserve'>{safe}</w:t></w:r></w:p>"


def export_docx(results: list[dict], path: Path) -> None:
    body_parts = [
        _paragraph_xml("Job Application Assistant Report"),
        _paragraph_xml(f"Generated: {date.today().isoformat()}"),
        _paragraph_xml(""),
    ]
    for i, r in enumerate(results, 1):
        body_parts += [
            _paragraph_xml(f"{i}. {r['title']}"),
            _paragraph_xml(f"Link: {r['link'] or 'N/A'}"),
            _paragraph_xml(f"Match score: {r['score']}/100"),
            _paragraph_xml("Missing keywords: " + ", ".join(r["missing_keywords"]) if r["missing_keywords"] else "Missing keywords: None"),
            _paragraph_xml("Resume summary:"),
            _paragraph_xml(r["resume_summary"]),
            _paragraph_xml("Key skills: " + ", ".join(r["key_skills"])),
            _paragraph_xml("Cover message:"),
        ]
        for line in r["cover_message"].splitlines():
            body_parts.append(_paragraph_xml(line))
        body_parts.append(_paragraph_xml(""))

    document_xml = f"""<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:document xmlns:wpc='http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'
 xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006'
 xmlns:o='urn:schemas-microsoft-com:office:office'
 xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
 xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'
 xmlns:v='urn:schemas-microsoft-com:vml'
 xmlns:wp14='http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'
 xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
 xmlns:w10='urn:schemas-microsoft-com:office:word'
 xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
 xmlns:w14='http://schemas.microsoft.com/office/word/2010/wordml'
 xmlns:wpg='http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
 xmlns:wpi='http://schemas.microsoft.com/office/word/2010/wordprocessingInk'
 xmlns:wne='http://schemas.microsoft.com/office/word/2006/wordml'
 xmlns:wps='http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
 mc:Ignorable='w14 wp14'>
<w:body>{''.join(body_parts)}<w:sectPr/></w:body></w:document>"""

    content_types = """<?xml version='1.0' encoding='UTF-8'?>
<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>
<Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>
<Default Extension='xml' ContentType='application/xml'/>
<Override PartName='/word/document.xml' ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'/>
</Types>"""

    rels = """<?xml version='1.0' encoding='UTF-8'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'>
<Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='word/document.xml'/>
</Relationships>"""

    with ZipFile(path, "w", compression=ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("word/document.xml", document_xml)


def load_jobs(input_path: Path | None) -> list[JobInput]:
    if input_path:
        data = json.loads(input_path.read_text(encoding="utf-8"))
        return [JobInput(title=i.get("title", "Untitled role"), description=i["description"], link=i.get("link", "")) for i in data]

    print("Paste one or more job descriptions. End each with a line containing only ---END---. Enter blank title to stop.")
    jobs: list[JobInput] = []
    while True:
        title = input("Job title (or Enter to finish): ").strip()
        if not title:
            break
        link = input("LinkedIn/job link (optional): ").strip()
        print("Paste job description, then type ---END--- on a new line:")
        lines = []
        while True:
            line = input()
            if line.strip() == "---END---":
                break
            lines.append(line)
        jobs.append(JobInput(title=title, link=link, description="\n".join(lines)))
    return jobs


def main() -> None:
    parser = argparse.ArgumentParser(description="Job Application Assistant")
    parser.add_argument("--input", type=Path, help="JSON file with list of jobs: [{title,link,description}, ...]")
    parser.add_argument("--output", type=Path, default=Path("job_application_report.docx"), help="Output .docx path")
    parser.add_argument("--print-json", action="store_true", help="Print structured JSON to stdout")
    args = parser.parse_args()

    jobs = load_jobs(args.input)
    if not jobs:
        print("No jobs provided.")
        return

    results = [score_job(j) for j in jobs]
    export_docx(results, args.output)

    if args.print_json:
        print(json.dumps(results, indent=2))
    else:
        for r in results:
            print(f"\n{r['title']}: {r['score']}/100")
            print("Missing:", ", ".join(r["missing_keywords"]) if r["missing_keywords"] else "None")
        print(f"\nWord report exported to: {args.output}")


if __name__ == "__main__":
    main()
