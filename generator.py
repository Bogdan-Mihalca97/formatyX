"""
AI-powered academic paper generator using Claude.
Generates Romanian conference papers section by section with context awareness.
"""

import os
import requests
from anthropic import Anthropic
from dotenv import load_dotenv
from docx import Document
from docx.shared import Pt, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

load_dotenv()

SECTIONS = [
    {
        "key": "title",
        "label": "Titlu",
        "max_tokens": 120,
        "instruction": (
            "Generează un titlu academic clar, precis și concis în limba română pentru această lucrare de cercetare. "
            "Titlul trebuie să reflecte exact subiectul și să fie potrivit pentru o conferință științifică. "
            "Returnează DOAR titlul, fără prefixe, ghilimele sau formatare suplimentară."
        ),
    },
    {
        "key": "rezumat",
        "label": "Rezumat",
        "max_tokens": 500,
        "instruction": (
            "Scrie un rezumat academic de aproximativ 200-250 de cuvinte în limba română. "
            "Prezintă: contextul cercetării, obiectivele, metodologia utilizată, principalele rezultate și concluzii. "
            "Scrie în stil academic formal. Nu include anteturi, prefixe sau text introductiv. "
            "Returnează DOAR textul rezumatului."
        ),
    },
    {
        "key": "keywords_ro",
        "label": "Cuvinte cheie",
        "max_tokens": 120,
        "instruction": (
            "Generează 5-7 cuvinte cheie relevante în română pentru această lucrare. "
            "Returnează DOAR cuvintele cheie separate prin virgulă și spațiu, fără alte texte."
        ),
    },
    {
        "key": "introduction",
        "label": "1. Introducere",
        "max_tokens": 900,
        "instruction": (
            "Scrie introducerea lucrării în română (450-550 cuvinte). "
            "Include: contextul teoretic și starea artei, motivația cercetării, "
            "relevanța temei în domeniu, obiectivele specifice și structura lucrării. "
            "Folosește stil academic formal, voce activă, fără referințe la 'această lucrare' sau 'acest studiu'. "
            "Nu include anteturi sau formatare suplimentară. Returnează DOAR conținutul introducerii."
        ),
    },
    {
        "key": "methodology",
        "label": "2. Metodologie",
        "max_tokens": 900,
        "instruction": (
            "Descrie metodologia cercetării în română (450-550 cuvinte). "
            "Include: modelul matematic sau cadrul teoretic, ipotezele și limitările, "
            "parametrii tehnici, procedurile de colectare și analiză a datelor. "
            "Folosește voce pasivă și limbaj tehnic precis. "
            "Nu include anteturi sau formatare suplimentară. Returnează DOAR conținutul metodologiei."
        ),
    },
    {
        "key": "results",
        "label": "3. Rezultate și discuții",
        "max_tokens": 900,
        "instruction": (
            "Prezintă și discută rezultatele obținute în română (450-550 cuvinte). "
            "Include: prezentarea datelor principale, analiza comparativă, interpretarea rezultatelor, "
            "implicațiile practice și teoretice. Conectează rezultatele cu obiectivele din introducere. "
            "Nu include anteturi sau formatare suplimentară. Returnează DOAR conținutul secțiunii."
        ),
    },
    {
        "key": "conclusions",
        "label": "4. Concluzii",
        "max_tokens": 500,
        "instruction": (
            "Scrie concluziile lucrării în română (200-280 cuvinte). "
            "Sintetizează: principalele descoperiri, contribuțiile originale, limitările cercetării "
            "și direcțiile viitoare de cercetare. "
            "Nu include anteturi sau formatare suplimentară. Returnează DOAR textul concluziilor."
        ),
    },
    {
        "key": "title_en",
        "label": "Title (EN)",
        "max_tokens": 120,
        "instruction": (
            "Translate the Romanian paper title into English. "
            "Keep it precise and academic. Return ONLY the translated title, no quotes or extra text."
        ),
    },
    {
        "key": "abstract_en",
        "label": "Abstract (EN)",
        "max_tokens": 500,
        "instruction": (
            "Translate the Romanian abstract (Rezumat) into English (200-250 words). "
            "Maintain academic tone and precision. Return ONLY the translated abstract text."
        ),
    },
    {
        "key": "keywords_en",
        "label": "Keywords (EN)",
        "max_tokens": 120,
        "instruction": (
            "Translate the Romanian keywords into English. "
            "Return ONLY the translated keywords separated by commas, no other text."
        ),
    },
    {
        "key": "references",
        "label": "Bibliografie",
        "max_tokens": 800,
        "instruction": (
            "Generează 10-12 referințe bibliografice relevante și realiste în format IEEE pentru această lucrare. "
            "Referințele trebuie să fie din domeniul temei cercetate, publicate în reviste sau conferințe recunoscute. "
            "Returnează DOAR lista numerotată de referințe, fără alte texte."
        ),
    },
]

SECTION_MAP = {s["key"]: s for s in SECTIONS}
SECTION_KEYS = [s["key"] for s in SECTIONS]

SYSTEM_PROMPT = (
    "Ești un expert în redactarea lucrărilor academice românești pentru conferințe științifice de inginerie și tehnologie. "
    "Scrii în stil academic formal, precis și coerent. Mențin consistența între secțiuni. "
    "Folosești terminologie tehnică corectă și referințe la standarde internaționale când este cazul. "
    "Nu adaugi niciodată anteturi, prefixe sau comentarii — returnezi DOAR conținutul cerut."
)


def _build_context(generated: dict) -> str:
    """Build a context string from already-generated sections."""
    parts = []
    for key in SECTION_KEYS:
        if key in generated and generated[key]:
            label = SECTION_MAP[key]["label"]
            parts.append(f"[{label}]\n{generated[key]}")
    return "\n\n".join(parts)


def generate_section(key: str, topic: str, domain: str, objectives: str,
                     keywords: str, generated: dict,
                     model: str = "claude-sonnet-4-6") -> str:
    """Generate a single section using Claude with full context."""
    client = Anthropic()
    section = SECTION_MAP[key]

    context = _build_context(generated)
    context_block = f"\n\nSecțiuni deja generate:\n{context}" if context else ""

    user_msg = (
        f"Subiect lucrare: {topic}\n"
        f"Domeniu: {domain}\n"
        f"Obiective cercetare: {objectives}\n"
        f"Cuvinte cheie sugerate: {keywords or 'nespecificate'}"
        f"{context_block}\n\n"
        f"Sarcină: {section['instruction']}"
    )

    response = client.messages.create(
        model=model,
        max_tokens=section["max_tokens"],
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": user_msg}],
    )
    return response.content[0].text.strip()


def check_grammar(text: str, language: str = "ro-RO") -> list[dict]:
    """Check grammar using LanguageTool free API. Returns list of issues."""
    try:
        resp = requests.post(
            "https://api.languagetool.org/v2/check",
            data={"text": text, "language": language},
            timeout=15,
        )
        resp.raise_for_status()
        matches = resp.json().get("matches", [])
        return [
            {
                "message": m["message"],
                "offset": m["offset"],
                "length": m["length"],
                "replacements": [r["value"] for r in m["replacements"][:3]],
                "context": m["context"]["text"],
            }
            for m in matches
        ]
    except Exception as e:
        return [{"message": f"Grammar check unavailable: {e}", "offset": 0, "length": 0, "replacements": [], "context": ""}]


def build_docx(paper: dict, output_path: str):
    """Build a conference-formatted DOCX from generated sections."""
    from create_template import create_template

    if not os.path.exists("template_conference.docx"):
        create_template("template_conference.docx")

    doc = Document("template_conference.docx")

    # Remove default empty paragraph
    for p in doc.paragraphs:
        p._element.getparent().remove(p._element)

    def add(text, style="Normal", bold=False, italic=False, center=False, size=None):
        p = doc.add_paragraph(style=style)
        run = p.add_run(text)
        if bold:
            run.bold = True
        if italic:
            run.italic = italic
        if center:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if size:
            run.font.size = Pt(size)
        return p

    def blank():
        add("", style="Normal")

    # Title (RO)
    if paper.get("title"):
        add(paper["title"].upper(), style="Title", bold=True, center=True)
        for _ in range(6):
            blank()

    # Authors
    if paper.get("authors"):
        for author in paper["authors"]:
            if author.strip():
                add(author.strip(), style="Author", center=True)
        blank()

    # Rezumat
    if paper.get("rezumat"):
        add("REZUMAT", style="Chapter Heading", bold=True)
        add(paper["rezumat"], style="Abstract Text", italic=True)
        blank()
        blank()

    # Keywords RO
    if paper.get("keywords_ro"):
        p = doc.add_paragraph(style="Abstract Text")
        run = p.add_run("Cuvinte cheie: ")
        run.bold = True
        run.italic = True
        p.add_run(paper["keywords_ro"]).italic = True
        blank()
        blank()

    # Body sections
    for key in ["introduction", "methodology", "results", "conclusions"]:
        if paper.get(key):
            label = SECTION_MAP[key]["label"]
            add(label, style="Chapter Heading", bold=True)
            for para in paper[key].split("\n\n"):
                para = para.strip()
                if para:
                    add(para, style="Normal")
            blank()

    # English title
    if paper.get("title_en"):
        blank()
        blank()
        add(paper["title_en"].upper(), style="Title", bold=True, center=True)
        blank()
        blank()

    # Abstract EN
    if paper.get("abstract_en"):
        add("ABSTRACT", style="Chapter Heading", bold=True)
        add(paper["abstract_en"], style="Abstract Text", italic=True)
        blank()
        blank()

    # Keywords EN
    if paper.get("keywords_en"):
        p = doc.add_paragraph(style="Abstract Text")
        run = p.add_run("Keywords: ")
        run.bold = True
        run.italic = True
        p.add_run(paper["keywords_en"]).italic = True
        blank()
        blank()
        blank()
        blank()

    # References
    if paper.get("references"):
        add("BIBLIOGRAFIE", style="Bibliography Header", bold=True, center=True)
        blank()
        blank()
        blank()
        for line in paper["references"].split("\n"):
            line = line.strip()
            if line:
                add(line, style="Bibliography Entry")

    doc.save(output_path)
