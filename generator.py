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

# Sections marked optional=True will be skipped if Claude returns "N/A"
SECTIONS = [
    {
        "key": "title",
        "label": "Titlu",
        "optional": False,
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
        "optional": False,
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
        "optional": False,
        "max_tokens": 120,
        "instruction": (
            "Generează 5-7 cuvinte cheie relevante în română pentru această lucrare. "
            "Returnează DOAR cuvintele cheie separate prin virgulă și spațiu, fără alte texte."
        ),
    },
    {
        "key": "nomenclature",
        "label": "Nomenclatură",
        "optional": True,
        "max_tokens": 500,
        "instruction": (
            "Dacă lucrarea implică simboluri matematice, mărimi fizice sau abrevieri tehnice specifice, "
            "listează-le ca 'Simbol — Definiție' (câte unul pe linie). "
            "Dacă lucrarea nu necesită o nomenclatură specifică sau nu se pot genera simboluri relevante "
            "pentru subiectul dat, returnează exact 'N/A'. "
            "Nu include texte introductive sau anteturi."
        ),
    },
    {
        "key": "introduction",
        "label": "1. Introducere",
        "optional": False,
        "max_tokens": 900,
        "instruction": (
            "Scrie introducerea lucrării în română (450-550 cuvinte). "
            "Include: contextul teoretic și starea artei, motivația cercetării, "
            "relevanța temei în domeniu, obiectivele specifice și structura lucrării. "
            "Folosește stil academic formal. Nu include anteturi. Returnează DOAR conținutul introducerii."
        ),
    },
    {
        "key": "relevance",
        "label": "2. Relevanța cercetării",
        "optional": False,
        "max_tokens": 700,
        "instruction": (
            "Descrie relevanța practică și teoretică a cercetării în română (300-400 cuvinte). "
            "Include: importanța domeniului, impactul potențial al rezultatelor, "
            "contribuțiile originale față de literatura existentă. "
            "Nu include anteturi. Returnează DOAR conținutul secțiunii."
        ),
    },
    {
        "key": "methodology",
        "label": "3. Metodologie (Model Matematic)",
        "optional": False,
        "max_tokens": 900,
        "instruction": (
            "Descrie metodologia și modelul matematic al cercetării în română (450-550 cuvinte). "
            "Include: ecuații și relații matematice relevante, ipotezele și limitările modelului, "
            "parametrii tehnici și condițiile de frontieră. "
            "Nu include anteturi. Returnează DOAR conținutul metodologiei."
        ),
    },
    {
        "key": "materials_methods",
        "label": "4. Materiale și Metode",
        "optional": False,
        "max_tokens": 800,
        "instruction": (
            "Descrie materialele, echipamentele și procedurile utilizate în română (400-500 cuvinte). "
            "Include: specificațiile tehnice ale materialelor/echipamentelor, "
            "procedurile de colectare și procesare a datelor, configurația experimentală. "
            "Folosește voce pasivă și limbaj tehnic precis. "
            "Nu include anteturi. Returnează DOAR conținutul secțiunii."
        ),
    },
    {
        "key": "technology_overview",
        "label": "5. Prezentare Tehnologică",
        "optional": True,
        "max_tokens": 800,
        "instruction": (
            "Dacă lucrarea implică o tehnologie specifică, prezintă o privire de ansamblu asupra "
            "tehnologiilor relevante în română (350-450 cuvinte): principii de funcționare, "
            "avantaje/dezavantaje comparativ cu alternativele, stadiul actual al tehnologiei. "
            "Dacă subiectul nu implică o componentă tehnologică distinctă față de metodologie, "
            "returnează exact 'N/A'. Nu include anteturi."
        ),
    },
    {
        "key": "case_study",
        "label": "6. Studiu de Caz",
        "optional": True,
        "max_tokens": 900,
        "instruction": (
            "Dacă lucrarea include un studiu de caz specific sau o aplicație concretă, "
            "descrie-l în română (450-550 cuvinte): contextul specific, datele de intrare, "
            "rezultatele aplicației practice, analiza performanței. "
            "Dacă subiectul nu se pretează unui studiu de caz distinct față de rezultate, "
            "returnează exact 'N/A'. Nu include anteturi."
        ),
    },
    {
        "key": "results",
        "label": "7. Rezultate și Discuții",
        "optional": False,
        "max_tokens": 900,
        "instruction": (
            "Prezintă și discută rezultatele obținute în română (450-550 cuvinte). "
            "Include: prezentarea datelor principale, analiza comparativă, "
            "interpretarea rezultatelor în raport cu obiectivele, implicațiile practice. "
            "Nu include anteturi. Returnează DOAR conținutul secțiunii."
        ),
    },
    {
        "key": "standards",
        "label": "8. Standarde și Reglementări",
        "optional": True,
        "max_tokens": 700,
        "instruction": (
            "Dacă lucrarea este relevantă pentru standarde tehnice, reglementări sau norme specifice, "
            "discută-le în română (300-400 cuvinte): standardele aplicabile, cerințele de conformitate, "
            "implicațiile pentru proiectare sau implementare. "
            "Dacă nu există standarde specific aplicabile subiectului, returnează exact 'N/A'. "
            "Nu include anteturi."
        ),
    },
    {
        "key": "future_challenges",
        "label": "9. Provocări Viitoare",
        "optional": False,
        "max_tokens": 700,
        "instruction": (
            "Identifică și discută provocările viitoare și direcțiile de cercetare în română (300-400 cuvinte). "
            "Include: limitările actuale ale cercetării, oportunități de îmbunătățire, "
            "direcții viitoare de cercetare și dezvoltare. "
            "Nu include anteturi. Returnează DOAR conținutul secțiunii."
        ),
    },
    {
        "key": "environmental",
        "label": "10. Sustenabilitate și Impact de Mediu",
        "optional": True,
        "max_tokens": 700,
        "instruction": (
            "Dacă lucrarea are implicații de mediu sau sustenabilitate relevante, "
            "analizează-le în română (300-400 cuvinte): impactul de mediu al soluțiilor propuse, "
            "beneficiile de sustenabilitate, amprenta de carbon sau eficiența energetică. "
            "Dacă subiectul nu are o componentă de mediu semnificativă, returnează exact 'N/A'. "
            "Nu include anteturi."
        ),
    },
    {
        "key": "conclusions",
        "label": "11. Concluzii",
        "optional": False,
        "max_tokens": 500,
        "instruction": (
            "Scrie concluziile lucrării în română (200-280 cuvinte). "
            "Sintetizează: principalele descoperiri, contribuțiile originale, "
            "limitările cercetării și recomandările pentru viitor. "
            "Nu include anteturi. Returnează DOAR textul concluziilor."
        ),
    },
    {
        "key": "title_en",
        "label": "Title (EN)",
        "optional": False,
        "max_tokens": 120,
        "instruction": (
            "Translate the Romanian paper title into English. "
            "Keep it precise and academic. Return ONLY the translated title, no quotes or extra text."
        ),
    },
    {
        "key": "abstract_en",
        "label": "Abstract (EN)",
        "optional": False,
        "max_tokens": 500,
        "instruction": (
            "Translate the Romanian abstract (Rezumat) into English (200-250 words). "
            "Maintain academic tone and precision. Return ONLY the translated abstract text."
        ),
    },
    {
        "key": "keywords_en",
        "label": "Keywords (EN)",
        "optional": False,
        "max_tokens": 120,
        "instruction": (
            "Translate the Romanian keywords into English. "
            "Return ONLY the translated keywords separated by commas, no other text."
        ),
    },
    {
        "key": "references",
        "label": "Bibliografie",
        "optional": False,
        "max_tokens": 900,
        "instruction": (
            "Generează 10-14 referințe bibliografice relevante și realiste în format IEEE pentru această lucrare. "
            "Referințele trebuie să fie din domeniul temei cercetate, publicate în reviste sau conferințe recunoscute. "
            "Returnează DOAR lista numerotată de referințe, fără alte texte."
        ),
    },
]

SECTION_MAP = {s["key"]: s for s in SECTIONS}
SECTION_KEYS = [s["key"] for s in SECTIONS]

SYSTEM_PROMPT = (
    "Ești un expert în redactarea lucrărilor academice românești pentru conferințe științifice de inginerie și tehnologie. "
    "Scrii în stil academic formal, precis și coerent. Menții consistența între secțiuni. "
    "Folosești terminologie tehnică corectă și referințe la standarde internaționale când este cazul. "
    "Nu adaugi niciodată anteturi, prefixe sau comentarii — returnezi DOAR conținutul cerut. "
    "Dacă o secțiune opțională nu este aplicabilă, returnezi exact 'N/A'."
)


def _build_context(generated: dict) -> str:
    """Build a context string from already-generated sections (excluding N/A ones)."""
    parts = []
    for key in SECTION_KEYS:
        text = generated.get(key, "")
        if text and text.strip() != "N/A":
            label = SECTION_MAP[key]["label"]
            parts.append(f"[{label}]\n{text}")
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
        f"Obiective cercetare: {objectives or 'nespecificate'}\n"
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
    text = response.content[0].text.strip()

    # Treat near-empty or explicit N/A as not applicable
    if section.get("optional") and text.upper() in ("N/A", "NA", "N/A.", "-", ""):
        return "N/A"
    return text


def check_grammar(text: str, language: str = "ro-RO") -> list[dict]:
    """Check grammar using LanguageTool free API."""
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

    def section_body(text):
        for para in text.split("\n\n"):
            para = para.strip()
            if para:
                add(para, style="Normal")

    def is_na(key):
        return (paper.get(key) or "").strip().upper() in ("N/A", "NA", "")

    # ── Title (RO) ──
    if not is_na("title"):
        add(paper["title"].upper(), style="Title", bold=True, center=True)
        for _ in range(6):
            blank()

    # ── Authors ──
    authors = paper.get("authors") or []
    if isinstance(authors, str):
        authors = [a.strip() for a in authors.splitlines() if a.strip()]
    for author in authors:
        add(author, style="Author", center=True)
    if authors:
        blank()

    # ── Rezumat ──
    if not is_na("rezumat"):
        add("REZUMAT", style="Chapter Heading", bold=True)
        add(paper["rezumat"], style="Abstract Text", italic=True)
        blank()
        blank()

    # ── Cuvinte cheie (RO) ──
    if not is_na("keywords_ro"):
        p = doc.add_paragraph(style="Abstract Text")
        r = p.add_run("Cuvinte cheie: ")
        r.bold = True
        r.italic = True
        p.add_run(paper["keywords_ro"]).italic = True
        blank()
        blank()

    # ── Nomenclatură ──
    if not is_na("nomenclature"):
        add("NOMENCLATURĂ", style="Chapter Heading", bold=True)
        section_body(paper["nomenclature"])
        blank()

    # ── Body sections (numbered) ──
    BODY_SECTIONS = [
        ("introduction",       "1. Introducere"),
        ("relevance",          "2. Relevanța Cercetării"),
        ("methodology",        "3. Metodologie (Model Matematic)"),
        ("materials_methods",  "4. Materiale și Metode"),
        ("technology_overview","5. Prezentare Tehnologică"),
        ("case_study",         "6. Studiu de Caz"),
        ("results",            "7. Rezultate și Discuții"),
        ("standards",          "8. Standarde și Reglementări"),
        ("future_challenges",  "9. Provocări Viitoare"),
        ("environmental",      "10. Sustenabilitate și Impact de Mediu"),
        ("conclusions",        "11. Concluzii"),
    ]

    for key, heading in BODY_SECTIONS:
        if not is_na(key):
            add(heading, style="Chapter Heading", bold=True)
            section_body(paper[key])
            blank()

    # ── English Title ──
    if not is_na("title_en"):
        blank()
        blank()
        add(paper["title_en"].upper(), style="Title", bold=True, center=True)
        blank()
        blank()

    # ── Abstract (EN) ──
    if not is_na("abstract_en"):
        add("ABSTRACT", style="Chapter Heading", bold=True)
        add(paper["abstract_en"], style="Abstract Text", italic=True)
        blank()
        blank()

    # ── Keywords (EN) ──
    if not is_na("keywords_en"):
        p = doc.add_paragraph(style="Abstract Text")
        r = p.add_run("Keywords: ")
        r.bold = True
        r.italic = True
        p.add_run(paper["keywords_en"]).italic = True
        for _ in range(4):
            blank()

    # ── References ──
    if not is_na("references"):
        add("BIBLIOGRAFIE", style="Bibliography Header", bold=True, center=True)
        blank()
        blank()
        blank()
        for line in paper["references"].split("\n"):
            line = line.strip()
            if line:
                add(line, style="Bibliography Entry")

    doc.save(output_path)
