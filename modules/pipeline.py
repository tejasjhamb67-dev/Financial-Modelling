"""End-to-end pipeline orchestration.

Runs the mandatory staged pipeline:

  documents -> ingestion -> understanding -> disclosure mapping -> extraction ->
  normalization -> source database -> GATE 1 -> model specification -> GATE 2 ->
  Excel generation -> formula/link validation -> model QC -> GATE 3 ->
  final Excel + build/diligence reports.

Each stage is explicit and its intermediate artifacts are persisted so the
RAW -> NORMALIZED -> VALIDATED -> MODEL chain is auditable.
"""

from __future__ import annotations

import glob
import json
import os
from dataclasses import dataclass, field
from datetime import date
from typing import Optional

from excel.qc.qc import run_qc
from excel.workbook_builder.builder import WorkbookBuilder
from modules.config_loader import model_config
from modules.disclosure_mapping.mapper import DisclosureMap, map_disclosures
from modules.ingestion.loader import load_package
from modules.model_spec import ModelSpec, build_spec
from modules.normalization import restatements as restate
from modules.normalization.normalizer import normalize
from modules.schemas import SourceDatabase
from modules.validation.gates import GateResult, gate1_data_validation, gate2_model_spec
from reports.build_report import build_report
from reports.diligence_report import diligence

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


@dataclass
class PipelineResult:
    company: str
    workbook_path: Optional[str] = None
    build_report_path: Optional[str] = None
    diligence_path: Optional[str] = None
    source_registry_path: Optional[str] = None
    gate_results: list[GateResult] = field(default_factory=list)
    qc_result: Optional[GateResult] = None
    db: Optional[SourceDatabase] = None
    spec: Optional[ModelSpec] = None
    stopped_reason: Optional[str] = None

    @property
    def success(self) -> bool:
        return self.workbook_path is not None and self.stopped_reason is None


def _ingest(package: Optional[str], documents: Optional[str],
            company: Optional[str]) -> SourceDatabase:
    if package:
        db = load_package(package)
        if company:
            db.company = company
        return db
    if documents:
        # lazy import so pdfplumber is only needed on the PDF path
        from modules.document_analysis.pdf_reader import parse_pdf
        from modules.extraction.table_extractor import extract_candidates
        pdfs = sorted(glob.glob(os.path.join(documents, "*.pdf")))
        if not pdfs:
            raise FileNotFoundError(f"No PDF documents found in {documents}")
        name = company or os.path.basename(os.path.dirname(documents.rstrip("/"))) or "Company"
        merged = SourceDatabase(name)
        for i, pdf in enumerate(pdfs):
            parsed = parse_pdf(pdf, raw_dir=os.path.join(REPO, "data", "raw"))
            parsed.document.doc_id = f"DOC{i+1:02d}"
            sub = extract_candidates(parsed, name,
                                     extracted_dir=os.path.join(REPO, "data", "extracted"))
            merged.add_document(parsed.document)
            for dp in sub.datapoints:
                merged.add(dp)
        return merged
    raise ValueError("Provide either a package file or a documents directory.")


def run(package: Optional[str] = None, documents: Optional[str] = None,
        output_dir: Optional[str] = None, company: Optional[str] = None,
        force: bool = False) -> PipelineResult:
    build_date = date.today().isoformat()
    output_dir = output_dir or os.path.join(REPO, "outputs", "models")
    os.makedirs(output_dir, exist_ok=True)

    # 1-6. ingest -> normalize -> restatements -> persist
    db = _ingest(package, documents, company)
    db_norm = normalize(db)
    restate.apply(db_norm)

    reg_dir = os.path.join(REPO, "sources", "source_registry")
    os.makedirs(reg_dir, exist_ok=True)
    safe = "".join(c if c.isalnum() else "_" for c in db_norm.company)
    reg_path = os.path.join(reg_dir, f"{safe}_source_registry.json")
    db_norm.save(reg_path)
    norm_dir = os.path.join(REPO, "data", "normalized")
    os.makedirs(norm_dir, exist_ok=True)
    db_norm.save(os.path.join(norm_dir, f"{safe}_normalized.json"))

    result = PipelineResult(company=db_norm.company, db=db_norm,
                            source_registry_path=reg_path)

    # GATE 1
    g1 = gate1_data_validation(db_norm)
    result.gate_results.append(g1)
    if not g1.passed and not force:
        result.stopped_reason = "Gate 1 (data validation) failed; use force to override."
        _write_reports(result, db_norm, None, [g1], None, build_date, output_dir, safe)
        return result

    # disclosure mapping + spec (GATE 2)
    dmap = map_disclosures(db_norm)
    spec = build_spec(db_norm, dmap)
    result.spec = spec
    g2 = gate2_model_spec(spec)
    result.gate_results.append(g2)
    if not g2.passed and not force:
        result.stopped_reason = "Gate 2 (model specification) failed; use force to override."
        _write_reports(result, db_norm, spec, [g1, g2], None, build_date, output_dir, safe)
        return result

    # persist model spec
    val_dir = os.path.join(REPO, "data", "validated")
    os.makedirs(val_dir, exist_ok=True)
    with open(os.path.join(val_dir, f"{safe}_model_spec.json"), "w", encoding="utf-8") as fh:
        json.dump(spec.to_dict(), fh, indent=2, default=str)

    # Excel generation
    wb_path = os.path.join(output_dir, f"{safe}_model.xlsx")
    builder = WorkbookBuilder(spec, db_norm, build_date=build_date)
    registry = builder.build(wb_path)
    result.workbook_path = wb_path

    # export lineage
    lineage_path = os.path.join(val_dir, f"{safe}_cell_lineage.json")
    with open(lineage_path, "w", encoding="utf-8") as fh:
        json.dump([l.__dict__ for l in registry.lineage], fh, indent=2, default=str)

    # GATE 3 QC
    qc = run_qc(wb_path, spec)
    result.qc_result = qc
    result.gate_results.append(qc)

    # reports
    _write_reports(result, db_norm, spec, [g1, g2], qc, build_date, output_dir, safe)
    return result


_MODULES = [
    "ingestion", "document_analysis", "disclosure_mapping", "extraction", "normalization",
    "validation", "financial_statements (accounting taxonomy + statement engine)",
    "three_statement", "revenue", "operating_drivers", "working_capital", "capex",
    "depreciation", "debt", "tax", "ratios", "valuation", "sensitivities", "checks",
    "excel.workbook_builder", "excel.formatting", "excel.linking", "excel.qc",
]


def _write_reports(result, db, spec, gates, qc, build_date, output_dir, safe):
    if spec is None:
        return
    qc_res = qc or GateResult("Gate 3 - Final Workbook QC (not run)")
    br = build_report.generate(db, spec, gates, qc_res, _MODULES, build_date)
    dr = diligence.generate(db, spec, build_date)
    br_path = os.path.join(output_dir, f"{safe}_MODEL_BUILD_REPORT.md")
    dr_path = os.path.join(output_dir, f"{safe}_MODEL_DILIGENCE.md")
    with open(br_path, "w", encoding="utf-8") as fh:
        fh.write(br)
    with open(dr_path, "w", encoding="utf-8") as fh:
        fh.write(dr)
    result.build_report_path = br_path
    result.diligence_path = dr_path
