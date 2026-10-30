#!/usr/bin/env python3
"""Offline batch de-identification for medical records or DICOM studies.

This tool is deliberately conservative: files that cannot be safely processed or
verified are copied to ``manual_review`` instead of ``share_ready``.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import logging
import os
import re
import secrets
import shutil
import sys
import tempfile
import traceback
import warnings
import zipfile
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable, Iterable

from openpyxl import Workbook, load_workbook


PID_RE = re.compile(r"Patient_[0-9A-F]{8}$")
ID_CARD_RE = re.compile(r"(?<![0-9Xx])\d{17}[0-9Xx](?![0-9Xx])")
PHONE_RE = re.compile(r"(?<!\d)(?:\+?86[- ]?)?1[3-9]\d{9}(?!\d)")
LANDLINE_RE = re.compile(r"(?<!\d)(?:0\d{2,3}[- ]?)?\d{7,8}(?!\d)")
LABELED_ID_RE = re.compile(
    r"(?i)(病案号|住院号|门诊号|患者编号|病历号|医保卡号|社保卡号|patient\s*id|medical\s*record\s*(?:no|number))"
    r"\s*[:：]?\s*(?!Patient_[0-9A-F]{8}\b)[A-Za-z0-9_-]{4,32}"
)
ADDRESS_RE = re.compile(
    r"(家庭住址|现住址|户籍地址|联系地址|地址)\s*[:：]\s*[^\n\r;；]{4,80}"
)
BIRTH_RE = re.compile(
    r"(出生日期|出生年月|出生时间|date\s*of\s*birth|DOB)\s*[:：]?\s*"
    r"(?:(?:19|20)\d{2}[-/.年]\d{1,2}[-/.月]\d{1,2}日?|[^,，\n\r;；]{4,32})",
    re.I,
)
LABELED_NAME_RE = re.compile(
    r"(病人姓名|患者姓名|姓名|patient\s*name)\s*[:：]?\s*"
    r"[^,，\n\r;；]{2,80}?"
    r"(?=\s*(?:[,，;；]|$)|\s+(?:诊断|手机|电话|联系电话|身份证|住院号|门诊号|病案号)\s*[:：]?)",
    re.I | re.M,
)
LABELED_DATE_RE = re.compile(
    r"(出生日期|出生年月|出生时间|检查日期|文件集创建时间|date\s*of\s*birth|DOB)"
    r"\s*[:：]?\s*[^,，\n\r;；]{2,80}", re.I
)

TEXT_EXTS = {".txt", ".md", ".log", ".ini", ".yaml", ".yml", ".html", ".htm"}
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp"}
UNSAFE_EXTS = {".doc", ".rtf", ".odt", ".zip", ".rar", ".7z"}


@dataclass
class Finding:
    kind: str
    detail: str


@dataclass
class Result:
    status: str
    findings: list[Finding] = field(default_factory=list)


def safe_log_value(value: str) -> str:
    """Return a non-reversible short fingerprint suitable for logs."""
    return hashlib.sha256(value.encode("utf-8", "ignore")).hexdigest()[:12]


def source_key(root: Path, folder: Path) -> str:
    # Include the dataset root so equally named folders from separate datasets never collide.
    identity = str(root.resolve()).casefold() + "|" + folder.relative_to(root).as_posix()
    return hashlib.sha256(identity.encode("utf-8")).hexdigest()


def paired_source_key(records_root: Path, dicom_root: Path, folder_name: str) -> str:
    """Stable pairing key for an explicitly matched records/DICOM folder pair."""
    identity = "both|" + str(records_root.resolve()).casefold() + "|" + str(dicom_root.resolve()).casefold()
    identity += "|" + folder_name
    return hashlib.sha256(identity.encode("utf-8")).hexdigest()


class MappingStore:
    HEADERS = ["原始患者姓名", "Patient ID", "SourceFolderKey", "创建时间"]

    def __init__(self, path: Path):
        self.path = path
        self.rows: dict[str, tuple[str, str, str]] = {}
        self.used: set[str] = set()
        if path.exists():
            wb = load_workbook(path, read_only=True, data_only=True)
            ws = wb.active
            headers = [c.value for c in next(ws.iter_rows())]
            required = set(self.HEADERS[:3])
            if not required.issubset(set(headers)):
                raise ValueError("映射表格式不兼容，缺少必要字段")
            idx = {v: headers.index(v) for v in required}
            for row in ws.iter_rows(values_only=True):
                name, pid, key = row[idx["原始患者姓名"]], row[idx["Patient ID"]], row[idx["SourceFolderKey"]]
                if key and pid and PID_RE.fullmatch(str(pid)):
                    self.rows[str(key)] = (str(name or ""), str(pid), "")
                    self.used.add(str(pid))
            wb.close()

    def get_or_create(self, key: str, name: str) -> str:
        if key in self.rows:
            return self.rows[key][1]
        while True:
            pid = "Patient_" + secrets.token_hex(4).upper()
            if pid not in self.used:
                break
        self.used.add(pid)
        self.rows[key] = (name, pid, datetime.now().isoformat(timespec="seconds"))
        self.save()
        return pid

    def get_or_create_paired(self, pair_key: str, component_keys: list[str], name: str) -> str:
        """Reuse a prior single-mode ID and bind all explicit pair keys to it."""
        keys = [pair_key, *component_keys]
        existing = {self.rows[key][1] for key in keys if key in self.rows}
        if len(existing) > 1:
            raise ValueError("该病历与影像目录已有相互冲突的 Patient ID 映射，拒绝自动合并")
        if existing:
            pid = next(iter(existing))
        else:
            pid = self.get_or_create(pair_key, name)
        changed = False
        for key in keys:
            if key not in self.rows:
                self.rows[key] = (name, pid, datetime.now().isoformat(timespec="seconds"))
                changed = True
        if changed:
            self.used.add(pid)
            self.save()
        return pid

    def save(self) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        wb = Workbook()
        ws = wb.active
        ws.title = "受保护的患者映射"
        ws.append(self.HEADERS)
        for key, (name, pid, created) in sorted(self.rows.items(), key=lambda x: x[1][1]):
            ws.append([name, pid, key, created])
        ws.freeze_panes = "A2"
        ws.column_dimensions["A"].width = 24
        ws.column_dimensions["B"].width = 22
        ws.column_dimensions["C"].width = 68
        ws.column_dimensions["D"].width = 22
        with tempfile.NamedTemporaryFile(suffix=".xlsx", dir=self.path.parent, delete=False) as f:
            tmp = Path(f.name)
        try:
            wb.save(tmp)
            os.replace(tmp, self.path)
        finally:
            tmp.unlink(missing_ok=True)


def redact_text(text: str, patient_name: str, patient_id: str) -> tuple[str, list[Finding]]:
    findings: list[Finding] = []

    def sub(pattern: re.Pattern, replacement: str, kind: str, value: str) -> str:
        nonlocal text
        count = len(pattern.findall(text))
        if count:
            text = pattern.sub(replacement, text)
            findings.append(Finding(kind, f"已处理 {count} 处"))
        return text

    if patient_name and not PID_RE.fullmatch(patient_name) and len(patient_name.strip()) >= 2:
        pattern = re.compile(re.escape(patient_name.strip()), re.I)
        sub(pattern, patient_id, "患者姓名", patient_name)
    sub(LABELED_NAME_RE, rf"\1：{patient_id}", "患者姓名字段", "")
    sub(LABELED_DATE_RE, r"\1：[日期已保护]", "日期字段", "")
    sub(ID_CARD_RE, "[身份证号已删除]", "身份证号", "")
    sub(PHONE_RE, "[联系电话已删除]", "手机号", "")
    sub(LANDLINE_RE, "[联系电话已删除]", "联系电话", "")
    sub(LABELED_ID_RE, lambda_placeholder("标识", patient_id), "原始患者标识", "")
    sub(ADDRESS_RE, r"\1：[地址已删除]", "地址", "")
    sub(BIRTH_RE, r"\1：[出生日期已删除]", "出生日期", "")
    return text, findings


def lambda_placeholder(label: str, patient_id: str) -> str:
    # Regex replacement intentionally preserves the label while standardizing value.
    return rf"\1：{patient_id}"


def scan_text(text: str, patient_name: str) -> list[str]:
    hits = []
    checks = [(ID_CARD_RE, "身份证号"), (PHONE_RE, "手机号"), (LABELED_ID_RE, "原始患者标识")]
    if patient_name and len(patient_name.strip()) >= 2 and not PID_RE.fullmatch(patient_name):
        checks.append((re.compile(re.escape(patient_name.strip()), re.I), "患者姓名"))
    for pattern, label in checks:
        if pattern.search(text):
            hits.append(label)
    return hits


def read_text(path: Path) -> tuple[str, str]:
    data = path.read_bytes()
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16"), "utf-16"
    if data.startswith(b"\xef\xbb\xbf"):
        return data.decode("utf-8-sig"), "utf-8-sig"
    if len(data) >= 4:
        even_nuls = data[0::2].count(0)
        odd_nuls = data[1::2].count(0)
        half = max(1, len(data) // 2)
        if odd_nuls / half > 0.25:
            return data.decode("utf-16-le"), "utf-16-le"
        if even_nuls / half > 0.25:
            return data.decode("utf-16-be"), "utf-16-be"
    for encoding in ("utf-8", "gb18030", "big5", "cp1252"):
        try:
            return data.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    from charset_normalizer import from_bytes
    match = from_bytes(data).best()
    if match is not None and match.encoding and match.percent_chaos <= 30:
        return str(match), match.encoding
    raise UnicodeError("无法可靠识别文本编码")


def process_plain(src: Path, dst: Path, name: str, pid: str) -> Result:
    text, encoding = read_text(src)
    cleaned, findings = redact_text(text, name, pid)
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_text(cleaned, encoding=encoding)
    residual = scan_text(cleaned, name)
    if residual:
        return Result("review", findings + [Finding("残留复扫", ",".join(residual))])
    return Result("success", findings)


def process_csv(src: Path, dst: Path, name: str, pid: str) -> Result:
    text, encoding = read_text(src)
    cleaned, findings = redact_text(text, name, pid)
    # Parse once to reject malformed CSV while preserving original dialect/content.
    list(csv.reader(cleaned.splitlines()))
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_text(cleaned, encoding=encoding)
    residual = scan_text(cleaned, name)
    return Result("review" if residual else "success", findings + ([Finding("残留复扫", ",".join(residual))] if residual else []))


def walk_docx_containers(doc) -> Iterable:
    yield doc
    for section in doc.sections:
        yield section.header
        yield section.footer
        yield section.first_page_header
        yield section.first_page_footer
        yield section.even_page_header
        yield section.even_page_footer


def redact_docx_container(container, name: str, pid: str) -> list[Finding]:
    findings: list[Finding] = []
    for paragraph in container.paragraphs:
        original = paragraph.text
        cleaned, found = redact_text(original, name, pid)
        if cleaned != original:
            # Runs are rebuilt; paragraph style remains, but mixed run formatting may simplify.
            for run in paragraph.runs:
                run.text = ""
            if paragraph.runs:
                paragraph.runs[0].text = cleaned
            else:
                paragraph.add_run(cleaned)
            findings.extend(found)
    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                findings.extend(redact_docx_container(cell, name, pid))
    return findings


def process_docx(src: Path, dst: Path, name: str, pid: str) -> Result:
    from docx import Document
    from docx.opc.constants import RELATIONSHIP_TYPE

    doc = Document(src)
    findings: list[Finding] = []
    for container in walk_docx_containers(doc):
        findings.extend(redact_docx_container(container, name, pid))
    props = doc.core_properties
    props.author = props.last_modified_by = props.comments = ""
    props.title, _ = redact_text(props.title or "", name, pid)
    props.subject, _ = redact_text(props.subject or "", name, pid)
    # python-docx 1.2 exposes comments separately from body/header containers.
    comments = getattr(doc, "comments", None)
    if comments is not None:
        for comment in comments:
            findings.extend(redact_docx_container(comment, name, pid))
            comment.author = ""
            comment.initials = ""
    # Hyperlinks can leak mail addresses/identifiers; flag, because python-docx cannot safely rewrite all fields.
    external_links = [r.target_ref for r in doc.part.rels.values() if r.reltype == RELATIONSHIP_TYPE.HYPERLINK and r.is_external]
    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(dst)
    verify = Document(dst)
    residual = []
    for container in walk_docx_containers(verify):
        for p in container.paragraphs:
            residual.extend(scan_text(p.text, name))
        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:
                    residual.extend(scan_text(cell.text, name))
    # Catch text boxes, custom XML, field codes, comments or other OOXML locations not
    # surfaced by python-docx. They are not modified blindly because that could corrupt
    # the package; any hit routes the document to manual review.
    with zipfile.ZipFile(dst) as package:
        for member in package.namelist():
            if member.endswith((".xml", ".rels")):
                raw = package.read(member).decode("utf-8", "ignore")
                residual.extend(scan_text(raw, name))
    if external_links:
        findings.append(Finding("外部超链接", "存在外部链接，需人工确认是否含身份信息"))
    if residual:
        findings.append(Finding("残留复扫", ",".join(sorted(set(residual)))))
    return Result("review" if external_links or residual else "success", findings)


def process_xlsx(src: Path, dst: Path, name: str, pid: str) -> Result:
    wb = load_workbook(src, data_only=False, keep_links=False)
    findings: list[Finding] = []
    review = False
    for ws in wb.worksheets:
        ws.sheet_properties.tabColor = None
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    if cell.data_type == "f":
                        # Formula strings can contain clinical logic/references; do not rewrite blindly.
                        if scan_text(cell.value, name):
                            review = True
                            findings.append(Finding("公式中的敏感信息", f"工作表 {ws.title} 存在待复核公式"))
                        continue
                    cell.value, found = redact_text(cell.value, name, pid)
                    findings.extend(found)
                if cell.comment:
                    cleaned, found = redact_text(cell.comment.text, name, pid)
                    cell.comment.text = cleaned
                    cell.comment.author = ""
                    findings.extend(found)
    wb.properties.creator = wb.properties.lastModifiedBy = wb.properties.description = ""
    dst.parent.mkdir(parents=True, exist_ok=True)
    wb.save(dst)
    check = load_workbook(dst, read_only=True, data_only=False, keep_links=False)
    residual = []
    for ws in check.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str):
                    residual.extend(scan_text(cell.value, name))
    check.close()
    if residual:
        review = True
        findings.append(Finding("残留复扫", ",".join(sorted(set(residual)))))
    return Result("review" if review else "success", findings)


def process_xls(src: Path, dst: Path, name: str, pid: str) -> Result:
    """Redact text cells in a legacy BIFF .xls while preserving cell formatting."""
    import xlrd
    from xlutils.copy import copy as copy_xls

    source = xlrd.open_workbook(src, formatting_info=True, on_demand=False)
    output = copy_xls(source)
    findings: list[Finding] = []
    for sheet_index in range(source.nsheets):
        read_sheet = source.sheet_by_index(sheet_index)
        write_sheet = output.get_sheet(sheet_index)
        for row_index in range(read_sheet.nrows):
            for col_index in range(read_sheet.ncols):
                cell = read_sheet.cell(row_index, col_index)
                if cell.ctype != xlrd.XL_CELL_TEXT or not isinstance(cell.value, str):
                    continue
                cleaned, found = redact_text(cell.value, name, pid)
                if cleaned != cell.value:
                    # xlutils has already copied the BIFF cell. Replacing only its
                    # shared-string index retains the original XF/style index.
                    row = write_sheet._Worksheet__rows[row_index]
                    out_cell = row._Row__cells[col_index]
                    out_cell.sst_idx = output.add_str(cleaned)
                    findings.extend(found)
    dst.parent.mkdir(parents=True, exist_ok=True)
    output.save(str(dst))
    source.release_resources()

    check = xlrd.open_workbook(dst, on_demand=False)
    residual: list[str] = []
    for sheet in check.sheets():
        for row_index in range(sheet.nrows):
            for col_index in range(sheet.ncols):
                value = sheet.cell_value(row_index, col_index)
                if isinstance(value, str):
                    residual.extend(scan_text(value, name))
    check.release_resources()
    if residual:
        findings.append(Finding("残留复扫", ",".join(sorted(set(residual)))))
    return Result("review" if residual else "success", findings)


def recursive_redact(value, name: str, pid: str, findings: list[Finding]):
    if isinstance(value, str):
        cleaned, found = redact_text(value, name, pid)
        findings.extend(found)
        return cleaned
    if isinstance(value, list):
        return [recursive_redact(v, name, pid, findings) for v in value]
    if isinstance(value, dict):
        result = {}
        for key, val in value.items():
            clean_key, found = redact_text(str(key), name, pid)
            findings.extend(found)
            result[clean_key] = recursive_redact(val, name, pid, findings)
        return result
    return value


def process_json(src: Path, dst: Path, name: str, pid: str) -> Result:
    text, encoding = read_text(src)
    data = json.loads(text)
    findings: list[Finding] = []
    data = recursive_redact(data, name, pid, findings)
    dst.parent.mkdir(parents=True, exist_ok=True)
    dst.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    residual = scan_text(json.dumps(data, ensure_ascii=False), name)
    return Result("review" if residual else "success", findings + ([Finding("残留复扫", ",".join(residual))] if residual else []))


def process_xml(src: Path, dst: Path, name: str, pid: str) -> Result:
    from defusedxml import ElementTree as ET

    tree = ET.parse(src)
    findings: list[Finding] = []
    for elem in tree.iter():
        if elem.text:
            elem.text, found = redact_text(elem.text, name, pid)
            findings.extend(found)
        if elem.tail:
            elem.tail, found = redact_text(elem.tail, name, pid)
            findings.extend(found)
        for key, value in list(elem.attrib.items()):
            elem.attrib[key], found = redact_text(value, name, pid)
            findings.extend(found)
    dst.parent.mkdir(parents=True, exist_ok=True)
    tree.write(dst, encoding="utf-8", xml_declaration=True)
    residual = scan_text(dst.read_text("utf-8"), name)
    return Result("review" if residual else "success", findings + ([Finding("残留复扫", ",".join(residual))] if residual else []))


def process_pdf(src: Path, dst: Path, name: str, pid: str) -> Result:
    import fitz

    doc = fitz.open(src)
    findings: list[Finding] = []
    review = False
    patterns = [ID_CARD_RE, PHONE_RE, LABELED_ID_RE, ADDRESS_RE, BIRTH_RE]
    if name and len(name.strip()) >= 2:
        patterns.insert(0, re.compile(re.escape(name.strip()), re.I))
    for page in doc:
        text = page.get_text("text")
        for pattern in patterns:
            for match in pattern.finditer(text):
                # Search exact text as rendered; multi-line/fragmented matches may fail and are caught by rescan.
                for rect in page.search_for(match.group(0)):
                    page.add_redact_annot(rect, fill=(1, 1, 1), text=pid if pattern.pattern == re.escape(name.strip()) else "[已删除]")
                    findings.append(Finding("PDF文本", "已遮盖敏感文本"))
        page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
        if page.get_images(full=True) or not text.strip():
            review = True
            findings.append(Finding("PDF图像/扫描页", f"第 {page.number + 1} 页需检查烧录文字或OCR遗漏"))
    metadata = {k: "" for k in doc.metadata}
    doc.set_metadata(metadata)
    dst.parent.mkdir(parents=True, exist_ok=True)
    doc.save(dst, garbage=4, deflate=True, clean=True)
    doc.close()
    check = fitz.open(dst)
    residual = []
    for page in check:
        residual.extend(scan_text(page.get_text("text"), name))
    check.close()
    if residual:
        review = True
        findings.append(Finding("残留复扫", ",".join(sorted(set(residual)))))
    return Result("review" if review else "success", findings)


RECORD_PROCESSORS: dict[str, Callable] = {
    ".txt": process_plain, ".md": process_plain, ".log": process_plain, ".ini": process_plain,
    ".yaml": process_plain, ".yml": process_plain, ".html": process_plain, ".htm": process_plain,
    ".csv": process_csv, ".docx": process_docx, ".xls": process_xls,
    ".xlsx": process_xlsx, ".json": process_json, ".xml": process_xml,
    ".pdf": process_pdf,
}


def safe_filename(name: str, patient_name: str, pid: str) -> tuple[str, bool]:
    stem, suffix = Path(name).stem, Path(name).suffix
    cleaned, _ = redact_text(stem, patient_name, pid)
    cleaned = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", cleaned).strip(" .") or "file"
    residual = bool(scan_text(cleaned, patient_name))
    return cleaned + suffix.lower(), residual


class UIDMapper:
    def __init__(self, secret_key: str):
        self.secret_key = secret_key
        self.cache: dict[str, str] = {}

    def map(self, old: str) -> str:
        from pydicom.uid import generate_uid
        if old not in self.cache:
            digest = hashlib.sha256((self.secret_key + "|" + old).encode()).hexdigest()
            self.cache[old] = generate_uid(entropy_srcs=[digest])
        return self.cache[old]


DICOM_CLEAR_KEYWORDS = {
    "PatientBirthDate", "PatientBirthTime", "PatientAddress", "PatientTelephoneNumbers",
    "OtherPatientIDs", "OtherPatientNames", "PatientMotherBirthName", "MedicalRecordLocator",
    "EthnicGroup", "Occupation", "AdditionalPatientHistory", "InstitutionAddress",
    "ReferringPhysicianAddress", "ReferringPhysicianTelephoneNumbers", "PersonAddress",
    "PersonTelephoneNumbers", "PatientComments", "InsurancePlanIdentification",
    "MilitaryRank", "BranchOfService", "CountryOfResidence", "RegionOfResidence",
}
DICOM_REPLACE_KEYWORDS = {
    "PatientName", "PatientID", "OtherPatientIDsSequence", "AdmissionID",
    "IssuerOfPatientID", "AccessionNumber", "StudyID", "RequestingPhysician",
    "ReferringPhysicianName", "PerformingPhysicianName", "OperatorsName",
    "PhysiciansOfRecord", "NameOfPhysiciansReadingStudy",
}
DATE_KEYWORDS = {
    "StudyDate", "SeriesDate", "AcquisitionDate", "ContentDate", "InstanceCreationDate",
    "PerformedProcedureStepStartDate", "PerformedProcedureStepEndDate", "ScheduledProcedureStepStartDate",
}
UID_KEYWORDS = {
    "StudyInstanceUID", "SeriesInstanceUID", "SOPInstanceUID", "MediaStorageSOPInstanceUID",
    "FrameOfReferenceUID", "SynchronizationFrameOfReferenceUID", "ReferencedSOPInstanceUID",
    "DimensionOrganizationUID", "ConcatenationUID", "IrradiationEventUID",
}
UID_EXEMPT_KEYWORDS = {
    "SOPClassUID", "ReferencedSOPClassUID", "MediaStorageSOPClassUID", "TransferSyntaxUID",
    "ReferencedSOPClassUIDInFile", "ReferencedTransferSyntaxUIDInFile",
    "ImplementationClassUID", "ImplementationVersionName", "CodingSchemeUID",
}


def shift_da(value: str, days: int) -> str:
    value = str(value)
    if len(value) < 8 or not value[:8].isdigit():
        return value
    try:
        shifted = datetime.strptime(value[:8], "%Y%m%d") + timedelta(days=days)
        return shifted.strftime("%Y%m%d") + value[8:]
    except ValueError:
        return value


def map_dicom_values(value, transform: Callable[[str], str]):
    """Apply a transform without collapsing pydicom MultiValue elements."""
    from pydicom.multival import MultiValue

    if isinstance(value, MultiValue):
        return [transform(str(item)) for item in value]
    return transform(str(value))


def _process_dicom_impl(src: Path, dst: Path, name: str, pid: str, uid_mapper: UIDMapper, day_offset: int) -> Result:
    import pydicom
    from pydicom.dataset import FileMetaDataset
    from pydicom.uid import (ExplicitVRBigEndian, ExplicitVRLittleEndian,
                             ImplicitVRLittleEndian, PYDICOM_IMPLEMENTATION_UID)
    ds = pydicom.dcmread(src, force=False)
    findings: list[Finding] = []
    review = False
    private_tag_count = 0

    def visit(dataset) -> None:
        nonlocal review, private_tag_count
        declared_charset = dataset.get("SpecificCharacterSet", None)
        if declared_charset:
            declared_value = getattr(declared_charset, "value", declared_charset)
            declared_text = str(declared_value).upper()
            if "GB18030" in declared_text or "GBK" in declared_text:
                # GB18030/GBK are standalone DICOM character repertoires and
                # cannot legally be combined with ISO 2022 code extensions.
                dataset.SpecificCharacterSet = "GB18030"
                implicit_vr, little_endian = dataset.original_encoding
                dataset.set_original_encoding(implicit_vr, little_endian, "GB18030")
        for elem in list(dataset):
            if elem.tag.is_private:
                del dataset[elem.tag]
                private_tag_count += 1
                continue
            if elem.VR == "SQ":
                for item in elem.value:
                    visit(item)
                continue
            keyword = elem.keyword
            if keyword in DICOM_CLEAR_KEYWORDS:
                elem.value = ""
            elif keyword in DICOM_REPLACE_KEYWORDS:
                elem.value = pid
            elif (keyword in DATE_KEYWORDS or elem.VR == "DA") and elem.value:
                elem.value = map_dicom_values(elem.value, lambda value: shift_da(value, day_offset))
            elif (keyword in UID_KEYWORDS or elem.VR == "UI") and keyword not in UID_EXEMPT_KEYWORDS and elem.value:
                elem.value = map_dicom_values(elem.value, uid_mapper.map)
            elif elem.VR in {"PN"} and elem.value:
                elem.value = pid
            elif elem.VR in {"LO", "LT", "SH", "ST", "UT"} and elem.value:
                def clean_text_value(value: str) -> str:
                    cleaned, found = redact_text(value, name, pid)
                    findings.extend(found)
                    return cleaned

                elem.value = map_dicom_values(elem.value, clean_text_value)
        if "SpecificCharacterSet" in dataset:
            dataset.SpecificCharacterSet = "ISO_IR 192"

    visit(ds)
    if private_tag_count:
        findings.append(Finding("DICOM私有标签", f"已删除 {private_tag_count} 个"))
    ds.PatientName = pid
    ds.PatientID = pid
    ds.PatientIdentityRemoved = "YES"
    ds.DeidentificationMethod = "Basic Profile; private tags removed; date shift"
    # Rebuild file meta from validated primitive values. Some vendor exports are
    # readable but contain malformed file-meta value types that fail on dcmwrite.
    old_meta = getattr(ds, "file_meta", None) or FileMetaDataset()
    transfer_syntax = getattr(old_meta, "TransferSyntaxUID", None)
    if not transfer_syntax:
        raw_encoding = getattr(ds, "original_encoding", (True, True))
        if not isinstance(raw_encoding, tuple) or len(raw_encoding) != 2:
            raw_encoding = (True, True)
        raw_implicit, raw_little = raw_encoding
        # A few vendor exports expose non-boolean proxy values here. Never ask
        # such values for truthiness because that can raise TypeError.
        implicit = raw_implicit if type(raw_implicit) is bool else True
        little = raw_little if type(raw_little) is bool else True
        if implicit and little:
            transfer_syntax = ImplicitVRLittleEndian
        elif not implicit and not little:
            transfer_syntax = ExplicitVRBigEndian
        else:
            transfer_syntax = ExplicitVRLittleEndian
    sop_class = getattr(ds, "SOPClassUID", None) or getattr(old_meta, "MediaStorageSOPClassUID", None)
    sop_instance = getattr(ds, "SOPInstanceUID", None)
    if sop_instance:
        sop_instance = str(sop_instance)
    else:
        old_instance = getattr(old_meta, "MediaStorageSOPInstanceUID", None)
        sop_instance = uid_mapper.map(str(old_instance)) if old_instance else None
    clean_meta = FileMetaDataset()
    clean_meta.TransferSyntaxUID = str(transfer_syntax)
    if sop_class:
        clean_meta.MediaStorageSOPClassUID = str(sop_class)
    if sop_instance:
        clean_meta.MediaStorageSOPInstanceUID = sop_instance
    clean_meta.ImplementationClassUID = PYDICOM_IMPLEMENTATION_UID
    clean_meta.ImplementationVersionName = "MEDDEID_1"
    ds.file_meta = clean_meta
    burned = str(getattr(ds, "BurnedInAnnotation", "")).upper()
    if "PixelData" in ds and burned != "NO":
        findings.append(Finding("像素数据", "按业务规则保持不变，未检查或修改烧录文字"))
    if src.name.upper() == "DICOMDIR" or getattr(ds, "DirectoryRecordSequence", None):
        findings.append(Finding("DICOMDIR", "已清理目录记录中的身份标签并完成写后读取验证"))
    dst.parent.mkdir(parents=True, exist_ok=True)
    try:
        pydicom.dcmwrite(dst, ds, enforce_file_format=True)
    except TypeError:
        # Some vendor media contain readable but non-conformant file-meta encoding.
        # Retry while preserving the source encoding; the mandatory read-back and
        # pixel/identity verification below still decides whether output is valid.
        dst.unlink(missing_ok=True)
        pydicom.dcmwrite(dst, ds, enforce_file_format=False)
        findings.append(Finding("DICOM兼容写入", "标准文件头重建失败，已保留源编码并通过写后验证"))
    check = pydicom.dcmread(dst, force=False)
    if str(check.PatientID) != pid or str(check.PatientName) != pid:
        raise ValueError("DICOM 写后验证失败")
    original_pixels = getattr(ds, "PixelData", None)
    checked_pixels = getattr(check, "PixelData", None)
    if original_pixels is not None:
        if checked_pixels is None or bytes(checked_pixels) != bytes(original_pixels):
            raise ValueError("DICOM 像素数据发生意外变化")
    elif checked_pixels is not None:
        raise ValueError("DICOM 空像素标签在写入后发生意外变化")
    return Result("review" if review else "success", findings)


def process_dicom(src: Path, dst: Path, name: str, pid: str, uid_mapper: UIDMapper, day_offset: int) -> Result:
    """Process DICOM while aggregating known source-conformance warnings."""
    known_fragments = (
        "Specific Character Set does not allow code extensions",
        "Failed to decode byte string with encoding",
        "exceeds the maximum length of 64 allowed for VR LO",
    )
    logged_messages: list[str] = []

    class DicomWarningFilter(logging.Filter):
        def filter(self, record: logging.LogRecord) -> bool:
            message = record.getMessage()
            if any(fragment in message for fragment in known_fragments):
                logged_messages.append(message)
                return False
            return True

    warning_filter = DicomWarningFilter()
    pydicom_logger = logging.getLogger("pydicom")
    handlers = list(dict.fromkeys([*logging.getLogger().handlers, *pydicom_logger.handlers]))
    for handler in handlers:
        handler.addFilter(warning_filter)
    caught = []
    try:
        with warnings.catch_warnings(record=True) as caught:
            warnings.simplefilter("always")
            result = _process_dicom_impl(src, dst, name, pid, uid_mapper, day_offset)
    finally:
        for handler in handlers:
            handler.removeFilter(warning_filter)

    caught_known_messages = []
    for warning in caught:
        message = str(warning.message)
        if any(fragment in message for fragment in known_fragments):
            caught_known_messages.append(message)
        else:
            warnings.warn_explicit(message, warning.category, warning.filename, warning.lineno)
    # pydicom emits the same event through both logging and warnings; prefer the
    # warning-channel list to avoid double-counting.
    known_warning_messages = caught_known_messages or logged_messages
    if known_warning_messages:
        charset_count = sum("Character Set" in msg or "decode byte string" in msg
                            for msg in known_warning_messages)
        lo_count = sum("VR LO" in msg for msg in known_warning_messages)
        details = []
        if charset_count:
            details.append(f"字符集不规范/无法完整解码 {charset_count} 次，输出已统一为 UTF-8")
        if lo_count:
            details.append(f"源 LO 字段超出标准长度 {lo_count} 次，原值未截断")
        result.findings.append(Finding("DICOM源数据规范警告", "；".join(details)))
    return result


def looks_like_dicom(path: Path) -> bool:
    if path.name.upper() == "DICOMDIR" or path.suffix.lower() in {".dcm", ".dicom"}:
        return True
    try:
        with path.open("rb") as f:
            f.seek(128)
            return f.read(4) == b"DICM"
    except OSError:
        return False


def date_offset_for(key: str) -> int:
    # Stable per source folder, in roughly +/- five-year range, never zero.
    n = int(hashlib.sha256(("date|" + key).encode()).hexdigest()[:8], 16)
    days = n % 3651 - 1825
    return days or 731


def write_report(path: Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fields = ["patient_id", "source_file_hash", "output_relative_path", "status", "findings"]
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


def read_report(path: Path) -> list[dict]:
    if not path.exists():
        return []
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            return list(csv.DictReader(f))
    except (OSError, csv.Error):
        return []


def copy_to_review(temp_dst: Path, review_dst: Path) -> None:
    review_dst.parent.mkdir(parents=True, exist_ok=True)
    if temp_dst.exists():
        os.replace(temp_dst, review_dst)


def safe_exception_detail(exc: Exception) -> str:
    """Diagnostic location without exception values, paths, or patient data."""
    frames = traceback.extract_tb(exc.__traceback__)
    if not frames:
        return type(exc).__name__
    frame = frames[-1]
    # pydicom wraps serialization failures with a safe numeric tag such as
    # ``(0010,0010)``. Include only that tag, never the element value.
    tag_match = re.search(r"\([0-9A-Fa-f]{4},[0-9A-Fa-f]{4}\)", str(exc))
    tag = f" tag={tag_match.group(0)}" if tag_match else ""
    return f"{type(exc).__name__} at {Path(frame.filename).name}:{frame.lineno} in {frame.name}{tag}"


def process_patient(folder: Path, input_root: Path, output_root: Path, mode: str, name: str, pid: str,
                    key: str, uid_mapper: UIDMapper, log: logging.Logger) -> list[dict]:
    rows: list[dict] = []
    share_root = output_root / "share_ready" / pid
    review_root = output_root / "manual_review" / pid
    dicom_failed = False
    for src in sorted(folder.rglob("*")):
        if not src.is_file():
            continue
        rel = src.relative_to(folder)
        parts = list(rel.parts)
        if mode == "dicom":
            # Non-target imaging media resources must remain byte-for-byte and
            # name-for-name unchanged. Only the patient root is pseudonymized.
            filename, filename_risk = parts[-1], False
            clean_rel = rel
        else:
            filename, filename_risk = safe_filename(parts[-1], name, pid)
            clean_rel = Path(*parts[:-1], filename)
        share_dst = share_root / clean_rel
        review_dst = review_root / clean_rel
        tmp_dst = output_root / ".working" / pid / clean_rel
        source_hash = safe_log_value(str(rel))
        try:
            if mode == "records":
                processor = RECORD_PROCESSORS.get(src.suffix.lower())
                if src.suffix.lower() in IMAGE_EXTS:
                    tmp_dst.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(src, tmp_dst)
                    result = Result("success", [Finding("安全图像", "按业务规则原样复制，未处理图像内容")])
                elif processor is None or src.suffix.lower() in UNSAFE_EXTS:
                    review_dst.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(src, review_dst)
                    result = Result("review", [Finding("不支持或含图像内容", "未修改原件副本，需人工脱敏")])
                else:
                    result = processor(src, tmp_dst, name, pid)
            else:
                upper_parts = {part.upper() for part in rel.parts[:-1]}
                is_label = src.name.upper() == "LABEL.TXT"
                is_dicomdir = src.name.upper() == "DICOMDIR"
                in_dicomobj = "DICOMOBJ" in upper_parts
                if is_label:
                    result = process_plain(src, tmp_dst, name, pid)
                elif is_dicomdir or in_dicomobj:
                    if not looks_like_dicom(src):
                        raise ValueError("DICOMOBJ/DICOMDIR 中存在不可读取的 DICOM 文件")
                    result = process_dicom(src, tmp_dst, name, pid, uid_mapper, date_offset_for(key))
                else:
                    # INDEX.files, PD-SViewer and all other viewer/media assets
                    # are explicitly outside the privacy-editing scope.
                    tmp_dst.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(src, tmp_dst)
                    result = Result("success", [Finding("非隐私目标文件", "按业务规则逐字节原样复制")])
            if filename_risk:
                result.status = "review"
                result.findings.append(Finding("文件名残留风险", "需人工复核"))
            if mode == "dicom" and result.status != "success":
                dicom_failed = True
            if mode == "records" and result.status != "success":
                log.warning("病历文件未通过 patient=%s file=%s status=%s",
                            pid, filename, result.status)
            if result.status == "success":
                share_dst.parent.mkdir(parents=True, exist_ok=True)
                os.replace(tmp_dst, share_dst)
                out_rel = share_dst.relative_to(output_root)
            else:
                if tmp_dst.exists():
                    copy_to_review(tmp_dst, review_dst)
                out_rel = review_dst.relative_to(output_root)
            rows.append({"patient_id": pid, "source_file_hash": source_hash,
                         "output_relative_path": out_rel.as_posix(), "status": result.status,
                         "findings": "; ".join(f"{x.kind}:{x.detail}" for x in result.findings)})
        except Exception as exc:
            if mode == "dicom":
                dicom_failed = True
            diagnostic = safe_exception_detail(exc)
            review_dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, review_dst)
            rows.append({"patient_id": pid, "source_file_hash": source_hash,
                         "output_relative_path": review_dst.relative_to(output_root).as_posix(),
                         "status": "error", "findings": f"处理异常:{diagnostic}"})
            if mode == "records":
                log.error("病历文件处理失败 patient=%s file=%s error=%s", pid, filename, diagnostic)
            else:
                log.error("文件处理失败 patient=%s file_hash=%s error=%s", pid, source_hash, diagnostic)
        finally:
            tmp_dst.unlink(missing_ok=True)
    if mode == "dicom" and dicom_failed and share_root.exists():
        # Patient-level transaction rule: every item under the image folder must
        # succeed; one failed/review item prevents the whole set from sharing.
        for staged in sorted(share_root.rglob("*")):
            if staged.is_file():
                target = review_root / staged.relative_to(share_root)
                target.parent.mkdir(parents=True, exist_ok=True)
                os.replace(staged, target)
        shutil.rmtree(share_root)
        share_prefix = (Path("share_ready") / pid).as_posix() + "/"
        review_prefix = (Path("manual_review") / pid).as_posix() + "/"
        for row in rows:
            if row["output_relative_path"].startswith(share_prefix):
                row["output_relative_path"] = review_prefix + row["output_relative_path"][len(share_prefix):]
                if row["status"] == "success":
                    row["status"] = "review"
                row["findings"] += "; 患者级影像事务:目录内存在未成功文件，整组不得进入share_ready"
        log.error("患者影像整组未通过 patient=%s；目录内至少一个文件未成功", pid)
    return rows


def validate_paths(input_root: Path, output_root: Path, mapping: Path) -> None:
    input_root = input_root.resolve()
    output_root = output_root.resolve()
    mapping = mapping.resolve()
    if not input_root.is_dir():
        raise ValueError("输入根目录不存在或不是目录")
    if output_root == input_root or input_root in output_root.parents:
        raise ValueError("输出目录不能等于输入目录，也不能位于输入目录内")
    if mapping == output_root or output_root in mapping.parents:
        raise ValueError("映射表必须保存在输出/科研共享目录之外的受保护位置")
    if mapping == input_root or input_root in mapping.parents:
        raise ValueError("映射表不能保存在原始患者数据目录内")


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description="医疗病历或 DICOM 数据离线批量隐私脱敏")
    p.add_argument("--mode", required=True, choices=["records", "dicom", "both"],
                   help="records=病历；dicom=影像；both=配对处理病历和影像")
    p.add_argument("--input", type=Path, help="records/dicom 单模式的患者输入根目录")
    p.add_argument("--records-input", type=Path, help="both 模式的病历患者根目录")
    p.add_argument("--dicom-input", type=Path, help="both 模式的影像患者根目录")
    p.add_argument("--output", required=True, type=Path, help="独立输出根目录（不可位于输入目录内）")
    p.add_argument("--mapping", required=True, type=Path, help="受保护位置的 Patient_ID_Mapping.xlsx")
    p.add_argument("--overwrite-patient", action="store_true", help="允许安全替换已存在的同一患者输出目录")
    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        if args.mode == "both":
            if args.input or not args.records_input or not args.dicom_input:
                raise ValueError("both 模式必须同时提供 --records-input 和 --dicom-input，且不使用 --input")
            validate_paths(args.records_input, args.output, args.mapping)
            validate_paths(args.dicom_input, args.output, args.mapping)
            record_dirs = {p.name: p for p in args.records_input.iterdir() if p.is_dir()}
            dicom_dirs = {p.name: p for p in args.dicom_input.iterdir() if p.is_dir()}
            if not record_dirs and not dicom_dirs:
                raise ValueError("病历和影像输入根目录下没有患者文件夹")
        else:
            if not args.input or args.records_input or args.dicom_input:
                raise ValueError("records/dicom 单模式只使用 --input")
            validate_paths(args.input, args.output, args.mapping)
        args.output.mkdir(parents=True, exist_ok=True)
        logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s", force=True,
                            handlers=[logging.FileHandler(args.output / "deidentification.log", encoding="utf-8"),
                                      logging.StreamHandler()])
        log = logging.getLogger("deidentifier")
        store = MappingStore(args.mapping)
        report_path = args.output / "deidentification_report.csv"
        previous_rows = read_report(report_path)
        report_rows: list[dict] = []
        processed_scopes: set[tuple[str, str]] = set()
        uid_mapper = UIDMapper(hashlib.sha256(str(args.mapping.resolve()).encode()).hexdigest())
        if args.mode == "both":
            jobs = []
            for name in sorted(set(record_dirs) | set(dicom_dirs)):
                key = paired_source_key(args.records_input, args.dicom_input, name)
                component_keys = []
                patient_jobs = []
                if name in record_dirs:
                    component_keys.append(source_key(args.records_input, record_dirs[name]))
                    patient_jobs.append((record_dirs[name], args.records_input, "records"))
                if name in dicom_dirs:
                    component_keys.append(source_key(args.dicom_input, dicom_dirs[name]))
                    patient_jobs.append((dicom_dirs[name], args.dicom_input, "dicom"))
                jobs.append((name, key, component_keys, patient_jobs))
        else:
            patient_dirs = sorted(p for p in args.input.iterdir() if p.is_dir())
            if not patient_dirs:
                raise ValueError("输入根目录下没有患者文件夹")
            jobs = [(folder.name, source_key(args.input, folder), [], [(folder, args.input, args.mode)])
                    for folder in patient_dirs]

        for name, key, component_keys, patient_jobs in jobs:
            pid = (store.get_or_create_paired(key, component_keys, name) if args.mode == "both"
                   else store.get_or_create(key, name))
            if not PID_RE.fullmatch(pid):
                raise ValueError("映射表中存在非标准 Patient ID")
            if args.mode == "both":
                available_modes = {job[2] for job in patient_jobs}
                if "records" not in available_modes:
                    log.warning("患者配对不完整 patient=%s missing=records；继续处理现有影像", pid)
                if "dicom" not in available_modes:
                    log.warning("患者配对不完整 patient=%s missing=dicom；继续处理现有病历", pid)
            pending = []
            for folder, input_root, job_mode in patient_jobs:
                base_output = args.output / job_mode if args.mode == "both" else args.output
                existing = ((base_output / "share_ready" / pid).exists() or
                            (base_output / "manual_review" / pid).exists())
                if existing and args.overwrite_patient:
                    shutil.rmtree(base_output / "share_ready" / pid, ignore_errors=True)
                    shutil.rmtree(base_output / "manual_review" / pid, ignore_errors=True)
                    existing = False
                if not existing:
                    pending.append((folder, input_root, job_mode, base_output))
            if not pending:
                log.info("跳过已有患者输出 patient=%s；如需重做请使用 --overwrite-patient", pid)
                continue
            for folder, input_root, job_mode, base_output in pending:
                processed_scopes.add((pid, job_mode if args.mode == "both" else ""))
                log.info("开始处理 patient=%s source_folder_hash=%s mode=%s",
                         pid, safe_log_value(folder.name), job_mode)
                new_rows = process_patient(folder, input_root, base_output, job_mode, name, pid, key, uid_mapper, log)
                if args.mode == "both":
                    for row in new_rows:
                        row["output_relative_path"] = job_mode + "/" + row["output_relative_path"]
                report_rows.extend(new_rows)
        def was_reprocessed(row: dict) -> bool:
            pid = row.get("patient_id", "")
            if (pid, "") in processed_scopes:
                return True
            rel = row.get("output_relative_path", "")
            return any(pid == scope_pid and rel.startswith(scope + "/")
                       for scope_pid, scope in processed_scopes if scope)

        retained_rows = [r for r in previous_rows if not was_reprocessed(r)]
        write_report(report_path, retained_rows + report_rows)
        errors = sum(r["status"] == "error" for r in report_rows)
        reviews = sum(r["status"] == "review" for r in report_rows)
        successes = sum(r["status"] == "success" for r in report_rows)
        log.info("完成 success=%d manual_review=%d error=%d；自动扫描通过不代表完全匿名化", successes, reviews, errors)
        return 2 if errors else (1 if reviews else 0)
    except Exception as exc:
        print(f"错误: {exc}", file=sys.stderr)
        return 2
    finally:
        # Important when main() is embedded by a scheduler/test runner on Windows:
        # release the log file so the output directory can be moved or archived.
        root_logger = logging.getLogger()
        for handler in root_logger.handlers[:]:
            root_logger.removeHandler(handler)
            handler.close()


if __name__ == "__main__":
    raise SystemExit(main())
