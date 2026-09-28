import os
import re
import shutil
import tempfile
import unicodedata
import warnings
from collections import Counter
from copy import copy
from datetime import datetime
from email import policy
from email.parser import BytesParser
from html.parser import HTMLParser

try:
	from openpyxl import load_workbook
except ImportError:
	load_workbook = None

try:
	from PyPDF2 import PdfReader
except ImportError:
	PdfReader = None

SERVICE_NAME = "RADIAN"
SERVICE_NOT_APPLICABLE_REQUIREMENTS = {"2.10", "2.11", "2.12", "4.3", "4.5"}
NOT_APPLICABLE_OBSERVATION = "No aplica para el servicio implementado"
SUPPORTED_EXTENSIONS = {".eml", ".pdf", ".xlsx", ".xls", ".doc", ".docx", ".zip"}
ANALYST_HINTS = ("analista", "prueba", "software", "junior", "qa", "test")
ANALYST_DOMAINS = ("dispapeles.com",)

CHECKS = [
	(11, "aprobación del cliente", ("aprobado", "aprobación", "visto bueno", "aceptamos", "pruebas")),
	(12, "credenciales de producción", ("credenciales", "producción", "portal", "usuario", "contraseña")),
	(13, "archivo ETF", ("etf",)),
	(15, "datos generales de la empresa", ("dirección", "telefono", "teléfono", "correo electrónico", "barrio", "responsabilidad tributaria")),
	(16, "servicio solicitado y aviso de creación", ("facturación", "facturacion", "servicio", "electrónica", "electronica")),
	(17, "usuario administrador", ("usuario administrador", "administrador de la empresa", "admin")),
	(18, "asociación de Dispapeles como proveedor tecnológico", ("proveedor tecnológico", "proveedor tecnologico", "dispapeles", "asociación", "asociacion")),
	(19, "asignación de documentos por usuario", ("asignar documentos", "documentos por usuario")),
	(20, "asignación de documentos por empresa", ("asignar documentos", "documentos por empresa")),
	(21, "representación gráfica", ("representación gráfica", "representacion grafica", "portal manual")),
	(22, "credenciales Web Service", ("web service", "webservice", "ws", "usuario", "contraseña")),
	(23, "servicios y eventos del ETF", ("facturación", "nómina", "nomina", "documento soporte", "eventos dian", "radian", "etf")),
	(25, "cargue de plantillas", ("plantilla", "plantillas", "cargue", "carga")),
	(26, "visto bueno del Analista de Procesos", ("visto bueno", "salida en vivo", "analista de procesos", "pmo")),
	(27, "envío de credenciales de producción", ("credenciales", "producción", "portal")),
]

REQUIREMENT_CATALOG = [
	{"id": "1.1", "section": "COMERCIAL", "folder": "COMERCIAL", "name": "MV-ER-040 Informe de Adjudicación", "kind": "document", "code_candidates": ["MV-ER-040"], "allow_na": True, "location": "{cliente}/COMERCIAL"},
	{"id": "1.2", "section": "COMERCIAL", "folder": "COMERCIAL", "name": "Email Asignación", "kind": "email", "code_candidates": [], "allow_na": False, "location": "{cliente}/COMERCIAL"},
	{"id": "2.1", "section": "PRELIMINARES", "folder": "IMPLEMENTACIONES/PRELIMINARES", "name": "SI-ER-020 Cálculo Tamaño Proyecto", "kind": "document", "code_candidates": ["SI-ER-020"], "allow_na": True, "location": "{cliente}/IMPLEMENTACIONES/PRELIMINARES"},
	{"id": "2.2", "section": "PRELIMINARES", "folder": "IMPLEMENTACIONES/PRELIMINARES", "name": "PMO-MD-002 Project Charter", "kind": "document", "code_candidates": ["PMO-MD-002"], "allow_na": True, "location": "{cliente}/IMPLEMENTACIONES/PRELIMINARES"},
	{"id": "2.3", "section": "PRELIMINARES", "folder": "IMPLEMENTACIONES/PRELIMINARES", "name": "PMO-ER-018 ETF", "kind": "document", "code_candidates": ["PMO-ER-018"], "allow_na": True, "location": "{cliente}/IMPLEMENTACIONES/PRELIMINARES"},
	{"id": "2.4", "section": "PRELIMINARES", "folder": "IMPLEMENTACIONES/PRELIMINARES", "name": "PMO-ER-012 Contro Cambios ETF", "kind": "document", "code_candidates": ["PMO-ER-012"], "allow_na": True, "location": "{cliente}/IMPLEMENTACIONES/PRELIMINARES"},
	{"id": "2.5", "section": "PRELIMINARES", "folder": "IMPLEMENTACIONES/PRELIMINARES", "name": "PMO-ER-025 Cronograma", "kind": "document", "code_candidates": ["PMO-ER-025"], "allow_na": True, "location": "{cliente}/IMPLEMENTACIONES/PRELIMINARES"},
	{"id": "2.6", "section": "PRELIMINARES", "folder": "IMPLEMENTACIONES/PRELIMINARES", "name": "CORREO APROBACIÓN", "kind": "email", "code_candidates": [], "allow_na": False, "location": "{cliente}/IMPLEMENTACIONES/PRELIMINARES"},
	{"id": "2.7", "section": "PRELIMINARES", "folder": "IMPLEMENTACIONES/PRELIMINARES", "name": "DIVULGACIÓN CRONOGRAMA", "kind": "email", "code_candidates": [], "allow_na": False, "location": "{cliente}/IMPLEMENTACIONES/PRELIMINARES"},
	{"id": "2.8", "section": "PRELIMINARES", "folder": "IMPLEMENTACIONES/PRELIMINARES", "name": "PMO-ER-009 Actas Reunión", "kind": "document", "code_candidates": ["PMO-ER-009"], "allow_na": True, "conditional": "SI-ER-020 indica proyecto Grande", "location": "{cliente}/IMPLEMENTACIONES/PRELIMINARES"},
	{"id": "2.9", "section": "PRELIMINARES", "folder": "IMPLEMENTACIONES/PRELIMINARES", "name": "CAPACITACIÓN", "kind": "email", "code_candidates": [], "allow_na": False, "location": "{cliente}/IMPLEMENTACIONES/PRELIMINARES"},
	{"id": "2.10", "section": "PRELIMINARES", "folder": "IMPLEMENTACIONES/PRELIMINARES", "name": "ARTES", "kind": "document", "code_candidates": [], "allow_na": True, "location": "{cliente}/IMPLEMENTACIONES/PRELIMINARES"},
	{"id": "2.11", "section": "PRELIMINARES", "folder": "IMPLEMENTACIONES/PRELIMINARES", "name": "APROBACIÓN ARTES", "kind": "email", "code_candidates": [], "allow_na": True, "location": "{cliente}/IMPLEMENTACIONES/PRELIMINARES"},
	{"id": "2.12", "section": "PRELIMINARES", "folder": "IMPLEMENTACIONES/PRELIMINARES", "name": "MAPEO DE DATOS", "kind": "email", "code_candidates": [], "allow_na": True, "location": "{cliente}/IMPLEMENTACIONES/PRELIMINARES"},
	{"id": "3.1", "section": "SALIDA EN VIVO", "folder": "IMPLEMENTACIONES/SALIDA EN VIVO", "name": "PMO-ER-010", "kind": "email", "code_candidates": ["PMO-ER-010"], "allow_na": False, "location": "{cliente}/IMPLEMENTACIONES/SALIDA EN VIVO"},
	{"id": "3.2", "section": "SALIDA EN VIVO", "folder": "IMPLEMENTACIONES/SALIDA EN VIVO", "name": "APROBACIÓN PRUEBAS CLIENTE", "kind": "email", "code_candidates": [], "allow_na": False, "location": "{cliente}/IMPLEMENTACIONES/SALIDA EN VIVO"},
	{"id": "4.1", "section": "SALIDA EN VIVO", "folder": "IMPLEMENTACIONES/SALIDA EN VIVO", "name": "PMO-ER-011 CyP.", "kind": "document", "code_candidates": ["PMO-ER-011"], "allow_na": True, "location": "{cliente}/IMPLEMENTACIONES/SALIDA EN VIVO"},
	{"id": "4.2", "section": "SALIDA EN VIVO", "folder": "IMPLEMENTACIONES/SALIDA EN VIVO", "name": "ACTA SALIDA PRODUCCIÓN", "kind": "document", "code_candidates": ["PMO-MD-001", "PMO-MD-005", "PMO-MD-006", "PMO-MD-007", "PMO-MD-009", "PMO-MD-010", "PMO-MD-011"], "allow_na": True, "location": "{cliente}/IMPLEMENTACIONES/SALIDA EN VIVO"},
	{"id": "4.3", "section": "SALIDA EN VIVO", "folder": "IMPLEMENTACIONES/SALIDA EN VIVO", "name": "Resolución", "kind": "document", "code_candidates": [], "allow_na": False, "location": "{cliente}/IMPLEMENTACIONES/SALIDA EN VIVO"},
	{"id": "4.4", "section": "SALIDA EN VIVO", "folder": "IMPLEMENTACIONES/SALIDA EN VIVO", "name": "RUT", "kind": "document", "code_candidates": [], "allow_na": False, "location": "{cliente}/IMPLEMENTACIONES/SALIDA EN VIVO"},
	{"id": "4.5", "section": "SALIDA EN VIVO", "folder": "IMPLEMENTACIONES/SALIDA EN VIVO", "name": "VERIFICACIÓN CORREO", "kind": "email", "code_candidates": [], "allow_na": True, "location": "{cliente}/IMPLEMENTACIONES/SALIDA EN VIVO"},
	{"id": "5.1", "section": "OPERACIONES Y FORMALIZACIONES", "folder": "IMPLEMENTACIONES/SALIDA EN VIVO", "name": "Reporte a Operaciones", "kind": "email", "code_candidates": [], "allow_na": False, "location": "{cliente}/IMPLEMENTACIONES/SALIDA EN VIVO"},
	{"id": "5.2", "section": "OPERACIONES Y FORMALIZACIONES", "folder": "IMPLEMENTACIONES/SALIDA EN VIVO", "name": "Formalización SV al Cliente", "kind": "email", "code_candidates": [], "allow_na": False, "location": "{cliente}/IMPLEMENTACIONES/SALIDA EN VIVO"},
	{"id": "5.3", "section": "OPERACIONES Y FORMALIZACIONES", "folder": "IMPLEMENTACIONES/SALIDA EN VIVO", "name": "Formalización SV a Comercial", "kind": "email", "code_candidates": [], "allow_na": False, "location": "{cliente}/IMPLEMENTACIONES/SALIDA EN VIVO"},
	{"id": "5.4", "section": "OPERACIONES Y FORMALIZACIONES", "folder": "IMPLEMENTACIONES/SALIDA EN VIVO", "name": "PMO-ER-024 Check List SV", "kind": "document", "code_candidates": ["PMO-ER-024"], "allow_na": True, "location": "{cliente}/IMPLEMENTACIONES/SALIDA EN VIVO"},
	{"id": "6.1", "section": "ENTREGA Y REMISIÓN A MDS", "folder": "IMPLEMENTACIONES/SALIDA EN VIVO", "name": "PMO-ER-007 Entrega a MDS", "kind": "document", "code_candidates": ["PMO-ER-007"], "allow_na": True, "location": "{cliente}/IMPLEMENTACIONES/SALIDA EN VIVO"},
	{"id": "6.2", "section": "ENTREGA Y REMISIÓN A MDS", "folder": "IMPLEMENTACIONES/SALIDA EN VIVO", "name": "Correo Remisión a MDS", "kind": "email", "code_candidates": [], "allow_na": False, "location": "{cliente}/IMPLEMENTACIONES/SALIDA EN VIVO"},
]

def _normalize(value):
	return re.sub(r"\s+", " ", str(value or "").strip().lower())

class _HtmlTextExtractor(HTMLParser):
	def __init__(self):
		super().__init__(convert_charrefs=True)
		self.parts = []

	def handle_starttag(self, tag, attrs):
		if tag.lower() in {"br", "div", "p", "li", "tr"}:
			self.parts.append("\n")

	def handle_endtag(self, tag):
		if tag.lower() in {"div", "p", "li", "tr"}:
			self.parts.append("\n")

	def handle_data(self, data):
		self.parts.append(data)

def _html_to_text(value):
	extractor = _HtmlTextExtractor()
	extractor.feed(str(value or ""))
	extractor.close()
	return "".join(extractor.parts)

def _read_eml(path):
	with open(path, "rb") as source:
		message = BytesParser(policy=policy.default).parse(source)
	parts = [str(message.get("subject", "")), str(message.get("from", "")), str(message.get("to", ""))]
	sender = str(message.get("from", "")).strip()
	body_parts = message.walk() if message.is_multipart() else (message,)
	for part in body_parts:
		if part.get_content_maintype() == "multipart" or part.get_filename():
			continue
		if part.get_content_type() == "text/plain":
			parts.append(part.get_content())
		elif part.get_content_type() == "text/html":
			parts.append(_html_to_text(part.get_content()))
	return "\n".join(parts), sender

def _read_pdf(path):
	if PdfReader is None:
		return ""
	try:
		reader = PdfReader(path)
		text = "\n".join(page.extract_text() or "" for page in reader.pages)
		if text.strip():
			return text
	except Exception:
		pass
	try:
		with open(path, "rb") as source:
			return source.read().decode("latin-1", errors="ignore")
	except OSError:
		return ""

def _read_xlsx(path):
	if load_workbook is None:
		return ""
	workbook = load_workbook(path, read_only=True, data_only=True)
	try:
		values = []
		for sheet in workbook.worksheets:
			for row in sheet.iter_rows(values_only=True):
				values.extend(str(value) for value in row if value is not None)
		return "\n".join(values)
	finally:
		workbook.close()

def _read_document(path):
	extension = os.path.splitext(path)[1].lower()
	if extension == ".eml":
		return _read_eml(path)
	if extension == ".pdf":
		return _read_pdf(path), ""
	if extension in {".xlsx", ".xls"}:
		return _read_xlsx(path), ""
	return "", ""

def _collect_documents(client_path):
	documents = []
	for root, _, files in os.walk(client_path):
		for name in files:
			if os.path.splitext(name)[1].lower() not in SUPPORTED_EXTENSIONS:
				continue
			path = os.path.join(root, name)
			try:
				text, sender = _read_document(path)
				is_email = os.path.splitext(name)[1].lower() == ".eml"
				documents.append({
					"path": path,
					"name": name,
					"text": _normalize(text),
					"sender": sender,
					"analyst_name": _extract_analyst_name(text) if is_email else "",
				})
			except Exception as exc:
				documents.append({"path": path, "name": name, "text": "", "sender": "", "analyst_name": "", "error": str(exc)})
	return documents

def _analyst_names(documents):
	names = []
	for document in documents:
		name = document.get("analyst_name", "")
		if name:
			names.append(name)
	return ", ".join(name for name, _ in Counter(names).most_common()) or "No identificado en los correos"

def _resolve_writable_cell(sheet, target_cell):
	cell = sheet[target_cell]
	for merged_range in sheet.merged_cells.ranges:
		if merged_range.min_row <= cell.row <= merged_range.max_row and merged_range.min_col <= cell.column <= merged_range.max_col:
			return sheet.cell(row=merged_range.min_row, column=merged_range.min_col)
	return cell

def _find_report_sheet(workbook):
	for sheet in workbook.worksheets:
		labels = {_normalize(sheet.cell(row=row, column=column).value) for row in range(1, min(sheet.max_row, 12) + 1) for column in range(1, min(sheet.max_column, 8) + 1)}
		if "fecha:" in labels and any(label.startswith("nombre cliente:") for label in labels):
			return sheet
	return workbook.active

def _write_value_next_to_label(sheet, label, value):
	target = _normalize(label)
	for row in sheet.iter_rows():
		for cell in row:
			if not _normalize(cell.value).startswith(target):
				continue
			label_range = next((merged_range for merged_range in sheet.merged_cells.ranges if cell.coordinate in merged_range), None)
			column = label_range.max_col + 1 if label_range else cell.column + 1
			value_cell = _resolve_writable_cell(sheet, sheet.cell(row=cell.row, column=column).coordinate)
			value_cell.value = value
			return value_cell
	return None

def _display_path(path, client_path):
	try:
		return os.path.relpath(path, client_path)
	except ValueError:
		return os.path.basename(path)

def _write_observation(sheet, row, observation):
	cell = _resolve_writable_cell(sheet, f"K{row}")
	cell.value = observation
	alignment = copy(cell.alignment)
	alignment.wrap_text = True
	alignment.vertical = "top"
	cell.alignment = alignment
	column_width = sheet.column_dimensions[cell.column_letter].width or 12
	line_count = max(1, (len(str(observation)) // max(1, int(column_width))) + 1)
	sheet.row_dimensions[row].height = max(sheet.row_dimensions[row].height or 15, min(120, line_count * 15))

def _find_pmo_row_for_requirement(sheet, requirement_id):
	if requirement_id is None:
		return None
	value = str(requirement_id).strip()
	for row in range(10, 80):
		for column in (2, 3):
			cell_value = sheet.cell(row=row, column=column).value
			if cell_value is None:
				continue
			if str(cell_value).strip() == value:
				return row
			try:
				if value.isdigit() and int(cell_value) == int(value):
					return row
			except (TypeError, ValueError):
				continue
	return None

def _write_result_to_pmo_sheet(sheet, requirement_result):
	if not isinstance(requirement_result, dict):
		return False
	requirement_id = requirement_result.get("id")
	row = _find_pmo_row_for_requirement(sheet, requirement_id)
	if row is None:
		return False
	status_value = str(requirement_result.get("status", "")).strip().upper()
	observation = str(requirement_result.get("observation", "")).strip()
	for status_key, column_letter in (("SI", "H"), ("NO", "I"), ("NA", "J")):
		cell = _resolve_writable_cell(sheet, f"{column_letter}{row}")
		cell.value = "X" if status_value == status_key else ""
	_write_observation(sheet, row, observation)
	return True

def _has_keywords(documents, keywords):
	normalized_keywords = tuple(_normalize(keyword) for keyword in keywords)
	return [document for document in documents if any(keyword in document["text"] for keyword in normalized_keywords)]

def _has_all_keywords(document, keyword_groups):
	return all(any(_normalize(keyword) in document["text"] for keyword in group) for group in keyword_groups)

def _matching_documents(documents, keyword_groups, source_documents=None):
	candidates = source_documents if source_documents is not None else documents
	return [document for document in candidates if _has_all_keywords(document, keyword_groups)]

def _analyst_evidence(documents):
	return [document for document in documents if document.get("analyst_name")]

def _client_email_evidence(documents):
	return [
		document for document in documents
		if os.path.splitext(document["name"])[1].lower() == ".eml"
		and not document.get("analyst_name")
		and not any(domain in _normalize(document.get("sender", "")) for domain in ANALYST_DOMAINS)
	]

def _extract_analyst_name(text):
	lines = [re.sub(r"\s+", " ", line).strip() for line in str(text or "").splitlines()]
	process_keywords = (
		"analista de procesos documentales",
		"analista de procesos",
		"analista de proyectos",
	)
	signature_markers = re.compile(r"^(?:--\s*$|(?:cordial(?:mente)?|atentamente|saludos(?: cordiales)?|gracias)[,:]?\s*$)", re.IGNORECASE)
	signature_metadata = re.compile(r"\b(?:versi[oó]n|fecha|configuraci[oó]n|cargue|carga|aprobaci[oó]n)\s*:", re.IGNORECASE)
	for index, line in enumerate(lines):
		is_signature_marker = bool(signature_markers.match(line))
		is_process_analyst = any(keyword in _normalize(line) for keyword in process_keywords)
		if not is_signature_marker and not is_process_analyst:
			continue
		if is_process_analyst and not is_signature_marker:
			for candidate in (lines[index - 1] if index > 0 else "", lines[index + 1] if index + 1 < len(lines) else ""):
				if not signature_metadata.search(candidate):
					continue
				name_candidate = re.split(r"\s*[,;|]\s*", candidate, maxsplit=1)[0].strip(" -:")
				if len(name_candidate.split()) >= 2 and "@" not in name_candidate:
					return name_candidate
			continue
		signature_lines = []
		empty_lines = 0
		for candidate in lines[index + 1:index + 9]:
			if not candidate:
				empty_lines += 1
				if signature_lines and empty_lines > 3:
					break
				continue
			empty_lines = 0
			signature_lines.append(candidate)
			if any(keyword in _normalize(candidate) for keyword in process_keywords):
				break
		if not any(any(keyword in _normalize(candidate) for keyword in process_keywords) for candidate in signature_lines):
			continue
		for candidate in signature_lines:
			if any(keyword in _normalize(candidate) for keyword in process_keywords):
				continue
			name_candidate = re.split(r"\s*[,;|]\s*", candidate, maxsplit=1)[0].strip(" -:")
			if re.match(r"^(re|fw|fwd)\s*:", name_candidate, re.IGNORECASE) or "@" in name_candidate:
				continue
			if len(name_candidate.split()) >= 2 and re.fullmatch(r"[\wÁÉÍÓÚáéíóúÑñÜü.'-]+(?:\s+[\wÁÉÍÓÚáéíóúÑñÜü.'-]+)+", name_candidate):
				return name_candidate
	return ""

def _evaluate(documents, keywords, description):
	matches = _has_keywords(documents, keywords)
	analyst_documents = _analyst_evidence(documents)
	if matches:
		return "SI", f"Evidencia encontrada en: {matches[0]['name']}"
	if not analyst_documents:
		return "NA", f"No se encontró evidencia enviada por el Analista para {description}."
	return "NO", f"No se encontró evidencia de {description}."

def _evaluate_with_groups(documents, keyword_groups, description):
	matches = _matching_documents(documents, keyword_groups)
	if matches:
		return "SI", f"Evidencia encontrada en: {matches[0]['name']}"
	if not _analyst_evidence(documents):
		return "NA", f"No se encontró evidencia enviada por el Analista para {description}."
	return "NO", f"No se encontró evidencia de {description}."

def _evaluate_across_groups(documents, keyword_groups, description):
	matched_groups = []
	for group in keyword_groups:
		matches = _has_keywords(documents, group)
		if not matches:
			if not _analyst_evidence(documents):
				return "NA", f"No se encontró evidencia enviada por el Analista para {description}."
			return "NO", f"No se encontró toda la evidencia de {description}."
		matched_groups.append(matches[0]["name"])
	return "SI", f"Evidencia encontrada en: {', '.join(dict.fromkeys(matched_groups))}"

def _implementation_type(documents):
	for document in documents:
		if "formato de adjudic" not in _normalize(document["name"]):
			continue
		if os.path.splitext(document["name"])[1].lower() != ".xlsx" or load_workbook is None:
			continue
		try:
			workbook = load_workbook(document["path"], read_only=True, data_only=True)
			for sheet in workbook.worksheets:
				rows = list(sheet.iter_rows(values_only=True))
				for row_index, row in enumerate(rows):
					values = [_normalize(value) for value in row if value is not None]
					label_index = next((index for index, value in enumerate(row) if "tipo de integrac" in _normalize(value)), None)
					if label_index is None:
						continue
					for candidate_row in rows[row_index + 1:row_index + 4]:
						candidate = _normalize(candidate_row[label_index] if label_index < len(candidate_row) else "")
						if "portal manual" in candidate:
							return "Portal manual"
						if "indirecta" in candidate:
							return "Indirecta"
						if "directa" in candidate:
							return "Directa"
		except Exception:
			continue
	return "Tipo de integración no identificado"

def _build_results(documents):
	implementation_type = _implementation_type(documents)
	results = {}
	for row, description, keywords in CHECKS:
		if row in (19, 20, 21) and implementation_type != "Portal manual":
			results[row] = "NA", "No aplica para este tipo de integración."
		elif row == 11:
			analyst_tests = _matching_documents(_analyst_evidence(documents), (("prueba", "pruebas", "testing"),))
			client_approval = _matching_documents(
				documents,
				(("aprobado", "aprobación", "aprobacion", "aceptado", "visto bueno", "conforme"),),
				source_documents=_client_email_evidence(documents),
			)
			if analyst_tests and client_approval:
				results[row] = "SI", f"Pruebas y aprobación encontradas en: {analyst_tests[0]['name']}, {client_approval[0]['name']}"
			elif not analyst_tests:
				results[row] = "NA", "No se encontró un correo enviado por el Analista con las pruebas al cliente."
			else:
				results[row] = "NO", "Se encontraron pruebas del Analista, pero no un correo de aprobación del cliente."
		elif row == 12:
			results[row] = _evaluate_with_groups(
				_analyst_evidence(documents),
				(("credencial", "credenciales"), ("producción", "produccion", "portal")),
				description,
			)
		elif row == 13:
			matches = [document for document in documents if os.path.splitext(document["name"])[1].lower() == ".pdf" and "etf" in _normalize(document["name"])]
			results[row] = (("SI", f"ETF encontrado: {matches[0]['name']}") if matches else ("NO", "No se encontró un archivo PDF, correo o Excel relacionado con ETF."))
		elif row == 16:
			results[row] = _evaluate_across_groups(
				documents,
				(("facturación", "facturacion", "servicio solicitado", "servicio contratado"), ("creación", "creacion", "creada", "habilitación", "habilitacion")),
				description,
			)
		elif row == 18:
			results[row] = _evaluate_with_groups(
				documents,
				(("dispapeles",), ("proveedor tecnológico", "proveedor tecnologico", "proveedor"), ("asociado", "asociación", "asociacion", "registrado")),
				description,
			)
		elif row == 17:
			results[row] = _evaluate_with_groups(
				documents,
				(("usuario administrador", "administrador de la empresa"),),
				description,
			)
		elif row == 22:
			results[row] = _evaluate_with_groups(
				documents,
				(("web service", "webservice"), ("usuario",), ("contraseña", "clave", "password")),
				description,
			)
		elif row == 23:
			results[row] = _evaluate_with_groups(
				documents,
				(("etf",), ("coinciden", "coincidan", "confirmación", "confirmacion", "validación", "validacion"), ("facturación", "facturacion", "nómina", "nomina", "documento soporte", "eventos dian", "radian")),
				description,
			)
		elif row == 26:
			results[row] = _evaluate_with_groups(
				documents,
				(("visto bueno", "aprobación", "aprobacion"), ("analista de procesos", "salida en vivo", "pmo")),
				description,
			)
		elif row == 27:
			results[row] = _evaluate_with_groups(
				_analyst_evidence(documents),
				(("credencial", "credenciales"), ("producción", "produccion", "portal")),
				description,
			)
		else:
			results[row] = _evaluate(documents, keywords, description)

	status, observation = results[23]
	results[23] = status, f"Tipo de integración: {implementation_type}. {observation}"
	return results

def _find_subdirectory(parent_path, directory_name):
	if not os.path.isdir(parent_path):
		return None
	target = _normalize(directory_name)
	for name in os.listdir(parent_path):
		path = os.path.join(parent_path, name)
		if os.path.isdir(path) and _normalize(name) == target:
			return path
	return None

def _normalize_name(value):
	text = unicodedata.normalize("NFKD", str(value or ""))
	text = "".join(ch for ch in text if not unicodedata.combining(ch))
	text = text.lower().replace("&", "y")
	text = re.sub(r"[^a-z0-9]+", " ", text)
	return " ".join(text.split())

def _file_has_exact_code(path, code_candidates):
	if not code_candidates:
		return True
	return _document_contains_exact_code(path, code_candidates)

def _service_siglas(service_name):
	text = str(service_name or "").strip()
	if not text:
		return "SERV"
	if _normalize_name(text) == "facturacion electronica":
		return "FE"
	matches = re.findall(r"[A-Za-zÁÉÍÓÚÑáéíóúñ]+", text)
	if not matches:
		return "SERV"
	siglas = "".join(part[0].upper() for part in matches if part)
	return siglas[:8] or "SERV"

def _safe_report_client_name(client_path):
	name = os.path.basename(str(client_path or "")).strip() or "cliente"
	name = unicodedata.normalize("NFKD", name)
	name = "".join(ch for ch in name if not unicodedata.combining(ch))
	name = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("_.")
	return name or "cliente"

def validate_uploaded_requirement_file(file_path, requirement):
	if not file_path or not os.path.exists(file_path):
		return False, "No se encontró el archivo cargado."
	req_id = requirement.get("id")
	name = os.path.basename(file_path)
	ext = os.path.splitext(name)[1].lower()
	candidate_name = requirement.get("name", "")
	code_candidates = requirement.get("code_candidates") or []
	normalized_name = _normalize_name(name)

	if req_id == "3.1":
		if ext != ".eml":
			return False, "La evidencia de 3.1 debe ser un correo (.eml) con adjunto Excel PMO-ER-010."
		if not _email_contains_attachment_code(file_path, ["PMO-ER-010"], allowed_extensions={".xlsx", ".xls"}):
			return False, "El correo no contiene un adjunto Excel con el código exacto PMO-ER-010."
		return True, ""

	if req_id == "4.1":
		if ext in {".pdf", ".xlsx", ".xls"}:
			if _document_contains_exact_code(file_path, ["PMO-ER-011"]):
				return True, ""
			return False, "El archivo cargado no contiene exactamente PMO-ER-011."
		if ext == ".eml":
			if _normalize_name(candidate_name) in normalized_name or "pmo er 011 cyp" in normalized_name:
				return True, ""
			return False, "El correo debe llevar el nombre PMO-ER-011 CyP."
		return False, "La evidencia de 4.1 debe ser un archivo o correo con el nombre PMO-ER-011 CyP."

	if req_id == "2.2":
		if ext in {".xlsx", ".xls", ".pdf", ".doc", ".docx"}:
			if "pmo md 002 project charter" in normalized_name or "pmo md 002" in normalized_name:
				if ext in {".xlsx", ".xls"}:
					return _file_has_exact_code(file_path, ["PMO-MD-002"]), "El código del Project Charter no corresponde a PMO-MD-002."
				return True, ""
			return False, "El archivo debe llevar el nombre de PMO-MD-002 Project Charter."
		return False, "El archivo requerido para 2.2 debe ser un PDF o Excel válido."

	if req_id == "2.12":
		if ext.lower() not in {".png", ".jpg", ".jpeg"}:
			return False, "La evidencia de 2.12 debe ser una imagen PNG con nombre que incluya Mapeo_de_datos."
		if "mapeo de datos" not in normalized_name and "mapeo_de_datos" not in normalized_name:
			return False, "La imagen debe llevar el nombre Mapeo_de_datos."
		return True, ""

	if req_id == "2.11":
		if ext != ".eml":
			return False, "La evidencia de 2.11 debe ser un correo (.eml)."
		if "informe de adjudicacion" not in normalized_name:
			return False, "El correo cargado debe contener 'INFORME DE ADJUDICACION' en el nombre."
		return True, ""

	if req_id == "2.12":
		if ext.lower() not in {".png", ".jpg", ".jpeg"}:
			return False, "La evidencia de 2.12 debe ser una imagen PNG con nombre que incluya Mapeo_de_datos."
		if "mapeo de datos" not in normalized_name and "mapeo_de_datos" not in normalized_name:
			return False, "La imagen debe llevar el nombre Mapeo_de_datos."
		return True, ""

	if code_candidates:
		if ext in {".xlsx", ".xls"}:
			ok = _file_has_exact_code(file_path, code_candidates)
			return ok, f"El código del archivo no corresponde a {', '.join(code_candidates)}."
		if ext == ".pdf":
			return True, ""
		if candidate_name and _normalize_name(candidate_name) not in normalized_name and not any(token in normalized_name for token in _normalize_name(candidate_name).split()):
			return False, f"El nombre del archivo no coincide con '{candidate_name}'."
		return True, ""
	return True, ""

def find_requirement_files(client_path, requirement_name, extensions=None):
	if not os.path.isdir(client_path):
		return []
	extensions = extensions or {".eml", ".pdf", ".xlsx", ".xls", ".doc", ".docx"}
	normalized_requirement = _normalize_name(requirement_name)
	matches = []
	for root, _, files in os.walk(client_path):
		for name in files:
			ext = os.path.splitext(name)[1].lower()
			if ext not in extensions:
				continue
			normalized_name = _normalize_name(name)
			if normalized_requirement in normalized_name or all(token in normalized_name for token in normalized_requirement.split() if len(token) >= 2):
				matches.append(os.path.join(root, name))
	return sorted(matches)

def _document_contains_exact_code(path, code_candidates):
	if not code_candidates:
		return True
	extension = os.path.splitext(path)[1].lower()
	if extension == ".pdf":
		return True
	if extension not in {".xlsx", ".xls"}:
		return False
	return _text_contains_exact_code(path, code_candidates)

def _text_contains_exact_code(path, code_candidates):
	text = ""
	try:
		text, _ = _read_document(path)
	except Exception:
		text = ""
	text = _normalize_name(text)
	for code in code_candidates:
		code_text = _normalize_name(code)
		if re.search(rf"(?<![a-z0-9]){re.escape(code_text)}(?![a-z0-9])", text):
			return True
	return False

def _email_contains_attachment_code(path, code_candidates, allowed_extensions=None):
	attachment_extensions = allowed_extensions or {".pdf", ".xlsx", ".xls", ".doc", ".docx"}
	content_type_extensions = {
		"application/pdf": ".pdf",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
		"application/vnd.ms-excel": ".xls",
		"application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
		"application/msword": ".doc",
	}
	try:
		with open(path, "rb") as source:
			message = BytesParser(policy=policy.default).parse(source)
		for part in message.walk():
			if part.get_content_maintype() == "multipart":
				continue
			filename = part.get_filename() or ""
			extension = os.path.splitext(filename)[1].lower()
			if extension not in attachment_extensions:
				extension = content_type_extensions.get(part.get_content_type(), "")
			if not extension:
				continue
			payload = part.get_payload(decode=True)
			if not payload:
				continue
			temporary = tempfile.NamedTemporaryFile(suffix=extension, delete=False)
			temporary_path = temporary.name
			try:
				temporary.write(payload)
				temporary.close()
				filename_matches = any(_normalize_name(code) in _normalize_name(filename) for code in code_candidates)
				code_matches = _text_contains_exact_code(temporary_path, code_candidates)
				if (extension == ".pdf" and (filename_matches or code_matches)) or (extension in {".xlsx", ".xls"} and code_matches):
					return filename or f"adjunto{extension}"
			finally:
				os.unlink(temporary_path)
		return ""
	except Exception:
		return ""

def evaluate_document_requirement(path, requirement_name, required_code=None, kind="document"):
	if not path or not os.path.exists(path):
		return {"status": "NO", "observation": f"No se encontró el archivo requerido: {requirement_name}"}
	file_name = os.path.basename(path)
	if required_code:
		code_ok = _document_contains_exact_code(path, [required_code])
		if code_ok:
			if os.path.splitext(file_name)[1].lower() == ".pdf":
				return {"status": "SI", "observation": f"Se encontró {file_name}; el nombre coincide con el formato requerido."}
			return {"status": "SI", "observation": f"Se encontró {file_name} y contiene {required_code}."}
		return {"status": "NO", "observation": f"Se encontró {file_name}, pero no contiene exactamente {required_code}."}
	return {"status": "SI", "observation": f"Se encontró {file_name} con el nombre requerido."}

def _detect_project_size(client_path):
	candidates = []
	for root, _, files in os.walk(client_path):
		for name in files:
			normalized_name = _normalize_name(name)
			if "si er 020" in normalized_name or "si er020" in normalized_name:
				candidates.append(os.path.join(root, name))
	if not candidates:
		return "desconocido"

	for path in sorted(candidates, key=lambda item: os.path.getmtime(item), reverse=True):
		try:
			text, _ = _read_document(path)
			normalized = _normalize_name(text)
			if any(marker in normalized for marker in (
				"califica como grande",
				"tipo de proyecto grande",
				"proyecto grande",
				"cliente grande",
				"calificacion grande",
			)):
				return "grande"
			if any(marker in normalized for marker in (
				"califica como mediano",
				"tipo de proyecto mediano",
				"proyecto mediano",
				"cliente mediano",
				"calificacion mediano",
			)):
				return "mediano"
			if any(marker in normalized for marker in (
				"califica como pequeno",
				"tipo de proyecto pequeno",
				"proyecto pequeno",
				"cliente pequeno",
				"calificacion pequeno",
			)):
				return "pequeno"
		except Exception:
			continue
	return "desconocido"

def _is_project_large(client_path):
	return _detect_project_size(client_path) == "grande"

def _has_project_size(client_path, target):
	return _detect_project_size(client_path) == target

def _project_size_for_2_1(client_path):
	return _detect_project_size(client_path)

def _propagate_not_applicable_override(manual_overrides, requirement_id):
	if requirement_id not in {"2.11", "2.12"}:
		return None
	parent_override = manual_overrides.get("2.10", {}) if isinstance(manual_overrides, dict) else {}
	if not parent_override.get("not_applicable"):
		return None
	return {
		"not_applicable": True,
		"observation": parent_override.get("observation") or "No aplica para este cliente.",
	}

def evaluate_requirement(client_path, requirement, not_applicable=False, not_applicable_observation=""):
	if requirement.get("id") in SERVICE_NOT_APPLICABLE_REQUIREMENTS:
		return {
			"id": requirement.get("id"),
			"section": requirement.get("section"),
			"folder": requirement.get("folder"),
			"status": "NA",
			"observation": NOT_APPLICABLE_OBSERVATION,
			"path": None,
		}
	if not_applicable:
		return {
			"id": requirement.get("id"),
			"section": requirement.get("section"),
			"folder": requirement.get("folder"),
			"status": "NA",
			"observation": not_applicable_observation or "No aplica para este cliente.",
			"path": None,
		}

	requirement_id = requirement.get("id")
	if requirement_id == "2.2":
		size = _project_size_for_2_1(client_path)
		if size == "pequeno":
			return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NA", "observation": "No aplica porque es un cliente pequeño.", "path": None}
		if size == "mediano":
			return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NA", "observation": "No aplica porque es un cliente mediano.", "path": None}
		name = requirement.get("name", "")
		matches = find_requirement_files(client_path, name, extensions={".pdf", ".xlsx", ".xls", ".doc", ".docx"})
		for path in matches:
			if _document_contains_exact_code(path, ["PMO-MD-002"]):
				return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "SI", "observation": f"Documento encontrado en: {_display_path(path, client_path)}", "path": path}
		if matches:
			return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": f"Se encontró {os.path.basename(matches[0])}, pero no contiene exactamente PMO-MD-002.", "path": matches[0]}
		return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": "No se encontró el Project Charter requerido.", "path": None}

	if requirement_id == "2.12":
		artes_dir = os.path.join(client_path, "IMPLEMENTACIONES", "PRELIMINARES", "ARTES")
		matches = []
		if os.path.isdir(artes_dir):
			for root, _, files in os.walk(artes_dir):
				for file_name in files:
					if os.path.splitext(file_name)[1].lower() in {".png", ".jpg", ".jpeg"} and "factura" in _normalize_name(file_name):
						matches.append(os.path.join(root, file_name))
		if matches:
			return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "SI", "observation": f"Documento encontrado en: {_display_path(matches[0], client_path)}", "path": matches[0]}
		return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": "No se encontró la o las PNG con nombre Factura_# en la carpeta ARTES.", "path": None}

	if requirement_id == "2.11":
		artes_dir = os.path.join(client_path, "IMPLEMENTACIONES", "PRELIMINARES", "ARTES")
		matches = find_requirement_files(artes_dir, "CORREO APROBACIÓN ARTES", extensions={".eml"})
		if matches:
			return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "SI", "observation": f"Correo encontrado en: {_display_path(matches[0], client_path)}", "path": matches[0]}
		return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": "No se encontró el correo requerido: CORREO APROBACIÓN ARTES.", "path": None}

	if requirement_id == "2.12":
		artes_dir = os.path.join(client_path, "IMPLEMENTACIONES", "PRELIMINARES", "ARTES")
		matches = []
		if os.path.isdir(artes_dir):
			for root, _, files in os.walk(artes_dir):
				for file_name in files:
					if os.path.splitext(file_name)[1].lower() in {".png", ".jpg", ".jpeg"} and ("mapeo de datos" in _normalize_name(file_name) or "mapeo_de_datos" in _normalize_name(file_name)):
						matches.append(os.path.join(root, file_name))
		if matches:
			return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "SI", "observation": f"PNG encontrado en: {_display_path(matches[0], client_path)}", "path": matches[0]}
		return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": "No se encontró la PNG requerida: Mapeo_de_datos.", "path": None}

	if requirement_id == "2.8":
		size = _project_size_for_2_1(client_path)
		if size == "pequeno":
			return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NA", "observation": "No aplica porque es un cliente pequeño.", "path": None}
		if size == "mediano":
			return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NA", "observation": "No aplica porque es un cliente mediano.", "path": None}
		if size != "grande":
			return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NA", "observation": "No aplica para este proyecto.", "path": None}
		name = requirement.get("name", "")
		matches = find_requirement_files(client_path, name, extensions={".pdf", ".xlsx", ".xls", ".doc", ".docx"})
		for path in matches:
			if _document_contains_exact_code(path, ["PMO-ER-009"]):
				return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "SI", "observation": f"Documento encontrado en: {_display_path(path, client_path)}", "path": path}
		if matches:
			return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": f"Se encontró {os.path.basename(matches[0])}, pero no contiene exactamente PMO-ER-009.", "path": matches[0]}
		return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": f"No se encontró el documento requerido: {name}", "path": None}

	if requirement_id == "2.9":
		name = requirement.get("name", "")
		candidates = [
			"CAPACITACIÓN",
			"CORREO APROBACIÓN",
		]
		matches = []
		for candidate in candidates:
			matches.extend(find_requirement_files(client_path, candidate, extensions={".eml"}))
		if matches:
			return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "SI", "observation": f"Correo encontrado en: {_display_path(matches[0], client_path)}", "path": matches[0]}
		return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": f"No se encontró el correo requerido: {name}", "path": None}

	if requirement_id == "3.1":
		matches = find_requirement_files(client_path, requirement.get("name", ""), extensions={".eml"})
		for path in matches:
			attachment_name = _email_contains_attachment_code(path, ["PMO-ER-010"], allowed_extensions={".xlsx", ".xls"})
			if attachment_name:
				return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "SI", "observation": f"Correo encontrado en: {_display_path(path, client_path)}; adjunto válido: {attachment_name}", "path": path}
		if matches:
			return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": "Se encontraron correos PMO-ER-010, pero ninguno contiene un adjunto con el código exacto PMO-ER-010.", "path": matches[0]}
		return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": "No se encontró el correo requerido: PMO-ER-010", "path": None}

	if requirement_id == "4.1":
		name = requirement.get("name", "")
		file_matches = find_requirement_files(client_path, name, extensions={".pdf", ".xlsx", ".xls"})
		email_matches = find_requirement_files(client_path, name, extensions={".eml"})
		valid_file = next((path for path in file_matches if _document_contains_exact_code(path, ["PMO-ER-011"])), None)
		if valid_file and email_matches:
			return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "SI", "observation": f"Documento encontrado en: {_display_path(valid_file, client_path)}; correo encontrado en: {_display_path(email_matches[0], client_path)}", "path": valid_file}
		if not valid_file and not file_matches:
			observation = "No se encontró el archivo requerido PMO-ER-011 CyP."
		elif not valid_file:
			observation = "Se encontró el archivo PMO-ER-011 CyP., pero no contiene exactamente PMO-ER-011."
		else:
			observation = "Se encontró el archivo PMO-ER-011 CyP., pero no el correo cuyo nombre contiene PMO-ER-011 CyP."
		return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": observation, "path": valid_file or (file_matches[0] if file_matches else None)}

	if requirement_id == "2.12":
		matches = find_requirement_files(client_path, "MAPEO DE DATOS", extensions={".eml"})
		if matches:
			return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "SI", "observation": f"Correo encontrado en: {_display_path(matches[0], client_path)}", "path": matches[0]}
		return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": "No se encontró el correo solicitado: MAPEO DE DATOS.", "path": None}

	if requirement_id == "4.3":
		pdf_matches = find_requirement_files(client_path, "Resolución", extensions={".pdf"})
		zip_matches = find_requirement_files(client_path, "Resoluciones", extensions={".zip"})
		if pdf_matches or zip_matches:
			location = pdf_matches[0] if pdf_matches else zip_matches[0]
			return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "SI", "observation": f"Documento encontrado en: {_display_path(location, client_path)}", "path": location}
		return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": "No se encontró resolución o ZIP de resoluciones.", "path": None}

	if requirement_id == "4.5":
		matches = find_requirement_files(client_path, "VERIFICACIÓN CORREO", extensions={".eml"})
		if matches:
			return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "SI", "observation": f"Correo encontrado en: {_display_path(matches[0], client_path)}", "path": matches[0]}
		return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": "No se encontró el correo solicitado: VERIFICACIÓN CORREO.", "path": None}

	if requirement_id == "2.2":
		name = requirement.get("name", "")
		matches = find_requirement_files(client_path, name, extensions={".pdf", ".xlsx", ".xls", ".doc", ".docx"})
		for path in matches:
			if _document_contains_exact_code(path, ["PMO-MD-002"]):
				return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "SI", "observation": f"Documento encontrado en: {_display_path(path, client_path)}", "path": path}
		if matches:
			return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": f"Se encontró {os.path.basename(matches[0])}, pero no contiene exactamente PMO-MD-002.", "path": matches[0]}
		return {"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": "No se encontró el Project Charter requerido.", "path": None}

	name = requirement.get("name", "")
	kind = requirement.get("kind", "document")
	code_candidates = requirement.get("code_candidates") or []
	allowed_extensions = {".eml"} if kind == "email" else {".eml", ".pdf", ".xlsx", ".xls", ".doc", ".docx"}
	matches = find_requirement_files(client_path, name, extensions=allowed_extensions)
	if not matches:
		return {"id": requirement.get("id"), "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": f"No se encontró el documento requerido: {name}", "path": None}
	for path in matches:
		if kind == "email":
			if _normalize_name(name) in _normalize_name(os.path.basename(path)):
				return {"id": requirement.get("id"), "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "SI", "observation": f"Correo encontrado en: {_display_path(path, client_path)}", "path": path}
			continue
		if code_candidates:
			if _document_contains_exact_code(path, code_candidates):
				return {"id": requirement.get("id"), "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "SI", "observation": f"Documento encontrado en: {_display_path(path, client_path)}", "path": path}
			return {"id": requirement.get("id"), "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": f"Se encontró el archivo {os.path.basename(path)}, pero no contiene exactamente los códigos requeridos: {', '.join(code_candidates)}", "path": path}
		return {"id": requirement.get("id"), "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "SI", "observation": f"Documento encontrado en: {_display_path(path, client_path)}", "path": path}
	return {"id": requirement.get("id"), "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": f"No se encontró {name}", "path": None}

def evaluate_all_requirements(client_path, manual_overrides=None):
	manual_overrides = manual_overrides or {}
	results = []
	for requirement in REQUIREMENT_CATALOG:
		requirement_id = requirement.get("id")
		override = manual_overrides.get(requirement_id, {}) if isinstance(manual_overrides, dict) else {}
		propagated_override = _propagate_not_applicable_override(manual_overrides, requirement_id)
		if propagated_override:
			results.append({
				"id": requirement_id,
				"section": requirement.get("section"),
				"folder": requirement.get("folder"),
				"status": "NA",
				"observation": propagated_override.get("observation") or "No aplica para este cliente.",
				"path": None,
			})
			continue
		if override.get("not_applicable"):
			results.append({
				"id": requirement_id,
				"section": requirement.get("section"),
				"folder": requirement.get("folder"),
				"status": "NA",
				"observation": override.get("observation") or "No aplica para este cliente.",
				"path": None,
			})
			continue
		if requirement_id == "4.1" and override.get("paths"):
			file_paths = [path for path in override["paths"] if os.path.splitext(path)[1].lower() in {".pdf", ".xlsx", ".xls"}]
			email_paths = [path for path in override["paths"] if os.path.splitext(path)[1].lower() == ".eml"]
			valid_file = next((path for path in file_paths if _normalize_name(requirement.get("name", "")) in _normalize_name(os.path.basename(path)) and _document_contains_exact_code(path, ["PMO-ER-011"])), None)
			valid_email = next((path for path in email_paths if _normalize_name(requirement.get("name", "")) in _normalize_name(os.path.basename(path))), None)
			if valid_file and valid_email:
				results.append({"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "SI", "observation": f"Archivo cargado: {valid_file}; correo cargado: {valid_email}", "path": valid_file})
			else:
				results.append({"id": requirement_id, "section": requirement.get("section"), "folder": requirement.get("folder"), "status": "NO", "observation": "La carga manual de 4.1 debe incluir un archivo válido y un correo PMO-ER-011 CyP.", "path": valid_file or valid_email})
			continue
		if override.get("path"):
			manual_result = evaluate_requirement(client_path, requirement)
			if manual_result.get("status") == "SI":
				manual_result["observation"] = override.get("observation") or manual_result.get("observation", "")
			results.append(manual_result)
			continue
		results.append(evaluate_requirement(client_path, requirement, not_applicable=bool(override.get("not_applicable")), not_applicable_observation=override.get("observation")))
	return results

def run_audit(client_path, service_name=SERVICE_NAME, reports_dir=None, template_path=None, manual_overrides=None):
	if _normalize(service_name) != _normalize(SERVICE_NAME):
		raise ValueError(f"Este flujo solo aplica al servicio '{SERVICE_NAME}'.")
	if not os.path.isdir(client_path):
		raise FileNotFoundError(f"No existe la carpeta del cliente: {client_path}")
	if load_workbook is None:
		raise RuntimeError("La librería openpyxl es necesaria para generar el reporte.")

	base_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
	template_path = template_path or os.path.join(base_dir, "Plantillas", "PMO-ER-005.xlsx")
	if not os.path.isfile(template_path):
		raise FileNotFoundError(f"No existe la plantilla del reporte: {template_path}")
	implementation_dir = _find_subdirectory(client_path, "IMPLEMENTACIONES") or os.path.join(client_path, "IMPLEMENTACIONES")
	output_dir = _find_subdirectory(implementation_dir, "SALIDA EN VIVO") or os.path.join(implementation_dir, "SALIDA EN VIVO")
	os.makedirs(output_dir, exist_ok=True)
	client_name = _safe_report_client_name(client_path)
	report_name = f"PMO-ER-005_{client_name}_{_service_siglas(service_name)}.xlsx"
	output_path = os.path.join(output_dir, report_name)
	if os.path.exists(output_path):
		os.remove(output_path)
	shutil.copy2(template_path, output_path)

	documents = _collect_documents(client_path)
	manual_overrides = manual_overrides or {}
	with warnings.catch_warnings():
		warnings.filterwarnings("ignore", message="Data Validation extension is not supported and will be removed")
		workbook = load_workbook(output_path)
		sheet = _find_report_sheet(workbook)
		for label, value in (
			("Fecha:", datetime.now().strftime("%d/%m/%Y")),
			("Analista de Procesos:", _analyst_names(documents)),
			("Nombre Cliente:", os.path.basename(client_path)),
			("Servicio/Producto Implementado:", service_name),
		):
			_write_value_next_to_label(sheet, label, value)
		for item in evaluate_all_requirements(client_path, manual_overrides=manual_overrides):
			_write_result_to_pmo_sheet(sheet, item)
		workbook.save(output_path)
		workbook.close()
	return output_path