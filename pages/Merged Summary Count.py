import re
from datetime import date
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path

import pandas as pd
import streamlit as st


st.header("Merged Summary Count")

DEFAULT_SERVER_MERGED_DIR = r"\\192.168.15.241\admin\ACTIVE\scperez\MAYA\MERGED ACCOUNTS"
ALLOWED_EXTENSIONS = {".xlsx", ".xls", ".csv"}
EXCLUDED_PLACEMENT = "Maya PayLater 181 DPD & UP"
REMARK_FILTERS = ["ACTIVE", "FULLY PAID", "REGULAR PULLED OUT"]
PLACEMENT_ORDER = [
	"Maya Credit 121 - 150 DPD",
	"Maya Credit 181 DPD & UP",
	"Maya Negosyo Advance 121 - 150 DPD",
	"Maya Negosyo Advance 181 DPD & UP",
	"Maya SME Flexi Loan 1 - 30 DPD",
	"Maya SME Flexi Loan 31 - 60 DPD",
	"Maya SME Flexi Loan 61 - 90 DPD",
	"Maya SME Flexi Loan 91 - 120 DPD",
	"Maya SME Flexi Loan 121 - 150 DPD",
	"Maya SME Flexi Loan 151 - 180 DPD",
	"Maya SME Flexi Loan 181 DPD & UP",
]
REMARKS_PLACEMENT_ORDER = [
	"Maya Credit 121 - 150 DPD",
	"Maya Credit 181 DPD & UP",
	"Maya Negosyo Advance 121 - 150 DPD",
	"Maya Negosyo Advance 181 DPD & UP",
	"Maya SME Flexi Loan 1 - 30 DPD",
	"Maya SME Flexi Loan 121 - 150 DPD",
	"Maya SME Flexi Loan 151 - 180 DPD",
	"Maya SME Flexi Loan 181 DPD & UP",
	"Maya SME Flexi Loan 31 - 60 DPD",
	"Maya SME Flexi Loan 61 - 90 DPD",
	"Maya SME Flexi Loan 91 - 120 DPD",
]


def normalize_key(text: str) -> str:
	return re.sub(r"[^A-Z0-9]+", "_", str(text).strip().upper()).strip("_")


def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
	lookup: dict[str, str] = {}
	for col in df.columns:
		key = normalize_key(col)
		if key and key not in lookup:
			lookup[key] = col

	for candidate in candidates:
		key = normalize_key(candidate)
		if key in lookup:
			return lookup[key]
	return None


TWO_DECIMAL_PLACES = Decimal("0.01")


def to_decimal_ob(value) -> Decimal:
	if pd.isna(value):
		return Decimal("0")

	if isinstance(value, Decimal):
		return value

	text = str(value).strip()
	if not text:
		return Decimal("0")

	negative = False
	if text.startswith("(") and text.endswith(")"):
		negative = True
		text = text[1:-1]

	cleaned = text.replace(",", "")

	try:
		parsed = Decimal(cleaned)
	except (InvalidOperation, ValueError):
		return Decimal("0")

	return -parsed if negative else parsed


def resolve_server_file(server_input: str) -> tuple[Path | None, str | None]:
	root_dir = Path(DEFAULT_SERVER_MERGED_DIR)
	if not root_dir.exists() or not root_dir.is_dir():
		return None, f"Server folder is not reachable: {root_dir}"

	requested = (server_input or "").strip()
	if not requested:
		return None, "Please provide a file name or relative path."

	normalized = requested.replace("/", "\\")
	requested_path = Path(normalized)

	if requested_path.exists() and requested_path.is_file():
		return requested_path, None

	direct_candidate = root_dir / normalized
	if direct_candidate.exists() and direct_candidate.is_file():
		return direct_candidate, None

	requested_name = requested_path.name
	if not requested_name:
		return None, "Please provide a valid file name or relative path."

	current_year_dir = root_dir / str(date.today().year)
	search_roots = [current_year_dir, root_dir] if current_year_dir.exists() else [root_dir]

	has_extension = Path(requested_name).suffix.lower() in ALLOWED_EXTENSIONS
	matches: list[Path] = []

	for base in search_roots:
		if not base.exists() or not base.is_dir():
			continue

		for file_path in base.rglob("*"):
			if not file_path.is_file() or file_path.suffix.lower() not in ALLOWED_EXTENSIONS:
				continue

			if has_extension:
				if file_path.name.lower() == requested_name.lower():
					matches.append(file_path)
			else:
				if file_path.stem.lower() == requested_name.lower() or file_path.name.lower() == requested_name.lower():
					matches.append(file_path)

	unique_matches = sorted(set(matches))
	if len(unique_matches) == 1:
		return unique_matches[0], None

	if len(unique_matches) > 1:
		preview = "\n".join(str(p) for p in unique_matches[:8])
		return None, (
			f"Multiple files matched '{requested_name}'. Please provide a more specific relative path.\n"
			f"Sample matches:\n{preview}"
		)

	return None, f"File not found from input: {requested}"


def read_dataframe_from_input(server_file: Path | None, uploaded_file) -> tuple[pd.DataFrame | None, str]:
	if server_file is not None:
		suffix = server_file.suffix.lower()
		if suffix == ".csv":
			return pd.read_csv(server_file), f"Server file: {server_file}"
		return pd.read_excel(server_file), f"Server file: {server_file}"

	if uploaded_file is not None:
		suffix = Path(uploaded_file.name).suffix.lower()
		if suffix == ".csv":
			return pd.read_csv(uploaded_file), f"Uploaded file: {uploaded_file.name}"
		return pd.read_excel(uploaded_file), f"Uploaded file: {uploaded_file.name}"

	return None, ""


def build_outputs(source_df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
	placement_col = find_column(source_df, ["PLACEMENT"])
	ob_col = find_column(source_df, ["OB", "OUTSTANDING BALANCE", "OSB", "BALANCE"])
	remarks_col = find_column(source_df, ["REMARKS", "REMARK", "STATUS"])

	missing = []
	if placement_col is None:
		missing.append("PLACEMENT")
	if ob_col is None:
		missing.append("OB")
	if remarks_col is None:
		missing.append("REMARKS")

	if missing:
		raise ValueError(f"Missing required column(s): {', '.join(missing)}")

	df = source_df[[placement_col, ob_col, remarks_col]].copy()
	df[placement_col] = df[placement_col].astype(str).str.strip()
	df = df[df[placement_col] != ""]
	df = df[df[placement_col].str.upper() != EXCLUDED_PLACEMENT.upper()]

	df[ob_col] = df[ob_col].apply(to_decimal_ob)
	df[remarks_col] = df[remarks_col].astype(str).str.strip().str.upper()

	counts_by_placement = df.groupby(placement_col, dropna=False).size().reindex(PLACEMENT_ORDER, fill_value=0).astype(int)
	ob_sums_by_placement = df.groupby(placement_col, dropna=False)[ob_col].apply(
		lambda values: sum(values, Decimal("0"))
	).reindex(PLACEMENT_ORDER)
	ob_sums_by_placement = ob_sums_by_placement.apply(lambda value: value if isinstance(value, Decimal) else Decimal("0"))

	columns = [*PLACEMENT_ORDER, "TOTAL"]
	count_values = [int(counts_by_placement.get(col, 0)) for col in PLACEMENT_ORDER]
	count_total = int(sum(count_values))

	ob_values = [
		float(ob_sums_by_placement.get(col, Decimal("0")).quantize(TWO_DECIMAL_PLACES, rounding=ROUND_HALF_UP))
		for col in PLACEMENT_ORDER
	]
	ob_total = float(sum(ob_sums_by_placement, Decimal("0")).quantize(TWO_DECIMAL_PLACES, rounding=ROUND_HALF_UP))

	placement_summary = pd.DataFrame(
		[count_values + [count_total], ob_values + [ob_total]],
		index=["COUNT OF PLACEMENT", "SUM OF OB"],
		columns=columns,
	)

	placement_columns = REMARKS_PLACEMENT_ORDER
	remarks_counts = (
		df[df[remarks_col].isin(REMARK_FILTERS)]
		.groupby([remarks_col, placement_col], dropna=False)
		.size()
		.unstack(fill_value=0)
	)
	remarks_counts = remarks_counts.reindex(index=REMARK_FILTERS, columns=placement_columns, fill_value=0).T
	remarks_counts.index.name = "PLACEMENT"
	remarks_display = remarks_counts.astype(object).mask(remarks_counts.eq(0), "")

	return placement_summary, remarks_display


default_server_name = f"maya_merged_accounts_{date.today().strftime('%m%d%y')}"

with st.form("merged_summary_count_form"):
	st.caption(f"Server base path: {DEFAULT_SERVER_MERGED_DIR}")
	use_server_file = st.checkbox("Use server file path", value=True)

	server_filename = st.text_input(
		"Server file name or relative path",
		value=default_server_name,
		placeholder=default_server_name,
		disabled=not use_server_file,
		help="Example name: maya_merged_accounts_041626 or relative path like 2026\\APRIL\\maya_merged_accounts_041626.xlsx",
	)

	uploaded_file = st.file_uploader(
		"Fallback file upload",
		type=["xlsx", "xls", "csv"],
		help="Used when server file cannot be reached or not found.",
	)

	submitted = st.form_submit_button("Generate Output", use_container_width=True)

if submitted:
	progress_bar = st.progress(0, text="Starting generation...")
	selected_server_file = None
	if use_server_file:
		progress_bar.progress(20, text="Resolving server file...")
		selected_server_file, server_error = resolve_server_file(server_filename)
		if server_error:
			st.warning(server_error)

	try:
		progress_bar.progress(50, text="Loading input file...")
		df_input, source_label = read_dataframe_from_input(selected_server_file, uploaded_file)
		if df_input is None:
			progress_bar.empty()
			st.error("No input file available. Provide a valid server file name/path or upload a fallback file.")
		else:
			progress_bar.progress(80, text="Computing summaries...")
			st.success(source_label)
			placement_ob_output, remarks_output = build_outputs(df_input)

			st.subheader("Count of Placement and Sum of OB")
			st.dataframe(placement_ob_output, use_container_width=True)

			st.subheader("Remarks Counts (rows: placements in required order; columns: ACTIVE / FULLY PAID / REGULAR PULLED OUT)")
			st.dataframe(remarks_output, use_container_width=True)
			progress_bar.progress(100, text="Done.")
    
	except Exception as exc:
		progress_bar.empty()
		st.error(f"Processing failed: {exc}")
