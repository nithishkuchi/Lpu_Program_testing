

# # import glob
# # import os
# # import platform
# # import re
# # import subprocess
# # import threading
# # import sys
# # import pandas as pd
# # import streamlit as st
# # import xlrd

# # st.set_page_config(page_title="LPU Testing", layout="wide")
# # st.title("LPU Programme Testing")

# # INCOMING_FEE = "lpu_monitor/incoming/fee_excel"
# # INCOMING_PROGLIST = "lpu_monitor/incoming/programme_list"
# # INCOMING_PROGRAM_ID = "lpu_monitor/incoming/program_id"
# # INCOMING_IMPORTANT_DATES = "lpu_monitor/incoming/important_dates"
# # INCOMING_INTERNATIONAL_FEE = "lpu_monitor/incoming/international_fee_excel"

# # SETUP_ACTIONS = {
# #     "Update Programme Seed": {
# #         "steps": [("Rebuild programme seed", "lpu_monitor.config.rebuild_programme_seed")],
# #         "needs_files": ["ProgramId File"],
# #     },
# #     "Update Fee Row Mapping": {
# #         "steps": [("Build fee row mapping", "lpu_monitor.config.fee_row_mapping_builder")],
# #         "needs_files": ["Fee Excel", "Programme List"],
# #     },
# #     "Update Programme-Dates Mapping": {
# #         "steps": [("Build programme-dates mapping", "lpu_monitor.config.important_dates_mapping_builder")],
# #         "needs_files": ["Important Dates File", "Programme List"],
# #     },
# #     "Update International Fee (Excel + Seed)": {
# #     "steps": [
# #         ("Parse International Fee Excel", "lpu_monitor.config.international_fee_excel_parser"),
# #         ("Build International Fee Row Mapping", "lpu_monitor.config.international_fee_row_mapping_builder"),
# #         ("Rebuild International Seed", "lpu_monitor.config.rebuild_international_seed"),
# #         ("Build International Fee Expected", "lpu_monitor.config.build_international_fee_expected"),
# #     ],
# #     "needs_files": ["International Fee Excel", "Programme List"],
# #     },
# # }

# # VALIDATION_CHECKS = {
# #     "Scholarship / Annexure": {
# #         "steps": [
# #             ("Parse Annexure lookup", "lpu_monitor.config.annexure_lookup"),
# #             ("Run Annexure validation", "lpu_monitor.run_annexure_validation"),
# #         ],
# #         # These write with a bare relative filename (e.g. out_path="phd_fee_results.csv"),
# #         # which resolves against the process's cwd -- the tool_Script root -- not a
# #         # lpu_monitor/ subfolder. The old "lpu_monitor/..." paths here never matched
# #         # where the files actually land, which is why Results always said "No results
# #         # yet" even after a run completed successfully and the file existed on disk.
# #         "results_file": "annexure_results.csv",
# #         "needs_files": ["Programme List"],
# #     },
# #     "Fee (non-PhD)": {
# #         "steps": [
# #             ("Parse Fee Excel", "lpu_monitor.config.fee_excel_parser"),
# #             ("Build fee_expected.csv", "lpu_monitor.config.build_fee_expected"),
# #             ("Run Fee validation", "lpu_monitor.run_fee_batch"),
# #         ],
# #         "results_file": "fee_check_results.csv",
# #         "needs_files": ["Fee Excel", "Programme List"],
# #     },
# #     "Fee (PhD)": {
# #         "steps": [
# #             ("Parse PhD Fee sheet", "lpu_monitor.config.phd_fee_parser"),
# #             ("Run PhD Fee validation", "lpu_monitor.run_phd_fee_batch"),
# #         ],
# #         "results_file": "phd_fee_results.csv",
# #         "needs_files": ["Fee Excel"],
# #     },
# #     "Important Dates": {
# #         "steps": [("Run Important Dates validation", "lpu_monitor.run_important_dates_batch")],
# #         # run_important_dates_batch.py's OUT_PATH is the bare filename
# #         # "important_dates_results.csv" -- same cwd-relative convention as
# #         # the fix already applied to the other three checks above.
# #         "results_file": "important_dates_results.csv",
# #         "needs_files": ["Programme List"],
# #     },
# #     "International Exposure": {
# #         "steps": [("Run International Exposure validation", "lpu_monitor.run_international_exposure_batch")],
# #         "results_file": "international_exposure_results.csv",
# #         "needs_files": ["Programme List"],
# #     },
# #     "International Fee": {
# #         "steps": [("Run International Fee validation", "lpu_monitor.run_international_fee_batch")],
# #         "results_file": "international_fee_results.csv",
# #         "needs_files": ["International Fee Excel", "Programme List"],
# #     },
# # }

# # RECORDS_ACTIONS = {
# #     "LPUNEST Links": {
# #         "steps": [("Fetch LPUNEST Links", "lpu_monitor.run_lpunest_links_report")],
# #         "results_file": "records_lpunest_links.xlsx",
# #         "needs_files": ["Programme List"],
# #     },
# #     "Discipline Links": {
# #         "steps": [("Fetch Discipline Links", "lpu_monitor.run_discipline_links_report")],
# #         "results_file": "records_discipline_links.xlsx",
# #         "needs_files": ["Programme List"],
# #     },
# # }


# # def save_uploaded_file(uploaded_file, folder):
# #     os.makedirs(folder, exist_ok=True)
# #     for old_file in glob.glob(os.path.join(folder, "*")):
# #         os.remove(old_file)
# #     dest_path = os.path.join(folder, uploaded_file.name)
# #     with open(dest_path, "wb") as f:
# #         f.write(uploaded_file.getbuffer())
# #     return dest_path


# # # Header text isn't assumed to be one exact string -- checked against a few
# # # common variants, same spirit as other header lookups already used in this
# # # project (e.g. "programme code" / "nest test code").
# # _NAME_HEADER_CANDIDATES = ("programme name", "program name", "course name", "name of programme")


# # @st.cache_data(show_spinner=False)
# # def _load_programme_name_lookup_cached(path, mtime):
# #     """mtime is only here to bust Streamlit's cache automatically when a new
# #     Programme List file is uploaded (same path, new content, new mtime)."""
# #     try:
# #         wb = xlrd.open_workbook(path)
# #         sheet = wb.sheet_by_index(0)
# #         header = sheet.row_values(0)
# #         code_col = next((i for i, h in enumerate(header) if "programme code" in str(h).lower()), None)
# #         name_col = next(
# #             (i for i, h in enumerate(header) if str(h).strip().lower() in _NAME_HEADER_CANDIDATES),
# #             None,
# #         )
# #         if code_col is None or name_col is None:
# #             return {}
# #         lookup = {}
# #         for r in range(1, sheet.nrows):
# #             row = sheet.row_values(r)
# #             code = row[code_col]
# #             if code:
# #                 lookup[str(code).strip()] = str(row[name_col]).strip()
# #         return lookup
# #     except Exception:
# #         return {}


# # def get_programme_name_lookup():
# #     """Builds an official_code -> programme name lookup from whatever
# #     Programme List Excel is currently uploaded. Centralized here so EVERY
# #     check's results table gets programme names for free -- individual
# #     run_xxxx.py batch scripts don't need to know about this at all, and
# #     any new check added later automatically gets it too."""
# #     files = glob.glob(os.path.join(INCOMING_PROGLIST, "*.xls"))
# #     if not files:
# #         return {}
# #     path = files[0]
# #     return _load_programme_name_lookup_cached(path, os.path.getmtime(path))


# # # Old dict-string / bare-list result formats use these substrings to mean
# # # "nothing wrong here" -- kept for backward compatibility with checks still
# # # writing that format (Fee, Annexure).
# # _LEGACY_CLEAN_SUBSTRING_PATTERN = r"^\s*\[\]\s*$|'mismatches':\s*\[\]|SKIPPED"


# # def is_flagged_result(value):
# #     """Whether a result cell should count as a flagged row, in a way that
# #     works for EVERY check's output format, not just the one that happened
# #     to be built first.

# #     The bug this fixes: a BLANK/NaN/"None" cell (what newer checks like
# #     Important Dates correctly write when there's nothing to report) was
# #     being treated as flagged, because it never CONTAINS the old "clean"
# #     substring pattern -- str.contains() on an empty/NaN cell simply can't
# #     match anything, so the old "~contains(clean pattern)" logic defaulted
# #     to "flagged" for every single blank row. Checking for blank/None FIRST,
# #     before ever looking at the legacy pattern, fixes this for both old- and
# #     new-format checks at once, and will keep working for any future check
# #     that follows the same "blank means nothing to report" convention.
# #     """
# #     if pd.isna(value):
# #         return False
# #     text = str(value).strip()
# #     if text == "" or text.lower() == "none":
# #         return False
# #     if re.search(_LEGACY_CLEAN_SUBSTRING_PATTERN, text):
# #         return False
# #     return True


# # def kill_process_tree(pid):
# #     """Kills a process AND all its children (e.g. a Playwright-launched
# #     Chromium browser). On Windows, Popen.terminate() alone only kills the
# #     direct child, leaving orphaned grandchild processes (and file locks)
# #     behind -- that's what caused the earlier PermissionError."""
# #     if platform.system() == "Windows":
# #         subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)], capture_output=True)
# #     else:
# #         import signal
# #         try:
# #             os.killpg(os.getpgid(pid), signal.SIGKILL)
# #         except Exception:
# #             pass


# # def launch_step(module):
# #     """-u = unbuffered: without it, a child script's stdout is block-buffered
# #     (not flushed line by line) whenever it's writing to a pipe instead of a
# #     real terminal -- so even a "live" reader would only see output arrive in
# #     occasional bursts, not as it's actually printed. stderr is merged into
# #     stdout (STDOUT) so both show up in the one live stream in the order they
# #     were printed, same as watching it run in a terminal.

# #     A background thread reads line-by-line into `lines` as they arrive --
# #     that list is what the UI polls every second, instead of waiting for the
# #     whole step to finish before showing anything."""
# #     proc = subprocess.Popen(
# #         [sys.executable, "-u", "-m", module],
# #         stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1,
# #         cwd=os.path.dirname(os.path.abspath(__file__)),
# #     )
# #     lines = []

# #     def _reader():
# #         for line in proc.stdout:
# #             lines.append(line.rstrip("\n"))

# #     thread = threading.Thread(target=_reader, daemon=True)
# #     thread.start()
# #     return proc, lines, thread


# # def start_pipeline(name, steps, results_file):
# #     """Called once, when the Run button is clicked."""
# #     if results_file and os.path.exists(results_file):
# #         try:
# #             os.remove(results_file)
# #         except PermissionError:
# #             st.session_state.pipeline_error = (
# #                 f"Could not clear {results_file} -- it's locked by another process. "
# #                 "Check Task Manager for a leftover python.exe/chrome.exe, end it, then try again."
# #             )
# #             return

# #     proc, lines, thread = launch_step(steps[0][1])
# #     st.session_state.pipeline_error = None
# #     st.session_state.pipeline = {
# #         "name": name,
# #         "steps": steps,
# #         "results_file": results_file,
# #         "idx": 0,
# #         "proc": proc,
# #         "live_lines": lines,       # output of the CURRENTLY running step, growing live
# #         "reader_thread": thread,
# #         "log": [],                 # list of (label, full_text) for FINISHED steps
# #         "status": "running",  # running | done | failed
# #         "message": "",
# #     }


# # def stop_pipeline():
# #     pl = st.session_state.get("pipeline")
# #     if pl and pl.get("proc") is not None:
# #         kill_process_tree(pl["proc"].pid)
# #         st.session_state.stopped_message = f"Stopped: {pl['name']}"
# #     st.session_state.pipeline = None


# # def force_full_rerun():
# #     """A rerun triggered from inside a @st.fragment only re-executes that
# #     fragment by default -- it does NOT re-run the rest of app.py. That was
# #     the actual bug behind "results generated on disk but page still says
# #     No results yet": the Results table, the success/error banner, and the
# #     Run button's disabled state all live in the OUTER script, which was
# #     never being told to look again. scope="app" forces a real full-page
# #     rerun instead of a fragment-only one. (Older Streamlit versions that
# #     don't support the scope= argument fall back to a plain rerun.)"""
# #     try:
# #         st.rerun(scope="app")
# #     except TypeError:
# #         st.rerun()


# # @st.fragment(run_every=1)
# # def pipeline_status_fragment():
# #     pl = st.session_state.get("pipeline")
# #     if not pl or pl["status"] != "running":
# #         return

# #     steps = pl["steps"]
# #     idx = pl["idx"]
# #     label, _module = steps[idx]

# #     st.progress(idx / len(steps), text=f"Running: {label}")
# #     # Only ONE button is created per rerun of this fragment -- no duplicate keys,
# #     # and because this is a real Streamlit rerun (triggered by run_every), a
# #     # click on it is actually detected.
# #     st.button("Stop this check", key="stop_running_check", on_click=stop_pipeline)

# #     # Live terminal output for the CURRENTLY running step -- re-rendered from
# #     # pl["live_lines"] on every 1s tick, so it fills in as the child script
# #     # actually prints, the same as watching it in a terminal, instead of
# #     # appearing all at once only after the step finishes.
# #     st.caption(f"Live output -- {label}")
# #     st.code("\n".join(pl["live_lines"]) or "(waiting for output...)", language=None)

# #     proc = pl["proc"]
# #     if proc.poll() is None:
# #         return  # still running -- check again on the next 1s tick

# #     # Step just finished: give the reader thread a moment to drain any last
# #     # buffered lines, then freeze this step's output into the permanent log.
# #     pl["reader_thread"].join(timeout=2)
# #     pl["log"].append((label, "\n".join(pl["live_lines"])))

# #     if proc.returncode != 0:
# #         pl["status"] = "failed"
# #         pl["message"] = f"{label} failed -- see log below."
# #         force_full_rerun()
# #         return

# #     next_idx = idx + 1
# #     if next_idx >= len(steps):
# #         pl["status"] = "done"
# #         pl["message"] = "Completed successfully."
# #         force_full_rerun()
# #         return

# #     next_label, next_module = steps[next_idx]
# #     next_proc, next_lines, next_thread = launch_step(next_module)
# #     pl["idx"] = next_idx
# #     pl["proc"] = next_proc
# #     pl["live_lines"] = next_lines
# #     pl["reader_thread"] = next_thread


# # # ---------- Shared file uploads ----------
# # st.header("Source files")
# # col1, col2, col3, col4 = st.columns(4)
# # with col1:
# #     fee_file = st.file_uploader("Fee & Scholarship Excel (.xls)", type=["xls"], key="fee")
# #     if fee_file:
# #         st.success(f"Saved to {save_uploaded_file(fee_file, INCOMING_FEE)}")
# # with col2:
# #     proglist_file = st.file_uploader("Final Programme List Excel (.xls)", type=["xls"], key="proglist")
# #     if proglist_file:
# #         st.success(f"Saved to {save_uploaded_file(proglist_file, INCOMING_PROGLIST)}")
# # with col3:
# #     programid_file = st.file_uploader("ProgramId Source File (.xlsx)", type=["xlsx"], key="programid")
# #     if programid_file:
# #         st.success(f"Saved to {save_uploaded_file(programid_file, INCOMING_PROGRAM_ID)}")
# # with col4:
# #     dates_file = st.file_uploader("Important Dates Schedule (.xlsx)", type=["xlsx"], key="dates")
# #     if dates_file:
# #         st.success(f"Saved to {save_uploaded_file(dates_file, INCOMING_IMPORTANT_DATES)}")
# # intl_fee_file = st.file_uploader("International Fee Excel (.xls)", type=["xls"], key="intlfee")
# # if intl_fee_file:
# #     st.success(f"Saved to {save_uploaded_file(intl_fee_file, INCOMING_INTERNATIONAL_FEE)}")

# # st.divider()

# # # ---------- Section picker ----------
# # st.header("Choose a section")
# # section = st.radio("Section", ["Setup / Data Refresh", "Validation", "Records"], horizontal=True, label_visibility="collapsed")
# # items = SETUP_ACTIONS if section == "Setup / Data Refresh" else (VALIDATION_CHECKS if section == "Validation" else RECORDS_ACTIONS)

# # st.header(f"Choose: {section}")
# # selected = st.selectbox("Action/Check", list(items.keys()), label_visibility="collapsed")
# # item = items[selected]
# # st.caption(f"Needs: {', '.join(item['needs_files'])}")

# # pipeline_running = st.session_state.get("pipeline") is not None

# # if st.button(f"Run: {selected}", type="primary", disabled=pipeline_running):
# #     start_pipeline(selected, item["steps"], item.get("results_file"))
# #     st.rerun()

# # if pipeline_running:
# #     st.caption("A check is already running -- stop it below before starting another.")

# # if st.session_state.get("pipeline_error"):
# #     st.error(st.session_state.pipeline_error)

# # if st.session_state.get("stopped_message"):
# #     st.warning(st.session_state.stopped_message)
# #     st.session_state.stopped_message = None

# # pl = st.session_state.get("pipeline")
# # if pl:
# #     pipeline_status_fragment()
# #     if pl["status"] in ("done", "failed"):
# #         if pl["status"] == "done":
# #             st.success(pl["message"])
# #         else:
# #             st.error(pl["message"])
# #         with st.expander("Log", expanded=(pl["status"] == "failed")):
# #             for label, full_text in pl["log"]:
# #                 st.text(f"--- {label} ---\n{full_text}")
# #         if st.button("Clear"):
# #             st.session_state.pipeline = None
# #             st.rerun()

# # st.divider()

# # # ---------- Results ----------
# # if section in ("Validation", "Records"):
# #     st.header(f"Results: {selected}")
# #     path = item["results_file"]

# #     # Don't offer the download while THIS check's own run is still writing to
# #     # that exact file -- only once it's fully generated. A finished run of a
# #     # different check, or an old leftover file from an earlier run, is fine
# #     # to download; a file this check is actively appending to right now is not.
# #     pl = st.session_state.get("pipeline")
# #     currently_writing_this_file = (
# #         pl is not None and pl.get("results_file") == path and pl.get("status") == "running"
# #     )

# #     if os.path.exists(path) and not currently_writing_this_file:
# #         with open(path, "rb") as f:
# #             st.download_button(
# #                 "Download " + ("Excel" if path.endswith(".xlsx") else "CSV"),
# #                 data=f.read(),
# #                 file_name=os.path.basename(path),
# #                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if path.endswith(".xlsx") else "text/csv",
# #             )
# #         if path.endswith(".xlsx"):
# #             df = pd.read_excel(path)
# #             st.dataframe(df, use_container_width=True)
# #         else:
# #             df = pd.read_csv(path, header=None, names=["official_code", "result"])

# #             name_lookup = get_programme_name_lookup()
# #             if name_lookup:
# #                 df.insert(
# #                     1, "programme_name",
# #                     df["official_code"].astype(str).str.strip().map(name_lookup).fillna(""),
# #                 )

# #             mismatch_mask = df["result"].apply(is_flagged_result)
# #             show_mismatches_only = st.checkbox("Show only mismatches/errors", value=True)
# #             if show_mismatches_only:
# #                 filtered = df[mismatch_mask]
# #                 st.write(f"{len(filtered)} of {len(df)} rows flagged")
# #                 st.dataframe(filtered, use_container_width=True)
# #             else:
# #                 st.dataframe(df, use_container_width=True)
# #     elif currently_writing_this_file:
# #         st.info("This check is still running -- results will appear here once it finishes.")
# #     else:
# #         st.info("No results yet for this check -- run it above.")
# import glob
# import os
# import platform
# import re
# import subprocess
# import threading
# import sys
# import pandas as pd
# import streamlit as st
# import xlrd

# st.set_page_config(page_title="LPU Testing", layout="wide")
# st.title("LPU Programme Testing")

# INCOMING_FEE = "lpu_monitor/incoming/fee_excel"
# INCOMING_PROGLIST = "lpu_monitor/incoming/programme_list"
# INCOMING_PROGRAM_ID = "lpu_monitor/incoming/program_id"
# INCOMING_IMPORTANT_DATES = "lpu_monitor/incoming/important_dates"
# INCOMING_INTERNATIONAL_FEE = "lpu_monitor/incoming/international_fee_excel"

# SETUP_ACTIONS = {
#     "Update Programme Seed": {
#         "steps": [("Rebuild programme seed", "lpu_monitor.config.rebuild_programme_seed")],
#         "needs_files": ["ProgramId File"],
#     },
#     "Update Fee Row Mapping": {
#         "steps": [("Build fee row mapping", "lpu_monitor.config.fee_row_mapping_builder")],
#         "needs_files": ["Fee Excel", "Programme List"],
#     },
#     "Update Programme-Dates Mapping": {
#         "steps": [("Build programme-dates mapping", "lpu_monitor.config.important_dates_mapping_builder")],
#         "needs_files": ["Important Dates File", "Programme List"],
#     },
#     "Update International Fee (Excel + Seed)": {
#     "steps": [
#         ("Parse International Fee Excel", "lpu_monitor.config.international_fee_excel_parser"),
#         ("Build International Fee Row Mapping", "lpu_monitor.config.international_fee_row_mapping_builder"),
#         ("Rebuild International Seed", "lpu_monitor.config.rebuild_international_seed"),
#         ("Build International Fee Expected", "lpu_monitor.config.build_international_fee_expected"),
#     ],
#     "needs_files": ["International Fee Excel", "Programme List"],
#     },
# }

# VALIDATION_CHECKS = {
#     "Scholarship / Annexure": {
#         "steps": [
#             ("Parse Annexure lookup", "lpu_monitor.config.annexure_lookup"),
#             ("Run Annexure validation", "lpu_monitor.run_annexure_validation"),
#         ],
#         # These write with a bare relative filename (e.g. out_path="phd_fee_results.csv"),
#         # which resolves against the process's cwd -- the tool_Script root -- not a
#         # lpu_monitor/ subfolder. The old "lpu_monitor/..." paths here never matched
#         # where the files actually land, which is why Results always said "No results
#         # yet" even after a run completed successfully and the file existed on disk.
#         "results_file": "annexure_results.csv",
#         "needs_files": ["Programme List"],
#     },
#     "Fee (non-PhD)": {
#         "steps": [
#             ("Parse Fee Excel", "lpu_monitor.config.fee_excel_parser"),
#             ("Build fee_expected.csv", "lpu_monitor.config.build_fee_expected"),
#             ("Run Fee validation", "lpu_monitor.run_fee_batch"),
#         ],
#         "results_file": "fee_check_results.csv",
#         "needs_files": ["Fee Excel", "Programme List"],
#     },
#     "Fee (PhD)": {
#         "steps": [
#             ("Parse PhD Fee sheet", "lpu_monitor.config.phd_fee_parser"),
#             ("Run PhD Fee validation", "lpu_monitor.run_phd_fee_batch"),
#         ],
#         "results_file": "phd_fee_results.csv",
#         "needs_files": ["Fee Excel"],
#     },
#     "Important Dates": {
#         "steps": [("Run Important Dates validation", "lpu_monitor.run_important_dates_batch")],
#         # run_important_dates_batch.py's OUT_PATH is the bare filename
#         # "important_dates_results.csv" -- same cwd-relative convention as
#         # the fix already applied to the other three checks above.
#         "results_file": "important_dates_results.csv",
#         "needs_files": ["Programme List"],
#     },
#     "International Exposure": {
#         "steps": [("Run International Exposure validation", "lpu_monitor.run_international_exposure_batch")],
#         "results_file": "international_exposure_results.csv",
#         "needs_files": ["Programme List"],
#     },
#     "International Fee": {
#         "steps": [("Run International Fee validation", "lpu_monitor.run_international_fee_batch")],
#         "results_file": "international_fee_results.csv",
#         "needs_files": ["International Fee Excel", "Programme List"],
#     },
# }

# RECORDS_ACTIONS = {
#     "LPUNEST Links": {
#         "steps": [("Fetch LPUNEST Links", "lpu_monitor.run_lpunest_links_report")],
#         "results_file": "records_lpunest_links.xlsx",
#         "needs_files": ["Programme List"],
#     },
#     "Discipline Links": {
#         "steps": [("Fetch Discipline Links", "lpu_monitor.run_discipline_links_report")],
#         "results_file": "records_discipline_links.xlsx",
#         "needs_files": ["Programme List"],
#     },
# }


# def save_uploaded_file(uploaded_file, folder):
#     os.makedirs(folder, exist_ok=True)
#     for old_file in glob.glob(os.path.join(folder, "*")):
#         os.remove(old_file)
#     dest_path = os.path.join(folder, uploaded_file.name)
#     with open(dest_path, "wb") as f:
#         f.write(uploaded_file.getbuffer())
#     return dest_path


# # Header text isn't assumed to be one exact string -- checked against a few
# # common variants, same spirit as other header lookups already used in this
# # project (e.g. "programme code" / "nest test code").
# _NAME_HEADER_CANDIDATES = ("programme name", "program name", "course name", "name of programme")


# @st.cache_data(show_spinner=False)
# def _load_programme_name_lookup_cached(path, mtime):
#     """mtime is only here to bust Streamlit's cache automatically when a new
#     Programme List file is uploaded (same path, new content, new mtime)."""
#     try:
#         wb = xlrd.open_workbook(path)
#         sheet = wb.sheet_by_index(0)
#         header = sheet.row_values(0)
#         code_col = next((i for i, h in enumerate(header) if "programme code" in str(h).lower()), None)
#         name_col = next(
#             (i for i, h in enumerate(header) if str(h).strip().lower() in _NAME_HEADER_CANDIDATES),
#             None,
#         )
#         if code_col is None or name_col is None:
#             return {}
#         lookup = {}
#         for r in range(1, sheet.nrows):
#             row = sheet.row_values(r)
#             code = row[code_col]
#             if code:
#                 lookup[str(code).strip()] = str(row[name_col]).strip()
#         return lookup
#     except Exception:
#         return {}


# def get_programme_name_lookup():
#     """Builds an official_code -> programme name lookup from whatever
#     Programme List Excel is currently uploaded. Centralized here so EVERY
#     check's results table gets programme names for free -- individual
#     run_xxxx.py batch scripts don't need to know about this at all, and
#     any new check added later automatically gets it too."""
#     files = glob.glob(os.path.join(INCOMING_PROGLIST, "*.xls"))
#     if not files:
#         return {}
#     path = files[0]
#     return _load_programme_name_lookup_cached(path, os.path.getmtime(path))


# # Old dict-string / bare-list result formats use these substrings to mean
# # "nothing wrong here" -- kept for backward compatibility with checks still
# # writing that format (Fee, Annexure).
# _LEGACY_CLEAN_SUBSTRING_PATTERN = r"^\s*\[\]\s*$|'mismatches':\s*\[\]|SKIPPED"


# def is_flagged_result(value):
#     """Whether a result cell should count as a flagged row, in a way that
#     works for EVERY check's output format, not just the one that happened
#     to be built first.

#     The bug this fixes: a BLANK/NaN/"None" cell (what newer checks like
#     Important Dates correctly write when there's nothing to report) was
#     being treated as flagged, because it never CONTAINS the old "clean"
#     substring pattern -- str.contains() on an empty/NaN cell simply can't
#     match anything, so the old "~contains(clean pattern)" logic defaulted
#     to "flagged" for every single blank row. Checking for blank/None FIRST,
#     before ever looking at the legacy pattern, fixes this for both old- and
#     new-format checks at once, and will keep working for any future check
#     that follows the same "blank means nothing to report" convention.
#     """
#     if pd.isna(value):
#         return False
#     text = str(value).strip()
#     if text == "" or text.lower() == "none":
#         return False
#     if re.search(_LEGACY_CLEAN_SUBSTRING_PATTERN, text):
#         return False
#     return True


# def kill_process_tree(pid):
#     """Kills a process AND all its children (e.g. a Playwright-launched
#     Chromium browser). On Windows, Popen.terminate() alone only kills the
#     direct child, leaving orphaned grandchild processes (and file locks)
#     behind -- that's what caused the earlier PermissionError."""
#     if platform.system() == "Windows":
#         subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)], capture_output=True)
#     else:
#         import signal
#         try:
#             os.killpg(os.getpgid(pid), signal.SIGKILL)
#         except Exception:
#             pass


# def launch_step(module):
#     """-u = unbuffered: without it, a child script's stdout is block-buffered
#     (not flushed line by line) whenever it's writing to a pipe instead of a
#     real terminal -- so even a "live" reader would only see output arrive in
#     occasional bursts, not as it's actually printed. stderr is merged into
#     stdout (STDOUT) so both show up in the one live stream in the order they
#     were printed, same as watching it run in a terminal.

#     A background thread reads line-by-line into `lines` as they arrive --
#     that list is what the UI polls every second, instead of waiting for the
#     whole step to finish before showing anything."""
#     proc = subprocess.Popen(
#         [sys.executable, "-u", "-m", module],
#         stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1,
#         cwd=os.path.dirname(os.path.abspath(__file__)),
#     )
#     lines = []

#     def _reader():
#         for line in proc.stdout:
#             lines.append(line.rstrip("\n"))

#     thread = threading.Thread(target=_reader, daemon=True)
#     thread.start()
#     return proc, lines, thread


# def start_pipeline(name, steps, results_file):
#     """Called once, when the Run button is clicked."""
#     if results_file and os.path.exists(results_file):
#         try:
#             os.remove(results_file)
#         except PermissionError:
#             st.session_state.pipeline_error = (
#                 f"Could not clear {results_file} -- it's locked by another process. "
#                 "Check Task Manager for a leftover python.exe/chrome.exe, end it, then try again."
#             )
#             return

#     proc, lines, thread = launch_step(steps[0][1])
#     st.session_state.pipeline_error = None
#     st.session_state.pipeline = {
#         "name": name,
#         "steps": steps,
#         "results_file": results_file,
#         "idx": 0,
#         "proc": proc,
#         "live_lines": lines,       # output of the CURRENTLY running step, growing live
#         "reader_thread": thread,
#         "log": [],                 # list of (label, full_text) for FINISHED steps
#         "status": "running",  # running | done | failed
#         "message": "",
#     }


# def stop_pipeline():
#     pl = st.session_state.get("pipeline")
#     if pl and pl.get("proc") is not None:
#         kill_process_tree(pl["proc"].pid)
#         st.session_state.stopped_message = f"Stopped: {pl['name']}"
#     st.session_state.pipeline = None


# def force_full_rerun():
#     """A rerun triggered from inside a @st.fragment only re-executes that
#     fragment by default -- it does NOT re-run the rest of app.py. That was
#     the actual bug behind "results generated on disk but page still says
#     No results yet": the Results table, the success/error banner, and the
#     Run button's disabled state all live in the OUTER script, which was
#     never being told to look again. scope="app" forces a real full-page
#     rerun instead of a fragment-only one. (Older Streamlit versions that
#     don't support the scope= argument fall back to a plain rerun.)"""
#     try:
#         st.rerun(scope="app")
#     except TypeError:
#         st.rerun()


# @st.fragment(run_every=1)
# def pipeline_status_fragment():
#     pl = st.session_state.get("pipeline")
#     if not pl or pl["status"] != "running":
#         return

#     steps = pl["steps"]
#     idx = pl["idx"]
#     label, _module = steps[idx]

#     st.progress(idx / len(steps), text=f"Running: {label}")
#     # Only ONE button is created per rerun of this fragment -- no duplicate keys,
#     # and because this is a real Streamlit rerun (triggered by run_every), a
#     # click on it is actually detected.
#     st.button("Stop this check", key="stop_running_check", on_click=stop_pipeline)

#     # Live terminal output for the CURRENTLY running step -- re-rendered from
#     # pl["live_lines"] on every 1s tick, so it fills in as the child script
#     # actually prints, the same as watching it in a terminal, instead of
#     # appearing all at once only after the step finishes.
#     st.caption(f"Live output -- {label}")
#     st.code("\n".join(pl["live_lines"]) or "(waiting for output...)", language=None)

#     proc = pl["proc"]
#     if proc.poll() is None:
#         return  # still running -- check again on the next 1s tick

#     # Step just finished: give the reader thread a moment to drain any last
#     # buffered lines, then freeze this step's output into the permanent log.
#     pl["reader_thread"].join(timeout=2)
#     pl["log"].append((label, "\n".join(pl["live_lines"])))

#     if proc.returncode != 0:
#         pl["status"] = "failed"
#         pl["message"] = f"{label} failed -- see log below."
#         force_full_rerun()
#         return

#     next_idx = idx + 1
#     if next_idx >= len(steps):
#         pl["status"] = "done"
#         pl["message"] = "Completed successfully."
#         force_full_rerun()
#         return

#     next_label, next_module = steps[next_idx]
#     next_proc, next_lines, next_thread = launch_step(next_module)
#     pl["idx"] = next_idx
#     pl["proc"] = next_proc
#     pl["live_lines"] = next_lines
#     pl["reader_thread"] = next_thread


# # ---------- Shared file uploads ----------
# st.header("Source files")
# col1, col2, col3, col4 = st.columns(4)
# with col1:
#     fee_file = st.file_uploader("Fee & Scholarship Excel (.xls)", type=["xls"], key="fee")
#     if fee_file:
#         st.success(f"Saved to {save_uploaded_file(fee_file, INCOMING_FEE)}")
# with col2:
#     proglist_file = st.file_uploader("Final Programme List Excel (.xls)", type=["xls"], key="proglist")
#     if proglist_file:
#         st.success(f"Saved to {save_uploaded_file(proglist_file, INCOMING_PROGLIST)}")
# with col3:
#     programid_file = st.file_uploader("ProgramId Source File (.xlsx)", type=["xlsx"], key="programid")
#     if programid_file:
#         st.success(f"Saved to {save_uploaded_file(programid_file, INCOMING_PROGRAM_ID)}")
# with col4:
#     dates_file = st.file_uploader("Important Dates Schedule (.xlsx)", type=["xlsx"], key="dates")
#     if dates_file:
#         st.success(f"Saved to {save_uploaded_file(dates_file, INCOMING_IMPORTANT_DATES)}")
# intl_fee_file = st.file_uploader("International Fee Excel (.xls)", type=["xls"], key="intlfee")
# if intl_fee_file:
#     st.success(f"Saved to {save_uploaded_file(intl_fee_file, INCOMING_INTERNATIONAL_FEE)}")

# st.divider()

# # ---------- Section picker ----------
# st.header("Choose a section")
# section = st.radio(
#     "Section",
#     ["Setup / Data Refresh", "Validation", "Records", "Previous Results"],
#     horizontal=True, label_visibility="collapsed",
# )

# if section == "Previous Results":
#     st.header("Previous Results")
    

#     # Non-recursive on purpose: only the tool_Script root, not lpu_monitor/,
#     # data/, python_embedded/, etc. -- those are internal working files, not
#     # results a person would want to download.
#     csv_files = sorted(glob.glob("*.csv"))

#     if not csv_files:
#         st.info("No CSV files in the project folder yet -- run a check first.")
#     else:
#         for fname in csv_files:
#             col_name, col_button = st.columns([3, 1])
#             with col_name:
#                 st.write(fname)
#             with col_button:
#                 with open(fname, "rb") as f:
#                     st.download_button(
#                         "Download",
#                         data=f.read(),
#                         file_name=fname,
#                         mime="text/csv",
#                         key=f"prev_dl_{fname}",
#                     )
# else:
#     items = SETUP_ACTIONS if section == "Setup / Data Refresh" else (VALIDATION_CHECKS if section == "Validation" else RECORDS_ACTIONS)

#     st.header(f"Choose: {section}")
#     selected = st.selectbox("Action/Check", list(items.keys()), label_visibility="collapsed")
#     item = items[selected]
#     st.caption(f"Needs: {', '.join(item['needs_files'])}")

#     pipeline_running = st.session_state.get("pipeline") is not None

#     if st.button(f"Run: {selected}", type="primary", disabled=pipeline_running):
#         start_pipeline(selected, item["steps"], item.get("results_file"))
#         st.rerun()

#     if pipeline_running:
#         st.caption("A check is already running -- stop it below before starting another.")

#     if st.session_state.get("pipeline_error"):
#         st.error(st.session_state.pipeline_error)

#     if st.session_state.get("stopped_message"):
#         st.warning(st.session_state.stopped_message)
#         st.session_state.stopped_message = None

#     pl = st.session_state.get("pipeline")
#     if pl:
#         pipeline_status_fragment()
#         if pl["status"] in ("done", "failed"):
#             if pl["status"] == "done":
#                 st.success(pl["message"])
#             else:
#                 st.error(pl["message"])
#             with st.expander("Log", expanded=(pl["status"] == "failed")):
#                 for label, full_text in pl["log"]:
#                     st.text(f"--- {label} ---\n{full_text}")
#             if st.button("Clear"):
#                 st.session_state.pipeline = None
#                 st.rerun()

#     st.divider()

#     # ---------- Results ----------
#     if section in ("Validation", "Records"):
#         st.header(f"Results: {selected}")
#         path = item["results_file"]

#         # Don't offer the download while THIS check's own run is still writing to
#         # that exact file -- only once it's fully generated. A finished run of a
#         # different check, or an old leftover file from an earlier run, is fine
#         # to download; a file this check is actively appending to right now is not.
#         pl = st.session_state.get("pipeline")
#         currently_writing_this_file = (
#             pl is not None and pl.get("results_file") == path and pl.get("status") == "running"
#         )

#         if os.path.exists(path) and not currently_writing_this_file:
#             with open(path, "rb") as f:
#                 st.download_button(
#                     "Download " + ("Excel" if path.endswith(".xlsx") else "CSV"),
#                     data=f.read(),
#                     file_name=os.path.basename(path),
#                     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if path.endswith(".xlsx") else "text/csv",
#                 )
#             if path.endswith(".xlsx"):
#                 df = pd.read_excel(path)
#                 st.dataframe(df, use_container_width=True)
#             else:
#                 df = pd.read_csv(path, header=None, names=["official_code", "result"])

#                 name_lookup = get_programme_name_lookup()
#                 if name_lookup:
#                     df.insert(
#                         1, "programme_name",
#                         df["official_code"].astype(str).str.strip().map(name_lookup).fillna(""),
#                     )

#                 mismatch_mask = df["result"].apply(is_flagged_result)
#                 show_mismatches_only = st.checkbox("Show only mismatches/errors", value=True)
#                 if show_mismatches_only:
#                     filtered = df[mismatch_mask]
#                     st.write(f"{len(filtered)} of {len(df)} rows flagged")
#                     st.dataframe(filtered, use_container_width=True)
#                 else:
#                     st.dataframe(df, use_container_width=True)
#         elif currently_writing_this_file:
#             st.info("This check is still running -- results will appear here once it finishes.")
#         else:
#             st.info("No results yet for this check -- run it above.")
import glob
import os
import platform
import re
import subprocess
import threading
import sys
import pandas as pd
import streamlit as st
import xlrd

st.set_page_config(page_title="LPU Testing", layout="wide")
st.title("LPU Programme Testing")

INCOMING_FEE = "lpu_monitor/incoming/fee_excel"
INCOMING_PROGLIST = "lpu_monitor/incoming/programme_list"
INCOMING_PROGRAM_ID = "lpu_monitor/incoming/program_id"
INCOMING_IMPORTANT_DATES = "lpu_monitor/incoming/important_dates"
INCOMING_ELIGIBILITY = "lpu_monitor/incoming/eligibility_excel"          
INCOMING_INTERNATIONAL_FEE = "lpu_monitor/incoming/international_fee_excel"

SETUP_ACTIONS = {
    "Update Programme Seed": {
        "steps": [("Rebuild programme seed", "lpu_monitor.config.rebuild_programme_seed")],
        "needs_files": ["ProgramId File"],
    },
    "Update Fee Row Mapping": {
        "steps": [("Build fee row mapping", "lpu_monitor.config.fee_row_mapping_builder")],
        "needs_files": ["Fee Excel", "Programme List"],
    },
    "Update Programme-Dates Mapping": {
        "steps": [("Build programme-dates mapping", "lpu_monitor.config.important_dates_mapping_builder")],
        "needs_files": ["Important Dates File", "Programme List"],
    },
        "Update International Fee (Excel + Seed)": {
    "steps": [
        ("Parse International Fee Excel", "lpu_monitor.config.international_fee_excel_parser"),
        ("Build International Fee Row Mapping", "lpu_monitor.config.international_fee_row_mapping_builder"),
        ("Rebuild International Seed", "lpu_monitor.config.rebuild_international_seed"),
        ("Build International Fee Expected", "lpu_monitor.config.build_international_fee_expected"),
    ],
    "needs_files": ["International Fee Excel", "Programme List"],
    },
    # ─────────────────────── ADD THIS BLOCK ───────────────────────
    "Update Eligibility Data": {
        "steps": [
            ("Parse Eligibility Excel", "lpu_monitor.config.eligibility_excel_parser"),
            ("Build Eligibility Row Mapping", "lpu_monitor.config.generate_programme_eligibility_mapping"),
            ("Build Eligibility Expected", "lpu_monitor.config.build_eligibility_expected"),
        ],
        "needs_files": ["Eligibility Excel", "Programme List"],
    },
    # ─────────────────────── END ADD ──────────────────────────────

}

VALIDATION_CHECKS = {
    "Scholarship / Annexure": {
        "steps": [
            ("Parse Annexure lookup", "lpu_monitor.config.annexure_lookup"),
            ("Run Annexure validation", "lpu_monitor.run_annexure_validation"),
        ],
        # These write with a bare relative filename (e.g. out_path="phd_fee_results.csv"),
        # which resolves against the process's cwd -- the tool_Script root -- not a
        # lpu_monitor/ subfolder. The old "lpu_monitor/..." paths here never matched
        # where the files actually land, which is why Results always said "No results
        # yet" even after a run completed successfully and the file existed on disk.
        "results_file": "annexure_results.csv",
        "needs_files": ["Programme List"],
    },
    "Fee (non-PhD)": {
        "steps": [
            ("Parse Fee Excel", "lpu_monitor.config.fee_excel_parser"),
            ("Build fee_expected.csv", "lpu_monitor.config.build_fee_expected"),
            ("Run Fee validation", "lpu_monitor.run_fee_batch"),
        ],
        "results_file": "fee_check_results.csv",
        "needs_files": ["Fee Excel", "Programme List"],
    },
    "Fee (PhD)": {
        "steps": [
            ("Parse PhD Fee sheet", "lpu_monitor.config.phd_fee_parser"),
            ("Run PhD Fee validation", "lpu_monitor.run_phd_fee_batch"),
        ],
        "results_file": "phd_fee_results.csv",
        "needs_files": ["Fee Excel"],
    },
    "Important Dates": {
        "steps": [("Run Important Dates validation", "lpu_monitor.run_important_dates_batch")],
        # run_important_dates_batch.py's OUT_PATH is the bare filename
        # "important_dates_results.csv" -- same cwd-relative convention as
        # the fix already applied to the other three checks above.
        "results_file": "important_dates_results.csv",
        "needs_files": ["Programme List"],
    },
    "International Exposure": {
        "steps": [("Run International Exposure validation", "lpu_monitor.run_international_exposure_batch")],
        "results_file": "international_exposure_results.csv",
        "needs_files": ["Programme List"],
    },
        "International Fee": {
        "steps": [("Run International Fee validation", "lpu_monitor.run_international_fee_batch")],
        "results_file": "international_fee_results.csv",
        "needs_files": ["International Fee Excel", "Programme List"],
    },
    # ─────────────────────── ADD THIS BLOCK ───────────────────────
    "Eligibility": {
        "steps": [("Run Eligibility validation", "lpu_monitor.run_eligibility_batch")],
        "results_file": "eligibility_check_results.csv",
        "needs_files": ["Eligibility Excel", "Programme List"],
    },
    # ─────────────────────── END ADD ──────────────────────────────

}

RECORDS_ACTIONS = {
    "Fee (non-PhD)": {
        "steps": [("Fetch Live Fee Records", "lpu_monitor.run_fee_report")],
        "results_file": "records_fee.xlsx",
        "needs_files": ["Programme List"],
    },
    "Fee (PhD)": {
        "steps": [("Fetch Live PhD Fee Records", "lpu_monitor.run_phd_fee_report")],
        "results_file": "records_phd_fee.xlsx",
        "needs_files": ["Programme List"],
    },
    "Important Dates": {
        "steps": [("Fetch Live Important Dates Records", "lpu_monitor.run_important_dates_report")],
        "results_file": "records_important_dates.xlsx",
        "needs_files": ["Programme List"],
    },
    "Eligibility": {
        "steps": [("Fetch Live Eligibility Records", "lpu_monitor.run_eligibility_report")],
        "results_file": "records_eligibility.xlsx",
        "needs_files": ["Programme List"],
    },
    "International Exposure": {
        "steps": [("Fetch Live International Exposure Records", "lpu_monitor.run_international_exposure_report")],
        "results_file": "records_international_exposure.xlsx",
        "needs_files": ["Programme List"],
    },
    "International Fee": {
        "steps": [("Fetch Live International Fee Records", "lpu_monitor.run_international_fee_report")],
        "results_file": "records_international_fee.xlsx",
        "needs_files": ["International Fee Excel", "Programme List"],
    },
    "Scholarship / Annexure": {
        "steps": [("Fetch Live Scholarship/Annexure Records", "lpu_monitor.run_annexure_report")],
        "results_file": "records_annexure.xlsx",
        "needs_files": ["Programme List"],
    },
    "LPUNEST Links": {
        "steps": [("Fetch LPUNEST Links", "lpu_monitor.run_lpunest_links_report")],
        "results_file": "records_lpunest_links.xlsx",
        "needs_files": ["Programme List"],
    },
    "Discipline Links": {
        "steps": [("Fetch Discipline Links", "lpu_monitor.run_discipline_links_report")],
        "results_file": "records_discipline_links.xlsx",
        "needs_files": ["Programme List"],
    },
    # ─────────────────────── ADD THESE FOUR ───────────────────────
    "Placement Highlights": {
        "steps": [("Build Placement Records", "lpu_monitor.run_placement_report")],
        "results_file": "records_placement.xlsx",
        "needs_files": [],  # No file upload needed — fetches from API
    },
    "Research": {
        "steps": [("Build Research Records", "lpu_monitor.run_research_report")],
        "results_file": "records_research.xlsx",
        "needs_files": [],  # No file upload needed — fetches from API
    },
    "Rankings": {
        "steps": [("Build Rankings Records", "lpu_monitor.run_rankings_report")],
        "results_file": "records_rankings.xlsx",
        "needs_files": [],  # No file upload needed — fetches from API
    },
    "Student Achievements": {
        "steps": [("Build Student Achievements", "lpu_monitor.run_student_achievements_report")],
        "results_file": "records_student_achievements.xlsx",
        "needs_files": [],  # No file upload needed — fetches from API
    },
    # ─────────────────────── END ADD ──────────────────────────────
}


def save_uploaded_file(uploaded_file, folder):
    os.makedirs(folder, exist_ok=True)
    for old_file in glob.glob(os.path.join(folder, "*")):
        os.remove(old_file)
    dest_path = os.path.join(folder, uploaded_file.name)
    with open(dest_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return dest_path


# Header text isn't assumed to be one exact string -- checked against a few
# common variants, same spirit as other header lookups already used in this
# project (e.g. "programme code" / "nest test code").
_NAME_HEADER_CANDIDATES = ("programme name", "program name", "course name", "name of programme")


@st.cache_data(show_spinner=False)
def _load_programme_name_lookup_cached(path, mtime):
    """mtime is only here to bust Streamlit's cache automatically when a new
    Programme List file is uploaded (same path, new content, new mtime)."""
    try:
        wb = xlrd.open_workbook(path)
        sheet = wb.sheet_by_index(0)
        header = sheet.row_values(0)
        code_col = next((i for i, h in enumerate(header) if "programme code" in str(h).lower()), None)
        name_col = next(
            (i for i, h in enumerate(header) if str(h).strip().lower() in _NAME_HEADER_CANDIDATES),
            None,
        )
        if code_col is None or name_col is None:
            return {}
        lookup = {}
        for r in range(1, sheet.nrows):
            row = sheet.row_values(r)
            code = row[code_col]
            if code:
                lookup[str(code).strip()] = str(row[name_col]).strip()
        return lookup
    except Exception:
        return {}


def get_programme_name_lookup():
    """Builds an official_code -> programme name lookup from whatever
    Programme List Excel is currently uploaded. Centralized here so EVERY
    check's results table gets programme names for free -- individual
    run_xxxx.py batch scripts don't need to know about this at all, and
    any new check added later automatically gets it too."""
    files = glob.glob(os.path.join(INCOMING_PROGLIST, "*.xls"))
    if not files:
        return {}
    path = files[0]
    return _load_programme_name_lookup_cached(path, os.path.getmtime(path))


# Old dict-string / bare-list result formats use these substrings to mean
# "nothing wrong here" -- kept for backward compatibility with checks still
# writing that format (Fee, Annexure).
_LEGACY_CLEAN_SUBSTRING_PATTERN = r"^\s*\[\]\s*$|'mismatches':\s*\[\]|SKIPPED"


def is_flagged_result(value):
    """Whether a result cell should count as a flagged row, in a way that
    works for EVERY check's output format, not just the one that happened
    to be built first.

    The bug this fixes: a BLANK/NaN/"None" cell (what newer checks like
    Important Dates correctly write when there's nothing to report) was
    being treated as flagged, because it never CONTAINS the old "clean"
    substring pattern -- str.contains() on an empty/NaN cell simply can't
    match anything, so the old "~contains(clean pattern)" logic defaulted
    to "flagged" for every single blank row. Checking for blank/None FIRST,
    before ever looking at the legacy pattern, fixes this for both old- and
    new-format checks at once, and will keep working for any future check
    that follows the same "blank means nothing to report" convention.
    """
    if pd.isna(value):
        return False
    text = str(value).strip()
    if text == "" or text.lower() == "none":
        return False
    if re.search(_LEGACY_CLEAN_SUBSTRING_PATTERN, text):
        return False
    return True


def kill_process_tree(pid):
    """Kills a process AND all its children (e.g. a Playwright-launched
    Chromium browser). On Windows, Popen.terminate() alone only kills the
    direct child, leaving orphaned grandchild processes (and file locks)
    behind -- that's what caused the earlier PermissionError."""
    if platform.system() == "Windows":
        subprocess.run(["taskkill", "/F", "/T", "/PID", str(pid)], capture_output=True)
    else:
        import signal
        try:
            os.killpg(os.getpgid(pid), signal.SIGKILL)
        except Exception:
            pass


def launch_step(module):
    """-u = unbuffered: without it, a child script's stdout is block-buffered
    (not flushed line by line) whenever it's writing to a pipe instead of a
    real terminal -- so even a "live" reader would only see output arrive in
    occasional bursts, not as it's actually printed. stderr is merged into
    stdout (STDOUT) so both show up in the one live stream in the order they
    were printed, same as watching it run in a terminal.

    A background thread reads line-by-line into `lines` as they arrive --
    that list is what the UI polls every second, instead of waiting for the
    whole step to finish before showing anything."""
    proc = subprocess.Popen(
        [sys.executable, "-u", "-m", module],
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1,
        cwd=os.path.dirname(os.path.abspath(__file__)),
    )
    lines = []

    def _reader():
        for line in proc.stdout:
            lines.append(line.rstrip("\n"))

    thread = threading.Thread(target=_reader, daemon=True)
    thread.start()
    return proc, lines, thread


def start_pipeline(name, steps, results_file):
    """Called once, when the Run button is clicked."""
    if results_file and os.path.exists(results_file):
        try:
            os.remove(results_file)
        except PermissionError:
            st.session_state.pipeline_error = (
                f"Could not clear {results_file} -- it's locked by another process. "
                "Check Task Manager for a leftover python.exe/chrome.exe, end it, then try again."
            )
            return

    proc, lines, thread = launch_step(steps[0][1])
    st.session_state.pipeline_error = None
    st.session_state.pipeline = {
        "name": name,
        "steps": steps,
        "results_file": results_file,
        "idx": 0,
        "proc": proc,
        "live_lines": lines,       # output of the CURRENTLY running step, growing live
        "reader_thread": thread,
        "log": [],                 # list of (label, full_text) for FINISHED steps
        "status": "running",  # running | done | failed
        "message": "",
    }


def stop_pipeline():
    pl = st.session_state.get("pipeline")
    if pl and pl.get("proc") is not None:
        kill_process_tree(pl["proc"].pid)
        st.session_state.stopped_message = f"Stopped: {pl['name']}"
    st.session_state.pipeline = None


def force_full_rerun():
    """A rerun triggered from inside a @st.fragment only re-executes that
    fragment by default -- it does NOT re-run the rest of app.py. That was
    the actual bug behind "results generated on disk but page still says
    No results yet": the Results table, the success/error banner, and the
    Run button's disabled state all live in the OUTER script, which was
    never being told to look again. scope="app" forces a real full-page
    rerun instead of a fragment-only one. (Older Streamlit versions that
    don't support the scope= argument fall back to a plain rerun.)"""
    try:
        st.rerun(scope="app")
    except TypeError:
        st.rerun()


@st.fragment(run_every=1)
def pipeline_status_fragment():
    pl = st.session_state.get("pipeline")
    if not pl or pl["status"] != "running":
        return

    steps = pl["steps"]
    idx = pl["idx"]
    label, _module = steps[idx]

    st.progress(idx / len(steps), text=f"Running: {label}")
    # Only ONE button is created per rerun of this fragment -- no duplicate keys,
    # and because this is a real Streamlit rerun (triggered by run_every), a
    # click on it is actually detected.
    st.button("Stop this check", key="stop_running_check", on_click=stop_pipeline)

    # Live terminal output for the CURRENTLY running step -- re-rendered from
    # pl["live_lines"] on every 1s tick, so it fills in as the child script
    # actually prints, the same as watching it in a terminal, instead of
    # appearing all at once only after the step finishes.
    st.caption(f"Live output -- {label}")
    st.code("\n".join(pl["live_lines"]) or "(waiting for output...)", language=None)

    proc = pl["proc"]
    if proc.poll() is None:
        return  # still running -- check again on the next 1s tick

    # Step just finished: give the reader thread a moment to drain any last
    # buffered lines, then freeze this step's output into the permanent log.
    pl["reader_thread"].join(timeout=2)
    pl["log"].append((label, "\n".join(pl["live_lines"])))

    if proc.returncode != 0:
        pl["status"] = "failed"
        pl["message"] = f"{label} failed -- see log below."
        force_full_rerun()
        return

    next_idx = idx + 1
    if next_idx >= len(steps):
        pl["status"] = "done"
        pl["message"] = "Completed successfully."
        force_full_rerun()
        return

    next_label, next_module = steps[next_idx]
    next_proc, next_lines, next_thread = launch_step(next_module)
    pl["idx"] = next_idx
    pl["proc"] = next_proc
    pl["live_lines"] = next_lines
    pl["reader_thread"] = next_thread


# ---------- Shared file uploads ----------
st.header("Source files")
col1, col2, col3, col4 = st.columns(4)
with col1:
    fee_file = st.file_uploader("Fee & Scholarship Excel (.xls)", type=["xls"], key="fee")
    if fee_file:
        st.success(f"Saved to {save_uploaded_file(fee_file, INCOMING_FEE)}")
with col2:
    proglist_file = st.file_uploader("Final Programme List Excel (.xls)", type=["xls"], key="proglist")
    if proglist_file:
        st.success(f"Saved to {save_uploaded_file(proglist_file, INCOMING_PROGLIST)}")
with col3:
    programid_file = st.file_uploader("ProgramId Source File (.xlsx)", type=["xlsx"], key="programid")
    if programid_file:
        st.success(f"Saved to {save_uploaded_file(programid_file, INCOMING_PROGRAM_ID)}")
with col4:
    dates_file = st.file_uploader("Important Dates Schedule (.xlsx)", type=["xlsx"], key="dates")
    if dates_file:
        st.success(f"Saved to {save_uploaded_file(dates_file, INCOMING_IMPORTANT_DATES)}")
intl_fee_file = st.file_uploader("International Fee Excel (.xls)", type=["xls"], key="intlfee")
if intl_fee_file:
    st.success(f"Saved to {save_uploaded_file(intl_fee_file, INCOMING_INTERNATIONAL_FEE)}")
# ─────────────────────── ADD THIS BLOCK ───────────────────────
elig_file = st.file_uploader("Eligibility Excel (.xls/.xlsx)", type=["xls", "xlsx"], key="elig")
if elig_file:
    st.success(f"Saved to {save_uploaded_file(elig_file, INCOMING_ELIGIBILITY)}")
# ─────────────────────── END ADD ──────────────────────────────

st.divider()

# ---------- Section picker ----------
st.header("Choose a section")
section = st.radio(
    "Section",
    ["Setup / Data Refresh", "Validation", "Records", "Previous Results"],
    horizontal=True, label_visibility="collapsed",
)

if section == "Previous Results":
    st.header("Previous Results")
    # Non-recursive on purpose: only the tool_Script root, not lpu_monitor/,
    # data/, python_embedded/, etc. -- those are internal working files, not
    # results a person would want to download. Add a pattern here if a
    # future check writes some other file type.
    RESULT_FILE_PATTERNS = ("*.csv", "*.xlsx")
    result_files = sorted(set().union(*(glob.glob(p) for p in RESULT_FILE_PATTERNS)))

    if not result_files:
        st.info("No result files in the project folder yet -- run a check first.")
    else:
        for fname in result_files:
            is_excel = fname.endswith(".xlsx")
            col_name, col_button = st.columns([3, 1])
            with col_name:
                st.write(fname)
            with col_button:
                with open(fname, "rb") as f:
                    st.download_button(
                        "Download",
                        data=f.read(),
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if is_excel else "text/csv",
                        key=f"prev_dl_{fname}",
                    )
else:
    items = SETUP_ACTIONS if section == "Setup / Data Refresh" else (VALIDATION_CHECKS if section == "Validation" else RECORDS_ACTIONS)

    st.header(f"Choose: {section}")
    selected = st.selectbox("Action/Check", list(items.keys()), label_visibility="collapsed")
    item = items[selected]
    st.caption(f"Needs: {', '.join(item['needs_files'])}")

    pipeline_running = st.session_state.get("pipeline") is not None

    if st.button(f"Run: {selected}", type="primary", disabled=pipeline_running):
        start_pipeline(selected, item["steps"], item.get("results_file"))
        st.rerun()

    if pipeline_running:
        st.caption("A check is already running -- stop it below before starting another.")

    if st.session_state.get("pipeline_error"):
        st.error(st.session_state.pipeline_error)

    if st.session_state.get("stopped_message"):
        st.warning(st.session_state.stopped_message)
        st.session_state.stopped_message = None

    pl = st.session_state.get("pipeline")
    if pl:
        pipeline_status_fragment()
        if pl["status"] in ("done", "failed"):
            if pl["status"] == "done":
                st.success(pl["message"])
            else:
                st.error(pl["message"])
            with st.expander("Log", expanded=(pl["status"] == "failed")):
                for label, full_text in pl["log"]:
                    st.text(f"--- {label} ---\n{full_text}")
            if st.button("Clear"):
                st.session_state.pipeline = None
                st.rerun()

    st.divider()

    # ---------- Results ----------
    if section in ("Validation", "Records"):
        st.header(f"Results: {selected}")
        path = item["results_file"]

        # Don't offer the download while THIS check's own run is still writing to
        # that exact file -- only once it's fully generated. A finished run of a
        # different check, or an old leftover file from an earlier run, is fine
        # to download; a file this check is actively appending to right now is not.
        pl = st.session_state.get("pipeline")
        currently_writing_this_file = (
            pl is not None and pl.get("results_file") == path and pl.get("status") == "running"
        )

        if os.path.exists(path) and not currently_writing_this_file:
            with open(path, "rb") as f:
                st.download_button(
                    "Download " + ("Excel" if path.endswith(".xlsx") else "CSV"),
                    data=f.read(),
                    file_name=os.path.basename(path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if path.endswith(".xlsx") else "text/csv",
                )
            if path.endswith(".xlsx"):
                df = pd.read_excel(path)
                st.dataframe(df, use_container_width=True)
            else:
                df = pd.read_csv(path, header=None, names=["official_code", "result"])

                name_lookup = get_programme_name_lookup()
                if name_lookup:
                    df.insert(
                        1, "programme_name",
                        df["official_code"].astype(str).str.strip().map(name_lookup).fillna(""),
                    )

                mismatch_mask = df["result"].apply(is_flagged_result)
                show_mismatches_only = st.checkbox("Show only mismatches/errors", value=True)
                if show_mismatches_only:
                    filtered = df[mismatch_mask]
                    st.write(f"{len(filtered)} of {len(df)} rows flagged")
                    st.dataframe(filtered, use_container_width=True)
                else:
                    st.dataframe(df, use_container_width=True)
        elif currently_writing_this_file:
            st.info("This check is still running -- results will appear here once it finishes.")
        else:
            st.info("No results yet for this check -- run it above.")