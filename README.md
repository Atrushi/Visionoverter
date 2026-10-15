# Visionoverter

A Windows desktop app for batch conversion between **Word, PDF, and Markdown**, with native file drag-and-drop, a white Light theme, and a Dark theme. Conversion and English/Arabic OCR happen locally; Visionoverter does not upload documents or download Markdown images.

| Direction | Input | Output | Desktop Microsoft Word required |
| --- | --- | --- | --- |
| Word → PDF | `.docx` | `.pdf` | Yes |
| PDF → Word | `.pdf` | `.docx` | Yes |
| Word → Markdown | `.docx` | `.md` | No |
| PDF → Markdown | `.pdf` | `.md` | Yes |
| Markdown → Word | `.md` or `.markdown` | `.docx` | No |
| Markdown → PDF | `.md` or `.markdown` | `.pdf` | Yes |

## Download and run (Windows x64)

Open this repository's **Releases**, choose the **Latest** release, and download **Visionoverter.exe** under **Assets**. Double-click it to start; no Python or Tesseract installation is needed. The icon and required application assets are embedded.

Desktop Microsoft Word is needed for the four PDF-related directions in the table above. Word → Markdown and Markdown → Word work without Word. Conversion and OCR run locally, without uploading your documents.

Choose the executable for the ready-to-run app. GitHub's separate **Source code (zip)** download contains the repository files for developers; it is not the Windows application download.

## Requirements

- Windows. The four directions marked **Yes** above need **desktop Microsoft Word installed, activated, and configured**. Open Word once and complete setup/sign-in prompts. Word Online is insufficient. Word → Markdown and Markdown → Word run without Microsoft Word.
- **Python 3.12 x64** to run the source or build the complete Windows package, including the supplied native OCR wheel.
- Built executables include Python, the OCR engine, English/Arabic language data, native drag-and-drop support, and application dependencies. End users need neither Python nor a separate Tesseract installation. The four Word-dependent directions still require Microsoft Word.

Visionoverter checks real Word COM automation before a Word-dependent batch and reports an actionable error if it is unavailable. The two pure Word ↔ Markdown directions do not initialize Word automation.

## Set up and run (developers)

Run these exact commands in Windows PowerShell from the project folder containing `main.py`:

```powershell
py -3.12-64 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe .\main.py
```

No virtual-environment activation or execution-policy change is needed for running the source.

The first command uses the Windows Python launcher to select Python 3.12 x64. If the launcher is unavailable, use the full path to a Python 3.12 x64 executable followed by `-m venv .venv`. Virtual environments are specific to their computer; recreate a copied environment if it refers to a missing Python installation.

## Use

1. Choose one of the six directions in the conversion selector. Changing direction clears the queue without deleting any files.
2. Use **Add Word Files**, **Add PDF Files**, or **Add Markdown Files**, or drop files onto the window or file list. Add multiple files from different folders. Drops accept only existing files matching the selected source type; folders are ignored and are not scanned recursively. Duplicate paths and Word temporary files beginning with `~$` are ignored. Both `.md` and `.markdown` are accepted for Markdown input.
3. Review filename, location, and status. Select rows to read complete paths and conversion details. **Remove Selected** and **Clear List** edit the queue.
4. Existing output files are skipped by default. Check **Overwrite existing output files** explicitly to replace them, including an existing `.docx` beside a PDF.
5. Choose **Convert** to process the entire queue sequentially. The interface stays responsive and shows processed counts, individual results, and a final successful/failed/skipped summary.
6. Switch **Light / Dark** at any time, including while conversion is running. The direction and queue controls are locked during a batch, and incoming drops are refused without changing the queue.

Each output is saved beside its source with the same base filename and the extension in the table above. For example, `Report.docx` produces `Report.md` in Word → Markdown; `Notes.markdown` produces `Notes.pdf` in Markdown → PDF. Sources are preserved. Output is staged in the same folder and published only after validation; failed conversions preserve existing output. Errors are recorded per file and processing continues. If native drag-and-drop cannot initialize, the interface explains this and the Add Files buttons remain available.

The footer displays **© 2026 Vision Dynamics**.

## Editable PDF → Word

Digital PDFs use Word's PDF reflow to preserve paragraphs, tables, and layout where possible. The app checks the resulting Word XML for visible editable text and source-text coverage. If Word produces picture-only or incomplete text, the app rebuilds editable text and records that layout may differ.

Scanned or mixed PDFs use bundled local Tesseract OCR for English and Arabic. Page images are used temporarily for recognition; they are **never inserted as full-page pictures in the generated Word document**. Recognized text becomes real Word paragraphs, retaining page sizes, page breaks, and simple indentation/spacing. An unreadable page fails clearly instead of producing a misleading image-only Word file.

PDFs do not store Word's paragraph/table structure, so complex layouts, equations, columns, handwriting, and scans can require correction. Review OCR spelling and reconstructed tables. Password-protected PDFs must be saved as an unencrypted copy first. See [Microsoft's PDF reflow explanation](https://support.microsoft.com/en-us/word/opening-pdfs-in-word).

## Markdown conversion

Word → Markdown converts document-body paragraphs and tables into UTF-8 Markdown. It preserves supported headings, bold/italic/strikethrough text, hyperlinks, ordered and bulleted lists, quotes, code, and simple tables. Merged cells, nested tables, and cell line breaks are flattened with a note in the result details; their supported text is retained. Embedded Word images become descriptive text placeholders. Page layout, fonts, colors, and unsupported Word features are not reproduced exactly.

PDF → Markdown first creates an editable temporary Word document through the same PDF reflow/OCR pipeline described above, then converts its supported structure to Markdown. The temporary document is removed after conversion. Layout and scan-recognition limitations also apply to the Markdown output.

Markdown → Word creates editable headings, paragraphs, emphasis, hyperlinks, nested lists, quotes, tables, code, and horizontal rules. Markdown → PDF renders that generated Word content through the dedicated Microsoft Word instance. Markdown input supports UTF-8, UTF-8 with a BOM, and UTF-16 with a BOM; generated Markdown is UTF-8. CommonMark syntax is supported with pipe tables and strikethrough. Raw HTML is treated as text rather than rendered HTML.

Supported local Markdown images are embedded in the generated Word/PDF content; relative image paths are resolved beside the Markdown file. Remote, unavailable, or unsupported images remain text placeholders containing their description and source. Visionoverter does not fetch HTTP(S) image URLs or open direct UNC image paths, including URL-encoded UNC paths. Hyperlinks remain links; conversion does not open them. Review the per-file details for image or layout limitations.

## Word session safety and limitations

`docx2pdf` is retained as a requested dependency. Its [Windows implementation](https://github.com/AlJohri/docx2pdf/blob/main/docx2pdf/__init__.py) uses shared `Dispatch("Word.Application")` and normally calls `Quit()`, which cannot reliably protect existing Word windows. Visionoverter therefore uses the permitted direct `pywin32` alternative: `DispatchEx`, a verified separate empty Word process, and PDF export through that dedicated instance.

For Word-dependent directions, the worker initializes and releases COM on its own thread. Word works on private copies with short temporary paths, then the app copies its result into the output staging path for validation and atomic publication. This avoids Office's [file-path limits](https://support.microsoft.com/en-us/onedrive/what-are-file-path-length-limits) and keeps Word's save/sync lifecycle away from the output folder. It closes only documents it opened and quits its instance only when no documents remain. PDF import retires each dedicated Word instance before copying its saved DOCX. Existing Word windows and unsaved work remain untouched; if isolation cannot be verified, conversion stops before changing Word settings. Pure Word → Markdown and Markdown → Word use `python-docx` and do not start Microsoft Word.

Office activation, add-ins, or modal Word dialogs can still block a COM call despite disabled alerts. The UI remains responsive, but closing is deferred until conversion and Word cleanup finish. Review Word for pending dialogs and allow the batch to finish. Visionoverter never force-kills Word.

## Build Windows executables

Build and verify `--onedir` first, then build `--onefile`. Both use `--windowed` and include CustomTkinter themes/fonts, native TkDND assets, the Markdown parser, OCR models/libraries, application icons, and COM dependencies. The helper script runs these same flags:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Mode onedir
.\.venv\Scripts\python.exe tools\verify_build.py onedir --launch
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\build.ps1 -Mode onefile
.\.venv\Scripts\python.exe tools\verify_build.py onefile --launch
```

The execution-policy override applies only to these helper-script processes. Equivalent full commands, from the project folder:

```powershell
$taskAssets = Join-Path (Get-Location).Path 'assets'
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --onedir --windowed `
    --name Visionoverter `
    --icon (Join-Path $taskAssets 'visionoverter.ico') `
    --collect-all customtkinter `
    --collect-all docx2pdf `
    --collect-all pypdf --collect-all pypdfium2 --collect-all pypdfium2_raw `
    --collect-all docx --collect-all PIL --collect-all tesserocr `
    --collect-all tkinterdnd2 --collect-all markdown_it `
    --copy-metadata tkinterdnd2 --copy-metadata markdown-it-py `
    --copy-metadata tesserocr --copy-metadata pypdfium2 --copy-metadata python-docx `
    --hidden-import pythoncom `
    --hidden-import pywintypes `
    --hidden-import win32com.client `
    --hidden-import win32timezone `
    --add-data "${taskAssets}:assets" `
    --distpath 'dist/Windows/onedir' --workpath 'build/onedir/manual' --specpath 'build/onedir' `
    main.py

.\.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --onefile --windowed `
    --name Visionoverter `
    --icon (Join-Path $taskAssets 'visionoverter.ico') `
    --collect-all customtkinter `
    --collect-all docx2pdf `
    --collect-all pypdf --collect-all pypdfium2 --collect-all pypdfium2_raw `
    --collect-all docx --collect-all PIL --collect-all tesserocr `
    --collect-all tkinterdnd2 --collect-all markdown_it `
    --copy-metadata tkinterdnd2 --copy-metadata markdown-it-py `
    --copy-metadata tesserocr --copy-metadata pypdfium2 --copy-metadata python-docx `
    --hidden-import pythoncom `
    --hidden-import pywintypes `
    --hidden-import win32com.client `
    --hidden-import win32timezone `
    --add-data "${taskAssets}:assets" `
    --distpath 'dist/Windows/onefile' --workpath 'build/onefile/manual' --specpath 'build/onefile' `
    main.py
```

Distribute the entire `dist\Windows\onedir\Visionoverter` folder for the directory build, or `dist\Windows\onefile\Visionoverter.exe` for the single-file build. Build on Windows; the target architecture follows the Python interpreter used to build. The helper creates a fresh intermediate work folder for every build and verifies bundled modules/assets. `--collect-all` includes CustomTkinter's non-Python assets in both formats; see [PyInstaller's collection options](https://pyinstaller.org/en/stable/usage.html).

Windowed startup errors use a native message box. GUI-library stdout/stderr writes are redirected to the log when no console is attached.

## Code and diagnostics

- `main.py`: logging, windowed startup, and application entry point.
- `visionoverter/ui.py`: six-direction selector, native file drops, themes, and queue controls; a main-thread event poll handles worker results.
- `visionoverter/models.py`: file queue, statuses, and immutable worker events.
- `visionoverter/conversion.py`: sequential batches, input validation, output staging, and per-file failures.
- `visionoverter/word_backend.py`: isolated Word automation and COM cleanup.
- `visionoverter/pdf_backend.py`: PDF preflight, Word reflow validation, local OCR, and editable paragraph reconstruction.
- `visionoverter/markdown_backend.py`: structural Word/Markdown conversion and local image handling; temporary Word adapters for PDF routes.
- `visionoverter/logging_setup.py`: rotating technical logs.

OCR data is supplied in `assets/tessdata` with its source commit, hashes, and license. Native OCR notices are in `assets/third-party`; PDFium license notices are included in the packaged metadata. Runtime conversion does not download models.

The log is `%LOCALAPPDATA%\Visionoverter\logs\visionoverter.log`, with a `%TEMP%\Visionoverter\logs` fallback if the application-data location is unavailable. Logs include paths and technical errors and remain local.

## Verification

Run automated checks:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\.venv\Scripts\python.exe tools\verify_windows.py --timeout-seconds 180
.\.venv\Scripts\python.exe tools\verify_build.py onedir
.\.venv\Scripts\python.exe tools\verify_build.py onefile
```

See [VERIFICATION.md](VERIFICATION.md) for actual test results, the version covered by each result, and remaining checks. `verify_windows.py` creates disposable documents for all six directions and a separate unsaved Word fixture; it never attaches to the user's Word instance. `verify_build.py` inspects assets, dependencies, and the Windows GUI subsystem. Add `--launch` to run the executable's internal verification mode, which exercises the UI, six conversion directions, and the registered drop callback with fresh local fixtures and records JSON evidence under `test-results`. Normal launches show only the application interface. Programmatic drop-callback checks do not by themselves verify a physical native drag gesture.

For a separate native drag test, run:

```powershell
.\.venv\Scripts\python.exe tools\verify_native_drop.py
```

Drag the helper's blue test-source row into the normal file queue. The helper creates its own disposable files and records native source/drop results. Also test dropping files from Windows Explorer; the helper's same-window source is a separate check.

Use this short manual checklist on Windows with Word configured:

- Add and drop files from multiple folders, including spaces and Arabic names. Include duplicates, a `~$` temporary Word file, a wrong source type, and a folder. Check filtering, remove/clear, sequential progress, and final counts.
- Exercise all six directions. Test both `.md` and `.markdown` input with headings, emphasis, links, lists, a table, code, Arabic text, and local/remote image references; inspect the generated structure and image placeholders.
- Convert with existing PDF, DOCX, and Markdown outputs: verify default skip and explicit overwrite while preserving originals. Mix valid, corrupt, encrypted/password-protected, invalid-text, and missing inputs; check continued processing and a locked output or unwritable folder.
- Convert paths containing spaces and non-English characters, including Arabic; review the resulting PDFs.
- In **PDF → Word**, convert a multi-page digital PDF, a table, an English scan, and an Arabic scan. Open the DOCX and edit/select the recognized text; verify all pages are represented and no page screenshots replace text.
- Switch Light/Dark and all directions; verify white default, readable table/status colors, correct file filters, safe queue clearing, and **© 2026 Vision Dynamics**. During a batch, verify that incoming drops are refused without changing the queue.
- Keep a separate Word window open with unsaved work. Verify that its contents, window, and modified state survive conversion and cleanup.
- On a machine without Word, verify that Word → Markdown and Markdown → Word work and that the other directions show the requirement message. For each packaged format, verify launch without Python installed or an attached console, native Explorer drops, themes/fonts/icons, conversion, and logging.
