"""Render a .docx to PDF so the converted layout can actually be inspected.

The generated Word file must be *seen* before it is delivered; page breaks,
clipped headings and mis-anchored pictures only show up when a real layout
engine paginates the document.  This wrapper drives Word (or LibreOffice)
headlessly and reports the resulting page count.

Two environment quirks are handled here:

* some Office installs cannot load the Word type library, making the usual
  ``$word.Documents.Open(...)`` call fail with ``TYPE_E_CANTLOADLIBRARY``.
  The generated PowerShell therefore drives Word through IDispatch
  (``InvokeMember``), which needs no type library.
* Windows PowerShell 5.1 reads ``.ps1`` files as ANSI, so a script containing a
  non-ASCII path is silently mangled.  Every path handed to PowerShell is
  therefore copied to an ASCII-only temp name first.

Examples
--------
python render_docx.py resume.docx                 # -> resume.pdf
python render_docx.py resume.docx --pdf out.pdf
python render_docx.py resume.docx --engine soffice
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile

PS_TEMPLATE = """$ErrorActionPreference = 'Stop'
$docx = '{docx}'
$pdf  = '{pdf}'
if (Test-Path $pdf) {{ Remove-Item -LiteralPath $pdf -Force }}
$word = New-Object -ComObject Word.Application
$wt = $word.GetType()
try {{
    $wt.InvokeMember('Visible', 'SetProperty', $null, $word, @($false)) | Out-Null
    $wt.InvokeMember('DisplayAlerts', 'SetProperty', $null, $word, @(0)) | Out-Null
    $docs = $wt.InvokeMember('Documents', 'GetProperty', $null, $word, @())
    $dt = $docs.GetType()
    $doc = $dt.InvokeMember('Open', 'InvokeMethod', $null, $docs, @($docx, $false, $true))
    $et = $doc.GetType()
    $et.InvokeMember('SaveAs', 'InvokeMethod', $null, $doc, @($pdf, 17)) | Out-Null
    $et.InvokeMember('Close', 'InvokeMethod', $null, $doc, @(0)) | Out-Null
}} finally {{
    $wt.InvokeMember('Quit', 'InvokeMethod', $null, $word, @()) | Out-Null
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
}}
"""


def find_soffice():
    for name in ("soffice", "soffice.exe", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    for candidate in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if os.path.exists(candidate):
            return candidate
    return None


def render_with_word(docx, pdf, keep_temp=False):
    temp = tempfile.mkdtemp(prefix="docxrender_")
    try:
        staged_docx = os.path.join(temp, "input.docx")
        staged_pdf = os.path.join(temp, "output.pdf")
        shutil.copy2(docx, staged_docx)
        script = os.path.join(temp, "render.ps1")
        with open(script, "w", encoding="ascii") as fh:
            fh.write(PS_TEMPLATE.format(docx=staged_docx, pdf=staged_pdf))
        completed = subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
             "-File", script],
            capture_output=True, text=True,
        )
        if completed.returncode != 0 or not os.path.exists(staged_pdf):
            raise RuntimeError(
                f"Word rendering failed (exit {completed.returncode}):\n"
                f"{completed.stdout}\n{completed.stderr}"
            )
        os.makedirs(os.path.dirname(os.path.abspath(pdf)) or ".", exist_ok=True)
        shutil.move(staged_pdf, pdf)
    finally:
        if not keep_temp:
            shutil.rmtree(temp, ignore_errors=True)
        else:
            print(f"[render] temp kept: {temp}")
    return pdf


def render_with_soffice(docx, pdf, soffice):
    temp = tempfile.mkdtemp(prefix="docxrender_")
    try:
        staged_docx = os.path.join(temp, "input.docx")
        shutil.copy2(docx, staged_docx)
        subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf", "--outdir", temp,
             staged_docx],
            check=True, capture_output=True, text=True,
        )
        produced = os.path.join(temp, "input.pdf")
        if not os.path.exists(produced):
            raise RuntimeError("LibreOffice produced no PDF")
        os.makedirs(os.path.dirname(os.path.abspath(pdf)) or ".", exist_ok=True)
        shutil.move(produced, pdf)
    finally:
        shutil.rmtree(temp, ignore_errors=True)
    return pdf


def page_count(pdf):
    try:
        import pdfplumber
        with pdfplumber.open(pdf) as handle:
            return len(handle.pages)
    except Exception:
        return None


def main(argv=None):
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("docx")
    parser.add_argument("--pdf", help="output PDF (default: alongside the docx)")
    parser.add_argument("--engine", choices=("auto", "word", "soffice"), default="auto")
    parser.add_argument("--keep-temp", action="store_true")
    args = parser.parse_args(argv)

    docx = os.path.abspath(args.docx)
    if not os.path.exists(docx):
        parser.error(f"no such file: {docx}")
    pdf = os.path.abspath(args.pdf) if args.pdf else os.path.splitext(docx)[0] + ".pdf"

    errors = []
    if args.engine in ("auto", "word"):
        try:
            render_with_word(docx, pdf, keep_temp=args.keep_temp)
            pages = page_count(pdf)
            print(f"[render] word   -> {pdf}" + (f" ({pages} pages)" if pages else ""))
            return 0
        except Exception as exc:  # noqa: BLE001 - fall through to the next engine
            errors.append(f"word: {exc}")
            if args.engine == "word":
                print("[render] " + errors[-1])
                return 1
    soffice = find_soffice()
    if args.engine in ("auto", "soffice") and soffice:
        try:
            render_with_soffice(docx, pdf, soffice)
            pages = page_count(pdf)
            print(f"[render] soffice -> {pdf}" + (f" ({pages} pages)" if pages else ""))
            return 0
        except Exception as exc:  # noqa: BLE001
            errors.append(f"soffice: {exc}")
    for error in errors:
        print("[render] " + error, file=sys.stderr)
    print("[render] no working renderer; verify the .docx by other means or state "
          "that visual verification was not possible", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
