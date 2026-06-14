from collections import defaultdict
from flask import Flask, request, send_file, render_template, jsonify
import os, sqlite3, hashlib, json, difflib, html
import re
import requests

import mimetypes
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

from library import startup
from library.config import config

app = Flask(__name__)
app.config.from_object(config.DevelopmentConfig)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app.config["STORAGE_FOLDER"] = os.path.join(BASE_DIR, app.config["STORAGE_FOLDER"])
app.config["DATABASE"] = os.path.join(BASE_DIR ,app.config["DATABASE"])

startup.check_db_exists_or_fail(app.config["DATABASE"])
startup.create_folder(app.config["STORAGE_FOLDER"])

TEXT_EXTENSIONS = {
    ".txt", ".md", ".py", ".js", ".ts", ".html", ".css", ".json", ".xml",
    ".yml", ".yaml", ".csv", ".sql", ".ini", ".cfg", ".conf", ".log",
    ".java", ".c", ".cpp", ".h", ".hpp", ".cs", ".php", ".rb", ".go",
    ".rs", ".sh", ".bat", ".ps1"
}
SQLITE_EXTENSIONS = {".db", ".sqlite", ".sqlite3"}
PDF_EXTENSIONS = {".pdf"}
DOCX_EXTENSIONS = {".docx"}
DOC_EXTENSIONS = {".doc"}
ZIP_EXTENSIONS = {".zip"}


def get_artifact_file_path(name, commit_hash):
    """
    Finds the real uploaded artifact path from the database.
    """
    conn, err = startup.connect_db(app.config["DATABASE"])
    if not conn:
        return None, str(err)

    try:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT artifact_storage.path
            FROM artifacts
            JOIN artifact_storage ON artifact_storage.artifact_id = artifacts.id
            WHERE artifacts.name = ? AND artifacts.commit_hash = ?
            ORDER BY artifacts.id DESC
            LIMIT 1
            """,
            (name, commit_hash)
        )

        row = cursor.fetchone()
        cursor.close()

        if not row:
            return None, f"Commit hash '{commit_hash}' does not exist for artifact '{name}'"

        return row[0], None

    except Exception as err:
        return None, f"Failed to find artifact path: {err}"

    finally:
        conn.close()
        

def get_db_dump_lines(db_path):
    """Connects to a SQLite database and returns its dump as a list of strings."""
    try:
        conn, err = startup.connect_db(db_path)
        if not conn:
            return None, err

        # We strip trailing newlines to ensure clean processing by difflib
        dump_lines = [line.strip() for line in conn.iterdump()]
        conn.close()

        # Sorting is okay for database dump because order can sometimes vary.
        dump_lines.sort()

        return dump_lines, ""

    except sqlite3.Error as err:
        error_message = f"Failed to fetch data from database '{db_path}'"
        app.logger.error(f"error_message: '{error_message}' | database: '{db_path}' | error: {err}")
        return None, error_message
    
    
def get_text_file_lines(file_path):
    """Reads normal text/code files."""
    try:
        try:
            with open(file_path, "r", encoding="utf-8") as file:
                return [line.rstrip("\n") for line in file.readlines()], ""
        except UnicodeDecodeError:
            with open(file_path, "r", encoding="cp1252") as file:
                return [line.rstrip("\n") for line in file.readlines()], ""

    except Exception as err:
        return None, f"Failed to read text/code file '{file_path}': {err}"


def get_pdf_lines(file_path):
    """Extracts text from a PDF file."""
    if PdfReader is None:
        return None, "PDF support requires pypdf. Install it with: python -m pip install pypdf"

    try:
        reader = PdfReader(file_path)
        lines = []

        for page_index, page in enumerate(reader.pages, start=1):
            text = page.extract_text() or ""
            lines.append(f"--- Page {page_index} ---")
            lines.extend(text.splitlines())
        
        return lines, ""

    except Exception as err:
        return None, f"Failed to read PDF file '{file_path}': {err}"


def get_docx_lines(file_path):
    """
    Extracts text from a .docx file.
    .docx files are ZIP files containing XML.
    """
    try:
        lines = []

        with zipfile.ZipFile(file_path, "r") as docx_zip:
            xml_content = docx_zip.read("word/document.xml")

        root = ET.fromstring(xml_content)

        namespace = {
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        }

        for paragraph in root.findall(".//w:p", namespace):
            text_parts = []

            for text_node in paragraph.findall(".//w:t", namespace):
                if text_node.text:
                    text_parts.append(text_node.text)

            paragraph_text = "".join(text_parts).strip()

            if paragraph_text:
                lines.append(paragraph_text)

        return lines, ""

    except Exception as err:
        return None, f"Failed to read DOCX file '{file_path}': {err}"


def get_zip_listing_lines(file_path):
    """
    Compares ZIP files by listing their contents.
    It compares filenames, sizes, and CRC values.
    """
    try:
        lines = []

        with zipfile.ZipFile(file_path, "r") as zip_file:
            for item in sorted(zip_file.infolist(), key=lambda x: x.filename):
                lines.append(
                    f"{item.filename} | size={item.file_size} | crc={item.CRC}"
                )

        return lines, ""

    except Exception as err:
        return None, f"Failed to read ZIP file '{file_path}': {err}"
    
    
def get_binary_file_info_lines(file_path):
    """
    Fallback for unknown binary files.
    It cannot show line-by-line diff, but it can compare file size and checksum.
    """
    try:
        file_size = os.path.getsize(file_path)

        with open(file_path, "rb") as file:
            checksum = hashlib.md5(file.read()).hexdigest()

        mime_type, _ = mimetypes.guess_type(file_path)

        return [
            f"File: {os.path.basename(file_path)}",
            f"Size: {file_size} bytes",
            f"MD5: {checksum}",
            f"MIME type: {mime_type or 'unknown'}",
        ], ""

    except Exception as err:
        return None, f"Failed to read binary file '{file_path}': {err}"


def get_artifacts():
    conn , err = startup.connect_db(app.config["DATABASE"])
    if not conn:
        return None, err
    try:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT * from artifacts a INNER JOIN artifact_storage s ON a.id = s.artifact_id
            ORDER BY a.name, a.created_at DESC
            """)
        rows = cursor.fetchall()
        grouped_artifacts = defaultdict(list)
        grouped_commit_hashes = defaultdict(list)
        for row in rows:
            grouped_artifacts[row["name"]].append({
                "id": row["id"],
                "commit_hash": row["commit_hash"],
                "tags": json.loads(row["tags"]) if row["tags"] else [],
                "checksum": row["checksum"],
                "path": row["path"],
                "created_at": row["created_at"]
            })
            grouped_commit_hashes[row["name"]].append(row["commit_hash"])
        return grouped_artifacts, grouped_commit_hashes, ""
    except sqlite3.Error as err:
        error_message = "Failed to fetch artifacts"
        app.logger.error(f"error_message: '{error_message}' | error: {err}")
        return None, None, err
    finally:
        conn.close()


def highlight_changes(old_row, new_row):
    highlighted_old = []
    highlighted_new = []
    
    max_len = max(len(old_row), len(new_row))
    for i in range(max_len):
        old_val = str(old_row[i]) if i < len(old_row) else ""
        new_val = str(new_row[i]) if i < len(new_row) else ""
        
        if old_val != new_val:
            highlighted_old.append(f"<span style='background-color: #ffcccc;'>{html.escape(old_val)}</span>")
            highlighted_new.append(f"<span style='background-color: #ccffcc;'>{html.escape(new_val)}</span>")
        else:
            highlighted_old.append(html.escape(old_val))
            highlighted_new.append(html.escape(new_val))
    
    return f"({', '.join(highlighted_old)})", f"({', '.join(highlighted_new)})"


def get_file_diff_lines(file_path):
    """
    Converts different file types into comparable text lines.
    Supports:
    - SQLite database files: .db, .sqlite, .sqlite3
    - Code/text files: .py, .js, .html, .css, .json, .txt, etc.
    - PDF files: .pdf
    - Word files: .docx
    - ZIP files: .zip
    - Unknown binary files: checksum comparison info
    """
    ext = Path(file_path).suffix.lower()

    if ext in SQLITE_EXTENSIONS:
        return get_db_dump_lines(file_path)

    if ext in TEXT_EXTENSIONS:
        return get_text_file_lines(file_path)

    if ext in PDF_EXTENSIONS:
        return get_pdf_lines(file_path)

    if ext in DOCX_EXTENSIONS:
        return get_docx_lines(file_path)

    if ext in DOC_EXTENSIONS:
        return None, "Old .doc files are not supported. Use .docx instead."

    if ext in ZIP_EXTENSIONS:
        return get_zip_listing_lines(file_path)

    return get_binary_file_info_lines(file_path)


def calculate_diff(name, commit_hash_a, commit_hash_b):
    path_a, err = get_artifact_file_path(name, commit_hash_a)
    if err:
        app.logger.error(f"error_message: '{err}' | name: '{name}' | commit_hash: '{commit_hash_a}'")
        return None, err

    path_b, err = get_artifact_file_path(name, commit_hash_b)
    if err:
        app.logger.error(f"error_message: '{err}' | name: '{name}' | commit_hash: '{commit_hash_b}'")
        return None, err

    if not os.path.exists(path_a):
        error_message = f"File does not exist for artifact '{name}' and commit '{commit_hash_a}'"
        app.logger.error(f"error_message: '{error_message}' | path: '{path_a}'")
        return None, error_message

    if not os.path.exists(path_b):
        error_message = f"File does not exist for artifact '{name}' and commit '{commit_hash_b}'"
        app.logger.error(f"error_message: '{error_message}' | path: '{path_b}'")
        return None, error_message

    rows_a, err = get_file_diff_lines(path_a)
    if err:
        app.logger.error(f"error: '{err}' | path: '{path_a}'")
        return None, err

    rows_b, err = get_file_diff_lines(path_b)
    if err:
        app.logger.error(f"error: '{err}' | path: '{path_b}'")
        return None, err

    if rows_a == rows_b:
        all_diffs = "(no differences found)"
    else:
        all_diffs = difflib.HtmlDiff(wrapcolumn=160).make_table(
            rows_a,
            rows_b,
            fromdesc=f"{name}@{commit_hash_a}",
            todesc=f"{name}@{commit_hash_b}",
            context=True,
            numlines=1
        )

    return all_diffs, None


def parse_tags(tags_text):
    """
    Safely converts the tags column into a dictionary.
    If tags is empty, invalid, or not a JSON object, it returns an empty dict.
    """
    if not tags_text:
        return {}

    try:
        tags = json.loads(tags_text)
    except json.JSONDecodeError:
        return {}

    if isinstance(tags, dict):
        return tags

    # If old tags were stored as a list like ["local", "test"],
    # convert them into a dictionary so we can add deployed=True/False.
    if isinstance(tags, list):
        return {"labels": tags}

    return {}


def set_uniquely_deployed_flag(artifact_id):
    """
    Marks the selected artifact as deployed.
    Marks all other artifacts with the same name as not deployed.
    """
    conn, err = startup.connect_db(app.config["DATABASE"])
    if not conn:
        return str(err), 500

    try:
        cursor = conn.cursor()

        cursor.execute(
            "SELECT name FROM artifacts WHERE id = ?",
            (artifact_id,)
        )

        row = cursor.fetchone()

        if not row:
            return "Artifact not found", 404

        name = row["name"] if isinstance(row, sqlite3.Row) else row[0]

        cursor.execute(
            "SELECT id, tags FROM artifacts WHERE name = ?",
            (name,)
        )

        artifacts = cursor.fetchall()

        for row in artifacts:
            current_artifact_id = row["id"] if isinstance(row, sqlite3.Row) else row[0]
            tags_text = row["tags"] if isinstance(row, sqlite3.Row) else row[1]

            tags = parse_tags(tags_text)

            if current_artifact_id == artifact_id:
                tags["deployed"] = True
            else:
                tags["deployed"] = False

            cursor.execute(
                "UPDATE artifacts SET tags = ? WHERE id = ?",
                (json.dumps(tags), current_artifact_id)
            )

        conn.commit()
        return "OK", 200

    except Exception as err:
        conn.rollback()
        error_message = "Failed to update deployed flag"
        app.logger.error(f"error_message: '{error_message}' | artifact_id: {artifact_id} | error: {err}")
        return error_message, 500

    finally:
        conn.close()


@app.route("/")
def index():
    artifacts, commit_hashes, err = get_artifacts()
    app.logger.info(f"artifacts: {commit_hashes}")
    if err:
        app.logger.error(f"error: '{err}'")
        return str(err), 500
    return render_template("index.html", artifacts=artifacts, commit_hashes=commit_hashes)


@app.route("/instructions")
def instructions():
    return render_template("instructions.html")


@app.route("/push", methods=["GET"])
def add():
    return render_template("add.html")


@app.route("/push", methods=["POST"])
def push():
    name = request.form.get("name")
    commit_hash = request.form.get("commit_hash")
    tags = request.form.get("tags", "")
    file_content = request.files.get("file")

    if not name or not commit_hash or not file_content:
        error_message = "Name and commit hash and file are required"
        app.logger.error(f"error: '{error_message}' | name: '{name}' | commit_hash: '{commit_hash}' | file_content: '{file_content}'") 
        return error_message, 400

    is_allowed_characters = re.match(r"[a-zA-Z0-9-_]+$", name)
    if not is_allowed_characters:
        error_message = "Invalid name. Name must be within this range of characters: (a-z, A-Z, 0-9, -, _)."
        app.logger.error(f"error_message: '{error_message}' | name: '{name}'")
        return error_message, 400
    
    try:
        parsed_tags = json.loads(tags) if tags else []
    except json.JSONDecodeError as err:
        error_message = "Invalid tags: Must be a valid JSON-encoded string."
        app.logger.error(f"error_message: '{error_message}' | tags: '{tags}' | error: {err}")
        return error_message, 400

    conn, err = startup.connect_db(app.config["DATABASE"])
    if not conn:
        return str(err), 500
    
    try:
        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO artifacts (name, commit_hash, tags) VALUES (?, ?, ?)
            """, (name, commit_hash, tags))
    
        artifact_id = cursor.lastrowid
    
        _, ext = os.path.splitext(file_content.filename)

        filepath = os.path.join(app.config["STORAGE_FOLDER"], f"{name}_{commit_hash}{ext}")
        file_content.save(filepath)

        with open(filepath, 'rb') as f:
            checksum = hashlib.md5(f.read()).hexdigest()

        cursor.execute(
            """
            INSERT INTO artifact_storage (artifact_id, checksum, path) VALUES (?, ?, ?)
            """, (artifact_id, checksum, filepath))

        conn.commit()
        cursor.close()
    except Exception as err:
        conn.rollback()
        error_message = "Failed to push artifact"
        app.logger.error(f"error_message: '{error_message}' | error: {err}")
        return error_message, 500
    finally:
        conn.close()
    return 'OK', 201


@app.route("/download", methods=["GET"])
def download_api():
    name = request.args.get("name")
    commit_hash = request.args.get("commit_hash", "")

    if not name:
        error_message = "Name is required"
        app.logger.error(f"error: '{error_message}' | name: '{name}'")
        return error_message, 400

    conn, err = startup.connect_db(app.config["DATABASE"])
    if not conn:
        return str(err), 500

    try:
        cursor = conn.cursor()

        if commit_hash:
            cursor.execute(
                """
                SELECT path FROM artifact_storage
                JOIN artifacts ON artifact_storage.artifact_id = artifacts.id
                WHERE artifacts.name = ? AND artifacts.commit_hash = ?
                """, (name, commit_hash))
        else:
            # if commit hash is not specified, we download the latest artifact.
            cursor.execute(
                """
                SELECT path FROM artifact_storage
                JOIN artifacts ON artifact_storage.artifact_id = artifacts.id
                WHERE artifacts.name = ? order by artifacts.id desc
                """, (name,))

        row = cursor.fetchone()
        if not row:
            error_message = "Artifact not found"
            app.logger.error(f"error_message: '{error_message}' | name: '{name}' | commit_hash: '{commit_hash}'")
            return error_message, 404
        filepath = row['path']

        return send_file(filepath, as_attachment=True), 200
    except Exception as err:
        error_message = "Failed to download artifact"
        app.logger.error(f"error_message: '{error_message}' | error: {err}")
        return error_message, 500
    finally:
        conn.close()
        
        
@app.route("/download/<int:artifact_id>")
def download_dashboard(artifact_id):
    conn, err = startup.connect_db(app.config["DATABASE"])
    if not conn:
        return str(err), 500
    
    try:
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT path FROM artifact_storage WHERE artifact_id = ?
            """, (artifact_id,))
        row = cursor.fetchone()
        if not row:
            error_message = "Artifact not found"
            app.logger.error(f"error_message: '{error_message}' | artifact_id: {artifact_id}")
            return error_message, 404
        filepath = row['path']

        return send_file(filepath, as_attachment=True), 200
    except Exception as err:
        error_message = "Failed to download artifact"
        app.logger.error(f"error_message: '{error_message}' | error: {err}")
        return error_message, 500
    finally:
        conn.close()


@app.route("/diff")
def diff_api():
    name = request.args.get("name")
    commit_hash_a = request.args.get("commit_hash_a")
    commit_hash_b = request.args.get("commit_hash_b")

    if not name or not commit_hash_a or not commit_hash_b:
        error_message = "Name, both commit hashes and table are required"
        app.logger.error(f"error: '{error_message}' | name: '{name}' | commit_hash_a: '{commit_hash_a}' | commit_hash_b: '{commit_hash_b}'")
        return error_message, 400

    diff, err = calculate_diff(name, commit_hash_a, commit_hash_b)

    if err:
        app.logger.error(f"error: '{err}' | name: '{name}' | commit_hash_a: '{commit_hash_a}' | commit_hash_b: '{commit_hash_b}'")
        return str(err), 500
    
    return jsonify({
        "name1": name,
        "commit_hash_a": commit_hash_a,
        "commit_hash_b": commit_hash_b,
        "diff_by_table": diff
    })


@app.route("/deploy/<int:artifact_id>")
def deploy_api(artifact_id):
    """
    Marks this artifact as deployed.
    All other artifacts with the same name become undeployed.
    """
    return set_uniquely_deployed_flag(artifact_id)


if __name__ == "__main__":
    app.run(debug=True)