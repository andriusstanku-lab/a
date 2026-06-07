"""
generate_odb.py

Sukuria realu LibreOffice Base failą (studentai.odb) su embedded HSQLDB 1.8.0.10
duomenu baze. Faile yra trys lenteles (Students, Courses, Enrollments) su
pirminiais ir isoriniais raktais ir 19 pavyzdziniu irasu.

Reikalavimai:
  - Java JDK su HSQLDB 1.8.0.10 jar
  - Python 3.x

Naudojimas:
  python generate_odb.py
  # Sukuriamas: studentai.odb (galima atidaryti su LibreOffice Base)

Veikimo principas:
  1. Paleidziamas Java programa CreateDB - sukuria HSQLDB failus (script + properties)
  2. Sugeneruojami ODF (Open Document Format) metaduomenys
  3. Visi failai supakuojami i ZIP archyva su .odb pletiniu
"""

import os
import shutil
import subprocess
import sys
import zipfile
import time
import uuid

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
HSQLDB_JAR = os.path.expanduser(
    "~/.m2/repository/org/hsqldb/hsqldb/1.8.0.10/hsqldb-1.8.0.10.jar"
)
WORK_DIR = os.path.join(SCRIPT_DIR, "studentai_db")
OUTPUT_FILE = os.path.join(SCRIPT_DIR, "studentai.odb")


# ============================================================
# ODF (Open Document Format) metaduomenu sablonai
# ============================================================

MIMETYPE = "application/vnd.oasis.opendocument.base"

MANIFEST_XML = """<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
 <manifest:file-entry manifest:full-path="/" manifest:version="1.2" manifest:media-type="application/vnd.oasis.opendocument.base"/>
 <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="settings.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/>
 <manifest:file-entry manifest:full-path="database/script" manifest:media-type=""/>
 <manifest:file-entry manifest:full-path="database/properties" manifest:media-type=""/>
</manifest:manifest>
"""

META_XML = """<?xml version="1.0" encoding="UTF-8"?>
<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 office:version="1.2">
 <office:meta>
  <meta:generator>PU4_VU_Andrius_Vargonas/2026.05</meta:generator>
  <meta:creation-date>{date}T00:00:00</meta:creation-date>
  <dc:date>{date}T00:00:00</dc:date>
  <dc:title>Studentu duomenu baze (PU4)</dc:title>
  <dc:creator>Andrius Vargonas</dc:creator>
  <dc:subject>Informacijos sistemos ir duomenu bazes</dc:subject>
 </office:meta>
</office:document-meta>
"""

SETTINGS_XML = """<?xml version="1.0" encoding="UTF-8"?>
<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.2">
 <office:settings/>
</office:document-settings>
"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8"?>
<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.2">
 <office:styles/>
 <office:automatic-styles/>
 <office:master-styles/>
</office:document-styles>
"""

# Pagrindinis content.xml - apraso, kad failas yra LibreOffice Base
# embedded HSQLDB tipo. Ypac svarbus xlink:href su sdbc:embedded:hsqldb URL.
CONTENT_XML = """<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:db="urn:oasis:names:tc:opendocument:xmlns:database:1.0"
 xmlns:xlink="http://www.w3.org/1999/xlink"
 office:version="1.2">
 <office:body>
  <office:database>
   <db:data-source>
    <db:connection-data>
     <db:connection-resource xlink:href="sdbc:embedded:hsqldb"/>
     <db:login db:user-name="SA" db:is-password-required="false"/>
    </db:connection-data>
    <db:driver-settings db:system-driver-settings="" db:base-dn="" db:parameter-name-substitution="false"/>
    <db:application-connection-settings db:is-table-name-length-limited="false" db:is-set-add-table-in-sql="true">
     <db:table-filter>
      <db:table-filter-pattern>%</db:table-filter-pattern>
     </db:table-filter>
    </db:application-connection-settings>
    <db:data-source-settings>
     <db:data-source-setting db:data-source-setting-name="ParameterNameSubstitution" db:data-source-setting-type="boolean">
      <db:data-source-setting-value>false</db:data-source-setting-value>
     </db:data-source-setting>
    </db:data-source-settings>
   </db:data-source>
  </office:database>
 </office:body>
</office:document-content>
"""


# ============================================================
# Pagalbines funkcijos
# ============================================================

def run_java_create_db():
    """Paleidzia CreateDB.java (kompiliuoja, jei reikia) ir sukuria HSQLDB failus."""
    if not os.path.exists(HSQLDB_JAR):
        sys.exit(
            "KLAIDA: HSQLDB 1.8.0.10 jar nerastas. Paleiskite:\n"
            "  mvn dependency:get -Dartifact=org.hsqldb:hsqldb:1.8.0.10"
        )

    # Pasalinti seną darbo direktorija
    if os.path.exists(WORK_DIR):
        shutil.rmtree(WORK_DIR)

    # Kompiliuoti CreateDB.java jeigu .class nera
    java_file = os.path.join(SCRIPT_DIR, "CreateDB.java")
    class_file = os.path.join(SCRIPT_DIR, "CreateDB.class")
    if not os.path.exists(class_file) or \
            os.path.getmtime(java_file) > os.path.getmtime(class_file):
        print("Kompiliuojama CreateDB.java...")
        result = subprocess.run(
            ["javac", "-cp", HSQLDB_JAR, java_file],
            capture_output=True, text=True, cwd=SCRIPT_DIR,
        )
        if result.returncode != 0:
            sys.exit("Kompiliavimo klaida:\n" + result.stderr)

    # Paleisti CreateDB
    print("Paleidziama CreateDB programa...")
    result = subprocess.run(
        ["java", "-cp", f".:{HSQLDB_JAR}", "CreateDB", WORK_DIR],
        capture_output=True, text=True, cwd=SCRIPT_DIR,
    )
    print(result.stdout)
    if result.returncode != 0:
        sys.exit("Java vykdymo klaida:\n" + result.stderr)


def build_odb_zip():
    """Supakuoja visus failus i .odb (ZIP) archyva pagal ODF specifikacija."""
    # Patikrinam, kad HSQLDB failai egzistuoja
    script_path = os.path.join(WORK_DIR, "database", "script")
    properties_path = os.path.join(WORK_DIR, "database", "properties")

    if not os.path.exists(script_path):
        sys.exit(f"KLAIDA: HSQLDB script failas nerastas: {script_path}")
    if not os.path.exists(properties_path):
        sys.exit(f"KLAIDA: HSQLDB properties failas nerastas: {properties_path}")

    today = time.strftime("%Y-%m-%d")

    print(f"Pakuojama i {OUTPUT_FILE}...")

    # SVARBU: mimetype turi buti PIRMAS faile ZIP archyve ir STORED (be kompresijos)
    # Tai oficialus ODF reikalavimas, kad failas butu atpazintas kaip Open Document.
    with zipfile.ZipFile(OUTPUT_FILE, "w") as zf:
        # 1. mimetype - pirmas, be kompresijos
        info = zipfile.ZipInfo("mimetype")
        info.compress_type = zipfile.ZIP_STORED
        zf.writestr(info, MIMETYPE)

        # 2. META-INF/manifest.xml
        zf.writestr("META-INF/manifest.xml", MANIFEST_XML,
                    compress_type=zipfile.ZIP_DEFLATED)

        # 3. content.xml, meta.xml, settings.xml, styles.xml
        zf.writestr("content.xml", CONTENT_XML,
                    compress_type=zipfile.ZIP_DEFLATED)
        zf.writestr("meta.xml", META_XML.format(date=today),
                    compress_type=zipfile.ZIP_DEFLATED)
        zf.writestr("settings.xml", SETTINGS_XML,
                    compress_type=zipfile.ZIP_DEFLATED)
        zf.writestr("styles.xml", STYLES_XML,
                    compress_type=zipfile.ZIP_DEFLATED)

        # 4. HSQLDB duomenu bazes failai
        with open(script_path, "rb") as f:
            zf.writestr("database/script", f.read(),
                        compress_type=zipfile.ZIP_DEFLATED)
        with open(properties_path, "rb") as f:
            zf.writestr("database/properties", f.read(),
                        compress_type=zipfile.ZIP_DEFLATED)

    # Patikrinam rezultata
    size = os.path.getsize(OUTPUT_FILE)
    print(f"Sukurta: {OUTPUT_FILE} ({size} B)")
    print("Failai archyve:")
    with zipfile.ZipFile(OUTPUT_FILE, "r") as zf:
        for info in zf.infolist():
            comp = "STORED" if info.compress_type == zipfile.ZIP_STORED else "DEFLATED"
            print(f"  {info.filename:30s}  {info.file_size:>6d} B  {comp}")


# ============================================================
# PAGRINDINE
# ============================================================

def main():
    print("=" * 60)
    print("LibreOffice Base failo (studentai.odb) generavimas")
    print("=" * 60)
    run_java_create_db()
    build_odb_zip()
    print()
    print("Baigta! Atidarykite studentai.odb su LibreOffice Base.")
    print("Vartotojo vardas: SA, slaptazodis tuscias.")


if __name__ == "__main__":
    main()
