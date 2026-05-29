"""
generate_odb_pu5.py

Sukuria tris realius LibreOffice Base failus (.odb):
  - eprekyba.odb
  - ligonine.odb
  - biblioteka.odb

Kiekvienam atlieka:
  1. Java + HSQLDB 1.8.0.10 sukuria duomenu bazes failus (script + properties)
  2. ODF metaduomenys + HSQLDB failai supakuojami i ZIP archyva su .odb pletiniu
"""

import os
import shutil
import subprocess
import sys
import zipfile
import time

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
HSQLDB_JAR = os.path.expanduser(
    "~/.m2/repository/org/hsqldb/hsqldb/1.8.0.10/hsqldb-1.8.0.10.jar"
)

DATABASES = [
    ("eprekyba_db", "eprekyba.odb"),
    ("ligonine_db", "ligonine.odb"),
    ("biblioteka_db", "biblioteka.odb"),
]

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
  <meta:generator>PU5_VU_Andrius_Vargonas/2026.05</meta:generator>
  <meta:creation-date>{date}T00:00:00</meta:creation-date>
  <dc:date>{date}T00:00:00</dc:date>
  <dc:title>{title}</dc:title>
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


def run_java():
    """Paleidzia CreateDB_pu5.java - sukuria 3 db duomenis."""
    if not os.path.exists(HSQLDB_JAR):
        sys.exit("KLAIDA: HSQLDB 1.8.0.10 jar nerastas. Paleiskite:\n"
                 "  mvn dependency:get -Dartifact=org.hsqldb:hsqldb:1.8.0.10")

    # Pasalinti senas DB direktorijas
    for dbdir, _ in DATABASES:
        full = os.path.join(SCRIPT_DIR, dbdir)
        if os.path.exists(full):
            shutil.rmtree(full)

    # Kompiliuoti CreateDB_pu5.java
    java_file = os.path.join(SCRIPT_DIR, "CreateDB_pu5.java")
    class_file = os.path.join(SCRIPT_DIR, "CreateDB_pu5.class")
    if not os.path.exists(class_file) or \
            os.path.getmtime(java_file) > os.path.getmtime(class_file):
        print("Kompiliuojama CreateDB_pu5.java...")
        result = subprocess.run(
            ["javac", "-cp", HSQLDB_JAR, java_file],
            capture_output=True, text=True, cwd=SCRIPT_DIR,
        )
        if result.returncode != 0:
            sys.exit("Kompiliavimo klaida:\n" + result.stderr)

    print("Paleidziama Java programa...")
    result = subprocess.run(
        ["java", "-cp", f".:{HSQLDB_JAR}", "CreateDB_pu5", SCRIPT_DIR],
        capture_output=True, text=True, cwd=SCRIPT_DIR,
    )
    print(result.stdout)
    if result.returncode != 0:
        sys.exit("Java vykdymo klaida:\n" + result.stderr)


def build_odb(db_dir, output_filename, title):
    """Supakuoja viena duomenu bazes direktorija i .odb (ZIP) archyva."""
    script_path = os.path.join(SCRIPT_DIR, db_dir, "database", "script")
    properties_path = os.path.join(SCRIPT_DIR, db_dir, "database", "properties")

    if not os.path.exists(script_path):
        sys.exit(f"KLAIDA: nerastas {script_path}")

    today = time.strftime("%Y-%m-%d")
    output_path = os.path.join(SCRIPT_DIR, output_filename)

    with zipfile.ZipFile(output_path, "w") as zf:
        info = zipfile.ZipInfo("mimetype")
        info.compress_type = zipfile.ZIP_STORED
        zf.writestr(info, MIMETYPE)
        zf.writestr("META-INF/manifest.xml", MANIFEST_XML,
                    compress_type=zipfile.ZIP_DEFLATED)
        zf.writestr("content.xml", CONTENT_XML,
                    compress_type=zipfile.ZIP_DEFLATED)
        zf.writestr("meta.xml", META_XML.format(date=today, title=title),
                    compress_type=zipfile.ZIP_DEFLATED)
        zf.writestr("settings.xml", SETTINGS_XML,
                    compress_type=zipfile.ZIP_DEFLATED)
        zf.writestr("styles.xml", STYLES_XML,
                    compress_type=zipfile.ZIP_DEFLATED)
        with open(script_path, "rb") as f:
            zf.writestr("database/script", f.read(),
                        compress_type=zipfile.ZIP_DEFLATED)
        with open(properties_path, "rb") as f:
            zf.writestr("database/properties", f.read(),
                        compress_type=zipfile.ZIP_DEFLATED)

    size = os.path.getsize(output_path)
    print(f"  Sukurta: {output_filename} ({size} B)")


def main():
    print("=" * 60)
    print("PU5: 3 LibreOffice Base failu generavimas")
    print("=" * 60)
    run_java()
    print("\nPakuojami .odb failai:")
    titles = {
        "eprekyba_db": "E-prekybos duomenu baze (PU5)",
        "ligonine_db": "Ligonines duomenu baze (PU5)",
        "biblioteka_db": "Bibliotekos duomenu baze (PU5)",
    }
    for db_dir, output_name in DATABASES:
        build_odb(db_dir, output_name, titles[db_dir])
    print("\nBaigta. Atidarykite kiekviena .odb su LibreOffice Base.")


if __name__ == "__main__":
    main()
