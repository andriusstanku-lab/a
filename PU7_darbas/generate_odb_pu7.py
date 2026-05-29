"""generate_odb_pu7.py - Sukuria DB1.odb is jau egzistuojanciu HSQLDB failu."""
import os, zipfile, time

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

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
  <meta:generator>PU7_VU_Andrius_Vargonas/2026.05</meta:generator>
  <meta:creation-date>{date}T00:00:00</meta:creation-date>
  <dc:date>{date}T00:00:00</dc:date>
  <dc:title>{title}</dc:title>
  <dc:creator>Andrius Vargonas</dc:creator>
 </office:meta>
</office:document-meta>
"""

SETTINGS_XML = '<?xml version="1.0" encoding="UTF-8"?>\n<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.2"><office:settings/></office:document-settings>\n'
STYLES_XML = '<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.2"><office:styles/><office:automatic-styles/><office:master-styles/></office:document-styles>\n'

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
     <db:table-filter><db:table-filter-pattern>%</db:table-filter-pattern></db:table-filter>
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


def build_odb(db_dir, output_filename, title):
    script_path = os.path.join(SCRIPT_DIR, db_dir, "database", "script")
    properties_path = os.path.join(SCRIPT_DIR, db_dir, "database", "properties")
    today = time.strftime("%Y-%m-%d")
    output_path = os.path.join(SCRIPT_DIR, output_filename)
    with zipfile.ZipFile(output_path, "w") as zf:
        info = zipfile.ZipInfo("mimetype")
        info.compress_type = zipfile.ZIP_STORED
        zf.writestr(info, MIMETYPE)
        for fname, content in [
            ("META-INF/manifest.xml", MANIFEST_XML),
            ("content.xml", CONTENT_XML),
            ("meta.xml", META_XML.format(date=today, title=title)),
            ("settings.xml", SETTINGS_XML),
            ("styles.xml", STYLES_XML),
        ]:
            zf.writestr(fname, content, compress_type=zipfile.ZIP_DEFLATED)
        with open(script_path, "rb") as f:
            zf.writestr("database/script", f.read(), compress_type=zipfile.ZIP_DEFLATED)
        with open(properties_path, "rb") as f:
            zf.writestr("database/properties", f.read(), compress_type=zipfile.ZIP_DEFLATED)
    print(f"Sukurta: {output_filename} ({os.path.getsize(output_path)} B)")


if __name__ == "__main__":
    # Sukursim 5 odb failus is to paties HSQLDB - bet skirtingu pavadinimu,
    # kad atspindetume DB1, DB3, DB4, DB5, DB6
    titles = {
        "DB1.odb": "DB1 - Lenteliu analize (PU7)",
        "DB3.odb": "DB3 - Formu kurimas (PU7)",
        "DB4.odb": "DB4 - Uzklausu kurimas (PU7)",
        "DB5.odb": "DB5 - Ataskaitu kurimas (PU7)",
        "DB6.odb": "DB6 - Skaiciavimai (PU7)",
    }
    for name, title in titles.items():
        build_odb("dbinventory_db", name, title)
