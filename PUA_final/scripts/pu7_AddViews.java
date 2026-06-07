/*
 * AddViews.java - Reads existing HSQLDB database in baseDir/database
 * and adds the CREATE VIEW statements for the given variant (db2, db4, db6),
 * then issues SHUTDOWN COMPACT so HSQLDB rewrites the script with views.
 */
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.Statement;
import java.io.File;

public class AddViews {

    static final String[] DB2_VIEWS = new String[] {
        "CREATE VIEW QRYDEPARTMENT AS SELECT PRODUCTCODE, SUPPLIERID FROM TBLINVENTORY WHERE DEPT = 'Hardware'",
        "CREATE VIEW QRYSUPPLIER AS SELECT PRODUCTCODE, SUPPLIERID, ITEMDESCRIPTION FROM TBLINVENTORY WHERE DEPT = 'Bathroom'",
        "CREATE VIEW QRYDATE AS SELECT PRODUCTCODE, DEPT, SUPPLIERID FROM TBLINVENTORY WHERE SUPPLIERID = 'WOODSTOCK' AND LASTORDERED = '2015-05-21'"
    };

    static final String[] DB4_VIEWS = new String[] {
        "CREATE VIEW QRYDUPLICATEPRODUCTS AS SELECT ITEMDESCRIPTION, COUNT(*) AS CNT FROM TBLINVENTORY GROUP BY ITEMDESCRIPTION HAVING COUNT(*) > 1",
        "CREATE VIEW QRYFINDSUPPLIER AS SELECT SUPPLIERID, COMPANY, CITY, COUNTRY FROM TBLSUPPLIERS",
        "CREATE VIEW QRYORIGIN AS SELECT ORIGIN, COUNT(*) AS PRODUCTCOUNT FROM TBLINVENTORY GROUP BY ORIGIN",
        "CREATE VIEW QRYPRODUCTSUPPLIERDETAIL AS SELECT I.PRODUCTCODE, I.ITEMDESCRIPTION, I.DEPT, S.COMPANY, S.COUNTRY FROM TBLINVENTORY I JOIN TBLSUPPLIERS S ON I.SUPPLIERID = S.SUPPLIERID",
        "CREATE VIEW QRYPROJECTPACKS AS SELECT PRODUCTCODE, ITEMDESCRIPTION, RETAILPRICE FROM TBLINVENTORY WHERE PRODUCTCODE LIKE 'pak-%'",
        "CREATE VIEW QRYREORDERDATE AS SELECT PRODUCTCODE, ITEMDESCRIPTION, LASTORDERED, REORDERLEVEL, UNITSINSTOCK FROM TBLINVENTORY ORDER BY LASTORDERED DESC",
        "CREATE VIEW QRYREORDERNOW AS SELECT PRODUCTCODE, ITEMDESCRIPTION, UNITSINSTOCK, REORDERLEVEL FROM TBLINVENTORY WHERE UNITSINSTOCK <= REORDERLEVEL",
        "CREATE VIEW QRYUNUSEDSUPPLIERS AS SELECT SUPPLIERID, COMPANY, CITY, COUNTRY FROM TBLSUPPLIERS WHERE COUNTRY = 'Canada'"
    };

    static final String[] DB6_EXTRA_VIEWS = new String[] {
        "CREATE VIEW QRYSUMBYDEPARTMENT AS SELECT DEPT, SUM(UNITSINSTOCK * RETAILPRICE) AS TOTALVALUE FROM TBLINVENTORY GROUP BY DEPT",
        "CREATE VIEW QRYAVGBYORIGIN AS SELECT ORIGIN, AVG(RETAILPRICE) AS AVGPRICE FROM TBLINVENTORY GROUP BY ORIGIN",
        "CREATE VIEW QRYCOUNTBYSUPPLIER AS SELECT SUPPLIERID, COUNT(*) AS PRODUCTCOUNT FROM TBLINVENTORY GROUP BY SUPPLIERID"
    };

    public static void main(String[] args) throws Exception {
        if (args.length < 2) {
            System.err.println("Usage: AddViews <db_dir> <variant: db2|db4|db5|db6>");
            System.exit(1);
        }
        String baseDir = args[0];
        String variant = args[1];

        // HSQLDB looks for files at <path>.script, <path>.properties etc.
        // The directory has files named "script" and "properties" (no "db." prefix).
        // We rename them temporarily so HSQLDB can open the database.
        File dir = new File(baseDir);
        File scriptIn = new File(dir, "script");
        File propsIn = new File(dir, "properties");
        File scriptHsql = new File(dir, "db.script");
        File propsHsql = new File(dir, "db.properties");
        if (!scriptIn.renameTo(scriptHsql)) throw new RuntimeException("Cannot rename script");
        if (!propsIn.renameTo(propsHsql)) throw new RuntimeException("Cannot rename properties");

        Class.forName("org.hsqldb.jdbcDriver");
        String dbPath = new File(dir, "db").getAbsolutePath();
        String url = "jdbc:hsqldb:file:" + dbPath + ";shutdown=true";
        Connection conn = DriverManager.getConnection(url, "SA", "");
        Statement st = conn.createStatement();
        st.execute("SET WRITE_DELAY FALSE");

        String[] views;
        if ("db2".equals(variant)) views = DB2_VIEWS;
        else if ("db4".equals(variant) || "db5".equals(variant)) views = DB4_VIEWS;
        else if ("db6".equals(variant)) {
            views = new String[DB4_VIEWS.length + DB6_EXTRA_VIEWS.length];
            System.arraycopy(DB4_VIEWS, 0, views, 0, DB4_VIEWS.length);
            System.arraycopy(DB6_EXTRA_VIEWS, 0, views, DB4_VIEWS.length, DB6_EXTRA_VIEWS.length);
        } else throw new RuntimeException("Unknown variant: " + variant);

        for (String sql : views) {
            try {
                st.execute(sql);
                System.out.println("OK: " + sql.substring(0, Math.min(70, sql.length())));
            } catch (Exception e) {
                System.err.println("FAIL: " + sql);
                System.err.println("   -> " + e.getMessage());
                throw e;
            }
        }

        st.execute("SHUTDOWN COMPACT");
        st.close();
        conn.close();

        // Rename back to script/properties
        if (!scriptHsql.renameTo(scriptIn)) throw new RuntimeException("Cannot rename db.script back");
        if (!propsHsql.renameTo(propsIn)) throw new RuntimeException("Cannot rename db.properties back");
        // Cleanup any leftover files
        new File(dir, "db.lck").delete();
        new File(dir, "db.log").delete();
        new File(dir, "db.data").delete();
        new File(dir, "db.backup").delete();
        System.out.println("Baigta variantui " + variant + " kataloge " + baseDir);
    }
}
