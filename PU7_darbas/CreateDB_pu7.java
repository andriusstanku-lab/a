/*
 * CreateDB_pu7.java
 */
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.Statement;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.util.ArrayList;
import java.util.List;

public class CreateDB_pu7 {

    public static void main(String[] args) throws Exception {
        String baseDir = args.length > 0 ? args[0] : ".";
        String csvDir = args.length > 1 ? args[1] : ".";
        Class.forName("org.hsqldb.jdbcDriver");

        File dir = new File(baseDir + "/dbinventory_db/database");
        dir.mkdirs();
        String dbPath = new File(baseDir + "/dbinventory_db/database/db").getAbsolutePath();
        String url = "jdbc:hsqldb:file:" + dbPath + ";shutdown=true";
        Connection conn = DriverManager.getConnection(url, "SA", "");
        Statement st = conn.createStatement();
        st.execute("SET WRITE_DELAY FALSE");

        st.execute("CREATE MEMORY TABLE tblDepartments (Department VARCHAR(50) NOT NULL PRIMARY KEY)");
        System.out.println("Sukurta: tblDepartments");

        st.execute("CREATE MEMORY TABLE tblSuppliers (" +
            "SupplierID VARCHAR(30) NOT NULL PRIMARY KEY," +
            "FirstName VARCHAR(50)," +
            "LastName VARCHAR(50)," +
            "ContactPhone VARCHAR(30)," +
            "Company VARCHAR(100)," +
            "ContactEmail VARCHAR(100)," +
            "Address VARCHAR(150)," +
            "City VARCHAR(60)," +
            "StateProvince VARCHAR(40)," +
            "Country VARCHAR(40)," +
            "PostalCode VARCHAR(20))");
        System.out.println("Sukurta: tblSuppliers");

        st.execute("CREATE MEMORY TABLE tblInventory (" +
            "ProductCode VARCHAR(30) NOT NULL PRIMARY KEY," +
            "Dept VARCHAR(50)," +
            "SupplierID VARCHAR(30)," +
            "ItemDescription VARCHAR(255)," +
            "Location VARCHAR(50)," +
            "Rack VARCHAR(20)," +
            "Origin VARCHAR(40)," +
            "UnitsInStock INTEGER," +
            "TargetInventory INTEGER," +
            "ReorderLevel INTEGER," +
            "LastOrdered DATE," +
            "OurUnitCost DECIMAL(10,2)," +
            "RetailPrice DECIMAL(10,2)," +
            "FOREIGN KEY (Dept) REFERENCES tblDepartments(Department)," +
            "FOREIGN KEY (SupplierID) REFERENCES tblSuppliers(SupplierID))");
        System.out.println("Sukurta: tblInventory");

        importDepartments(conn, csvDir + "/DB4_tblDepartments.csv");
        importSuppliers(conn, csvDir + "/DB4_tblSuppliers.csv");
        importInventory(conn, csvDir + "/DB4_tblInventory.csv");

        st.execute("SHUTDOWN COMPACT");
        st.close();
        conn.close();

        File dbDir = new File(baseDir + "/dbinventory_db/database");
        new File(dbDir, "db.script").renameTo(new File(dbDir, "script"));
        new File(dbDir, "db.properties").renameTo(new File(dbDir, "properties"));
        new File(dbDir, "db.lck").delete();
        new File(dbDir, "db.log").delete();
        new File(dbDir, "db.data").delete();
        new File(dbDir, "db.backup").delete();
        System.out.println("Baigta.");
    }

    private static List<String[]> readCsv(String path) throws Exception {
        List<String[]> rows = new ArrayList<>();
        BufferedReader br = new BufferedReader(new FileReader(path));
        String line;
        while ((line = br.readLine()) != null) {
            List<String> fields = new ArrayList<>();
            StringBuilder cur = new StringBuilder();
            boolean inQuote = false;
            for (int i = 0; i < line.length(); i++) {
                char c = line.charAt(i);
                if (inQuote) {
                    if (c == '"') {
                        if (i + 1 < line.length() && line.charAt(i + 1) == '"') {
                            cur.append('"'); i++;
                        } else { inQuote = false; }
                    } else { cur.append(c); }
                } else {
                    if (c == ',') { fields.add(cur.toString()); cur.setLength(0); }
                    else if (c == '"') inQuote = true;
                    else cur.append(c);
                }
            }
            fields.add(cur.toString());
            rows.add(fields.toArray(new String[0]));
        }
        br.close();
        return rows;
    }

    private static void importDepartments(Connection conn, String path) throws Exception {
        List<String[]> rows = readCsv(path);
        PreparedStatement ps = conn.prepareStatement("INSERT INTO tblDepartments VALUES(?)");
        int n = 0;
        for (int i = 1; i < rows.size(); i++) {
            ps.setString(1, rows.get(i)[0]);
            ps.addBatch(); n++;
        }
        ps.executeBatch();
        ps.close();
        System.out.println("  Departments: " + n);
    }

    private static void importSuppliers(Connection conn, String path) throws Exception {
        List<String[]> rows = readCsv(path);
        PreparedStatement ps = conn.prepareStatement(
            "INSERT INTO tblSuppliers VALUES(?,?,?,?,?,?,?,?,?,?,?)");
        int n = 0;
        for (int i = 1; i < rows.size(); i++) {
            String[] r = rows.get(i);
            for (int j = 0; j < 11; j++) {
                ps.setString(j + 1, j < r.length ? r[j] : null);
            }
            ps.addBatch(); n++;
        }
        ps.executeBatch();
        ps.close();
        System.out.println("  Suppliers: " + n);
    }

    private static void importInventory(Connection conn, String path) throws Exception {
        List<String[]> rows = readCsv(path);
        PreparedStatement ps = conn.prepareStatement(
            "INSERT INTO tblInventory (ProductCode,Dept,SupplierID,ItemDescription,Location,Rack,Origin,UnitsInStock,TargetInventory,ReorderLevel,LastOrdered,OurUnitCost,RetailPrice) " +
            "VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)");
        int n = 0;
        for (int i = 1; i < rows.size(); i++) {
            String[] r = rows.get(i);
            try {
                int unitsInStock = Integer.parseInt(r[0].trim());
                int targetInventory = Integer.parseInt(r[1].trim());
                int reorderLevel = Integer.parseInt(r[2].trim());
                String lastOrdered = r[3].trim();
                if (lastOrdered.contains(" ")) lastOrdered = lastOrdered.split(" ")[0];
                String ourCostStr = r[4].replace("$", "").replace(",", "").trim();
                String retailStr = r[5].replace("$", "").replace(",", "").trim();
                double ourCost = ourCostStr.isEmpty() ? 0 : Double.parseDouble(ourCostStr);
                double retail = retailStr.isEmpty() ? 0 : Double.parseDouble(retailStr);
                ps.setString(1, r[6]);
                ps.setString(2, r[7]);
                ps.setString(3, r[8]);
                ps.setString(4, r[9]);
                ps.setString(5, r[10]);
                ps.setString(6, r[11]);
                ps.setString(7, r[12]);
                ps.setInt(8, unitsInStock);
                ps.setInt(9, targetInventory);
                ps.setInt(10, reorderLevel);
                ps.setDate(11, java.sql.Date.valueOf(lastOrdered));
                ps.setDouble(12, ourCost);
                ps.setDouble(13, retail);
                ps.addBatch();
                n++;
            } catch (Exception e) {
                System.err.println("  Praleidziama eilute " + i + ": " + e.getMessage());
            }
        }
        ps.executeBatch();
        ps.close();
        System.out.println("  Inventory: " + n);
    }
}
