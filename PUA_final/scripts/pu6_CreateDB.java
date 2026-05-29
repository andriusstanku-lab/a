/*
 * CreateDB_pu6.java
 * 3 normalizuotos duomenu bazes pagal 3NF:
 *   - studentai.odb (Students, Instructors, Courses, Enrollments)
 *   - pardavimai.odb (Customers, Products, Salespeople, Sales)
 *   - sandeliai.odb (Warehouses, Products, Suppliers, Deliveries)
 */
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.Statement;
import java.io.File;

public class CreateDB_pu6 {

    public static void main(String[] args) throws Exception {
        String baseDir = args.length > 0 ? args[0] : ".";
        Class.forName("org.hsqldb.jdbcDriver");
        createStudentai(baseDir);
        createPardavimai(baseDir);
        createSandeliai(baseDir);
    }

    private static Connection openDb(String dbDirName, String baseDir) throws Exception {
        File dir = new File(baseDir + "/" + dbDirName + "/database");
        dir.mkdirs();
        String dbPath = new File(baseDir + "/" + dbDirName + "/database/db").getAbsolutePath();
        String url = "jdbc:hsqldb:file:" + dbPath + ";shutdown=true";
        return DriverManager.getConnection(url, "SA", "");
    }

    private static void cleanup(String dbDirName, String baseDir) {
        File dbDir = new File(baseDir + "/" + dbDirName + "/database");
        new File(dbDir, "db.script").renameTo(new File(dbDir, "script"));
        new File(dbDir, "db.properties").renameTo(new File(dbDir, "properties"));
        new File(dbDir, "db.lck").delete();
        new File(dbDir, "db.log").delete();
        new File(dbDir, "db.data").delete();
        new File(dbDir, "db.backup").delete();
    }

    private static void createStudentai(String baseDir) throws Exception {
        System.out.println("\n[1/3] Kuriama: studentai_db (3NF)");
        Connection conn = openDb("studentai_db", baseDir);
        Statement st = conn.createStatement();
        st.execute("SET WRITE_DELAY FALSE");

        st.execute("CREATE MEMORY TABLE Students (StudentID INTEGER IDENTITY, StudentName VARCHAR(50) NOT NULL)");
        st.execute("CREATE MEMORY TABLE Instructors (InstructorID INTEGER IDENTITY, InstructorName VARCHAR(50) NOT NULL, InstructorPhone VARCHAR(20))");
        st.execute("CREATE MEMORY TABLE Courses (CourseID VARCHAR(10) PRIMARY KEY, CourseName VARCHAR(50) NOT NULL, InstructorID INTEGER, FOREIGN KEY (InstructorID) REFERENCES Instructors(InstructorID))");
        st.execute("CREATE MEMORY TABLE Enrollments (EnrollmentID INTEGER IDENTITY, StudentID INTEGER NOT NULL, CourseID VARCHAR(10) NOT NULL, Grade VARCHAR(2), FOREIGN KEY (StudentID) REFERENCES Students(StudentID), FOREIGN KEY (CourseID) REFERENCES Courses(CourseID))");
        System.out.println("  4 lenteles sukurtos");

        st.execute("INSERT INTO Students VALUES(1,'Alice')");
        st.execute("INSERT INTO Students VALUES(2,'Bob')");
        st.execute("INSERT INTO Students VALUES(3,'Charlie')");
        st.execute("INSERT INTO Students VALUES(4,'Diana')");
        st.execute("INSERT INTO Students VALUES(5,'Edward')");

        st.execute("INSERT INTO Instructors VALUES(1,'Dr. Smith','123-456-7890')");
        st.execute("INSERT INTO Instructors VALUES(2,'Dr. Jones','234-567-8901')");
        st.execute("INSERT INTO Instructors VALUES(3,'Dr. Brown','345-678-9012')");
        st.execute("INSERT INTO Instructors VALUES(4,'Dr. Wilson','456-789-0123')");
        st.execute("INSERT INTO Instructors VALUES(5,'Dr. Davis','567-890-1234')");

        st.execute("INSERT INTO Courses VALUES('C101','Database',1)");
        st.execute("INSERT INTO Courses VALUES('C102','Networks',2)");
        st.execute("INSERT INTO Courses VALUES('C103','AI',3)");
        st.execute("INSERT INTO Courses VALUES('C104','Web Development',4)");
        st.execute("INSERT INTO Courses VALUES('C105','Statistics',5)");

        st.execute("INSERT INTO Enrollments VALUES(1,1,'C101','A')");
        st.execute("INSERT INTO Enrollments VALUES(2,2,'C102','B')");
        st.execute("INSERT INTO Enrollments VALUES(3,1,'C103','A')");
        st.execute("INSERT INTO Enrollments VALUES(4,3,'C101','B')");
        st.execute("INSERT INTO Enrollments VALUES(5,2,'C101','A')");
        System.out.println("  20 irasu (5 i kiekviena lentele)");

        st.execute("SHUTDOWN COMPACT");
        st.close(); conn.close();
        cleanup("studentai_db", baseDir);
    }

    private static void createPardavimai(String baseDir) throws Exception {
        System.out.println("\n[2/3] Kuriama: pardavimai_db (3NF)");
        Connection conn = openDb("pardavimai_db", baseDir);
        Statement st = conn.createStatement();
        st.execute("SET WRITE_DELAY FALSE");

        st.execute("CREATE MEMORY TABLE Customers (CustomerID INTEGER IDENTITY, CustomerName VARCHAR(50) NOT NULL)");
        st.execute("CREATE MEMORY TABLE Products (ProductID VARCHAR(10) PRIMARY KEY, ProductName VARCHAR(50) NOT NULL, Price DECIMAL(8,2))");
        st.execute("CREATE MEMORY TABLE Salespeople (SalespersonID INTEGER IDENTITY, SalespersonName VARCHAR(50) NOT NULL, SalespersonPhone VARCHAR(20))");
        st.execute("CREATE MEMORY TABLE Sales (TransactionID INTEGER IDENTITY, CustomerID INTEGER NOT NULL, ProductID VARCHAR(10) NOT NULL, SalespersonID INTEGER NOT NULL, Quantity INTEGER, FOREIGN KEY (CustomerID) REFERENCES Customers(CustomerID), FOREIGN KEY (ProductID) REFERENCES Products(ProductID), FOREIGN KEY (SalespersonID) REFERENCES Salespeople(SalespersonID))");
        System.out.println("  4 lenteles sukurtos");

        st.execute("INSERT INTO Customers VALUES(101,'John Doe')");
        st.execute("INSERT INTO Customers VALUES(102,'Jane Smith')");
        st.execute("INSERT INTO Customers VALUES(103,'Mike Brown')");
        st.execute("INSERT INTO Customers VALUES(104,'Sarah Wilson')");
        st.execute("INSERT INTO Customers VALUES(105,'Tom Davis')");

        st.execute("INSERT INTO Products VALUES('P001','Laptop',1000.00)");
        st.execute("INSERT INTO Products VALUES('P002','Mouse',50.00)");
        st.execute("INSERT INTO Products VALUES('P003','Keyboard',75.00)");
        st.execute("INSERT INTO Products VALUES('P004','Monitor',300.00)");
        st.execute("INSERT INTO Products VALUES('P005','Headphones',150.00)");

        st.execute("INSERT INTO Salespeople VALUES(1,'Alice','123-456-7890')");
        st.execute("INSERT INTO Salespeople VALUES(2,'Bob','234-567-8901')");
        st.execute("INSERT INTO Salespeople VALUES(3,'Charlie','345-678-9012')");
        st.execute("INSERT INTO Salespeople VALUES(4,'Diana','456-789-0123')");
        st.execute("INSERT INTO Salespeople VALUES(5,'Edward','567-890-1234')");

        st.execute("INSERT INTO Sales VALUES(1,101,'P001',1,1)");
        st.execute("INSERT INTO Sales VALUES(2,102,'P002',2,2)");
        st.execute("INSERT INTO Sales VALUES(3,101,'P003',1,1)");
        st.execute("INSERT INTO Sales VALUES(4,103,'P001',3,1)");
        st.execute("INSERT INTO Sales VALUES(5,102,'P001',1,1)");
        System.out.println("  20 irasu (5 i kiekviena lentele)");

        st.execute("SHUTDOWN COMPACT");
        st.close(); conn.close();
        cleanup("pardavimai_db", baseDir);
    }

    private static void createSandeliai(String baseDir) throws Exception {
        System.out.println("\n[3/3] Kuriama: sandeliai_db (3NF)");
        Connection conn = openDb("sandeliai_db", baseDir);
        Statement st = conn.createStatement();
        st.execute("SET WRITE_DELAY FALSE");

        st.execute("CREATE MEMORY TABLE Warehouses (WarehouseID INTEGER IDENTITY, WarehouseLocation VARCHAR(80) NOT NULL)");
        st.execute("CREATE MEMORY TABLE Products (ProductID VARCHAR(10) PRIMARY KEY, ProductName VARCHAR(50) NOT NULL)");
        st.execute("CREATE MEMORY TABLE Suppliers (SupplierID VARCHAR(10) PRIMARY KEY, SupplierName VARCHAR(50) NOT NULL, SupplierPhone VARCHAR(20))");
        st.execute("CREATE MEMORY TABLE Deliveries (DeliveryID INTEGER IDENTITY, WarehouseID INTEGER NOT NULL, ProductID VARCHAR(10) NOT NULL, SupplierID VARCHAR(10) NOT NULL, Quantity INTEGER, DeliveryDate DATE, FOREIGN KEY (WarehouseID) REFERENCES Warehouses(WarehouseID), FOREIGN KEY (ProductID) REFERENCES Products(ProductID), FOREIGN KEY (SupplierID) REFERENCES Suppliers(SupplierID))");
        System.out.println("  4 lenteles sukurtos");

        st.execute("INSERT INTO Warehouses VALUES(1,'New York')");
        st.execute("INSERT INTO Warehouses VALUES(2,'Los Angeles')");
        st.execute("INSERT INTO Warehouses VALUES(3,'Chicago')");
        st.execute("INSERT INTO Warehouses VALUES(4,'Houston')");
        st.execute("INSERT INTO Warehouses VALUES(5,'Miami')");

        st.execute("INSERT INTO Products VALUES('P001','Laptop')");
        st.execute("INSERT INTO Products VALUES('P002','Mouse')");
        st.execute("INSERT INTO Products VALUES('P003','Keyboard')");
        st.execute("INSERT INTO Products VALUES('P004','Monitor')");
        st.execute("INSERT INTO Products VALUES('P005','Headphones')");

        st.execute("INSERT INTO Suppliers VALUES('S001','TechCorp','123-456-7890')");
        st.execute("INSERT INTO Suppliers VALUES('S002','SupplyCo','234-567-8901')");
        st.execute("INSERT INTO Suppliers VALUES('S003','KeyMasters','345-678-9012')");
        st.execute("INSERT INTO Suppliers VALUES('S004','DisplayInc','456-789-0123')");
        st.execute("INSERT INTO Suppliers VALUES('S005','AudioPro','567-890-1234')");

        st.execute("INSERT INTO Deliveries VALUES(1,1,'P001','S001',50,'2025-04-01')");
        st.execute("INSERT INTO Deliveries VALUES(2,2,'P002','S002',200,'2025-04-02')");
        st.execute("INSERT INTO Deliveries VALUES(3,1,'P003','S003',100,'2025-04-03')");
        st.execute("INSERT INTO Deliveries VALUES(4,3,'P001','S001',75,'2025-04-04')");
        st.execute("INSERT INTO Deliveries VALUES(5,2,'P001','S001',50,'2025-04-05')");
        System.out.println("  20 irasu (5 i kiekviena lentele)");

        st.execute("SHUTDOWN COMPACT");
        st.close(); conn.close();
        cleanup("sandeliai_db", baseDir);
    }
}
