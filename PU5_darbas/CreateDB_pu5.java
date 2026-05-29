/*
 * CreateDB_pu5.java
 *
 * Sukuria tris HSQLDB 1.8.0.10 embedded duomenu bazes, suderinamas su
 * LibreOffice Base:
 *   - eprekyba.odb (Klientai, Produktai, Uzsakymai)
 *   - ligonine.odb (Pacientai, Gydytojai, Vizitai)
 *   - biblioteka.odb (Knygos, Nariai, Skolinimasi)
 */
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.Statement;
import java.io.File;

public class CreateDB_pu5 {

    public static void main(String[] args) throws Exception {
        String baseDir = args.length > 0 ? args[0] : ".";
        Class.forName("org.hsqldb.jdbcDriver");

        createEPrekyba(baseDir);
        createLigonine(baseDir);
        createBiblioteka(baseDir);
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
        File scriptOld = new File(dbDir, "db.script");
        File scriptNew = new File(dbDir, "script");
        if (scriptOld.exists()) scriptOld.renameTo(scriptNew);
        File propOld = new File(dbDir, "db.properties");
        File propNew = new File(dbDir, "properties");
        if (propOld.exists()) propOld.renameTo(propNew);
        new File(dbDir, "db.lck").delete();
        new File(dbDir, "db.log").delete();
        new File(dbDir, "db.data").delete();
        new File(dbDir, "db.backup").delete();
    }

    private static void createEPrekyba(String baseDir) throws Exception {
        System.out.println("\n[1/3] Kuriama: eprekyba_db");
        Connection conn = openDb("eprekyba_db", baseDir);
        Statement st = conn.createStatement();
        st.execute("SET WRITE_DELAY FALSE");

        st.execute(
            "CREATE MEMORY TABLE Klientai (" +
            "  KlientoID INTEGER IDENTITY," +
            "  Vardas VARCHAR(100) NOT NULL," +
            "  ElPastas VARCHAR(100)," +
            "  CONSTRAINT uq_klientai_email UNIQUE (ElPastas)" +
            ")"
        );
        System.out.println("  Sukurta: Klientai");

        st.execute(
            "CREATE MEMORY TABLE Produktai (" +
            "  ProduktoID INTEGER IDENTITY," +
            "  Pavadinimas VARCHAR(100) NOT NULL," +
            "  Kaina DECIMAL(8,2)," +
            "  Kategorija VARCHAR(50)" +
            ")"
        );
        System.out.println("  Sukurta: Produktai");

        st.execute(
            "CREATE MEMORY TABLE Uzsakymai (" +
            "  UzsakymoID INTEGER IDENTITY," +
            "  KlientoID INTEGER NOT NULL," +
            "  ProduktoID INTEGER NOT NULL," +
            "  UzsakymoData DATE," +
            "  Kiekis INTEGER," +
            "  PristatymoAdresas VARCHAR(150)," +
            "  FOREIGN KEY (KlientoID) REFERENCES Klientai(KlientoID)," +
            "  FOREIGN KEY (ProduktoID) REFERENCES Produktai(ProduktoID)" +
            ")"
        );
        System.out.println("  Sukurta: Uzsakymai");

        st.execute("INSERT INTO Klientai VALUES(1,'Andrius Vargonas','andrius@example.lt')");
        st.execute("INSERT INTO Klientai VALUES(2,'Egle Kazlauskaite','egle@example.lt')");
        st.execute("INSERT INTO Klientai VALUES(3,'Tomas Petrauskas','tomas@example.lt')");
        st.execute("INSERT INTO Klientai VALUES(4,'Ruta Jonaityte','ruta@example.lt')");
        st.execute("INSERT INTO Klientai VALUES(5,'Mantas Bagdonas','mantas@example.lt')");

        st.execute("INSERT INTO Produktai VALUES(1,'Knyga Programavimas Python',25.99,'Knygos')");
        st.execute("INSERT INTO Produktai VALUES(2,'Klaviatura Logitech K380',49.99,'Elektronika')");
        st.execute("INSERT INTO Produktai VALUES(3,'Kava 250g',8.50,'Maistas')");
        st.execute("INSERT INTO Produktai VALUES(4,'Marskineliai medvilniniai',19.99,'Drabuziai')");
        st.execute("INSERT INTO Produktai VALUES(5,'Ausines Sony WH-1000XM4',289.99,'Elektronika')");

        st.execute("INSERT INTO Uzsakymai VALUES(1,1,1,'2025-01-15',2,'Vilnius g. 1, Vilnius')");
        st.execute("INSERT INTO Uzsakymai VALUES(2,2,2,'2025-01-20',1,'K. Donelaicio g. 5, Kaunas')");
        st.execute("INSERT INTO Uzsakymai VALUES(3,1,3,'2025-02-01',5,'Vilnius g. 1, Vilnius')");
        st.execute("INSERT INTO Uzsakymai VALUES(4,3,4,'2025-02-10',2,'Taikos pr. 10, Klaipeda')");
        st.execute("INSERT INTO Uzsakymai VALUES(5,4,5,'2025-02-15',1,'Tilzes g. 8, Siauliai')");

        System.out.println("  Iterpta 15 irasu");

        st.execute("SHUTDOWN COMPACT");
        st.close();
        conn.close();
        cleanup("eprekyba_db", baseDir);
    }

    private static void createLigonine(String baseDir) throws Exception {
        System.out.println("\n[2/3] Kuriama: ligonine_db");
        Connection conn = openDb("ligonine_db", baseDir);
        Statement st = conn.createStatement();
        st.execute("SET WRITE_DELAY FALSE");

        st.execute(
            "CREATE MEMORY TABLE Pacientai (" +
            "  PacientoID INTEGER IDENTITY," +
            "  PacientoVardas VARCHAR(100) NOT NULL" +
            ")"
        );
        System.out.println("  Sukurta: Pacientai");

        st.execute(
            "CREATE MEMORY TABLE Gydytojai (" +
            "  GydytojoID INTEGER IDENTITY," +
            "  GydytojoVardas VARCHAR(100) NOT NULL," +
            "  Specializacija VARCHAR(80)" +
            ")"
        );
        System.out.println("  Sukurta: Gydytojai");

        st.execute(
            "CREATE MEMORY TABLE Vizitai (" +
            "  VizitoID INTEGER IDENTITY," +
            "  PacientoID INTEGER NOT NULL," +
            "  GydytojoID INTEGER NOT NULL," +
            "  VizitoData DATE," +
            "  Diagnoze VARCHAR(150)," +
            "  GydymoPastabos VARCHAR(255)," +
            "  FOREIGN KEY (PacientoID) REFERENCES Pacientai(PacientoID)," +
            "  FOREIGN KEY (GydytojoID) REFERENCES Gydytojai(GydytojoID)" +
            ")"
        );
        System.out.println("  Sukurta: Vizitai");

        st.execute("INSERT INTO Pacientai VALUES(1,'Jonas Kazlauskas')");
        st.execute("INSERT INTO Pacientai VALUES(2,'Petras Jonaitis')");
        st.execute("INSERT INTO Pacientai VALUES(3,'Marija Petraitiene')");
        st.execute("INSERT INTO Pacientai VALUES(4,'Antanas Bagdonas')");
        st.execute("INSERT INTO Pacientai VALUES(5,'Ona Stankaite')");

        st.execute("INSERT INTO Gydytojai VALUES(1,'Dr. Linas Stankus','Kardiologas')");
        st.execute("INSERT INTO Gydytojai VALUES(2,'Dr. Daiva Vaitkute','Pediatre')");
        st.execute("INSERT INTO Gydytojai VALUES(3,'Dr. Tomas Sereika','Chirurgas')");
        st.execute("INSERT INTO Gydytojai VALUES(4,'Dr. Egle Karaliene','Neurologe')");
        st.execute("INSERT INTO Gydytojai VALUES(5,'Dr. Mindaugas Birutis','Ortopedas')");

        st.execute("INSERT INTO Vizitai VALUES(1,1,1,'2025-01-10','Hipertenzija','Skirtas gydymas vaistais')");
        st.execute("INSERT INTO Vizitai VALUES(2,2,3,'2025-01-12','Apendicitas','Skubi operacija')");
        st.execute("INSERT INTO Vizitai VALUES(3,3,2,'2025-01-15','Sloga','Vaistai 7 dienoms')");
        st.execute("INSERT INTO Vizitai VALUES(4,4,5,'2025-02-01','Sanario nudegimas','Konsultacija')");
        st.execute("INSERT INTO Vizitai VALUES(5,5,4,'2025-02-05','Migrenos priepuolis','Vaistai')");

        System.out.println("  Iterpta 15 irasu");

        st.execute("SHUTDOWN COMPACT");
        st.close();
        conn.close();
        cleanup("ligonine_db", baseDir);
    }

    private static void createBiblioteka(String baseDir) throws Exception {
        System.out.println("\n[3/3] Kuriama: biblioteka_db");
        Connection conn = openDb("biblioteka_db", baseDir);
        Statement st = conn.createStatement();
        st.execute("SET WRITE_DELAY FALSE");

        st.execute(
            "CREATE MEMORY TABLE Knygos (" +
            "  KnygosID INTEGER IDENTITY," +
            "  Pavadinimas VARCHAR(150) NOT NULL," +
            "  Autorius VARCHAR(100)," +
            "  ISBN VARCHAR(20)," +
            "  CONSTRAINT uq_isbn UNIQUE (ISBN)" +
            ")"
        );
        System.out.println("  Sukurta: Knygos");

        st.execute(
            "CREATE MEMORY TABLE Nariai (" +
            "  NarioID INTEGER IDENTITY," +
            "  NarioVardas VARCHAR(100) NOT NULL," +
            "  NarystesData DATE" +
            ")"
        );
        System.out.println("  Sukurta: Nariai");

        st.execute(
            "CREATE MEMORY TABLE Skolinimasi (" +
            "  SkolinimoID INTEGER IDENTITY," +
            "  KnygosID INTEGER NOT NULL," +
            "  NarioID INTEGER NOT NULL," +
            "  SkolinimosiData DATE," +
            "  GrazinimoData DATE," +
            "  FOREIGN KEY (KnygosID) REFERENCES Knygos(KnygosID)," +
            "  FOREIGN KEY (NarioID) REFERENCES Nariai(NarioID)" +
            ")"
        );
        System.out.println("  Sukurta: Skolinimasi");

        st.execute("INSERT INTO Knygos VALUES(1,'Programavimas Python','J. Smith','9780123456789')");
        st.execute("INSERT INTO Knygos VALUES(2,'Duomenu bazes','R. Brown','9789876543210')");
        st.execute("INSERT INTO Knygos VALUES(3,'Marketingo strategija','A. Jonaitis','9781234567890')");
        st.execute("INSERT INTO Knygos VALUES(4,'Verslo statistika','P. Kazlauskas','9783210987654')");
        st.execute("INSERT INTO Knygos VALUES(5,'Web kurimas','M. Petrauskas','9781357924680')");

        st.execute("INSERT INTO Nariai VALUES(1,'Andrius Vargonas','2024-09-01')");
        st.execute("INSERT INTO Nariai VALUES(2,'Egle Kazlauskaite','2024-09-15')");
        st.execute("INSERT INTO Nariai VALUES(3,'Tomas Petrauskas','2025-01-10')");
        st.execute("INSERT INTO Nariai VALUES(4,'Ruta Jonaityte','2025-02-01')");
        st.execute("INSERT INTO Nariai VALUES(5,'Mantas Bagdonas','2025-02-10')");

        st.execute("INSERT INTO Skolinimasi VALUES(1,1,1,'2025-03-01','2025-03-15')");
        st.execute("INSERT INTO Skolinimasi VALUES(2,2,2,'2025-03-05','2025-03-20')");
        st.execute("INSERT INTO Skolinimasi VALUES(3,3,1,'2025-03-10','2025-03-24')");
        st.execute("INSERT INTO Skolinimasi VALUES(4,4,4,'2025-03-15','2025-03-29')");
        st.execute("INSERT INTO Skolinimasi VALUES(5,5,5,'2025-03-20','2025-04-03')");

        System.out.println("  Iterpta 15 irasu");

        st.execute("SHUTDOWN COMPACT");
        st.close();
        conn.close();
        cleanup("biblioteka_db", baseDir);
    }
}
