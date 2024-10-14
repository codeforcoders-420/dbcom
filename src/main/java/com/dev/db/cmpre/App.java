import javafx.application.Application;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.layout.VBox;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.Statement;
import java.util.Iterator;

public class ExcelToAccessApp extends Application {

    private static final String DB_URL = "jdbc:ucanaccess://path_to_your_sharepoint/Customdb.acc";

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Excel to Access Loader");

        Button selectFileButton = new Button("Select Excel File");
        selectFileButton.setOnAction(event -> {
            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("Open Excel File");
            fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));
            File selectedFile = fileChooser.showOpenDialog(primaryStage);
            if (selectedFile != null) {
                try {
                    loadExcelToDatabase(selectedFile);
                    compareTablesAndFindChanges();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        });

        VBox vbox = new VBox(selectFileButton);
        Scene scene = new Scene(vbox, 400, 200);
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    private void loadExcelToDatabase(File excelFile) throws Exception {
        try (Connection conn = DriverManager.getConnection(DB_URL)) {
            // Clear the currentweek_tempdata table before inserting new data
            Statement stmt = conn.createStatement();
            stmt.executeUpdate("DELETE FROM currentweek_tempdata");

            String insertSQL = "INSERT INTO currentweek_tempdata (Column1, Column2) VALUES (?, ?)";
            try (PreparedStatement pstmt = conn.prepareStatement(insertSQL);
                 FileInputStream fis = new FileInputStream(excelFile);
                 Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheetAt(0); // Assuming data is in the first sheet
                Iterator<Row> rowIterator = sheet.iterator();
                int batchSize = 1000; // Batch size for batch insert
                int count = 0;

                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    if (row.getRowNum() == 0) continue; // Skip header row

                    String col1 = row.getCell(0).getStringCellValue(); // Replace with actual columns
                    String col2 = row.getCell(1).getStringCellValue(); // Replace with actual columns

                    pstmt.setString(1, col1);
                    pstmt.setString(2, col2);
                    pstmt.addBatch();

                    if (++count % batchSize == 0) {
                        pstmt.executeBatch(); // Execute batch
                    }
                }
                pstmt.executeBatch(); // Execute any remaining batch
            }
        }
    }


    private void compareTablesAndFindChanges() throws Exception {
        try (Connection conn = DriverManager.getConnection(DB_URL)) {
            // Clear the changefile table before inserting new changes
            Statement stmt = conn.createStatement();
            stmt.executeUpdate("DELETE FROM changefile");

            // Compare the tables and find mismatches
            String sql = "INSERT INTO changefile (Column1, Column2) " +
                         "SELECT lw.Column1, lw.Column2 " +
                         "FROM lastweek_tempdata lw " +
                         "LEFT JOIN currentweek_tempdata cw " +
                         "ON lw.Column1 = cw.Column1 " +
                         "WHERE (lw.Column2 <> cw.Column2 OR cw.Column2 IS NULL)";
            stmt.executeUpdate(sql);
        }
    }
}
