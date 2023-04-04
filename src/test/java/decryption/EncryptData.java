package decryption;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Base64;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.testng.annotations.Test;

public class EncryptData {
    @Test
    public void main() throws IOException {
        // Open the input Excel file
        File inputFile = new File("C:\\Users\\DELL\\eclipse-workspace\\S.Grid\\Book.xlsx");
        FileInputStream fis = new FileInputStream(inputFile);

        // Create a Workbook object
        Workbook workbook = WorkbookFactory.create(fis);

        // Get the sheet with the data to encrypt
        Sheet sheet = workbook.getSheetAt(0);

        // Get the column index for the data to encrypt
        int dataColumnIndex = 1; // Index starts from 0

        // Loop through all the rows in the sheet
        for (Row row : sheet) {
            // Get the cell in the data column
            Cell cellToEncrypt = row.getCell(dataColumnIndex);

            // Check if the cell is not null and not empty
            if (cellToEncrypt != null && !cellToEncrypt.getStringCellValue().isEmpty()) {
                // Get the cell value
                String dataToEncrypt = cellToEncrypt.getStringCellValue();

                // Encode the data
                byte[] encodedBytes = Base64.getEncoder().encode(dataToEncrypt.getBytes());
                String encryptedData = new String(encodedBytes);

                // Write the encrypted data back to the same cell
                cellToEncrypt.setCellValue(encryptedData);

                // Set the column header for the encrypted data
                if (row.getRowNum() == 0) {
                    sheet.getRow(0).createCell(dataColumnIndex).setCellValue("Encrypted Data");
                }
            }
        }

        // Save the updated Excel file
        FileOutputStream fos = new FileOutputStream(inputFile);
        workbook.write(fos);
        workbook.close();
        fis.close();
        fos.close();
    }
}
