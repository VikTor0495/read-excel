import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class Main {
    public static void main(String args[]) throws IOException {
        Workbook wb = new XSSFWorkbook(new FileInputStream("src/main/resources/TABELLE ATECO.xlsx"));

        Sheet sheet = wb.getSheetAt(1);

        FileOutputStream fos = new FileOutputStream(new File("src/main/resources/output.txt"));

        for (Row row : sheet) {
            fos.write(
                    "INSERT INTO CODICI_ATECO (CODICE, DESCRIZIONE, REDDITIVITA, GESTIONE_PREVIDENZIALE, MACROCATEGORIA) VALUES ("
                            .getBytes());
            for (Cell cell : row) {

                switch (cell.getCellType()) {
                    case STRING:
                        fos.write("'".getBytes());
                        fos.write(String.valueOf(cell.getStringCellValue()).getBytes());
                        fos.write("'".getBytes());
                        break;
                    case NUMERIC:
                        fos.write(String.valueOf(cell.getNumericCellValue()).getBytes());
                        break;
                    case BOOLEAN:
                        fos.write(String.valueOf(cell.getBooleanCellValue()).getBytes());
                        break;
                    default:
                        fos.write("''".getBytes());
                        break;
                }
                fos.write(",".getBytes());
            }
            fos.write(") \n".getBytes());
        }
    }
}
