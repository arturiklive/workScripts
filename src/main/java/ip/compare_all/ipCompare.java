package ip.compare_all;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ipCompare {
    public static void main(String[] args) {
        String fileName = "C:\\src\\import.xlsx";
        int columnIndex = 2; // index of column C (starting from 0)
        Map<String, String[]> duplicates = new HashMap<>();
        try (FileInputStream excelFile = new FileInputStream(fileName);
             XSSFWorkbook workbook = new XSSFWorkbook(excelFile)) {
            Sheet sheet = workbook.getSheetAt(0); // get first sheet
            for (Row row : sheet) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null) {
                    String value = getStringValue(cell);
                    if (!value.isEmpty()) {
                        String[] rowValues = new String[2];
                        rowValues[0] = getStringValue(row.getCell(0));
                        rowValues[1] = getStringValue(row.getCell(1));
                        if (duplicates.containsKey(value)) {
                            String[] firstRowValues = duplicates.get(value);
                            System.out.printf("IP adrese \"%s\", \"%s\" и \"%s\" tika atrasts failā ar nosaukumu \"%s\".xlsx, un ar datumu \"%s\"%n",
                                    value, firstRowValues[0], firstRowValues[1], rowValues[0], rowValues[1] );
                        } else {
                            duplicates.put(value, rowValues);
                        }
                    }
                }
            }
        } catch (IOException e) {
            System.err.println("Kļūda lasot Excel failu: " + e.getMessage());
        }
    }

    private static String getStringValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
}
