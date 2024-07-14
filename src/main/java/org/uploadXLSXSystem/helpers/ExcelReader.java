package org.uploadXLSXSystem.helpers;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.uploadXLSXSystem.models.CellRequest;

import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelReader {

    public Map<String , Object> readExcelData(InputStream inputStream ,
                                                     List<CellRequest> cellRequests)
    {
        Map<String, Object> results = new HashMap<>();

        try(Workbook workbook = new XSSFWorkbook(inputStream))
        {
            /*
             * I will search for each cell in the Excel based on the label of the cellRequests
             */
            cellRequests.forEach(request -> {
                Sheet sheet = workbook.getSheet(request.getSheetName());
                String adjacentCellValue = findAndReadAdjacentByLabel(sheet, request.getCellLabel(), request.getDirection());
                if (adjacentCellValue != null) {
                    /*
                     * If I found the value for the label I wil save it in my map
                     * EG : key : Segment General!LRN
                     */
                    results.put(request.getSheetName() + "!" + request.getCellLabel(), adjacentCellValue);

                }
            });

        } catch (Exception e) {
            e.printStackTrace();
        }
        return results;
    }

    private String findAndReadAdjacentByLabel(Sheet sheet, String cellLabel, String direction) {
        /*
         * Iterate over all the rows and columns (cells)
         * - If I find the correctLabel for my cellRequest
         * ...
         */
        for(Row row : sheet)
        {
            for (Cell cell : row)
            {
                if (getCellStringValue(cell).equalsIgnoreCase(cellLabel)) {
                    CellRangeAddress mergedRegion = getMergedRegion(sheet, cell);

                    if ("right".equals(direction) || "right-boolean".equals(direction)) {
                        int colIndexRight = (mergedRegion != null) ? mergedRegion.getLastColumn() + 1 : cell.getColumnIndex() + 1;
                        Cell firstRightCell = row.getCell(colIndexRight, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        Cell secondRightCell = row.getCell(colIndexRight + 1, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

                        if ("right".equals(direction)) {
                            if (firstRightCell != null) {
                                return getCellStringValue(firstRightCell);
                            }
                        } else { // "right-boolean"
                            String firstCellValue = getCellStringValue(firstRightCell).trim();
                            String secondCellValue = getCellStringValue(secondRightCell).trim();

                            // Return "true" if the first cell to the right is non-empty and not just spaces
                            if (!firstCellValue.isEmpty() && !firstCellValue.equals("    ")) {
                                return "true";
                            }
                            // Return "false" if the second cell to the right meets the criteria, assuming first cell did not
                            else if (!secondCellValue.isEmpty() && !secondCellValue.equals("    ")) {
                                return "false";
                            }
                            // Optional: Return a default value if neither cell meets the criteria
                        }
                    }

                }
            }
        }
        return null;
    }

    private static String getCellStringValue(Cell cell) {
        if (cell == null) return "";

        /*
         * The DataFormatter it converts the data
         * from the cell into a string (even Numbers and Dates)
         */
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell);
    }

    private static CellRangeAddress getMergedRegion(Sheet sheet, Cell cell) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.isInRange(cell.getRowIndex(), cell.getColumnIndex())) {
                return merged;
            }
        }
        return null; // Return null if the cell is not within a merged region.
    }

}
