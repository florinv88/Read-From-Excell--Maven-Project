package org.uploadXLSXSystem.models;

public class CellRequest {

    private final String sheetName;
    private final String cellLabel;
    private final String direction; // "right", "bottom" or "right-boolean" (2 adjacent cells with true/false)

    public CellRequest(String sheetName, String cellLabel, String direction) {
        this.sheetName = sheetName;
        this.cellLabel = cellLabel;
        this.direction = direction;
    }

    public String getSheetName() {
        return sheetName;
    }

    public String getCellLabel() {
        return cellLabel;
    }

    public String getDirection() {
        return direction;
    }
}
