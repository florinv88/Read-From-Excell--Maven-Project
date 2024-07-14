package org.uploadXLSXSystem.beans;

import org.uploadXLSXSystem.helpers.ExcelReader;
import org.uploadXLSXSystem.models.CellRequest;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class UploadBean {
    private static final String segmentGeneralSheet = "Segment general";
    private static final List<CellRequest> segmentGeneralCells = List.of(
            new CellRequest(segmentGeneralSheet,"LRN","right"),
            new CellRequest(segmentGeneralSheet, "MRN", "right"),
            new CellRequest(segmentGeneralSheet, "Tip", "right"),
            new CellRequest(segmentGeneralSheet, "Tip sup.", "right")
    );

    public static void main(String[] args) throws IOException {

        String excelFilePath = ".\\datafiles\\ExcelProject.xlsx";

        // After I populate the results with excelData (LABEL + VALUE)
        // I will use to map to my model the values
        Map<String , Object> results = null;

        try(FileInputStream inputStream = new FileInputStream(excelFilePath);)
        {
            ExcelReader excelReader= new ExcelReader();
            /*
             * inputStream is the low level stream
             * segmentGeneralCells --> represents the data that I want to find it in the stream
             */
            results = excelReader.readExcelData(inputStream,segmentGeneralCells);
        }
        // Check if the map is not null
        if (results != null) {
            // Iterate and print each key-value pair
            results.forEach((key, value) -> {
                System.out.println(key + " : " + value);
            });
        } else {
            System.out.println("The map is null.");
        }


    }
}
