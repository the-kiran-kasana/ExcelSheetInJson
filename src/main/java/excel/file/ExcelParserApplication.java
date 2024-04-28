package excel.file;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@SpringBootApplication
public class ExcelParserApplication {

    public static void main(String[] args) {
        SpringApplication.run(ExcelParserApplication.class, args);
    }
}

@RestController
class ExcelParserController {

    @PostMapping("/upload")
    public ResponseEntity<?> uploadFile(@RequestParam("file") MultipartFile file) {
        try {
            // Parse Excel file from uploaded MultipartFile
            List<Map<String, Object>> excelData = parseExcel(file);
            // Create JSON array to hold the row data
            JSONArray dataArray = new JSONArray();
            // Group data by date
            Map<String, List<String>> groupedData = groupDataByDate(excelData);
            // Iterate over grouped data and construct JSON objects
            for (Map.Entry<String, List<String>> entry : groupedData.entrySet()) {
                JSONObject rowObject = new JSONObject();
                rowObject.put("date", entry.getKey());
                rowObject.put("values", new JSONArray(entry.getValue()));
                dataArray.put(rowObject);
            }
            // Create JSON object with "data" array
            JSONObject jsonObject = new JSONObject();
            jsonObject.put("data", dataArray);
            return ResponseEntity.ok(jsonObject.toString());
        } catch (final IOException e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
                    .body("Exception occurred while processing Excel file. Exception : " + e.getMessage());
        }
    }

    private List<Map<String, Object>> parseExcel(MultipartFile file) throws IOException {
        List<Map<String, Object>> dataList = new ArrayList<>();
        try (Workbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0); // Assuming data is on the first sheet
            for (Row row : sheet) {
                Cell dateCell = row.getCell(0);
                Cell valueCell = row.getCell(1);
                String date = getDateCellValue(dateCell);
                String value = getValueCellValue(valueCell);
                Map<String, Object> map = new HashMap<>();
                map.put("date", date);
                map.put("value", value);
                dataList.add(map);
            }
        }
        return dataList;
    }

    private String getDateCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        return cell.toString(); 
    }

    private String getValueCellValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        return cell.toString(); 
    }

    private Map<String, List<String>> groupDataByDate(List<Map<String, Object>> dataList) {
        Map<String, List<String>> groupedData = new HashMap<>();
        for (Map<String, Object> data : dataList) {
            String date = (String) data.get("date");
            String value = (String) data.get("value");
            if (!groupedData.containsKey(date)) {
                groupedData.put(date, new ArrayList<>());
            }
            groupedData.get(date).add(value);
        }
        return groupedData;
    }
}

