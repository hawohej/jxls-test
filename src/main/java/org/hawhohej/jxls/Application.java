package org.hawhohej.jxls;

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.regex.Pattern;

public class Application {

    public static void main(String[] args) throws IOException {
        var objectMapper = new ObjectMapper();

        // Read files from resources
        var classLoader = Application.class;
        var jsonDataInputStream = classLoader.getClassLoader().getResourceAsStream("data.json");
        var templateInputStream = classLoader.getClassLoader().getResourceAsStream("template.xlsx");

        // Process template with data
        var jsonData = objectMapper.readValue(jsonDataInputStream, Map.class);
        var out = processDocument(templateInputStream, jsonData);

        // Create result file and generate report
        DateTimeFormatter dateTimeFormat = DateTimeFormatter.ofPattern("yyyy-MM-dd_hh-mm-ss");
        String fileName = String.format("src/main/resources/generated/template_%s.xlsx", LocalDateTime.now().format(dateTimeFormat));
        try (var file = new FileOutputStream(fileName)) {
            file.write(out.toByteArray());
        }
    }

    private static ByteArrayOutputStream processDocument(InputStream templateFile, Map<String, Object> jsonData) throws IOException {
        var printFormOutputStream = new ByteArrayOutputStream();
        var context = new Context(jsonData);
        JxlsHelper.getInstance()
                .setProcessFormulas(true)
                .setEvaluateFormulas(true)
                .processTemplate(templateFile, printFormOutputStream, context);
        return printFormOutputStream;
    }

    private static List<String> extractTemplateFields(InputStream templateInputStream) {
        try {
            var templateFields = new ArrayList<String>();

            Workbook workbook = new XSSFWorkbook(templateInputStream);
            Sheet sheet = workbook.getSheetAt(0);

            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        templateFields.addAll(extractVariables(cell.getStringCellValue()));
                    }
                }
            }

            return templateFields;
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    private static Set<String> extractVariables(String cellContent) {
        var variables = new HashSet<String>();
        var pattern = Pattern.compile("\\$\\{([^${}]+)}");
        var matcher = pattern.matcher(cellContent);

        while (matcher.find()) {
            variables.add(matcher.group(1));
        }

        return variables;
    }
}
