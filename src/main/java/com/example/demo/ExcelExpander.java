package com.example.demo;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelExpander {

    /**
     * 从输入 Excel 读取数据，根据 Q 列拆分系统类型，生成多行，并写入新 Excel
     */
    public static void expandAndGenerate(String inputPath, String outputPath) throws IOException {
        List<ExcelRowData> allExpandedRows = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(inputPath);
             Workbook inputWorkbook = new XSSFWorkbook(fis)) {

            // 遍历所有 Sheet（或者指定 Sheet）
            for (Sheet sheet : inputWorkbook) {
                String sheetName = sheet.getSheetName();
                System.out.println("正在处理 Sheet: " + sheetName);
                if(sheetName.equals("BRD & EPIC")) {

                    // 遍历每一行（从第2行开始，假设第1行是表头）
                    for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
                        Row row = sheet.getRow(rowNum);
                        if (row == null) continue;

                        // 读取 E 列（索引4）和 Q 列（索引16）
                        Cell eCell = row.getCell(4);  // E列
                        Cell qCell = row.getCell(15); // P列

                        String eValue = getCellValueAsString(eCell);
                        String qValue = getCellValueAsString(qCell);

                        // 条件：E列有值，Q列有值且包含多个系统类型&& !qValue.equals("system type")
                        if (isNotBlank(eValue) && isNotBlank(qValue) ) {
                            // 拆分 Q 列，如 "Android & Front-End & QA" → ["Android", "Front-End", "QA"]
                            List<String> systemTypes = Arrays.stream(qValue.split("&"))
                                    .map(String::trim)
                                    .filter(s -> !s.isEmpty())
                                    .collect(Collectors.toList());
                            if (!systemTypes.isEmpty()) {
                                // 获取原始行所有列的数据（列索引 -> 值）
                                Map<Integer, String> originalValues = getRowValuesAsMap(row);

                                // 为每个系统类型生成一行
                                for (int i = 0; i < systemTypes.size(); i++) {
                                    String systemType = systemTypes.get(i);
                                    int systemIndex = i + 1;
                                    allExpandedRows.add(new ExcelRowData(originalValues, systemType, systemIndex));
                                }
                            }
                        }
                    }
                }
            }
        }

        // 写入新的 Excel 文件
        writeExpandedExcel(outputPath, allExpandedRows);
    }

    /**
     * 将所有展开的数据写入新的 Excel 文件
     */
    private static void writeExpandedExcel(String outputPath, List<ExcelRowData> expandedRows) throws IOException {
        try (Workbook outputWorkbook = new XSSFWorkbook()) {
            Sheet outputSheet = outputWorkbook.createSheet("Expanded Data");

            // 创建表头
            Row headerRow = outputSheet.createRow(0);
            headerRow.createCell(0).setCellValue("序号");
            headerRow.createCell(1).setCellValue("系统类型");
            // 假设你还要保留原数据列，比如 A(0) ~ P(15)，这里可以根据需求设置更多表头
            // 比如：headerRow.createCell(2).setCellValue("E列内容");
            // 你可以根据 originalValues 的列索引来动态设置表头

            int dataRowIndex = 1;
            for (ExcelRowData rowData : expandedRows) {
                Row dataRow = outputSheet.createRow(dataRowIndex++);

                // 写入序号
                dataRow.createCell(0).setCellValue(rowData.getSystemIndex());
                // 写入系统类型，如 "1. Android"
                dataRow.createCell(1).setCellValue(rowData.getSystemIndex() + ". " + rowData.getSystemType());

                // TODO：这里可以继续写入原始数据的其他列
                // 比如 E列（原数据中的某个重要字段）
                Map<Integer, String> originalValues = rowData.getOriginalValues();
                // 假设 E列是原数据中重要的列，可以取出展示
                String eValue = originalValues.getOrDefault(4, ""); // E列索引是4
                dataRow.createCell(2).setCellValue(eValue);

                // 如果你有更多列要展示，可以类似处理：
                // dataRow.createCell(3).setCellValue(originalValues.getOrDefault(5, ""));
            }

            // 自动调整列宽
            for (int i = 0; i < 3; i++) { // 假定表头有3列（序号、系统类型、E列内容）
                outputSheet.autoSizeColumn(i);
            }

            // 写入文件
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                outputWorkbook.write(fos);
            }
        }
    }

    /**
     * 将一行 Excel 数据转成 Map<列索引, 值>
     */
    private static Map<Integer, String> getRowValuesAsMap(Row row) {
        Map<Integer, String> values = new HashMap<>();
        if (row == null) return values;

        for (Cell cell : row) {
            int colIndex = cell.getColumnIndex();
            String value = getCellValueAsString(cell);
            values.put(colIndex, value);
        }
        return values;
    }

    /**
     * 获取单元格的字符串值
     */
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue().trim();
            case NUMERIC -> DateUtil.isCellDateFormatted(cell)
                    ? cell.getDateCellValue().toString()
                    : String.valueOf((long) cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            case BLANK -> "";
            default -> "";
        };
    }

    /**
     * 判断字符串是否非空
     */
    private static boolean isNotBlank(String s) {
        return s != null && !s.trim().isEmpty();
    }
}