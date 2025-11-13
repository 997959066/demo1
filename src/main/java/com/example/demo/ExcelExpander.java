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
                        Cell eCell13 = row.getCell(13);  // N列
//                        Cell eCell14 = row.getCell(14);  // 0列
                        Cell qCell = row.getCell(15); // P列

                        String eValue = getCellValueAsString(eCell);
                        String qValue = getCellValueAsString(qCell);
                        String qValue13Task = getCellValueAsString(eCell13);

                        // 条件：E列有值，Q列有值且包含多个系统类型&& !qValue.equals("system type")
                        if (isNotBlank(eValue) && isNotBlank(qValue) && !qValue.equals("system type")) {
                            qValue=qValue13Task +"&" + qValue+" & Quality Assurance & Product Manager & Delivery Manager & Product Designer";
                            // 拆分 Q 列，如 "Android & Front-End & QA" → ["Android Developer", "Front-End Developer", "QA"]
                            List<String> systemTypes = Arrays.stream(qValue.split("&")).map(String::trim).filter(s -> !s.isEmpty()).collect(Collectors.toList());
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
            headerRow.createCell(0).setCellValue("EPIC Story");
            headerRow.createCell(1).setCellValue("Task / Description");
            headerRow.createCell(2).setCellValue("CAPEX/OTE");
            headerRow.createCell(3).setCellValue("Release");
            headerRow.createCell(4).setCellValue("Vendor/SHDR");
            headerRow.createCell(5).setCellValue("Cost Type");
            headerRow.createCell(6).setCellValue("Asset Function"); // In-House
            headerRow.createCell(7).setCellValue("Team");
            headerRow.createCell(8).setCellValue("Role list");
            headerRow.createCell(9).setCellValue("Contractor (Y/N)自动带出列");
            headerRow.createCell(10).setCellValue("Contractor Level");
            headerRow.createCell(11).setCellValue("New Hiring (Y/N)");
            headerRow.createCell(12).setCellValue("Refresh/Replacement(Y/N) 不填写列");
            headerRow.createCell(13).setCellValue("Labor Hours/Quantities");
            headerRow.createCell(14).setCellValue("Rate / Unit Price RMB");
            headerRow.createCell(15).setCellValue("Amount RMB");
            headerRow.createCell(16).setCellValue("Amount USD");
            // 假设你还要保留原数据列，比如 A(0) ~ P(15)，这里可以根据需求设置更多表头
            // 比如：headerRow.createCell(2).setCellValue("E列内容");
            // 你可以根据 originalValues 的列索引来动态设置表头

            int dataRowIndex = 1;
            for (ExcelRowData rowData : expandedRows) {
                Row dataRow = outputSheet.createRow(dataRowIndex++);
                int systemIndex = rowData.getSystemIndex();
                // EPIC Story
                Map<Integer, String> originalValues = rowData.getOriginalValues();
                // 假设 E列是原数据中重要的列，可以取出展示
                String eValue = originalValues.getOrDefault(4, ""); // E列索引是4
                dataRow.createCell(0).setCellValue(eValue);
                // 写入系统类型，如 "1. Android" //Task / Description
                String systemType = rowData.getSystemType();
                dataRow.createCell(1).setCellValue(systemType);
                String capex =systemType.equals("Delivery Manager")?"OTE":"CAPEX";
                dataRow.createCell(2).setCellValue(capex);

                 dataRow.createCell(3).setCellValue("Release 1");

                 dataRow.createCell(4).setCellValue("Vendor");
                 //Cost Type
                String costType = systemIndex==1?"Vendor Service":"Labor";
                dataRow.createCell(5).setCellValue(costType);
                //Asset Function
                String assetFunction = systemIndex==1?"Vendor Enhance Software":"In-House";
                dataRow.createCell(6).setCellValue(assetFunction);
                //Team
                dataRow.createCell(7).setCellValue("Digital Engineering");
                //Role list"
                String role = systemIndex==1?"":systemType;
                dataRow.createCell(8).setCellValue(role);
                //Contractor
                String contractor = systemIndex==1?"":"Y";
                dataRow.createCell(9).setCellValue(contractor);
                //Level
//                String contractorLevel = systemIndex==1?"":"Y";
                dataRow.createCell(10).setCellValue("Level 4 (5-8Years)");

                //Hiring
                String hiring = systemIndex==1?"":"Y";
                dataRow.createCell(11).setCellValue(hiring);

                //空列，不需要生成
                dataRow.createCell(12).setCellValue("");
                //Labor Hours/Quantities 带入公式

                dataRow.createCell(13).setCellValue("");
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
