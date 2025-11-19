package com.example.demo;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Collectors;

public class ExcelExpander {
//
    public static String Product_Manager="Product Manager";
    public static String Delivery_Manager="Delivery Manager";
    public static String Quality_Assurance="Quality Assurance";
    public static String Sr_Quality_Assurance="Sr Quality Assurance";
    public static String Android_Developer="Android Developer";
    public static String Front_end_Developer="Front-end Developer";
    public static String Back_end_Developer="Back-end Developer";
    public static String SR_Back_end_Developer="SR Back-end Developer";
    public static String Product_Designer="Product Designer";

    /**
     * 从输入 Excel 读取数据，根据 Q 列拆分系统类型，生成多行，并写入新 Excel
     */
    public static void expandAndGenerate(String inputPath, String outputPath) throws IOException {
        List<ExcelRowData> allExpandedRows = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(inputPath);
             Workbook inputWorkbook = new XSSFWorkbook(fis)) {
            for (Sheet sheet : inputWorkbook) {
                String sheetName = sheet.getSheetName();
                if(sheetName.equals("BRD & EPIC")) {
                    System.out.println("正在处理 Sheet: " + sheetName);
                    // 遍历每一行（从第2行开始，假设第1行是表头）
                    for (int rowNum = 6; rowNum <= sheet.getLastRowNum(); rowNum++) {
                        Row row = sheet.getRow(rowNum);
                        if (row == null) continue;
                        // 读取 E 列（索引4）和 Q 列（索引16）
                        Cell eCell = row.getCell(4);  // E列
                        Cell eCell14 = row.getCell(14);  // 0列
                        Cell qCell = row.getCell(15); // P列

                        String eValue = getCellValueAsString(eCell);
                        String qValue = getCellValueAsString(qCell);
                        String qValue14Task = getCellValueAsString(eCell14);

                        // 条件：E列有值，Q列有值且包含多个系统类型
                        if (isNotBlank(eValue) && isNotBlank(qValue) && !qValue.equals("system type")) {
                            qValue= qValue14Task +"&" + qValue;
                            List<String> systemTypes = Arrays.stream(qValue.split("&")).map(String::trim).filter(s -> !s.isEmpty()).collect(Collectors.toList());
                            //多流出一条
                            if (systemTypes.contains(Back_end_Developer)) {
                                systemTypes.add(SR_Back_end_Developer);
                            }
                            if (systemTypes.contains(Quality_Assurance)) {
                                systemTypes.add(Sr_Quality_Assurance);
                            }

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
            Sheet outputSheet = outputWorkbook.createSheet("Cost Breakdown");

            // 创建表头
            Row headerRow0 = outputSheet.createRow(0);
            headerRow0.createCell(0).setCellValue("Project Category");
            Row headerRow1 = outputSheet.createRow(1);
            headerRow1.createCell(0).setCellValue("N/A");

            Row headerRow = outputSheet.createRow(3);
            headerRow.createCell(0).setCellValue("BRD Priority");
            headerRow.createCell(1).setCellValue("BRD Overview");
            headerRow.createCell(2).setCellValue("BRD Detailed Requirement");
            headerRow.createCell(3).setCellValue("EPIC Title");


            headerRow.createCell(4).setCellValue("EPIC Story");
            headerRow.createCell(5).setCellValue("Task / Description");
            headerRow.createCell(6).setCellValue("CAPEX/OTE");
            headerRow.createCell(7).setCellValue("Release");
            headerRow.createCell(8).setCellValue("Vendor/SHDR");
            headerRow.createCell(9).setCellValue("Cost Type");//J
            headerRow.createCell(10).setCellValue("Asset"); //     K

            headerRow.createCell(11).setCellValue("Asset Function"); // In-House L
            headerRow.createCell(12).setCellValue("Team");//M
            headerRow.createCell(13).setCellValue("Role list");//N
            headerRow.createCell(14).setCellValue("Contractor (Y/N)自动带出列");//O
            headerRow.createCell(15).setCellValue("Contractor Level");//P
            headerRow.createCell(16).setCellValue("New Hiring (Y/N)");//Q
            headerRow.createCell(17).setCellValue("Refresh/Replacement(Y/N) 不填写列");//R

            headerRow.createCell(18).setCellValue("Labor Hours/Quantities");//S
            headerRow.createCell(19).setCellValue("Rate / Unit Price RMB");//T
            headerRow.createCell(20).setCellValue("Amount RMB");//U
            headerRow.createCell(21).setCellValue("Amount USD");//V
            // 假设你还要保留原数据列，比如 A(0) ~ P(15)，这里可以根据需求设置更多表头
            // 比如：headerRow.createCell(2).setCellValue("E列内容");
            // 你可以根据 originalValues 的列索引来动态设置表头

            int dataRowIndex = 4;
            for (ExcelRowData rowData : expandedRows) {
                Row dataRow = outputSheet.createRow(dataRowIndex++);
                int systemIndex = rowData.getSystemIndex();
                // EPIC Story
                Map<Integer, String> originalValues = rowData.getOriginalValues();
                // 假设 E列是原数据中重要的列，可以取出展示
                String eValue = originalValues.getOrDefault(4, ""); // E列索引是4
                dataRow.createCell(4).setCellValue(eValue);
                // 写入系统类型，如 "1. Android" //Task / Description
                String systemType = rowData.getSystemType();
                dataRow.createCell(5).setCellValue(systemType);
                String capex =systemType.equals("Delivery Manager")?"OTE":"CAPEX";
                dataRow.createCell(6).setCellValue(capex);

                 dataRow.createCell(7).setCellValue("Release 1");

                 dataRow.createCell(8).setCellValue("Vendor");
                 //Cost Type
                String costType = systemIndex==1?"Vendor Service":"Labor";
                dataRow.createCell(9).setCellValue(costType);
                //Asset
                dataRow.createCell(10).setCellValue("");
                //Asset Function
                String assetFunction = systemIndex==1?"Vendor Enhance Software":"In-House";
                dataRow.createCell(11).setCellValue(assetFunction);
                //Team
                dataRow.createCell(12).setCellValue("Digital Engineering");
                //Role list"
//                String role = systemIndex==1?"":systemType;
                String role =  "";
                String roleLevel =  "Level 4 (5-8Years)";
                if(systemIndex!=1){
                    role=systemType;
                    if(systemType.equals(Sr_Quality_Assurance)){
                        role=Quality_Assurance;
                        roleLevel="Level 3 (8-10Years)";
                    }
                    if(systemType.equals(SR_Back_end_Developer)){
                        role=Back_end_Developer;
                        roleLevel="Level 3 (8-10Years)";
                    }
                }

                dataRow.createCell(13).setCellValue(role);
                //Contractor
                String contractor = systemIndex==1?"":"Y";
                dataRow.createCell(14).setCellValue(contractor);
                //contractor Level
//                String contractorLevel = systemIndex==1?"":"Y";
                dataRow.createCell(15).setCellValue(roleLevel);

                //New Hiring
                String hiring = systemIndex==1?"":"Y";
                dataRow.createCell(16).setCellValue(hiring);

                //Refresh/Replacement(Y/N) 不填写列
                dataRow.createCell(17).setCellValue("");
                //Labor Hours/Quantities 带入公式
                if(systemIndex==1){
                    dataRow.createCell(18).setCellValue("1");
                }else {
                    if(systemType.equals(Product_Manager)){
//                        dataRow.createCell(18).setCellFormula("ROUNDUP(INDEX('BRD & EPIC'!T:T,MATCH(TEXTBEFORE($E"+dataRowIndex+",\" \")&\"*\",'BRD & EPIC'!E:E,0)) * DE_Cost!$C$2, 0)");
                        dataRow.createCell(18).setCellFormula("DE_Cost!D2");
                    }else if (systemType.equals(Delivery_Manager)){
//                        dataRow.createCell(18).setCellFormula("ROUNDUP(INDEX('BRD & EPIC'!T:T,MATCH(TEXTBEFORE($E"+dataRowIndex+",\" \")&\"*\",'BRD & EPIC'!E:E,0)) * DE_Cost!$C$3, 0)");
                        dataRow.createCell(18).setCellFormula("DE_Cost!D3");
                    }else if (systemType.equals(Quality_Assurance) || systemType.equals(Sr_Quality_Assurance)){
                        dataRow.createCell(18).setCellFormula("ROUNDUP(INDEX('BRD & EPIC'!T:T,MATCH(TEXTBEFORE($E"+dataRowIndex+",\" \")&\"*\",'BRD & EPIC'!E:E,0)) * DE_Cost!$C$4, 0)");
                    }else if (systemType.equals(Android_Developer)){
                        dataRow.createCell(18).setCellFormula("ROUNDUP(INDEX('BRD & EPIC'!T:T,MATCH(TEXTBEFORE($E"+dataRowIndex+",\" \")&\"*\",'BRD & EPIC'!E:E,0)) * DE_Cost!$C$5, 0)");
                    }else if (systemType.equals(Back_end_Developer) || systemType.equals(SR_Back_end_Developer)){
                        dataRow.createCell(18).setCellFormula("ROUNDUP(INDEX('BRD & EPIC'!T:T,MATCH(TEXTBEFORE($E"+dataRowIndex+",\" \")&\"*\",'BRD & EPIC'!E:E,0)) * DE_Cost!$C$7, 0)");
                    }else if (systemType.equals(Product_Designer)){
//                        dataRow.createCell(18).setCellFormula("ROUNDUP(INDEX('BRD & EPIC'!T:T,MATCH(TEXTBEFORE($E"+dataRowIndex+",\" \")&\"*\",'BRD & EPIC'!E:E,0)) * DE_Cost!$C$9, 0)");
                        dataRow.createCell(18).setCellFormula("DE_Cost!D9");
                    }else if (systemType.equals(Front_end_Developer)){
                        dataRow.createCell(18).setCellFormula("ROUNDUP(INDEX('BRD & EPIC'!T:T,MATCH(TEXTBEFORE($E"+dataRowIndex+",\" \")&\"*\",'BRD & EPIC'!E:E,0)) * DE_Cost!$C$6, 0)");
                    }
                }

                //Rate / Unit Price RMB
                if(systemIndex==1){
                    dataRow.createCell(19).setCellValue("1");
                    dataRow.createCell(19).setCellFormula("INDEX('BRD & EPIC'!AE:AE,MATCH(TEXTBEFORE($E"+dataRowIndex+",\" \")&\"*\",'BRD & EPIC'!E:E,0))");
                }
                //Amount RMB 20 U
                //Amount USD 21 V
                //Comments W


            }

            // 自动调整列宽
            for (int i = 0; i < 3; i++) { // 假定表头有3列（序号、系统类型、E列内容）
                outputSheet.autoSizeColumn(i);
            }
//            outputWorkbook.setForceFormulaRecalculation(true);
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
