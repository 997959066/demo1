package com.example.demo;


import java.util.Map;

/**
 * 封装从 Excel 读取的一行数据，以及要生成的系统类型
 */
public class ExcelRowData {
    private Map<Integer, String> originalValues;  // 原始行数据（列索引 -> 值）
    private String systemType;                    // 当前系统类型，如 "Android"
    private int systemIndex;                      // 序号，如 1, 2, 3...

    public ExcelRowData(Map<Integer, String> originalValues, String systemType, int systemIndex) {
        this.originalValues = originalValues;
        this.systemType = systemType;
        this.systemIndex = systemIndex;
    }

    // Getters
    public Map<Integer, String> getOriginalValues() { return originalValues; }
    public String getSystemType() { return systemType; }
    public int getSystemIndex() { return systemIndex; }
}