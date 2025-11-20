package com.example.demo;

import org.junit.jupiter.api.Test;

import java.io.File;

import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelExpanderTest {

    String property = System.getProperty("user.dir");
    String userHome =  System.getProperty("user.home") ;

    @Test
    void testExpandExcelWithSystemTypes() {
        try {

            String inputFileName = property + "/Split up the Epic Story tasks.xlsx";

            File outputFile = File.createTempFile("expanded-output-", ".xlsx");

            String outputExcelPath = userHome + "/Downloads/AA/" + outputFile.getName();

            ExcelExpander.expandAndGenerate(inputFileName, outputExcelPath);

            // 5. 验证输出文件已生成
            assertTrue(outputFile.exists(), "❌ 输出 Excel 文件应该被生成: " + outputExcelPath);
            System.out.println("✅ 输出文件已成功生成: " + outputExcelPath);

        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("测试失败", e);
        }
    }
}
