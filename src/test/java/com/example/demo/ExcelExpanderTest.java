package com.example.demo;

import com.example.demo.ExcelExpander;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelExpanderTest {

    @Test
    void testExpandExcelWithSystemTypes() {
        try {
            // 1. 从 classpath（src/test/resources/）读取输入的 Excel 文件
            String inputFileName = System.getProperty("user.dir")+"/Split up the Epic Story tasks.xlsx";

            File outputFile = File.createTempFile("expanded-output-", ".xlsx");
            String outputExcelPath = outputFile.getAbsolutePath();

            ExcelExpander.expandAndGenerate(inputFileName, outputExcelPath);

            // 5. 验证输出文件已生成
            assertTrue(outputFile.exists(), "❌ 输出 Excel 文件应该被生成: " + outputExcelPath);
            System.out.println("✅ 输出文件已成功生成: " + outputExcelPath);

            // 6. （可选）删除临时文件（清理）
            // inputFile.deleteOnExit();
            // outputFile.deleteOnExit();

        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("测试失败", e);
        }
    }
}
