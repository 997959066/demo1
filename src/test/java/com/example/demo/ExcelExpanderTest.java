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
            String inputFileName = "PET_Project Estimation Tool.xlsm";
            InputStream inputStream = getClass().getClassLoader().getResourceAsStream(inputFileName);

            if (inputStream == null) {
                throw new RuntimeException("❌ 请确保文件 '" + inputFileName + "' 已放在项目的 src/test/resources/ 目录下");
            }

            // 2. 定义一个临时输入文件路径（可选，也可以直接用 InputStream，但为了兼容你的 API，我们先写出到临时文件）
            // 如果你的 ExcelExpander.expandAndGenerate() 只接受文件路径（String），而不是 InputStream，那么我们需要先将该 Excel 文件写入一个临时文件
            File inputFile = File.createTempFile("test-input-", ".xlsm");
            try (OutputStream fos = new FileOutputStream(inputFile)) {
                byte[] buffer = new byte[1024];
                int bytesRead;
                while ((bytesRead = inputStream.read(buffer)) != -1) {
                    fos.write(buffer, 0, bytesRead);
                }
            }

            System.out.println("✅ 已将测试输入文件写入临时位置: " + inputFile.getAbsolutePath());

            // 3. 定义输出文件（也写到临时目录，或固定路径，比如项目根目录下的 output/）
            File outputFile = File.createTempFile("expanded-output-", ".xlsx");
            String outputExcelPath = outputFile.getAbsolutePath();

            // 4. 调用你的核心方法：读取输入文件，生成带拆分系统类型的输出文件
            ExcelExpander.expandAndGenerate(inputFile.getAbsolutePath(), outputExcelPath);

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