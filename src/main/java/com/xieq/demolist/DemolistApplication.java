package com.xieq.demolist;

import com.itextpdf.text.DocumentException;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

@SpringBootApplication
public class DemolistApplication {

    public static void main(String[] args) throws IOException, DocumentException {
        SpringApplication.run(DemolistApplication.class, args);
        // excel文件路径
        String excelFilePath = "/Users/xiequan/Desktop/test/未命名.xlsx";
        // 生成PDF地址
        String outputPdfPath = "/Users/xiequan/Desktop/test";
        // 生成PDF模版位置
        String templatePath = "/Users/xiequan/Desktop/test/template.pdf";
        // PDF字体中文设置
        String fontPath = "/Library/Fonts/Arial Unicode.ttf";

        // 生成PDF模版
        createPdfTemplate(templatePath, fontPath);
        // 读取excel文件内容，结合模版生成PDF文件
        readExcelAndCreateNewPdf(excelFilePath, templatePath, fontPath);
    }


    /**
     * 读取Excel文件内容，生成PDF文件
     *
     * @param excelFilePath excel文件地址
     * @param templatePath  pdf模版地址
     * @param fontPath      自定义字符文件地址
     * @throws IOException
     * @throws DocumentException
     */
    public static void readExcelAndCreateNewPdf(String excelFilePath, String templatePath, String fontPath) throws IOException, DocumentException {
        // 1.读取excel文件
        FileInputStream excelFile = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(excelFile);
        // 获取工作表1
        Sheet sheet = workbook.getSheetAt(0);

        // 2.遍历excel表每行数据，根据每行的数据生成不同内容的PDF文件
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);

            // 3.获取每行具体单元格数据: 第一列：名字；第二列：年龄；第三列：性别
            String name = row.getCell(0).getStringCellValue();
            double age = row.getCell(1).getNumericCellValue();
            String gender = row.getCell(2).getStringCellValue();

            // 4.使用PDF模版创建PDF并填充数据
            // 输出路径
            String outputPdfPath = "/Users/xiequan/Desktop/test/output_" + name + "_" + rowIndex + ".pdf";

            fillTemplateWithData(templatePath, outputPdfPath, name, age, gender, fontPath);
        }
        // 5.关闭资源
        workbook.close();
        excelFile.close();
        System.out.println("PDF文件已生成！！！");
    }


    /**
     * 生成新PDF文件
     *
     * @param templatePath
     * @param outputPdfPath
     * @param name
     * @param age
     * @param gender
     */
    public static void fillTemplateWithData(String templatePath, String outputPdfPath, String name, double age, String gender, String fontPath) throws IOException {
        // 1.模版加载
        PDDocument document = PDDocument.load(new java.io.File(templatePath));

        // 2.加载自定义字体
        PDType0Font customFont = PDType0Font.load(document, new File(fontPath));

        PDPageTree pages = document.getPages();
        PDPage page = pages.get(0);

        // 3.开始填充
        PDPageContentStream contentStream = new PDPageContentStream(document, page, PDPageContentStream.AppendMode.APPEND, true);

        // 4.填充姓名字段
        contentStream.beginText();
        contentStream.setFont(customFont, 12);
        // 对应“姓名”的坐标
        contentStream.newLineAtOffset(200, 700);
        contentStream.showText(name);
        contentStream.endText();

        // 5.填充年龄字段
        contentStream.beginText();
        contentStream.setFont(customFont, 12);
        contentStream.newLineAtOffset(200, 670);
        contentStream.showText(String.valueOf(age));
        contentStream.endText();

        // 6.填充性别字段
        contentStream.beginText();
        contentStream.setFont(customFont, 12);
        contentStream.newLineAtOffset(200, 640);
        contentStream.showText(gender);
        contentStream.endText();

        // 7.关闭内容流并保存新 PDF 文件
        contentStream.close();
        document.save(outputPdfPath);
        document.close();
    }


    /**
     * 手动创建PDF模版
     *
     * @param templatePath
     * @throws FileNotFoundException
     * @throws DocumentException
     */
    public static void createPdfTemplate(String templatePath, String fontPath) throws FileNotFoundException, DocumentException {
        try {
            // 1.创建 PDF 文档
            PDDocument document = new PDDocument();
            PDPage page = new PDPage();
            document.addPage(page);

            // 2.加载自定义字体
            PDType0Font customFont = PDType0Font.load(document, new File(fontPath));

            // 2.创建内容流
            PDPageContentStream contentStream = new PDPageContentStream(document, page);

            // 3.写入标题
            contentStream.beginText();
            contentStream.setFont(customFont, 18);
            contentStream.newLineAtOffset(100, 750);
            contentStream.showText("用户信息表单");
            contentStream.endText();

            // 4.写入占位符 “姓名” 字段
            contentStream.beginText();
            contentStream.setFont(customFont, 12);
            contentStream.newLineAtOffset(100, 700);
            contentStream.showText("姓名: _______________________");
            contentStream.endText();

            // 5.写入占位符 “年龄” 字段
            contentStream.beginText();
            contentStream.setFont(customFont, 12);
            contentStream.newLineAtOffset(100, 670);
            contentStream.showText("年龄: _______________________");
            contentStream.endText();

            // 6.写入占位符 “性别” 字段
            contentStream.beginText();
            contentStream.setFont(customFont, 12);
            contentStream.newLineAtOffset(100, 640);
            contentStream.showText("性别: _______________________");
            contentStream.endText();

            // 7.关闭内容流并保存文档
            contentStream.close();
            document.save(templatePath);
            document.close();

            System.out.println("PDF 模板创建成功：" + templatePath);
        } catch (IOException e) {
            e.printStackTrace();
        }

    }


}
