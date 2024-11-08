package com.inesat;


import lombok.extern.log4j.Log4j2;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.math3.util.MathUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import org.springframework.context.annotation.Bean;

import javax.annotation.Resource;
import java.io.*;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@SpringBootApplication
@EnableConfigurationProperties
@Log4j2
public class Application {

    public static void main(String[] args) {
        SpringApplication.run(Application.class, args);
    }

    @Resource
    private Mappings mappings;
    private String dateStr;

    @Bean
    public CommandLineRunner commandLineRunner() {
        log.info("Full Command: java -jar salary-0.0.1-SNAPSHOT.jar 202401 工资模板.xls 工资条202401.xls");
        return args -> {
            try {
                dateStr = LocalDate.now().plusMonths(-1).format(DateTimeFormatter.ofPattern("yyyyMM"));
                if (args.length > 0) {
                    dateStr = args[0];
                }
                String inputFileName = "工资模板.xls";
                if (args.length > 1) {
                    inputFileName = args[1];
                }
                String outputFileName = "工资条" + dateStr + ".xls";
                if (args.length > 2) {
                    outputFileName = args[2];
                }
                log.info("start transform from file {}.", inputFileName);
                List<Map<String, String>> values = readExcel(inputFileName);
                if (values.size() > 0) {
                    writeExcel(outputFileName, values);
                }
            } catch (Exception e) {
                log.error(e);
                int b = System.in.read();
                log.debug("input byte {} ", b);
            }
        };
    }

    public List<Map<String, String>> readExcel(String inputFileName) throws IOException {
        List<ColMapping> cols = mappings.getCols();
        HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(inputFileName));
        HSSFSheet sheet = wb.getSheetAt(0);
        log.info("sheet rows: {}", sheet.getLastRowNum());
        updateMappingIndexByTitleRow(cols, sheet);

        List<Map<String, String>> values = new ArrayList<>();

        for (int i = 1; i < sheet.getLastRowNum() + 1; i++) {
            HSSFRow row = sheet.getRow(i);
            if (row == null || row.getCell(0) == null || StringUtils.isBlank(row.getCell(0).toString())) {
                log.info("Break read rows, current row index is: {}", i);
                break;
            }

            Map<String, String> rowValue = new LinkedHashMap<>();
            cols.forEach(e -> {
                HSSFCell cell = null;
                if (e.getColIndex() != null) {
                    cell = row.getCell(e.getColIndex());
                }
                rowValue.put(e.getTo(), getCellStringValue(cell));
            });
            values.add(rowValue);
        }

        return values;
    }

    private String getCellStringValue(HSSFCell cell) {
        if (cell != null) {
            if (CellType.NUMERIC.equals(cell.getCellType())) {
                double value = cell.getNumericCellValue();
                return BigDecimal.valueOf(value).setScale(2, RoundingMode.HALF_UP).toString();
            } else {
                cell.setCellType(CellType.STRING);
                return cell.toString();
            }
        }
        return "";
    }
    private void updateMappingIndexByTitleRow(List<ColMapping> cols, HSSFSheet sheet) {
        HSSFRow titleRow = sheet.getRow(0);
        if (titleRow == null) return;
        log.info("titleRow is : "+titleRow);
        List<String> titleRowValue = new ArrayList<>();
        for (int i = 0; i < titleRow.getLastCellNum(); i++) {
            HSSFCell cell = titleRow.getCell(i);
            titleRowValue.add(cell == null ? null : cell.toString());
        }
        System.out.println(titleRowValue);
        cols.forEach(e -> {
            if (StringUtils.isNotBlank(e.getFrom())) {
                int index = titleRowValue.indexOf(e.getFrom());
                if (index > -1) {
                    e.setColIndex(index);
                }
            }
        });
    }

    public void writeExcel(String outputFileName, List<Map<String, String>> values) throws IOException {
        Workbook wb = new HSSFWorkbook();
        /*Excel文件创建完毕*/
        CreationHelper createHelper = wb.getCreationHelper();  //创建帮助工具

        /*创建表单*/
        Sheet sheet = wb.createSheet("工资条");

        //设置字体
        Font headFont = wb.createFont();
        headFont.setFontHeightInPoints((short) 12);
        headFont.setFontName("宋体");

        /*设置数据单元格格式*/
        CellStyle dataStyle = wb.createCellStyle();
        dataStyle.setBorderBottom(BorderStyle.DOUBLE);  //设置单元格线条
        dataStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());   //设置单元格颜色
        dataStyle.setBorderLeft(BorderStyle.THIN);
        dataStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        dataStyle.setBorderRight(BorderStyle.THIN);
        dataStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        dataStyle.setBorderTop(BorderStyle.THIN);
        dataStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        dataStyle.setAlignment(HorizontalAlignment.CENTER);    //设置水平对齐方式
        dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);  //设置垂直对齐方式
        dataStyle.setFont(headFont);  //设置字体

        //设置头部单元格样式
        CellStyle headStyle = wb.createCellStyle();
        headStyle.setBorderBottom(BorderStyle.THIN);  //设置单元格线条
        headStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());   //设置单元格颜色
        headStyle.setBorderLeft(BorderStyle.THIN);
        headStyle.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        headStyle.setBorderRight(BorderStyle.THIN);
        headStyle.setRightBorderColor(IndexedColors.BLACK.getIndex());
        headStyle.setBorderTop(BorderStyle.DOUBLE);
        headStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());
        headStyle.setAlignment(HorizontalAlignment.CENTER);    //设置水平对齐方式
        headStyle.setVerticalAlignment(VerticalAlignment.CENTER);  //设置垂直对齐方式
        headStyle.setFont(headFont);  //设置字体

        /*设置列宽度*/
        for (int i = 0; i <= values.get(0).size(); i++) {
            sheet.setColumnWidth(i, 12 * 256);
        }

        int startRowIndex = 0;
        for (Map<String, String> value : values) {
            Row headRow = sheet.createRow(startRowIndex++);
            headRow.setHeight((short) 400);
            Row valueRow = sheet.createRow(startRowIndex++);
            valueRow.setHeight((short) 400);
            int col = 0;
            createTextCell(createHelper, headStyle, headRow, col, "日期");
            createTextCell(createHelper, dataStyle, valueRow, col++, dateStr);
            for (Map.Entry<String, String> entry : value.entrySet()) {
                createTextCell(createHelper, headStyle, headRow, col, entry.getKey());
                createTextCell(createHelper, dataStyle, valueRow, col++, entry.getValue());
            }
        }

        File file = new File(outputFileName);
        try (OutputStream fileOut = new FileOutputStream(file)) {
            wb.write(fileOut);   //将workbook写入文件流
        }
        log.info("file write to : {}", file.getAbsolutePath());

    }

    private void createTextCell(CreationHelper createHelper, CellStyle cellStyle, Row row, int i, Object text) {
        Cell cell;
        cell = row.createCell(i);
        cell.setCellValue(createHelper.createRichTextString(text == null ? "" : text.toString()));
        cell.setCellStyle(cellStyle);
    }

}
