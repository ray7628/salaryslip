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
        return args -> {
            try {
                dateStr = LocalDate.now().plusMonths(-1).format(DateTimeFormatter.ofPattern("yyyyMM"));
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

        //设置头部单元格样式
        CellStyle headStyle = wb.createCellStyle();
        headStyle.setBorderBottom(BorderStyle.THIN);  //设置单元格线条
        headStyle.setBottomBorderColor(IndexedColors.BLACK.getIndex());   //设置单元格颜色
        headStyle.setBorderLeft(BorderStyle.THIN);

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