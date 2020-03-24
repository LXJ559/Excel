package com.how2java.springboot.util;

import com.how2java.springboot.entity.Student;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.ArrayList;

import java.util.List;



public class ExcelUtil {

    private static Logger log = Logger.getLogger(ExcelUtil.class);
    /*导入数据*/
    public static List<Object[]> importExcel(String fileName) {
        log.info("导入解析开始，fileName:"+fileName);
        try {
            List<Object[]> list = new ArrayList<>();
            InputStream inputStream = new FileInputStream(fileName);
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet sheet = workbook.getSheetAt(0);
            //获取sheet的行数
            int rows = sheet.getPhysicalNumberOfRows();
            for (int i = 0; i < rows; i++) {
                //过滤表头行
                if (i == 0) {
                    continue;
                }
                //获取当前行的数据
                Row row = sheet.getRow(i);
                Object[] objects = new Object[row.getPhysicalNumberOfCells()];
                int index = 0;
                for (Cell cell : row) {
                    if (cell.getCellType().equals(CellType.NUMERIC)) {
                        objects[index] = (int) cell.getNumericCellValue();
                    }
                    if (cell.getCellType().equals(CellType.STRING)) {
                        objects[index] = cell.getStringCellValue();
                    }
                    if (cell.getCellType().equals(CellType.BOOLEAN)) {
                        objects[index] = cell.getBooleanCellValue();
                    }
                    if (cell.getCellType().equals(CellType.ERROR)) {
                        objects[index] = cell.getErrorCellValue();
                    }
                    index++;
                }
                list.add(objects);
            }
            log.info("导入文件解析成功！");
            return list;
        }catch (Exception e){
            log.info("导入文件解析失败！");
            e.printStackTrace();
        }
        return null;
    }


    /*导出数据*/
    public static void exportExcel(HttpServletResponse response,List<Student> listStudent , String[] tableHeader) throws UnsupportedEncodingException {
        /*表的列数*/
        int cellNumber = tableHeader.length;
        HSSFWorkbook workbook = new HSSFWorkbook(); //创建一个Excel
        HSSFSheet sheet = workbook.createSheet("sheet1"); //创建一个sheet
        setTitle(workbook,sheet,tableHeader);
        setData(sheet,listStudent,cellNumber);
        String filename = new String("学生信息表.xls".getBytes("utf-8"), "ISO-8859-1");
        setBrowser(response,workbook, filename);
    }

    private static void setTitle(HSSFWorkbook workbook, HSSFSheet sheet, String[] str) {
        try {
            HSSFRow row = sheet.createRow(0);
            //设置列宽，setColumnWidth的第二个参数要乘以256，这个参数的单位是1/256个字符宽度
            for (int i = 0; i <= str.length; i++) {
                sheet.setColumnWidth(i, 15 * 256);
            }
            //设置为居中加粗,格式化时间格式
            HSSFCellStyle style = workbook.createCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER);
            HSSFFont font = workbook.createFont();
            font.setBold(true);
            style.setFont(font);

            //创建表头每个单元格的value 与 style
            HSSFCell cell;
            for (int j = 0; j < str.length; j++) {
                cell = row.createCell(j);
                cell.setCellValue(str[j]);
                cell.setCellStyle(style);
            }
            log.info("设置表头成功！");
        } catch (Exception e) {
            log.info("导出时设置表头失败！");
            e.printStackTrace();
        }
    }

    /**
     * 方法名：setData
     * 功能：表格赋值
     */
    private static void setData(HSSFSheet sheet, List<Student> studentList ,int cellNumber) {
        try{
            int rowNum = 1;
            for (int i = 0; i < studentList.size(); i++) {
                HSSFRow row = sheet.createRow(rowNum);
                row.createCell(0).setCellValue(studentList.get(i).getId());
                row.createCell(1).setCellValue(studentList.get(i).getName());
                row.createCell(2).setCellValue(studentList.get(i).getAge());
                rowNum++;
            }
            log.info("表格赋值成功！");
        }catch (Exception e){
            log.info("表格赋值失败！");
            e.printStackTrace();
        }
    }

    /**
     * 方法名：setBrowser
     * 功能：使用浏览器下载
     */
    private static void setBrowser(HttpServletResponse response, HSSFWorkbook workbook, String fileName) {
        try {
            //设置response的Header
            response.addHeader("Content-Disposition", "attachment;filename=" + fileName);
            OutputStream os = new BufferedOutputStream(response.getOutputStream());
            response.setContentType("application/vnd.ms-excel;charset=gb2312");
            //将excel写入到输出流中
            workbook.write(os);
            os.flush();
            os.close();
            log.info("设置浏览器下载成功！");
        } catch (Exception e) {
            log.info("设置浏览器下载失败！");
            e.printStackTrace();
        }

    }

}
