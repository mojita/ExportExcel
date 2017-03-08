package com.sobey.model;

import java.io.FileOutputStream;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.util.StringUtils;

/**
 * Created by lijunhong on 17/3/8.
 *
 * 表头《档案借(查)阅登记单》
 * 序号:
 * 档号:
 * 文件编号:
 * 题目:
 * 密级:
 * 页数:
 * 文件状态:
 *上面信息中有的就写没有的就空着
 * 档案借阅登记单model为:0
 *
 * 本类主要是档案借阅登记单
 */
public class RecordBorrowRegistry {


    private static Logger logger = Logger.getLogger(ThemePreviewExcel.class);

    private final String EXCEL_FILE_SHEET_NAME = "借阅查登记单";    //Excel文件中的工作薄默认创建名字

    private String fileName;                    //文件名
    private String path;                        //文件存放路径
    private String filePath;                    //文件存放路径和文件名
    private HSSFWorkbook workbook = null;       //Excel对象
    private HSSFSheet sheet = null;             //Excel的工作簿对象



    /**
     * 这里需要提供文件名文件路径和文件类型然后才能创建Excel
     * @param fileName
     * @param path
     */
    public RecordBorrowRegistry(String fileName,String path){
        this.path = path;
        this.fileName = fileName;
        this.filePath = path + fileName;
        workbook = new HSSFWorkbook();
        sheet = workbook.createSheet(EXCEL_FILE_SHEET_NAME);
    }


    /**
     * 本方法主要是根据提供的数据构造胡一张Excel表
     */
    public void createExcelTableHeader(){

        sheet.setColumnWidth(0,20*256);              //设置序号
        sheet.setColumnWidth(1,28*256);              //设置档号
        sheet.setColumnWidth(2,28*256);              //设置文件编号
        sheet.setColumnWidth(3,28*256);              //设置题目
        sheet.setColumnWidth(4,15*256);              //设置密级
        sheet.setColumnWidth(5,15*256);              //设置页数
        sheet.setColumnWidth(6,15*256);              //设置文件状态



        //样式
        HSSFCellStyle style = workbook.createCellStyle();

        HSSFRow row = null;
        HSSFCell cell = null;

        //设置第一行,表头信息,现将第一行第单元格进行合并
        row = sheet.createRow(0);
        CellRangeAddress range = new CellRangeAddress(0,0,0,6);
        sheet.addMergedRegion(range);
        cell = row.createCell(0);
        cell.setCellValue("档案借(查)阅登记单");
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        cell.setCellStyle(style);



        //设置第二行,元数据
        row = sheet.createRow(1);
        //设置边框
        HSSFCellStyle borderAndCenter = getHssfCellStyle();

        cell = row.createCell(0);
        cell.setCellValue("序号");
        cell.setCellStyle(borderAndCenter);

        cell = row.createCell(1);
        cell.setCellValue("档号");
        cell.setCellStyle(borderAndCenter);

        cell = row.createCell(2);
        cell.setCellValue("文件编号");
        cell.setCellStyle(borderAndCenter);

        cell = row.createCell(3);
        cell.setCellValue("题目");
        cell.setCellStyle(borderAndCenter);

        cell = row.createCell(4);
        cell.setCellValue("密级");
        cell.setCellStyle(borderAndCenter);

        cell = row.createCell(5);
        cell.setCellValue("页数");
        cell.setCellStyle(borderAndCenter);

        cell = row.createCell(6);
        cell.setCellValue("文件状态");
        cell.setCellStyle(borderAndCenter);


        writerExcel();

    }


    /**
     * 生成Excel文件的主体内容
     * @param metaDataLsits
     */
    public int createExcelTableBody(List<List<MetaData>> metaDataLsits){
        int index = 0;
        for (List<MetaData> metaDatas:metaDataLsits){
            for (MetaData metaData:metaDatas){
                HSSFRow row = sheet.createRow(index+2);
                HSSFCell cell = null;

                //设置档序号
                cell = row.createCell(0);
                cell.setCellValue(index+1);
                cell.setCellStyle(getHssfCellStyle());      //设置边框

                //设置档号
                cell = row.createCell(1);
                cell.setCellValue(!StringUtils.isEmpty(metaData.getFileNo())?metaData.getFileNo():"");
                cell.setCellStyle(getHssfCellStyle());

                //设置文件编号
                cell = row.createCell(2);
                cell.setCellValue(!StringUtils.isEmpty(metaData.getFileId())?metaData.getFileId():"");
                cell.setCellStyle(getHssfCellStyle());

                //设置题名
                cell = row.createCell(3);
                cell.setCellValue(!StringUtils.isEmpty(metaData.getTitle())?metaData.getTitle():"");
                cell.setCellStyle(getHssfCellStyle());

                //设置密级
                cell = row.createCell(4);
                cell.setCellValue(!StringUtils.isEmpty(metaData.getSecretLevel())?metaData.getSecretLevel():"");
                cell.setCellStyle(getHssfCellStyle());

                //设置页数
                cell = row.createCell(5);
                cell.setCellValue(!StringUtils.isEmpty(metaData.getPageNum())?metaData.getPageNum():"");
                cell.setCellStyle(getHssfCellStyle());

                //设置文件状态
                cell = row.createCell(6);
                cell.setCellValue(!StringUtils.isEmpty(metaData.getFileState())?metaData.getFileState():"");
                cell.setCellStyle(getHssfCellStyle());
                index++;
            }
        }
        writerExcel();
        return index;
    }

    /**
     * 设置边框
     * @return
     */
    private HSSFCellStyle getHssfCellStyle() {
        HSSFCellStyle style = workbook.createCellStyle();
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        return style;
    }

    /**
     * 用来将内存中的Excel写入到指定路劲的文件中
     */
    private void writerExcel() {
        try {
            FileOutputStream out = new FileOutputStream(filePath);
            workbook.write(out);
            out.close();
        } catch (Exception e) {
            e.printStackTrace();
            if(logger.isErrorEnabled()) logger.error("生成借阅查Excel文件出现错误"+e);
        }
    }




}
