package com.sobey.model;

import java.io.FileOutputStream;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.util.StringUtils;

import com.sobey.util.ExcelFileUtils;
import com.sobey.util.ExcelViewJudge;

/**
 * Created by lijunhong on 17/3/8.
 *model为:1
 * 主题一览Excel模板
 * 表明:《主题一览》
 * 档号:
 * 题名:
 * 密级:
 * 片长:
 * 摄录日期:
 * 拍摄者:
 * 摄录地点:
 *
 * 本类主要是主题一览模板
 */
public class ThemePreviewExcel {


    private static Logger logger = Logger.getLogger(ThemePreviewExcel.class);

    private final String EXCEL_FILE_SHEET_NAME = "主题一览";    //Excel文件中的工作薄默认创建名字

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
    public ThemePreviewExcel(String fileName,String path){
        this.path = path;
        this.fileName = fileName;
        if(ExcelFileUtils.isWindosSystem){
            this.filePath = path + "/" +fileName;
        }else {
            this.filePath = path + fileName;
        }
        workbook = new HSSFWorkbook();
        sheet = workbook.createSheet(EXCEL_FILE_SHEET_NAME);
    }


    /**
     * 本方法主要是根据提供的数据构造胡一张Excel表
     */
    public void createExcelTableHeader(){

        sheet.setColumnWidth(0,20*256);             //设置档号的宽度
        sheet.setColumnWidth(1,60*256);              //设置题名
        sheet.setColumnWidth(2,18*256);              //设置密级的款的
        sheet.setColumnWidth(3,18*256);              //设置片长
        sheet.setColumnWidth(4,20*256);              //设置拍摄时间
        sheet.setColumnWidth(5,20*256);              //设置拍摄地点
        sheet.setColumnWidth(6,20*256);              //设置拍摄者

        HSSFRow row = null;
        HSSFCell cell = null;

        HSSFCellStyle style = getHssfCellStyle();     //获取边框

        CellRangeAddress range = null;
        //设置表名
        row = sheet.createRow(0);
        range = new CellRangeAddress(0,0,0,6);
        sheet.addMergedRegion(range);
        cell = row.createCell(0);
        cell.setCellValue("主题一览");
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
        cell.setCellStyle(style);

        //设置第一行的元数据
        row = sheet.createRow(1);
        cell = row.createCell(0);
        cell.setCellValue("档号");
        cell.setCellStyle(style);


        cell = row.createCell(1);
        cell.setCellValue("题名");
        cell.setCellStyle(style);

        cell = row.createCell(2);
        cell.setCellValue("密级");
        cell.setCellStyle(style);

        cell = row.createCell(3);
        cell.setCellValue("片长");
        cell.setCellStyle(style);

        cell = row.createCell(4);
        cell.setCellValue("摄录时间");
        cell.setCellStyle(style);

        cell = row.createCell(5);
        cell.setCellValue("摄录者");
        cell.setCellStyle(style);

        cell = row.createCell(6);
        cell.setCellValue("摄录地点");
        cell.setCellStyle(style);

        writerExcel();

    }


    public int createExcelTableBody(List<List<MetaData>> metaDataLsits){
        int index = 0;
        for (List<MetaData> metaDatas:metaDataLsits){
            for (MetaData metaData:metaDatas){
                HSSFRow row = sheet.createRow(index+2);
                HSSFCell cell = null;

                //设置档案号
                cell = row.createCell(0);
                cell.setCellValue(!StringUtils.isEmpty(metaData.getFileNo())?metaData.getFileNo():"");
                cell.setCellStyle(getHssfCellStyle());      //设置边框

                //设置题名
                cell = row.createCell(1);
                cell.setCellValue(!StringUtils.isEmpty(metaData.getTitle())?metaData.getTitle():"");
                cell.setCellStyle(getHssfCellStyle());

                //设置密级
                cell = row.createCell(2);
                cell.setCellValue(ExcelViewJudge.getPrivilegeTemplateCodeString(!StringUtils.isEmpty(metaData.getSecretLevel())?metaData.getSecretLevel():""));
                cell.setCellStyle(getHssfCellStyle());

                //设置片长
                cell = row.createCell(3);
                if(!StringUtils.isEmpty(metaData.getVideoLength())){
                    Long l100 = Long.parseLong(metaData.getVideoLength());
                    String timeStr = ExcelViewJudge.getTimeLength(l100);
                    cell.setCellValue(timeStr);
                }
                cell.setCellStyle(getHssfCellStyle());

                //设置拍摄时间
                cell = row.createCell(4);
                cell.setCellValue(!StringUtils.isEmpty(metaData.getShootingDate())?metaData.getShootingDate():"");
                cell.setCellStyle(getHssfCellStyle());

                //设置拍摄者
                cell = row.createCell(5);
                cell.setCellValue(!StringUtils.isEmpty(metaData.getPhotographer())?metaData.getPhotographer():"");
                cell.setCellStyle(getHssfCellStyle());

                //设置拍摄地点
                cell = row.createCell(6);
                cell.setCellValue(!StringUtils.isEmpty(metaData.getShootingAddress())?metaData.getShootingAddress():"");
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
            if(logger.isErrorEnabled()) logger.error("生成主题一览Excel文件出现错误"+e);
        }
    }


}
