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
 * 声像移交目录model为:2
 * /**
 * 表明:《声像档案归档移交目录》
 * 个数:
 * 序号:
 * 题名:
 * 摄录日期:
 * 密级:
 * 片场:
 * 拍摄地点:
 * 摄录单位:
 * 摄录者:
 * 表格下面内容哦那个:
 * 移交人:
 * 领导签字:
 * 移交时间:
 * 接收人:
 * 领导签字:
 * 移交时间:
 *
 * 本类主要是创建声像档案移交的Excel文件模板
 *
 */
public class AudioAndVideoHandOver {
    private static Logger logger = Logger.getLogger(ThemePreviewExcel.class);

    private final String EXCEL_FILE_SHEET_NAME = "声像档案移交目录";    //Excel文件中的工作薄默认创建名字

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
    public AudioAndVideoHandOver(String fileName,String path){
        this.path = path;
        this.fileName = fileName;
        if(ExcelFileUtils.isWindosSystem){
            this.filePath = path + "/" + fileName;
        }else {
            this.filePath = path + fileName;
        }
        workbook = new HSSFWorkbook();
        sheet = workbook.createSheet(EXCEL_FILE_SHEET_NAME);
    }


    /**
     * 本方法主要是根据提供的数据构造胡一张Excel表
     */
    public void createExcelTableHeader(int count){

        sheet.setColumnWidth(0,10*256);              //设置序号
        sheet.setColumnWidth(1,60*256);              //设置题名
        sheet.setColumnWidth(2,20*256);              //设置摄录时间
        sheet.setColumnWidth(3,15*256);              //设置密级
        sheet.setColumnWidth(4,15*256);              //设置片长
        sheet.setColumnWidth(5,20*256);              //设置拍摄地点
        sheet.setColumnWidth(6,20*256);              //设置拍摄单位
        sheet.setColumnWidth(7,15*256);              //设置摄录者



        HSSFRow row = null;
        HSSFCell cell = null;

        HSSFCellStyle style = null;



        //合并单元格
        CellRangeAddress range = null;

        //创建表头
        row = sheet.createRow(0);
        range = new CellRangeAddress(0,0,0,7);
        sheet.addMergedRegion(range);
        cell = row.createCell(0);
        cell.setCellValue("声像档案归档移交目录");
        style = workbook.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        cell.setCellStyle(style);



        //创建页眉信息
        row = sheet.createRow(1);
        range = new CellRangeAddress(1,1,0,6);
        sheet.addMergedRegion(range);
        cell = row.createCell(0);
        cell.setCellValue("个数:");
        style = workbook.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_RIGHT);
        cell.setCellStyle(style);

        cell = row.createCell(7);
        //TODO 这里的条数暂时不知道是什么字段数据,不知道是不是总条数的意思?
        cell.setCellValue(count);
        style = workbook.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_LEFT);
        cell.setCellStyle(style);


        //创建单元格边框
        style = getHssfCellStyle();

        //创建第一行元数据
        row = sheet.createRow(2);

        cell = row.createCell(0);
        cell.setCellValue("序号");
        cell.setCellStyle(style);

        cell = row.createCell(1);
        cell.setCellValue("题名");
        cell.setCellStyle(style);

        cell = row.createCell(2);
        cell.setCellValue("摄录日期");
        cell.setCellStyle(style);

        cell = row.createCell(3);
        cell.setCellValue("密级");
        cell.setCellStyle(style);

        cell = row.createCell(4);
        cell.setCellValue("片长");
        cell.setCellStyle(style);

        cell = row.createCell(5);
        cell.setCellValue("拍摄地点");
        cell.setCellStyle(style);


        cell = row.createCell(6);
        cell.setCellValue("摄录单位");
        cell.setCellStyle(style);

        cell = row.createCell(7);
        cell.setCellValue("摄录者");
        cell.setCellStyle(style);

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
                HSSFRow row = sheet.createRow(index+3);
                HSSFCell cell = null;

                //设置档序号
                cell = row.createCell(0);
                cell.setCellValue(index+1);
                cell.setCellStyle(getHssfCellStyle());      //设置边框

                //设置题名
                cell = row.createCell(1);
                cell.setCellValue(!StringUtils.isEmpty(metaData.getTitle())?metaData.getTitle():"");
                cell.setCellStyle(getHssfCellStyle());

                //设置摄录时间
                cell = row.createCell(2);
                cell.setCellValue(!StringUtils.isEmpty(metaData.getShootingDate())?metaData.getShootingDate():"");
                cell.setCellStyle(getHssfCellStyle());

                //设置密级
                cell = row.createCell(3);
                cell.setCellValue(ExcelViewJudge.getPrivilegeTemplateCodeString(!StringUtils.isEmpty(metaData.getSecretLevel())?metaData.getSecretLevel():""));
                cell.setCellStyle(getHssfCellStyle());

                //设置片长
                cell = row.createCell(4);
                if(!StringUtils.isEmpty(metaData.getVideoLength())){
                    Long l100 = Long.parseLong(metaData.getVideoLength());
                    String timeStr = ExcelViewJudge.getTimeLength(l100);
                    cell.setCellValue(timeStr);
                }
                cell.setCellStyle(getHssfCellStyle());

                //设置摄录地点
                cell = row.createCell(5);
                cell.setCellValue(!StringUtils.isEmpty(metaData.getShootingAddress())?metaData.getShootingAddress():"");
                cell.setCellStyle(getHssfCellStyle());

                //设置拍摄单位
                cell = row.createCell(6);
                cell.setCellValue(!StringUtils.isEmpty(metaData.getShootingUnit())?metaData.getShootingUnit():"");
                cell.setCellStyle(getHssfCellStyle());

                //设置拍摄者
                cell = row.createCell(7);
                cell.setCellValue(!StringUtils.isEmpty(metaData.getPhotographer())?metaData.getPhotographer():"");
                cell.setCellStyle(getHssfCellStyle());
                index++;
            }
        }
        writerExcel();
        return index;
    }


    /**
     * 设置页为,添加领导签字等行
     * @param indexRow
     */
    public void createExcelFooter(int indexRow){
        indexRow += 3;
        HSSFRow row = null;
        CellRangeAddress range = null;
        HSSFCell cell = null;
        row = sheet.createRow(indexRow);
        range = new CellRangeAddress(indexRow,indexRow,1,3);
        sheet.addMergedRegion(range);
        range = new CellRangeAddress(indexRow,indexRow,5,7);
        sheet.addMergedRegion(range);
        indexRow++;


        cell = row.createCell (1);
        cell.setCellValue("移交人  :  ____________");

        cell = row.createCell(5);
        cell.setCellValue("接收人  :  ____________");


        row = sheet.createRow(indexRow);
        range = new CellRangeAddress(indexRow,indexRow,1,3);
        sheet.addMergedRegion(range);
        range = new CellRangeAddress(indexRow,indexRow,5,7);
        sheet.addMergedRegion(range);

        cell = row.createCell(1);
        cell.setCellValue("领导签字: ____________");

        cell = row.createCell(5);
        cell.setCellValue("领导签字: ____________");
        indexRow++;

        row = sheet.createRow(indexRow);
        range = new CellRangeAddress(indexRow,indexRow,1,3);
        sheet.addMergedRegion(range);
        range = new CellRangeAddress(indexRow,indexRow,5,7);
        sheet.addMergedRegion(range);

        cell = row.createCell(1);
        cell.setCellValue("移交时间: ____________");

        cell = row.createCell(5);
        cell.setCellValue("移交时间: ____________");
        writerExcel();
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
            if(logger.isErrorEnabled()) logger.error("生成声像档案移交目录Excel文件出现错误"+e);
        }
    }


}
