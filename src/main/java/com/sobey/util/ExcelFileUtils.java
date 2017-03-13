package com.sobey.util;

import java.io.File;
import java.util.Date;
import java.util.List;
import java.util.Properties;

import org.apache.log4j.Logger;
import org.springframework.util.StringUtils;

import com.sobey.model.AudioAndVideoHandOver;
import com.sobey.model.MetaData;
import com.sobey.model.RecordBorrowRegistry;
import com.sobey.model.ThemePreviewExcel;

/**
 * Created by lijunhong on 17/3/8.
 * 获取配置一些配置信息
 */
public class ExcelFileUtils {

    public static final String RECORD_BORROW_REGISTRY_MODEL = "0";          //代表借阅登记Excel
    public static final String THEME_PREVIEWEXCEL_MODEL = "1";              //代表主题一览Excel
    public static final String AUDIO_AND_VIDEO_HANDOVER_MODEL = "2";        //代表声像移交目录Excel

    private static Logger logger = Logger.getLogger(ExcelFileUtils.class);
    public static String filePath;          //这里是获取到jar所在的文件路径:文件路径后面带有"/"
    public static String fileName = "";     //记录了根据model匹配的文件名
    public static Boolean isWindosSystem = false;

    static {

        //获取当前运行的系统名,判断是在什么系统下运行
        Properties properties = System.getProperties();
        String name = properties.getProperty("os.name");
        System.out.println(name);

        String allPath = ExcelFileUtils.class
                .getProtectionDomain()
                .getCodeSource()
                .getLocation()
                .getPath();
        String path = null;


        if(name.startsWith("win") || name.startsWith("Win")){
            //windows操作系统  file:/c:/test/test.jar!/BOOT/..
            int lastIndexOf = allPath.lastIndexOf(".jar")+4;
            String jarPath = allPath.substring(0,lastIndexOf);
            int pathIndexOf = jarPath.lastIndexOf("/")+1;
            int startFileIndexOf = jarPath.indexOf("/")+1;
            path = jarPath.substring(startFileIndexOf,pathIndexOf);
            isWindosSystem = true;
        }else {
            //mac os x file:/Users/libai/test.jar!/BOOT/..
            int lastIndexOf = allPath.lastIndexOf(".jar")+4;
            String jarPath = allPath.substring(0,lastIndexOf);
            int pathIndexOf = jarPath.lastIndexOf(File.separator)+1;
            int startFileIndexOf = jarPath.indexOf(":")+1;
            if(startFileIndexOf>0) {
                path = jarPath.substring(startFileIndexOf, pathIndexOf);
            }else {
                path = jarPath.substring(0,pathIndexOf);

            }
        }


        if(!StringUtils.isEmpty(path)) {
            filePath = path;
            logger.info("filePath:"+filePath);
            //TODO 这里实测是的数据
//            filePath = "/Users/lijunhong/Desktop/jsonp/";
        }else {
            if(logger.isErrorEnabled()) logger.error("Not found .jar path");
        }
    }

    /**
     * 将获取到的文件路径返
     * @return
     */
    public static String getFilePath(){
        return filePath;
    }

    /**
     * model是Excel的文件类型0:借阅登记,1:主题一览,2:声像移交目录
     * @param model
     */
    public static String getFileName(String model){
        Date date = new Date();
        Long time = date.getTime();
        System.out.println(time);
        System.out.println("0getFileName"+model);
        if(!StringUtils.isEmpty(model) && RECORD_BORROW_REGISTRY_MODEL.equals(model)) {
            fileName = time + "registry.xls";
            System.out.println("1"+fileName);
        }
        if(!StringUtils.isEmpty(model) && THEME_PREVIEWEXCEL_MODEL.equals(model)){
            fileName = time + "theme.xls";
            System.out.println("2"+fileName);
        }
        if(!StringUtils.isEmpty(model) && AUDIO_AND_VIDEO_HANDOVER_MODEL.equals(model)){
            fileName = time + "handOver.xls";
            System.out.println("3"+fileName);
        }
        System.out.println("into getFileName:"+fileName);
        return fileName;
    }


    /**
     * 根据Model和数据创建Excel文件
     * @param model 文件模板类型
     * @param metaDataList  生成Excel的数据
     */
    public static String createExcelByModel(String model, List<List<MetaData>> metaDataList,int count){
        String fileName = getFileName(model);
        System.out.println("fileName:(create)"+fileName);
        //生成借查阅等级单Excel
        if(!StringUtils.isEmpty(model) && RECORD_BORROW_REGISTRY_MODEL.equals(model) && metaDataList != null){
            RecordBorrowRegistry recordBorrowRegistry = new RecordBorrowRegistry(fileName,filePath);
            recordBorrowRegistry.createExcelTableHeader();
            recordBorrowRegistry.createExcelTableBody(metaDataList);
            if(logger.isInfoEnabled()) logger.info("生成借查阅登记Excel");
        }
        //生成主题一览Excel
        if(!StringUtils.isEmpty(model) && THEME_PREVIEWEXCEL_MODEL.equals(model) && metaDataList != null){
            ThemePreviewExcel themePreviewExcel = new ThemePreviewExcel(fileName,filePath);
            themePreviewExcel.createExcelTableHeader();
            themePreviewExcel.createExcelTableBody(metaDataList);
            if(logger.isInfoEnabled()) logger.info("生成主题一览Excel");
        }
        //生成声像移交目录Excel
        if(!StringUtils.isEmpty(model) && AUDIO_AND_VIDEO_HANDOVER_MODEL.equals(model) && metaDataList != null){
            AudioAndVideoHandOver audioAndVideoHandOver = new AudioAndVideoHandOver(fileName,filePath);
            audioAndVideoHandOver.createExcelTableHeader(count);
            int indexRow = audioAndVideoHandOver.createExcelTableBody(metaDataList);
            audioAndVideoHandOver.createExcelFooter(indexRow);
            if(logger.isInfoEnabled()) logger.info("生成声像移交目录Excel");
        }

        return fileName;
    }


    /**
     * 删除指定文件名的文件
     * @param fileName
     */
    public static void deleteExcelFile(String fileName){
        File file = new File(filePath+fileName);
        if(file.exists()){
            file.delete();
        }
    }


}
