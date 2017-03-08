package com.sobey.util;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.log4j.Logger;
import org.springframework.util.StringUtils;

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

    static {
        String allPath = ExcelFileUtils.class
                .getProtectionDomain()
                .getCodeSource()
                .getLocation()
                .getPath();

        int lastIndexOf = allPath.lastIndexOf(File.separator)+1;
        String path = allPath.substring(0,lastIndexOf);
        if(!StringUtils.isEmpty(path)) {
            filePath = path;
        }else {
            if(logger.isErrorEnabled()) logger.error("获取不到jar文件路径");
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
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String dateStr = simpleDateFormat.format(date);
        String fileName = "";
        if(!StringUtils.isEmpty(model)&&"0".equals(model)){
            fileName = dateStr+"registry.xls";
        }
        if(!StringUtils.isEmpty(model)&&"1".equals(model)){
            fileName = dateStr+"theme.xls";
        }
        if(!StringUtils.isEmpty(model)&&"2".equals(model)){
            fileName = dateStr+"handOver.xls";
        }
        return fileName;
    }


}
