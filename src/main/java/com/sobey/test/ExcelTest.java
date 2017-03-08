package com.sobey.test;

import java.util.ArrayList;
import java.util.List;

import org.testng.annotations.Test;

import com.sobey.model.AudioAndVideoHandOver;
import com.sobey.model.MetaData;
import com.sobey.model.RecordBorrowRegistry;
import com.sobey.model.ThemePreviewExcel;

/**
 * Created by lijunhong on 17/3/8.
 */
public class ExcelTest {

    String filePath = "/Users/lijunhong/Desktop/importexcle/";
    String fileName = "TestTheme.xls";



    public List<List<MetaData>> initMetaData(){
        List<List<MetaData>> metaDataList = new ArrayList<>();

        int count = 1;
        for (int j=1;j<=4;j++) {
            List<MetaData> metaDatas = new ArrayList<>();
            for (int i = 0; i < 10; i++) {
                MetaData metaData = new MetaData();
                metaData.setFileId(count+ "");
                metaData.setTitle("中国"+count);
                metaData.setSecretLevel("保密");
                metaData.setFileState("在线");
                metaData.setPageNum(count + "");
                metaData.setFileNo("");
                metaData.setVideoLength("00:29:30:00");
                metaData.setShootingDate("2017-09-09");
                metaData.setShootingAddress("北京");
                metaData.setShootingUnit("NHK");
                metaData.setPhotographer("李白");
                metaDatas.add(metaData);
                count++;
            }
            metaDataList.add(metaDatas);
        }

        return metaDataList;
    }


    @Test
    public void testTableWidth(){

        ThemePreviewExcel themePreviewExcel = new ThemePreviewExcel(fileName,filePath);
        List<List<MetaData>> meListList = initMetaData();
        System.out.println(meListList.size());
        System.out.println(meListList.get(0).size());

        themePreviewExcel.createExcelTableHeader();
        themePreviewExcel.createExcelTableBody(meListList);
    }

    @Test
    public void testRecodRegistry(){
        RecordBorrowRegistry recordBorrowRegistry = new RecordBorrowRegistry(fileName,filePath);
        recordBorrowRegistry.createExcelTableHeader();
        recordBorrowRegistry.createExcelTableBody(initMetaData());
    }

    @Test
    public void testAudioAnd(){
        AudioAndVideoHandOver audioAndVideoHandOver = new AudioAndVideoHandOver(fileName,filePath);
        audioAndVideoHandOver.createExcelTableHeader();
        int index = audioAndVideoHandOver.createExcelTableBody(initMetaData());
        System.out.println(index);
        audioAndVideoHandOver.createExcelFooter(index);
    }





}
