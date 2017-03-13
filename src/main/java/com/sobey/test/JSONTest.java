package com.sobey.test;

import java.util.*;

import org.testng.annotations.Test;

import com.alibaba.fastjson.JSON;
import com.sobey.model.MetaData;
import com.sobey.service.ExportExcelApplication;

/**
 * Created by lijunhong on 17/3/7.
 */
public class JSONTest {

    public static List<List<MetaData>> metaDataList = new ArrayList<>();

    @Test
    public void test1(){
        String data= "[{'name':'长萨','age':'23','sex':'男'}]";
        List<Person> json1 =  JSON.parseArray(data,Person.class);

        for (Person p:json1){
            System.out.println(p.getName());
        }
    }


    //判断是否可以执行创建Excel并且提供下载功能测试
    @Test
    public void testRequestGetValue(){

//        while (true) {
//            if (ExportExcelApplication.isCreateExcel) {
//                metaDataList = ExportExcelApplication.metaDatas;
//                ExportExcelApplication.isCreateExcel = false;
//                if(metaDataList!=null){
//                    for (List<MetaData> metaDatas:metaDataList){
//                        for (MetaData metaData:metaDatas){
//                            System.out.println(metaData.getTitle());
//                        }
//                    }
//                }
//                return;
//            }
//        }

        if(ExportExcelApplication.isCreateExcel){

            for (List<MetaData> metaDatas:metaDataList) {
                for (MetaData metaData:metaDatas) {
                    System.out.println(metaData.getTitle());
                }
            }
        }




    }


}
