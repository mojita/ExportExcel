package com.sobey.test;

import java.util.List;

import org.testng.annotations.Test;

import com.alibaba.fastjson.JSON;

/**
 * Created by lijunhong on 17/3/7.
 */
public class JSONTest {
    @Test
    public void test1(){
        String data= "[{'name':'长萨','age':'23','sex':'男'}]";
        List<Person> json1 =  JSON.parseArray(data,Person.class);

        for (Person p:json1){
            System.out.println(p.getName());
        }


    }

}
