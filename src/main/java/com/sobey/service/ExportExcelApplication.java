package com.sobey.service;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.log4j.Logger;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.*;

import com.alibaba.fastjson.JSON;
import com.sobey.model.MetaData;

/**
 * 提供对DCMP检索素材封装并生成Excel文件,提供客户端下载功能
 */
@SpringBootApplication
@RestController
public class ExportExcelApplication {


	private Logger logger = Logger.getLogger(ExportExcelApplication.class);

	public static List<List<MetaData>> metaDatas = new ArrayList<>();		//存放前台需要生成Excel的元数据

	//下面这两个参数是用来计算请求数量是否满足条件,当满足条件时才开始生成Excel文档并且提供下载
	public static int requestCount;		//这里记录请求的次数
	public static int metaDataCount;	//数据发送的总条数

	@RequestMapping(value = "/")
	public String index(@RequestParam(value = "data",defaultValue = "hello")String data){
		return "hello"+data;
	}


	/**
	 * 获得从前台传输的数据并且进行数据的封装
	 * @param request
     * @return
     */
	@RequestMapping(value = "/getAjaxValue")
	@ResponseBody
	public String getAjaxValue(HttpServletRequest request,HttpServletResponse response){


		String reslut = "{reslut:'true'}";
		String dataList = request.getParameter("datalist");
		String callBack = request.getParameter("callback");
		String count = request.getParameter("metaDataCount");
		if(!StringUtils.isEmpty(metaDataCount))
			metaDataCount = Integer.parseInt(count);
		System.out.println(count);


		if(dataList!=null) {
			List<MetaData> metaDatasTemp = JSON.parseArray(dataList, MetaData.class);
			if(logger.isInfoEnabled()) logger.info("获取ExcelMetaData数据条数:"+metaDatasTemp.size());
			requestCount++;
		}else {
			if(logger.isWarnEnabled()) logger.warn("前台构造的Excel数据为null,没有货取到MeataData数据");
		}

		if(!StringUtils.isEmpty(callBack)){
			reslut = callBack+"("+reslut+")";
		}
		return reslut;
	}


	/**
	 * 该方法主要是提供下载功能
	 * @param res
	 * @throws IOException
     */
	@RequestMapping(value="/testDownload",method= RequestMethod.GET)
	public void testDownload(HttpServletResponse res) throws IOException {
		String fileName="test.xls";
		res.setHeader("content-type", "application/octet-stream");
		res.setContentType("application/octet-stream");
		res.setHeader("Content-Disposition", "attachment;filename="+fileName);
		File file=new File("c:/test.xls");
		res.setContentLengthLong(file.length());

		InputStream in = new FileInputStream(file);
		OutputStream out = res.getOutputStream();

		int len = 0;
		byte[] b = new byte[10*1024];
		while ((len=in.read(b))!=-1){
			out.write(b,0,len);
		}

		in.close();
		out.flush();
		out.close();
	}



	public static void main(String[] args) {
		SpringApplication.run(ExportExcelApplication.class, args);
	}
}
