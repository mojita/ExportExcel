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
import com.sobey.util.ExcelFileUtils;

/**
 * 提供对DCMP检索素材封装并生成Excel文件,提供客户端下载功能
 */
@SpringBootApplication
@RestController
public class ExportExcelApplication {


	private Logger logger = Logger.getLogger(ExportExcelApplication.class);

	public static List<List<MetaData>> metaDatas = new ArrayList<>();		//存放前台需要生成Excel的元数据
	public int metaDataCount = 0;											//传输过来的数据总条数
	public int meataDataLength = 0;											//传输过来的数据条数的统计,用于判断是否传输完毕
	public static Boolean isCreateExcel = false;							//是否能创建Excel文件标记
	public String model;													//用于判断生成的Excel模板类型



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
	public String getAjaxValue(HttpServletRequest request,HttpServletResponse response) throws InterruptedException {


		String reslut = "{reslut:'true'}";

		String dataList = request.getParameter("datalist");
		String callBack = request.getParameter("callback");
		model = request.getParameter("excelModelModelValue");
//		metaDataCount = Integer.parseInt(request.getParameter("metaDataCount"));


		if(dataList!=null) {
			List<MetaData> metaDatasTemp = JSON.parseArray(dataList, MetaData.class);
			meataDataLength += metaDatasTemp.size();
			metaDatas.add(metaDatasTemp);
			if(logger.isInfoEnabled()) logger.info("获取ExcelMetaData数据条数:"+metaDatasTemp.size());
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
	public void testDownload(HttpServletResponse res,HttpServletRequest request) throws IOException, InterruptedException {

		Thread.sleep(500L);									//为了防止下载的请求快于数据请求

		//封装数据,生成Excel文档
		if(StringUtils.isEmpty(model)) {
			if(logger.isErrorEnabled()) logger.error("用户没有选择Excel的生成模板,不能生成模板并且提供下载");
		}
		String fileName = ExcelFileUtils.createExcelByModel(model, metaDatas,meataDataLength);
		System.out.println("11111"+fileName);
		if(StringUtils.isEmpty(fileName)){
		    logger.error("fileName is null");
		}
		downloadExcel(res, ExcelFileUtils.getFilePath(),fileName);
		ExcelFileUtils.deleteExcelFile(fileName);



		metaDatas.clear();  	//将集合中的数据清零
		meataDataLength = 0; 	//将统计的条数清零

	}

	/**
	 * 提供下载功能
	 * @param res
	 * @param filePath
	 * @param fileName
	 * @throws IOException
     */
	private void downloadExcel(HttpServletResponse res,String filePath,String fileName) throws IOException {
		
		System.out.println("filename"+fileName);
		System.out.println("filepath"+filePath);
		
		res.setHeader("content-type", "application/octet-stream");
		res.setContentType("application/octet-stream");
		res.setHeader("Content-Disposition", "attachment;filename="+fileName);
		File file = null;
		if(ExcelFileUtils.isWindosSystem){
			file = new File(filePath+"/"+fileName);
		}else {
			file=new File(filePath+fileName);
		}
		System.out.println("filePath:--"+filePath+":::"+fileName);
//		File file = new File("/Users/lijunhong/Desktop/jsonp/Test.xml");
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
