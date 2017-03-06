package com.example;

import java.io.*;
import javax.servlet.http.HttpServletResponse;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

/**
 * 提供对DCMP检索素材封装并生成Excel文件,提供客户端下载功能
 */
@SpringBootApplication
@RestController
public class ExportExcelApplication {

	@RequestMapping(value = "/")
	public String index(){
		return "hello";
	}



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
