package com.example.demo.Controller;
/**
 * 
 * @author g1996
 * 	项目中文件下载
 *
 */

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

@Controller
public class DownloadController {

	
	
	/**
	 * 下载文件
	 * @throws IOException 
	 */
	@RequestMapping("/fileDownload")
	@ResponseBody
	public ResponseEntity<String> Download(HttpServletResponse response) throws IOException{
		File file = new File(ResourceUtils.getFile(ResourceUtils.CLASSPATH_URL_PREFIX),"/templates/1.xls");
		if(file.exists()) {
			response.setContentType("application/force-download");
			response.addHeader("Content-Disposition", "attachment;fileName="+"1.xls");
			FileInputStream input = new FileInputStream(file);
			ServletOutputStream out = response.getOutputStream();
			int real;
			byte[] bytes = new byte[1024];
			while ((real=input.read(bytes))!=-1) {
				out.write(bytes, 0, real);
			}
			input.close();
			out.close();
			return ResponseEntity.status(HttpStatus.OK).body("成功");
		}
		return ResponseEntity.status(HttpStatus.NOT_FOUND).body("文件不存在");
		
		
	}
}
