package com.example.demo.Controller;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import com.example.demo.pojo.User;
import com.example.demo.utils.ExcelUtils;

@Controller
public class ExcelController {

	@Autowired
	private ExcelUtils excelUtils;
	/**
	 * 下载
	 * @param request
	 * @param response
	 * @return
	 * @throws IOException
	 */
	 @RequestMapping("/download")
	public ResponseEntity<String> download(HttpServletRequest request,HttpServletResponse response) throws IOException{
		 ArrayList<User> arrayList = new ArrayList<User>();
		User user = new User();
		user.setName("gxy");
		user.setAge("12");
		arrayList.add(user);
		 excelUtils.download(request, response, "gxy",arrayList );
		 
		 return new ResponseEntity<String>(HttpStatus.OK);
		
		 
	 }
	 
	 
	 /**
	  * 上传
	  * @param request
	  * @return
	 * @throws Exception 
	  */
	 @RequestMapping("/upload")
	 public ResponseEntity<String> upload(HttpServletRequest request) throws Exception{
	 //拿到名为file的excel
	 MultipartHttpServletRequest multipartRequest= (MultipartHttpServletRequest)request;
	 MultipartFile file = multipartRequest.getFile("file");
	 //文件作为输入流
	InputStream in = new ByteArrayInputStream(file.getBytes());
	 //得到文件名
	 String fileName = file.getOriginalFilename();
	 //判断类型+得到excel 对象
	 Workbook wb = excelUtils.getWorkbook(in, fileName);
	 //得到第一个表单；
	 Sheet sheet = wb.getSheetAt(0);
	 //存数据的list
	 ArrayList<User> list = new ArrayList<User>();
	 //行循环
	 for(int i=sheet.getFirstRowNum();i<=sheet.getLastRowNum();i++) {
		 User u = new User();
		 int j=0;
		 u.setName(excelUtils.getCellValue(sheet.getRow(i).getCell(j++)));
		 u.setAge(excelUtils.getCellValue(sheet.getRow(i).getCell(j++)));
		 list.add(u);
	 }
	 
	 //关闭输入流
	 in.close();
	 //根据需求存到数据库   等一系列操作
	 return ResponseEntity.ok("成功");
	 
	 }
	 
	 
}
