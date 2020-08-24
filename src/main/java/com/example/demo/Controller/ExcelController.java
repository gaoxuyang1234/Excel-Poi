package com.example.demo.Controller;

import java.io.ByteArrayInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.PathVariable;
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
	   *     第一种：下载 ：根据项目中模板下载
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
		 excelUtils.download(request, response, "gxy",arrayList);
		 return new ResponseEntity<String>(HttpStatus.OK);
		
		 
	 }
	 /**
	      * 第二种：下载：生成excel文件到服务器，然后前端获得路径 下载
	  * @throws IOException
	  */
	 	@RequestMapping("/download2")
	    public void  testExcel2() throws IOException {
	        //创建工作薄对象
	        HSSFWorkbook workbook=new HSSFWorkbook();//这里也可以设置sheet的Name
	        //创建工作表对象
	        HSSFSheet sheet = workbook.createSheet();
	        //创建工作表的行
	        HSSFRow row = sheet.createRow(0);//设置第一行，从零开始
	        row.createCell(2).setCellValue("aaaaaaaaaaaa");//第一行第三列为aaaaaaaaaaaa
	        row.createCell(0).setCellValue(new Date());//第一行第一列为日期
	        workbook.setSheetName(0,"sheet的Name");//设置sheet的Name
	 
	        //文档输出
	        FileOutputStream out = new FileOutputStream("D:/123.xls");
	        workbook.write(out);
	        workbook.close();
	        out.close();
	        //返回路径给页面 页面直接下载
	        //return ResponseEntity.ok("D:/123.xls");
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
	 //这里-------根据需求存到数据库   等一系列操作
	 return ResponseEntity.ok("成功");
	 
	 }
	 
	 
}
