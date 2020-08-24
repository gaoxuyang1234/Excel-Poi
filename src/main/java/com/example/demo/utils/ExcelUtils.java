package com.example.demo.utils;

import java.io.BufferedOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import com.example.demo.pojo.User;

@Component
public class ExcelUtils {
	 @Value("${FileTemplate}")
	 private String FileTemplate;
	
	/**
	 * excel下载 --（1）
	 * @author gxy
	 * @throws IOException 
	 */
	public void download(HttpServletRequest request,HttpServletResponse response,String fileName,List<User> list) throws IOException {
		
		//获得项目中模板
		InputStream in= this.getClass().getClassLoader().getResourceAsStream(FileTemplate);
		POIFSFileSystem fs=new POIFSFileSystem(in);
		//excel的文档对象
		HSSFWorkbook wb=new HSSFWorkbook(fs);
		//excel表单
		HSSFSheet sheet= wb.createSheet();
		//将模板中数据添加到生成的表单中
		sheet=copyRows(wb,wb.getSheetAt(0),sheet);
		int sheetIx=1;
		//将数据加到表单中
		fileName=setExportExcel(request,response,fileName,list,wb,sheetIx);
		if(StringUtils.isBlank(fileName)) {
			fileName=new SimpleDateFormat("yyyyMMdd").format(new Date());
		}
		
		response.setContentType("application/vnd.ms-excel;charset=UTF-8");
		response.setHeader("Content-Disposition","attachment;filename="+new String(fileName.getBytes("gb2312"),"ISO8859-1")+".xls");
		BufferedOutputStream buff=null;
		ServletOutputStream sos=null;
		try {
			sos=response.getOutputStream();
			buff=new BufferedOutputStream(sos);
			wb.removeSheetAt(0);
			wb.write(sos);
			buff.flush();
		}catch(Exception e) {
			e.printStackTrace();
		}finally {
			try {
				wb.close();
				buff.close();
				sos.close();
			}catch(Exception e) {
				e.printStackTrace();
			}
			
		}
	}

	
	/**
	 * 
	 * 获得excel 模板中的数据-2
	 * @author gxy
	 */
	private HSSFSheet copyRows(HSSFWorkbook wb, HSSFSheet sheetFrom, HSSFSheet sheetTo) {

		CellRangeAddress region=null;
		Row rowFrom=null,rowTo=null;
		Cell cellFrom=null,cellTo=null;
		//得到所有合并单元格
		int numMergedRegions = sheetFrom.getNumMergedRegions();
		if(numMergedRegions!=0) {
			for(int i=0;i<numMergedRegions;i++) {
				//得到某一个单元格
				region=sheetFrom.getMergedRegion(i);
				if(region.getFirstColumn()>=sheetFrom.getFirstRowNum()&&region.getLastRow()<= sheetFrom.getLastRowNum()) {
					sheetTo.addMergedRegion(region);
				}
			}
		}
	
		for(int intRow=sheetFrom.getFirstRowNum();intRow<=sheetFrom.getLastRowNum();intRow++) {
			rowFrom=sheetFrom.getRow(intRow);
			rowTo=sheetTo.createRow(intRow);
			if(rowFrom==null) {
				continue;
			}
			rowTo.setHeight(rowFrom.getHeight());
			for(int intCol=0;intCol<rowFrom.getLastCellNum();intCol++) {
				if(sheetFrom.getColumnStyle(intCol)!=null) {
					sheetTo.setDefaultColumnStyle(intCol, sheetFrom.getColumnStyle(intCol));
				}
				sheetTo.setColumnWidth(intCol, sheetFrom.getColumnWidth(intCol));
				cellFrom=rowFrom.getCell(intCol);
				cellTo=rowTo.createCell(intCol);
				if(cellFrom==null) {
					continue;
				}
				cellTo.setCellStyle(cellFrom.getCellStyle());
				String stringCellValue = cellFrom.getStringCellValue();//纯数字的话这可能会有错 解决办法： https://blog.csdn.net/ysughw/article/details/9288307
				if(!StringUtils.isBlank(stringCellValue)) {
					cellTo.setCellValue(stringCellValue);
					
				}
			}
			
		}
		sheetTo.setDisplayGridlines(true);
		sheetTo.setZoom(100);
		return sheetTo;
	}
	
	
	/**
	 * 
	 * 根据模板填充数据并生成excel -3
	 * @author gxy
	 */
	
	private String setExportExcel(HttpServletRequest request, HttpServletResponse response, String fileName,
			List<User> list, HSSFWorkbook wb, int sheetIx) {
		HSSFSheet sheet = wb.getSheetAt(sheetIx);
		fileName=fileName+new SimpleDateFormat("yyyMMdd").format(new Date());
		wb.setSheetName(sheetIx, fileName);
		sheet.setForceFormulaRecalculation(true);
		for(int i=0;i<list.size();i++) {
			User user = list.get(i);
			HSSFRow rowCurrent = sheet.getRow(i+1);
			rowCurrent.createCell(1).setCellValue(user.getAge());
			rowCurrent.createCell(2).setCellValue(user.getName());
		}
		
		return fileName;
	}

	
	
	
	//====================================上传需要的工具类=========================================
	

	 
	 //加载excel 模板+判断类型
	 public Workbook getWorkbook(InputStream is,String fileName) throws Exception{
	 Workbook wb=null;
	 String fileType=fileName.substring(fileName.lastIndexOf("."));
	 if(".xls".equals(fileType)){
	 wb=new HSSFWorkbook(is);
	 }else if(".xlsx".equals(fileType)){
	 wb=new XSSFWorkbook(is);
	 }else{
	 throw new Exception("格式解析错误");
	 }
	 return wb;
	 }
	 
	 
	 //格式化单元格值
	public String getCellValue(Cell cell) {
		String result = "";
		if (cell != null) {
			switch (cell.getCellTypeEnum()) {
			case NUMERIC:
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					SimpleDateFormat sdf = null;
					if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {
						sdf = new SimpleDateFormat("HH:mm");
					} else {
						sdf = new SimpleDateFormat("yyyy-MM-dd");
					}
					Date date = cell.getDateCellValue();
					result = sdf.format(date);
				}

				else if (cell.getCellStyle().getDataFormat() == 58) {
					SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
					double value = cell.getNumericCellValue();
					Date date = DateUtil.getJavaDate(value);
					result = sdf.format(date);
				}
				else {
					DecimalFormat df = new DecimalFormat();
					df.setGroupingUsed(false);
					result = String.valueOf(df.format(cell.getNumericCellValue()));
				}
				break;

			case STRING:
				result = String.valueOf(cell.getStringCellValue());
				break;

			case BLANK:
				result = "";
				break;
			default:
				result = "";
				break;
			}

		}
		return result;
	}
	
}
