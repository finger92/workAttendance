package workAttendance;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	static long nd = 1000 * 24 * 60 * 60;// 一天的毫秒数    
    static long nh = 1000 * 60 * 60;// 一小时的毫秒数    
    static long nm = 1000 * 60;// 一分钟的毫秒数
	static List<PersonAttendBean> pab = new ArrayList<PersonAttendBean>();
	
	/**
	 * 对外提供读取excel 的方法
	 * */
	public static List<List<Object>> readExcel(File file) throws IOException {
		String fileName = file.getName();
		String extension = fileName.lastIndexOf(".") == -1 ? "" : fileName
				.substring(fileName.lastIndexOf(".") + 1);
		if ("xls".equals(extension)) {
			return read2003Excel(file);
		} else if ("xlsx".equals(extension)) {
			return read2007Excel(file);
		} else {
			throw new IOException("不支持的文件类型");
		}
	}

	/**
	 * 读取 office 2003 excel
	 * 
	 * @throws IOException
	 * @throws FileNotFoundException
	 */
	private static List<List<Object>> read2003Excel(File file)
			throws IOException {
		List<List<Object>> list = new LinkedList<List<Object>>();
		HSSFWorkbook hwb = new HSSFWorkbook(new FileInputStream(file));
		HSSFSheet sheet = hwb.getSheetAt(0);
		Object value = null;
		HSSFRow row = null;
		HSSFCell cell = null;
		int counter = 0;
		for (int i = sheet.getFirstRowNum(); counter < sheet
				.getPhysicalNumberOfRows(); i++) {
			row = sheet.getRow(i);
			if (row == null) {
				continue;
			} else {
				counter++;
			}
			List<Object> linked = new LinkedList<Object>();
			for (int j = 0; j <= row.getLastCellNum(); j++) {
				cell = row.getCell(j);
				if (cell == null) {
					linked.add("");
					continue;
				}
				DecimalFormat df = new DecimalFormat("0");// 格式化 number String
															// 字符
				SimpleDateFormat sdf = new SimpleDateFormat(
						"yyyy-MM-dd HH:mm:ss");// 格式化日期字符串
				DecimalFormat nf = new DecimalFormat("0.00");// 格式化数字
				switch (cell.getCellType()) {
				case XSSFCell.CELL_TYPE_STRING:
					//System.out.println(i + "行" + j + " 列 is String type");
					value = cell.getStringCellValue();
					break;
				case XSSFCell.CELL_TYPE_NUMERIC:
					//System.out.println(i + "行" + j
					//		+ " 列 is Number type ; DateFormt:"
					//		+ cell.getCellStyle().getDataFormatString());
					if ("@".equals(cell.getCellStyle().getDataFormatString())) {
						value = df.format(cell.getNumericCellValue());
					} else if ("General".equals(cell.getCellStyle()
							.getDataFormatString())) {
						value = nf.format(cell.getNumericCellValue());
					} else {
						value = sdf.format(HSSFDateUtil.getJavaDate(cell
								.getNumericCellValue()));
					}
					break;
				case XSSFCell.CELL_TYPE_BOOLEAN:
					//System.out.println(i + "行" + j + " 列 is Boolean type");
					value = cell.getBooleanCellValue();
					break;
				case XSSFCell.CELL_TYPE_BLANK:
					//System.out.println(i + "行" + j + " 列 is Blank type");
					value = "";
					break;
				default:
					//System.out.println(i + "行" + j + " 列 is default type");
					value = cell.toString();
				}
//				if (value == null || " ".equals(value)) {
//					continue;
//				}
				linked.add(value);
			}
			list.add(linked);
		}
		return list;
	}

	/**
	 * 读取Office 2007 excel
	 * */
	private static List<List<Object>> read2007Excel(File file)
			throws IOException {
		List<List<Object>> list = new LinkedList<List<Object>>();
		// 构造 XSSFWorkbook 对象，strPath 传入文件路径
		XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(file));
		// 读取第一章表格内容
		XSSFSheet sheet = xwb.getSheetAt(0);
		Object value = null;
		XSSFRow row = null;
		XSSFCell cell = null;
		int counter = 0;
		for (int i = sheet.getFirstRowNum(); counter < sheet
				.getPhysicalNumberOfRows(); i++) {
			row = sheet.getRow(i);
			if (row == null) {
				continue;
			} else {
				counter++;
			}
			List<Object> linked = new LinkedList<Object>();
			for (int j = 0; j <= row.getLastCellNum(); j++) {
				cell = row.getCell(j);
				if (cell == null) {
					linked.add("");
					continue;
				}
				DecimalFormat df = new DecimalFormat("0");// 格式化 number String
															// 字符
				SimpleDateFormat sdf = new SimpleDateFormat(
						"yyyy-MM-dd HH:mm:ss");// 格式化日期字符串
				DecimalFormat nf = new DecimalFormat("0.00");// 格式化数字
				switch (cell.getCellType()) {
				case XSSFCell.CELL_TYPE_STRING:
					//System.out.println(i + "行" + j + " 列 is String type");
					value = cell.getStringCellValue();
					break;
				case XSSFCell.CELL_TYPE_NUMERIC:
					//System.out.println(i + "行" + j
					//		+ " 列 is Number type ; DateFormt:"
					//		+ cell.getCellStyle().getDataFormatString());
					if ("@".equals(cell.getCellStyle().getDataFormatString())) {
						value = df.format(cell.getNumericCellValue());
					} else if ("General".equals(cell.getCellStyle()
							.getDataFormatString())) {
						value = nf.format(cell.getNumericCellValue());
					} else {
						value = sdf.format(HSSFDateUtil.getJavaDate(cell
								.getNumericCellValue()));
					}
					break;
				case XSSFCell.CELL_TYPE_BOOLEAN:
					//System.out.println(i + "行" + j + " 列 is Boolean type");
					value = cell.getBooleanCellValue();
					break;
				case XSSFCell.CELL_TYPE_BLANK:
					//System.out.println(i + "行" + j + " 列 is Blank type");
					value = "";
					break;
				default:
					//System.out.println(i + "行" + j + " 列 is default type");
					value = cell.toString();
				}
//				if (value == null || "".equals(value)) {
//					continue;
//				}
				linked.add(value);
			}
			list.add(linked);
		}
		return list;
	}

	public static Boolean passClearWords(String ts){
		String[] clearWords = PropertyUtil.get("config.properties", "CLEAR_WORDS").split(",");
		for(int i=0;i<clearWords.length;i++){
			if(ts.equals(clearWords[i])){
				return false;
			}
		}
		return true;
	}
	
	public static PersonAttendBean getPABFromList(String name){
		for(int i=0;i<pab.size();i++){
			if(pab.get(i).getName().equals(name)){
				return pab.get(i);
			}
		}
		PersonAttendBean temp = new PersonAttendBean(name);
		pab.add(temp);
		return temp;
	}
	
	private static String readDataFromConsole(String prompt) {  
        BufferedReader br = new BufferedReader(new InputStreamReader(System.in));  
        String str = null;  
        try {  
            System.out.println(prompt);  
            str = br.readLine();  
  
        } catch (IOException e) {  
            e.printStackTrace();  
        }  
        return str;  
    }
	
	private static boolean testYesterdayExtraWork(List<PersonBean> pbList, String date){
		Calendar c = Calendar.getInstance();
		try {
			c.setTime(new SimpleDateFormat("yyyy-MM-dd").parse(date));
		} catch (ParseException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		c.add(Calendar.DAY_OF_MONTH, -1);  
        String yesDay = new SimpleDateFormat("yyyy-MM-dd").format(c.getTime());
		for(PersonBean pb : pbList){
			if(pb.attendDate.equals(yesDay)&&pb.getFourOut()!=null){
				return true;
			}
		}
		return false;
	}
	
	public static void main(String[] args) {
		Date shangbanTime = null;
		Date xiabanTime = null;
		try{
			shangbanTime = new SimpleDateFormat("HH:mm").parse("08:30");
		    xiabanTime = new SimpleDateFormat("HH:mm").parse("17:30");
	    }catch(Exception e){
	    	e.printStackTrace();
	    }
		List<PersonBean> pList = new ArrayList<PersonBean>();
		try {
			String str = readDataFromConsole("请输入需要录入的考勤记录表，如 D:\\考勤记录（2016.9.23-2016.10.09）.xls：");
			String search = readDataFromConsole("请输入检索筛选的时间范围，如2016-09-26 2016-09-30（包含边界），为空则检索全部信息：");
			System.out.println("正在读取考勤信息...");
			List<List<Object>> list = readExcel(new File(str));
			System.out.println("正在统计缺勤情况...");
			for(int i=0;i<list.size();i++){
				if(list.get(i).size()==0
						||!passClearWords((String) list.get(i).get(0))
						||"".equals(list.get(i).get(0))){
					continue;
				}
				PersonBean pb = new PersonBean();
				pb.setName((String)list.get(i).get(Integer.parseInt(
						PropertyUtil.get("config.properties", "NAME"))));
				pb.setAttendDate((String)list.get(i).get(Integer.parseInt(
						PropertyUtil.get("config.properties", "ATTENDDATE"))));
				if(!((String)list.get(i).get(Integer.parseInt(
						PropertyUtil.get("config.properties", "SECONDIN")))).equals("")){
					pb.setSecondIn(new SimpleDateFormat("HH:mm").parse((String)list.get(i).get(Integer.parseInt(
							PropertyUtil.get("config.properties", "SECONDIN")))));
				}else{
					pb.setSecondIn(null);
				}
				if(!((String)list.get(i).get(Integer.parseInt(
						PropertyUtil.get("config.properties", "SECONDOUT")))).equals("")){
					pb.setSecondOut(new SimpleDateFormat("HH:mm").parse((String)list.get(i).get(Integer.parseInt(
							PropertyUtil.get("config.properties", "SECONDOUT")))));
				}else{
					pb.setSecondOut(null);
				}
				if(!((String)list.get(i).get(Integer.parseInt(
						PropertyUtil.get("config.properties", "THIRDIN")))).equals("")){
					pb.setThirdIn(new SimpleDateFormat("HH:mm").parse((String)list.get(i).get(Integer.parseInt(
							PropertyUtil.get("config.properties", "THIRDIN")))));
				}else{
					pb.setThirdIn(null);
				}
				if(!((String)list.get(i).get(Integer.parseInt(
						PropertyUtil.get("config.properties", "THIRDOUT")))).equals("")){
					pb.setThirdOut(new SimpleDateFormat("HH:mm").parse((String)list.get(i).get(Integer.parseInt(
							PropertyUtil.get("config.properties", "THIRDOUT")))));
				}else{
					pb.setThirdOut(null);
				}
				if(!((String)list.get(i).get(Integer.parseInt(
						PropertyUtil.get("config.properties", "FOURIN")))).equals("")){
					pb.setFourIn(new SimpleDateFormat("HH:mm").parse((String)list.get(i).get(Integer.parseInt(
							PropertyUtil.get("config.properties", "FOURIN")))));
				}else{
					pb.setFourIn(null);
				}
				if(!((String)list.get(i).get(Integer.parseInt(
						PropertyUtil.get("config.properties", "FOUROUT")))).equals("")){
					pb.setFourOut(new SimpleDateFormat("HH:mm").parse((String)list.get(i).get(Integer.parseInt(
							PropertyUtil.get("config.properties", "FOUROUT")))));
				}else{
					pb.setFourOut(null);
				}
				pList.add(pb);
			}
			
			//设置检索信息;
			Date dateBegin = null;
			Date dateEnd = null;
			long days = 0;
			if(!search.equals("")){
				dateBegin = new SimpleDateFormat("yyyy-MM-dd").parse(search.split(" ")[0]);
				dateEnd = new SimpleDateFormat("yyyy-MM-dd").parse(search.split(" ")[1]);	
				days = (dateEnd.getTime()-dateBegin.getTime())/(24*60*60*1000);
			}
			
			for(int j=0;j<pList.size();j++){
				PersonBean tempPb = pList.get(j);
				//加入检索信息		
				Date infoDate = new SimpleDateFormat("yyyy-MM-dd").parse(tempPb.getAttendDate());
				if(dateBegin!=null&&dateEnd!=null&&(
						infoDate.getTime()<dateBegin.getTime()||infoDate.getTime()>dateEnd.getTime())){
					continue;
				}
				PersonAttendBean tempPab = getPABFromList(pList.get(j).getName());
				String addString = "";
				Boolean weishua=false,chidao=false,zaotui=false;
				if(tempPb.getSecondIn()==null){
					if(!testYesterdayExtraWork(pList, tempPb.getAttendDate())){
						weishua=true;
						tempPab.weishua++;
						addString="上午未刷，";
					}
				}
//				if(tempPb.getFourIn()==null
//						&&tempPb.getThirdOut()==null){
//					weishua=true;
//					tempPab.weishua++;
//					addString+="中午未刷，";
//				}
				if(tempPb.getSecondOut()==null){
					weishua=true;
					tempPab.weishua++;
					addString+="下午未刷，";
				}
				if(tempPb.getSecondIn()==null
						&&tempPb.getSecondOut()==null){
					weishua=true;
					addString = "全天未刷，";
				}
				if(tempPb.getSecondIn()!=null && tempPb.getSecondIn().getTime()>shangbanTime.getTime()){
					long chidaoTime = tempPb.getSecondIn().getTime()-shangbanTime.getTime();  
					long day=chidaoTime/(24*60*60*1000);
					long hour=(chidaoTime/(60*60*1000)-day*24);
					long min=((chidaoTime/(60*1000))-day*24*60-hour*60);
					chidao=true;
					addString+="上午迟到"+(hour*60+min)+"分钟，";
				}
				if(tempPb.getSecondOut()!=null && tempPb.getSecondOut().getTime()<xiabanTime.getTime()){
					long chidaoTime = xiabanTime.getTime()-tempPb.getSecondOut().getTime();  
					long day=chidaoTime/(24*60*60*1000);
					long hour=(chidaoTime/(60*60*1000)-day*24);
					long min=((chidaoTime/(60*1000))-day*24*60-hour*60);
					zaotui=true;
					addString+="下午早退"+(hour*60+min)+"分钟，";
				}
				if(chidao)tempPab.chidao++;
				if(zaotui)tempPab.zaotui++;
				//没有异常记录记满勤一天,一周清一次
				if(!weishua&&!chidao&&!zaotui){
					tempPab.mqDays++;
				}
				
				if(!addString.equals("")){
					String formatDate = Integer.parseInt(tempPb.getAttendDate().split("-")[1])
							+"."+Integer.parseInt(tempPb.getAttendDate().split("-")[2]);
					tempPab.kqyc+=formatDate+addString;
				}
			}
			
			for(int j=0;j<pab.size();j++){
				PersonAttendBean tempPab = pab.get(j);
				if(tempPab.getWeishua()!=0){
					tempPab.kqkk+="未刷"+tempPab.getWeishua()+"次，";
				}
				if(tempPab.getChidao()!=0){
					tempPab.kqkk+="迟到"+tempPab.getChidao()+"次，";				
				}
				if(tempPab.getZaotui()!=0){
					tempPab.kqkk+="早退"+tempPab.getZaotui()+"次";
				}
				if(tempPab.mqDays>=5){
					tempPab.setHege("合格");
				}else{
					tempPab.setHege("不合格");
				}
			}

			// 输出结果到excel
			ExportExcel<PersonAttendBean> ex = new ExportExcel<PersonAttendBean>();
			String[] headers = { "姓名", "职务", "考勤扣款类别", "考勤异常情况","未刷","迟到","早退","满勤天数","合格情况"};
			try {
				System.out.println("正在导出缺勤情况统计表...");
				OutputStream out = new FileOutputStream("output.xls");
				ex.exportExcel(headers, pab, out);
				out.close();
				System.out.println("导出成功！");
			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			}
			
		}catch (ParseException e){
			e.printStackTrace();
		}catch (IOException e) {
			e.printStackTrace();
		}
	}
}

//package workAttendance;
//
//import java.io.BufferedReader;
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileNotFoundException;
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.io.InputStreamReader;
//import java.io.OutputStream;
//import java.text.DecimalFormat;
//import java.text.ParseException;
//import java.text.SimpleDateFormat;
//import java.util.ArrayList;
//import java.util.Date;
//import java.util.LinkedList;
//import java.util.List;
//
//import org.apache.poi.hssf.usermodel.HSSFCell;
//import org.apache.poi.hssf.usermodel.HSSFDateUtil;
//import org.apache.poi.hssf.usermodel.HSSFRow;
//import org.apache.poi.hssf.usermodel.HSSFSheet;
//import org.apache.poi.hssf.usermodel.HSSFWorkbook;
//import org.apache.poi.xssf.usermodel.XSSFCell;
//import org.apache.poi.xssf.usermodel.XSSFRow;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//public class ExcelUtils {
//	static long nd = 1000 * 24 * 60 * 60;// 一天的毫秒数    
//    static long nh = 1000 * 60 * 60;// 一小时的毫秒数    
//    static long nm = 1000 * 60;// 一分钟的毫秒数
//	static List<PersonAttendBean> pab = new ArrayList<PersonAttendBean>();
//	
//	/**
//	 * 对外提供读取excel 的方法
//	 * */
//	public static List<List<Object>> readExcel(File file) throws IOException {
//		String fileName = file.getName();
//		String extension = fileName.lastIndexOf(".") == -1 ? "" : fileName
//				.substring(fileName.lastIndexOf(".") + 1);
//		if ("xls".equals(extension)) {
//			return read2003Excel(file);
//		} else if ("xlsx".equals(extension)) {
//			return read2007Excel(file);
//		} else {
//			throw new IOException("不支持的文件类型");
//		}
//	}
//
//	/**
//	 * 读取 office 2003 excel
//	 * 
//	 * @throws IOException
//	 * @throws FileNotFoundException
//	 */
//	private static List<List<Object>> read2003Excel(File file)
//			throws IOException {
//		List<List<Object>> list = new LinkedList<List<Object>>();
//		HSSFWorkbook hwb = new HSSFWorkbook(new FileInputStream(file));
//		HSSFSheet sheet = hwb.getSheetAt(0);
//		Object value = null;
//		HSSFRow row = null;
//		HSSFCell cell = null;
//		int counter = 0;
//		for (int i = sheet.getFirstRowNum(); counter < sheet
//				.getPhysicalNumberOfRows(); i++) {
//			row = sheet.getRow(i);
//			if (row == null) {
//				continue;
//			} else {
//				counter++;
//			}
//			List<Object> linked = new LinkedList<Object>();
//			for (int j = 0; j <= row.getLastCellNum(); j++) {
//				cell = row.getCell(j);
//				if (cell == null) {
//					linked.add("");
//					continue;
//				}
//				DecimalFormat df = new DecimalFormat("0");// 格式化 number String
//															// 字符
//				SimpleDateFormat sdf = new SimpleDateFormat(
//						"yyyy-MM-dd HH:mm:ss");// 格式化日期字符串
//				DecimalFormat nf = new DecimalFormat("0.00");// 格式化数字
//				switch (cell.getCellType()) {
//				case XSSFCell.CELL_TYPE_STRING:
//					//System.out.println(i + "行" + j + " 列 is String type");
//					value = cell.getStringCellValue();
//					break;
//				case XSSFCell.CELL_TYPE_NUMERIC:
//					//System.out.println(i + "行" + j
//					//		+ " 列 is Number type ; DateFormt:"
//					//		+ cell.getCellStyle().getDataFormatString());
//					if ("@".equals(cell.getCellStyle().getDataFormatString())) {
//						value = df.format(cell.getNumericCellValue());
//					} else if ("General".equals(cell.getCellStyle()
//							.getDataFormatString())) {
//						value = nf.format(cell.getNumericCellValue());
//					} else {
//						value = sdf.format(HSSFDateUtil.getJavaDate(cell
//								.getNumericCellValue()));
//					}
//					break;
//				case XSSFCell.CELL_TYPE_BOOLEAN:
//					//System.out.println(i + "行" + j + " 列 is Boolean type");
//					value = cell.getBooleanCellValue();
//					break;
//				case XSSFCell.CELL_TYPE_BLANK:
//					//System.out.println(i + "行" + j + " 列 is Blank type");
//					value = "";
//					break;
//				default:
//					//System.out.println(i + "行" + j + " 列 is default type");
//					value = cell.toString();
//				}
////				if (value == null || " ".equals(value)) {
////					continue;
////				}
//				linked.add(value);
//			}
//			list.add(linked);
//		}
//		return list;
//	}
//
//	/**
//	 * 读取Office 2007 excel
//	 * */
//	private static List<List<Object>> read2007Excel(File file)
//			throws IOException {
//		List<List<Object>> list = new LinkedList<List<Object>>();
//		// 构造 XSSFWorkbook 对象，strPath 传入文件路径
//		XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(file));
//		// 读取第一章表格内容
//		XSSFSheet sheet = xwb.getSheetAt(0);
//		Object value = null;
//		XSSFRow row = null;
//		XSSFCell cell = null;
//		int counter = 0;
//		for (int i = sheet.getFirstRowNum(); counter < sheet
//				.getPhysicalNumberOfRows(); i++) {
//			row = sheet.getRow(i);
//			if (row == null) {
//				continue;
//			} else {
//				counter++;
//			}
//			List<Object> linked = new LinkedList<Object>();
//			for (int j = 0; j <= row.getLastCellNum(); j++) {
//				cell = row.getCell(j);
//				if (cell == null) {
//					linked.add("");
//					continue;
//				}
//				DecimalFormat df = new DecimalFormat("0");// 格式化 number String
//															// 字符
//				SimpleDateFormat sdf = new SimpleDateFormat(
//						"yyyy-MM-dd HH:mm:ss");// 格式化日期字符串
//				DecimalFormat nf = new DecimalFormat("0.00");// 格式化数字
//				switch (cell.getCellType()) {
//				case XSSFCell.CELL_TYPE_STRING:
//					//System.out.println(i + "行" + j + " 列 is String type");
//					value = cell.getStringCellValue();
//					break;
//				case XSSFCell.CELL_TYPE_NUMERIC:
//					//System.out.println(i + "行" + j
//					//		+ " 列 is Number type ; DateFormt:"
//					//		+ cell.getCellStyle().getDataFormatString());
//					if ("@".equals(cell.getCellStyle().getDataFormatString())) {
//						value = df.format(cell.getNumericCellValue());
//					} else if ("General".equals(cell.getCellStyle()
//							.getDataFormatString())) {
//						value = nf.format(cell.getNumericCellValue());
//					} else {
//						value = sdf.format(HSSFDateUtil.getJavaDate(cell
//								.getNumericCellValue()));
//					}
//					break;
//				case XSSFCell.CELL_TYPE_BOOLEAN:
//					//System.out.println(i + "行" + j + " 列 is Boolean type");
//					value = cell.getBooleanCellValue();
//					break;
//				case XSSFCell.CELL_TYPE_BLANK:
//					//System.out.println(i + "行" + j + " 列 is Blank type");
//					value = "";
//					break;
//				default:
//					//System.out.println(i + "行" + j + " 列 is default type");
//					value = cell.toString();
//				}
////				if (value == null || "".equals(value)) {
////					continue;
////				}
//				linked.add(value);
//			}
//			list.add(linked);
//		}
//		return list;
//	}
//
//	public static Boolean passClearWords(String ts){
//		String[] clearWords = PropertyUtil.get("config.properties", "CLEAR_WORDS").split(",");
//		for(int i=0;i<clearWords.length;i++){
//			if(ts.equals(clearWords[i])){
//				return false;
//			}
//		}
//		return true;
//	}
//	
//	public static PersonAttendBean getPABFromList(String name){
//		for(int i=0;i<pab.size();i++){
//			if(pab.get(i).getName().equals(name)){
//				return pab.get(i);
//			}
//		}
//		PersonAttendBean temp = new PersonAttendBean(name);
//		pab.add(temp);
//		return temp;
//	}
//	
//	private static String readDataFromConsole(String prompt) {  
//        BufferedReader br = new BufferedReader(new InputStreamReader(System.in));  
//        String str = null;  
//        try {  
//            System.out.println(prompt);  
//            str = br.readLine();  
//  
//        } catch (IOException e) {  
//            e.printStackTrace();  
//        }  
//        return str;  
//    }
//	
//	public static void main(String[] args) {
//		Date shangbanTime = null;
//		Date xiabanTime = null;
//		try{
//			shangbanTime = new SimpleDateFormat("HH:mm").parse("08:30");
//		    xiabanTime = new SimpleDateFormat("HH:mm").parse("17:30");
//	    }catch(Exception e){
//	    	e.printStackTrace();
//	    }
//		List<PersonBean> pList = new ArrayList<PersonBean>();
//		try {
//			String str = readDataFromConsole("请输入需要录入的考勤记录表，如 D:\\考勤记录（2016.9.23-2016.10.09）.xls：");
//			String search = readDataFromConsole("请输入检索筛选的时间范围，如2016-09-26 2016-09-30（包含边界），为空则检索全部信息：");
//			System.out.println("正在读取考勤信息...");
//			List<List<Object>> list = readExcel(new File(str));
//			System.out.println("正在统计缺勤情况...");
//			for(int i=0;i<list.size();i++){
//				if(list.get(i).size()==0
//						||!passClearWords((String) list.get(i).get(0))
//						||"".equals(list.get(i).get(0))){
//					continue;
//				}
//				PersonBean pb = new PersonBean();
//				pb.setName((String)list.get(i).get(Integer.parseInt(
//						PropertyUtil.get("config.properties", "NAME"))));
//				pb.setAttendDate((String)list.get(i).get(Integer.parseInt(
//						PropertyUtil.get("config.properties", "ATTENDDATE"))));
//				if(!((String)list.get(i).get(Integer.parseInt(
//						PropertyUtil.get("config.properties", "SECONDIN")))).equals("")){
//					pb.setSecondIn(new SimpleDateFormat("HH:mm").parse((String)list.get(i).get(Integer.parseInt(
//							PropertyUtil.get("config.properties", "SECONDIN")))));
//				}else{
//					pb.setSecondIn(null);
//				}
//				if(!((String)list.get(i).get(Integer.parseInt(
//						PropertyUtil.get("config.properties", "SECONDOUT")))).equals("")){
//					pb.setSecondOut(new SimpleDateFormat("HH:mm").parse((String)list.get(i).get(Integer.parseInt(
//							PropertyUtil.get("config.properties", "SECONDOUT")))));
//				}else{
//					pb.setSecondOut(null);
//				}
//				if(!((String)list.get(i).get(Integer.parseInt(
//						PropertyUtil.get("config.properties", "THIRDIN")))).equals("")){
//					pb.setThirdIn(new SimpleDateFormat("HH:mm").parse((String)list.get(i).get(Integer.parseInt(
//							PropertyUtil.get("config.properties", "THIRDIN")))));
//				}else{
//					pb.setThirdIn(null);
//				}
//				if(!((String)list.get(i).get(Integer.parseInt(
//						PropertyUtil.get("config.properties", "THIRDOUT")))).equals("")){
//					pb.setThirdOut(new SimpleDateFormat("HH:mm").parse((String)list.get(i).get(Integer.parseInt(
//							PropertyUtil.get("config.properties", "THIRDOUT")))));
//				}else{
//					pb.setThirdOut(null);
//				}
//				if(!((String)list.get(i).get(Integer.parseInt(
//						PropertyUtil.get("config.properties", "FOURIN")))).equals("")){
//					pb.setFourIn(new SimpleDateFormat("HH:mm").parse((String)list.get(i).get(Integer.parseInt(
//							PropertyUtil.get("config.properties", "FOURIN")))));
//				}else{
//					pb.setFourIn(null);
//				}
//				if(!((String)list.get(i).get(Integer.parseInt(
//						PropertyUtil.get("config.properties", "FOUROUT")))).equals("")){
//					pb.setFourOut(new SimpleDateFormat("HH:mm").parse((String)list.get(i).get(Integer.parseInt(
//							PropertyUtil.get("config.properties", "FOUROUT")))));
//				}else{
//					pb.setFourOut(null);
//				}
//				pList.add(pb);
//			}
//			
//			//设置检索信息;
//			Date dateBegin = null;
//			Date dateEnd = null;
//			long days = 0;
//			if(!search.equals("")){
//				dateBegin = new SimpleDateFormat("yyyy-MM-dd").parse(search.split(" ")[0]);
//				dateEnd = new SimpleDateFormat("yyyy-MM-dd").parse(search.split(" ")[1]);	
//				days = (dateEnd.getTime()-dateBegin.getTime())/(24*60*60*1000);
//			}
//			
//			for(int j=0;j<pList.size();j++){
//				PersonBean tempPb = pList.get(j);
//				//加入检索信息		
//				Date infoDate = new SimpleDateFormat("yyyy-MM-dd").parse(tempPb.getAttendDate());
//				if(dateBegin!=null&&dateEnd!=null&&(
//						infoDate.getTime()<dateBegin.getTime()||infoDate.getTime()>dateEnd.getTime())){
//					continue;
//				}
//				PersonAttendBean tempPab = getPABFromList(pList.get(j).getName());
//				String addString = "";
//				Boolean weishua=false,chidao=false,zaotui=false;
//				if(tempPb.getSecondIn()==null){
//					weishua=true;
//					tempPab.weishua++;
//					addString="上午未刷，";
//				}
//				if(tempPb.getThirdIn()==null
//						&&tempPb.getThirdOut()==null){
//					weishua=true;
//					tempPab.weishua++;
//					addString+="中午未刷，";
//				}
//				if(tempPb.getFourOut()==null){
//					weishua=true;
//					tempPab.weishua++;
//					addString+="下午未刷，";
//				}
//				if(tempPb.getSecondIn()==null
//						&&tempPb.getThirdIn()==null
//						&&tempPb.getThirdOut()==null
//						&&tempPb.getFourOut()==null){
//					weishua=true;
//					addString = "全天未刷，";
//				}
//				if(tempPb.getSecondIn()!=null && tempPb.getSecondIn().getTime()>shangbanTime.getTime()){
//					long chidaoTime = tempPb.getSecondIn().getTime()-shangbanTime.getTime();  
//					long day=chidaoTime/(24*60*60*1000);
//					long hour=(chidaoTime/(60*60*1000)-day*24);
//					long min=((chidaoTime/(60*1000))-day*24*60-hour*60);
//					chidao=true;
//					addString+="上午迟到"+(hour*60+min)+"分钟，";
//				}
//				if(tempPb.getFourOut()!=null && tempPb.getFourOut().getTime()<xiabanTime.getTime()){
//					long chidaoTime = xiabanTime.getTime()-tempPb.getFourOut().getTime();  
//					long day=chidaoTime/(24*60*60*1000);
//					long hour=(chidaoTime/(60*60*1000)-day*24);
//					long min=((chidaoTime/(60*1000))-day*24*60-hour*60);
//					zaotui=true;
//					addString+="下午早退"+(hour*60+min)+"分钟，";
//				}
//				if(chidao)tempPab.chidao++;
//				if(zaotui)tempPab.zaotui++;
//				//没有异常记录记满勤一天,一周清一次
//				if(!weishua&&!chidao&&!zaotui){
//					tempPab.mqDays++;
//				}
//				
//				if(!addString.equals("")){
//					String formatDate = Integer.parseInt(tempPb.getAttendDate().split("-")[1])
//							+"."+Integer.parseInt(tempPb.getAttendDate().split("-")[2]);
//					tempPab.kqyc+=formatDate+addString;
//				}
//			}
//			
//			for(int j=0;j<pab.size();j++){
//				PersonAttendBean tempPab = pab.get(j);
//				if(tempPab.getWeishua()!=0){
//					tempPab.kqkk+="未刷"+tempPab.getWeishua()+"次，";
//				}
//				if(tempPab.getChidao()!=0){
//					tempPab.kqkk+="迟到"+tempPab.getChidao()+"次，";				
//				}
//				if(tempPab.getZaotui()!=0){
//					tempPab.kqkk+="早退"+tempPab.getZaotui()+"次";
//				}
//				if(tempPab.mqDays>=5){
//					tempPab.setHege("合格");
//				}else{
//					tempPab.setHege("不合格");
//				}
//			}
//
//			// 输出结果到excel
//			ExportExcel<PersonAttendBean> ex = new ExportExcel<PersonAttendBean>();
//			String[] headers = { "姓名", "职务", "考勤扣款类别", "考勤异常情况","未刷","迟到","早退","满勤天数","合格情况"};
//			try {
//				System.out.println("正在导出缺勤情况统计表...");
//				OutputStream out = new FileOutputStream("output.xls");
//				ex.exportExcel(headers, pab, out);
//				out.close();
//				System.out.println("导出成功！");
//			} catch (FileNotFoundException e) {
//				e.printStackTrace();
//			} catch (IOException e) {
//				e.printStackTrace();
//			}
//			
//		}catch (ParseException e){
//			e.printStackTrace();
//		}catch (IOException e) {
//			e.printStackTrace();
//		}
//	}
//}
