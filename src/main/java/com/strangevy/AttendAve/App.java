package com.strangevy.AttendAve;

import java.io.FileOutputStream;
import java.net.URL;
import java.net.URLDecoder;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.strangevy.AttendAve.util.ExcelReadUtil;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws Exception
    {
    	String filePath = "";
		//获取当前运行类路径
		URL url = App.class.getProtectionDomain().getCodeSource().getLocation();
        filePath = URLDecoder.decode(url.getPath(), "utf-8");// 转化为utf-8编码  
        if (filePath.endsWith(".jar")) {// 可执行jar包运行的结果里包含".jar"  
            // 截取路径中的jar包名  
            filePath = filePath.substring(0, filePath.lastIndexOf("/") + 1);  
        }  
        
		Scanner scanner=new Scanner(System.in);
		System.out.println("请输入文件名（不包含xlsx后缀）");
		String name = scanner.nextLine();
		scanner.close();
        long startDate = System.currentTimeMillis();
        System.out.println("处理中...");
        List<String[]> rows = ExcelReadUtil.excelToArrayList(filePath+name+".xlsx", null);
        
        Workbook wb = new SXSSFWorkbook(1000);
	    FileOutputStream fileOut = new FileOutputStream(filePath+"result.xlsx");
	    
		
	    //总列数
        int cols = rows.get(2).length;
		Sheet sheet = wb.createSheet();
		SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");
		int i = 0;
		long aveUp = 0;
		long aveDown = 0;
		int countDay = 0;
        for(String[] row:rows){
        	//创建Excel工作表的行     
		    Row r = sheet.createRow(i);
        	if(i<3){
        		for(int j=0;j<5;j++){
    		    	if(i==2&&j==3){
    		    		r.createCell(j).setCellValue("上班平均时间");
    		    	}else if(i==2&&j==4){
    		    		r.createCell(j).setCellValue("下班平均时间");
    		    	}else{
    		    		r.createCell(j).setCellValue(row[j]);
    		    	}
    		    }
        	}else{
        		long tempUp = 0;
        		long tempDown = 0;
        		int days = 0;
        		for(int j=0;j<cols;j++){
        			if(j<3){
        				r.createCell(j).setCellValue(row[j]);
        			}else{
        				String[] dateArr = row[j].split("\n");
            			if(dateArr.length>=2){
            				days++;
            				tempUp += sdf.parse(dateArr[0].trim().substring(0, 5)).getTime();
            				int downDate = Integer.parseInt(dateArr[dateArr.length-1].trim().substring(0, 2));
            				if(0<=downDate&&8>downDate){
                				tempDown += sdf.parse("23:59").getTime();
            				}else{
                				tempDown += sdf.parse(dateArr[dateArr.length-1].trim().substring(0, 5)).getTime();
            				}
            			}else{
            				continue;
            			}
        			}
    		    }
        		if(days!=0){
        			countDay += days;
        			aveUp += tempUp;
        			aveDown += tempDown;
        			r.createCell(3).setCellValue(sdf.format(new Date(tempUp/days)));
            		r.createCell(4).setCellValue(sdf.format(new Date(tempDown/days)));
        		}
        	}
        	i++;
        }
        sheet.getRow(0).getCell(0).setCellValue("整月平均上班时间："+sdf.format(new Date(aveUp/countDay)));
        sheet.getRow(1).getCell(0).setCellValue("整月平均下班时间："+sdf.format(new Date(aveDown/countDay)));
	    wb.write(fileOut);
	    fileOut.close();
	    long endDate = System.currentTimeMillis();
	    System.out.println("转换完成："+i+"行\n耗时："+((endDate-startDate)/1000)+"秒");
    }
}
