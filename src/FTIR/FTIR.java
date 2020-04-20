package FTIR;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

import javax.swing.JFileChooser;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.xy.XYSeries;
import org.jfree.data.xy.XYSeriesCollection;

public class FTIR {

	public static void main(String[] args){
		long startTime = System.nanoTime();
		String[] add=path();
		String pre_address=add[0];
		String post_address=add[1];
		String res_address=add[2];
		//cvs to Excel, sheet0 add pre, post, absorbance
		csvToExcel(pre_address, post_address, res_address);
		//sheet 1 peak area
		peakAreaSheet(res_address);
		ftirPlot(res_address);
		long endTime   = System.nanoTime();
		long totalTime = (endTime - startTime)/1000000;
		System.out.println(totalTime+"ms");
	}
	//get file path
	public static String[] path() {
		String[] res=new String[3];
		File pre_file = null;
		File post_file = null;
		String pre_filePath=null;
		String post_filePath=null;
		String res_filePath=null;
		 JFileChooser chooser1 = new JFileChooser();
		 JFileChooser chooser2 = new JFileChooser();
		 chooser1.setDialogTitle("Select prefile");
		 chooser2.setDialogTitle("Select postfile");
         int returnValue1 = chooser1.showOpenDialog( null ) ;
         int returnValue2 = chooser2.showOpenDialog( null ) ;
		 if( returnValue1 == JFileChooser.APPROVE_OPTION ) {
		        pre_file = chooser1.getSelectedFile() ;
		 }
		 if(pre_file != null)
		 {
		      pre_filePath = pre_file.getPath();
		 }
		 try {
			 res_filePath=new File(".").getCanonicalPath()+"/delta.xls";
		 if( returnValue2 == JFileChooser.APPROVE_OPTION ) {
		        post_file = chooser2.getSelectedFile() ;
		 }
		 if(post_file != null)
		 {
		      post_filePath = post_file.getPath();
		 }
		 res[0]=pre_filePath;
		 res[1]=post_filePath;
		 res[2]=res_filePath;
		 } catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		 return res;
	}
	//convert csv to xls
	public static void csvToExcel(String pre_address, String post_address, String xlsFileAddress){
		String pre_csvFileAddress=pre_address;
		String post_csvFileAddress=post_address;	
		HSSFWorkbook workBook = new HSSFWorkbook();
		HSSFSheet sheet = workBook.createSheet("raw_data");
		String[] title= {"wavenumber", "pre","post","absorbance"};
		String pre_currentLine=null;
		String post_currentLine=null;
		int pre_RowNum=1;
		int post_RowNum=1;
		int col_pt=0;;
		BufferedReader pre_br;	
		BufferedReader post_br;	
			try {
				//add title
				HSSFRow firstRow=sheet.createRow(0);
				 for(int i=0;i<title.length;i++){
				        firstRow.createCell(i).setCellValue(title[i]);
				    }
				//add pre data
			pre_br = new BufferedReader(new FileReader(pre_csvFileAddress));
		while ((pre_currentLine = pre_br.readLine()) != null) {
		    String pre_str[] = pre_currentLine.split(",");
		    if(pre_str[0].charAt(0)>='A'&&pre_str[0].charAt(0)<='z') continue;
		    HSSFRow currentRow=sheet.createRow(pre_RowNum);
		    for(int i=0;i<pre_str.length;i++){
		        currentRow.createCell(i).setCellValue(pre_str[i]);
		    }
		    col_pt=pre_str.length;
		    pre_RowNum++;
		}
		//add post data
		  post_br = new BufferedReader(new FileReader(post_csvFileAddress));
		  while ((post_currentLine = post_br.readLine()) != null) {
			    String post_str[] = post_currentLine.split(",");
			    if(post_str[0].charAt(0)>='A'&&post_str[0].charAt(0)<='z') continue;
			    HSSFRow currentRow=sheet.getRow(post_RowNum);
			    currentRow.createCell(col_pt).setCellValue(post_str[1]);
			    post_RowNum++;
			}
		  //add absorbance data
		  int lastRow=sheet.getLastRowNum();
		  int col_pt_v1=++col_pt;
		  for(int i=1;i<=lastRow;i++) {
			  HSSFRow currentRow=sheet.getRow(i);
			  float pre_value=Float.valueOf(currentRow.getCell(1).getStringCellValue());
			  float post_value=Float.valueOf(currentRow.getCell(2).getStringCellValue());
			  float abs_value=(float)Math.log10(pre_value/post_value);
			  currentRow.createCell(col_pt_v1).setCellValue(abs_value);
		  }
		File xlsFile = new File(xlsFileAddress);//新建文件路径
		FileOutputStream fileOutputStream =  new FileOutputStream(xlsFile);
		workBook.write(fileOutputStream);
		fileOutputStream.close();
		System.out.println("Done");
	} catch (Exception e) {
		// TODO Auto-generated catch block
		System.out.println(e.getMessage()+"Exception in try");
	}
}
	public static void peakAreaSheet(String add) {
		try {
		Workbook workbook = WorkbookFactory.create(new File(add));
		Sheet sheet = workbook.createSheet("peak_area");
		String peak_labels[]= {"Si-N", "Si-H", "N-H1", "N-H2"};
		double[] peak_area=peakArea(add);
		//add title
		Row row = sheet.createRow(0);
		for (int i = 0; i < peak_labels.length; i++) {
			Cell cell = row.createCell(i);
			cell.setCellValue(peak_labels[i]);
		}
		//add data
		row = sheet.createRow(1);
		for (int i = 0; i < peak_area.length; i++) {
			Cell cell = row.createCell(i);
			cell.setCellValue(peak_area[i]);
		}
		FileOutputStream fileOutputStream = new FileOutputStream(add);
		workbook.write(fileOutputStream);
		fileOutputStream.close();
			System.out.println("Done");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			System.out.println(e.getMessage()+"Exception in try");
		}	
	}
	// plot, JFreeChart (XY Plot)
	public static void ftirPlot(String address){
		File file = new File(address);
		 XYSeries plot = new XYSeries("FTIR Plot"); //plot title
		 try {
				//创建excel
				HSSFWorkbook workbook = 
					new HSSFWorkbook(FileUtils.openInputStream(file));
				//获得第一个sheet
				HSSFSheet sheet = workbook.getSheetAt(0);
				int firstRowNum = 1;
				//获得sheet中的最后一行的行数
				int lastRowNum = sheet.getLastRowNum();
				for (int i = firstRowNum; i <=lastRowNum; i++) {
					HSSFRow row = sheet.getRow(i);
					//获得一行中随后一个cell的cell数
						HSSFCell x_cell = row.getCell(0);
						HSSFCell abs_cell = row.getCell(3);
						float f1 = Float.valueOf(x_cell.getStringCellValue());
						float f2 = (float)abs_cell.getNumericCellValue();
						plot.add(f1,f2);
				}
				 XYSeriesCollection dataset = new XYSeriesCollection( );
			      dataset.addSeries(plot);
				 JFreeChart xylineChart = ChartFactory.createXYLineChart(
						 "FTIR Plot","wavenumber(cm^{-1})",
				         "absorbance", // title, x-title, y-title 
				        dataset,
				         PlotOrientation.VERTICAL, 
				         true, true, false);
				  int width = 640;   /* Width of the image */
			      int height = 480;  /* Height of the image */ 
			      File XYChart = new File( "/Users/yi/Desktop/Projects/Java Amat Project/FTIR/ftir/Raw Data/plot.jpeg" ); 
			      ChartUtilities.saveChartAsJPEG( XYChart, xylineChart, width, height);
			} catch (IOException e) {
				e.printStackTrace();
			}
}
	//find peak area
	public static double[] peakArea(String address){
		double[][] peak_positions = {{649.473, 1163.64}, {1124.93, 1297.3},{2092.09, 2286.27}, {3058.72, 3524.22}};
		double x_data[][]=new double[4][2]; //peak position x_value
		double y_data[][]=new double[4][2]; //peak position y_value
		int[][] row=new int[4][2];
		double[] result=new double[4];
		//get peak positions in sheet x_axis
		File file = new File(address);
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(FileUtils.openInputStream(file));
			HSSFSheet sheet1 = workbook.getSheetAt(0);
			int lastRowNum = sheet1.getLastRowNum();
			for(int j=0;j<4;j++) {
				for(int k=0;k<2;k++) {
					double[]arr=xyAxis(address,peak_positions[j][k], 1, lastRowNum);
					x_data[j][k]=arr[0];
					row[j][k]=(int)arr[2];
					y_data[j][k]=arr[1];
				}
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		//Trapezoidal Rule to get peak area
		for(int i=0;i<4;i++) {
			result[i]=trapezoidal(x_data[i][0],x_data[i][1],y_data[i][0],y_data[i][1], 1.0, address, row[i][0], row[i][1]);			
		}	
		return result;
	}
	//get x_axis by closest x value
	public static double[] xyAxis(String address, double x, int startRow, int endRow) {
		File file = new File(address);
		double x_value=Integer.MAX_VALUE;
		double result[]=new double[3];//result[0] x_value, result[1] y_value, result[2] rowNum
		int row=0;
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(FileUtils.openInputStream(file));
			HSSFSheet sheet1 = workbook.getSheetAt(0);
			double min=Integer.MAX_VALUE;		
			    for (int i = startRow; i <=endRow; i++) {
				   HSSFCell x_cell = sheet1.getRow(i).getCell(0);
				   double x_axis=Float.valueOf(x_cell.getStringCellValue());
			    if(Math.abs(x-x_axis)<min) {
			    	min=Math.abs(x-x_axis);
				    x_value=x_axis;
				    row=i;
			}
		}
			 result[0]=x_value;
		     result[1]=sheet1.getRow(row).getCell(3).getNumericCellValue();
			 result[2]=row;
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		 return result;
		   
}   
	//get y_axis by closest x value
	public static double yAxis(String address, double x,int startRow, int endRow) {
		File file = new File(address);
		double y_value=Integer.MAX_VALUE;
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(FileUtils.openInputStream(file));
			HSSFSheet sheet1 = workbook.getSheetAt(0);
			double min=Integer.MAX_VALUE;	
			for (int i = startRow; i <=endRow; i++) {
				if(x==(Float.valueOf(sheet1.getRow(i).getCell(0).getStringCellValue()))){
					y_value=sheet1.getRow(i).getCell(3).getNumericCellValue();
				}	
			}	
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		 return y_value;		   
} 
	
	//Trapezoidal Rule
	public static double trapezoidal(double x0, double x1, double y0, double y1, double interval, String address, int startRow, int endRow) 
    {   
		int n=(int)((x1-x0)/interval);
        // Computing sum of first and last terms 
        // in above formula 
        double s = (y0+y1)/2.0; 
        // Adding middle terms in above formula 
        for (int i = 1; i <n; i++) {
        	double[] arr=xyAxis(address,x0 + i*interval,startRow, endRow);
        	s += arr[1]; 
        }              

        return Math.abs(interval* s); 
    } 
}
