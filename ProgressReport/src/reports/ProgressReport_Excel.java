package reports;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Map;
import java.util.Properties;
import java.util.Scanner;
import java.util.Set;
import java.util.StringTokenizer;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Multipart;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ProgressReport_Excel {
	static Scanner scan=new Scanner(System.in);
	public static void main(String[] args)throws Exception {
		ProgressReport_Excel progress=new ProgressReport_Excel();
		
		String[] details=new String[4];
		details[0]="Progress Report";//Header of the excel
		
		System.out.print("Enter Institue name: ");
		details[3]="Institute: "+scan.nextLine();

		System.out.print("Enter Course: ");
		details[2]="Course: "+scan.nextLine();
		
		String [] subjects=progress.readSubjects();
		
		Excel report=new Excel(subjects,details);
		
		ReadFile txtFile=new ReadFile("src/resource/NameList");

		Properties students=txtFile.readStudents();
		
		progress.enterMark(students,report);
	}
	
	private String[] readSubjects() {
		
		System.out.print("Enter number of subjects: ");
		int num=Integer.parseInt(scan.nextLine());
		String[] subjects=new String[num];
		
		for(int i=0;i<num;i++) {
			System.out.print("Enter the subject name: ");
			subjects[i]=scan.nextLine();
		}
		return subjects;
	}
	
	private void enterMark(Properties students, Excel report) {
		Set set=students.entrySet();
		System.out.println();
		Iterator iter=set.iterator();
		
		while(iter.hasNext()) {
			Map.Entry<String, String> map=(Map.Entry<String, String>)iter.next();
			report.reportCard(map.getValue(),map.getKey());
		}
		
		System.out.println("Copy of the report is saved in your system");
	}

}


class Excel{

	Scanner scan=new Scanner(System.in);
	private String[] subjects,details;
	public int len;
	Workbook workbook=new XSSFWorkbook();
	
	public Excel(String[] subjects,String[] details) {
		this.subjects=subjects;
		this.len=subjects.length;
		this.details=details;
	}
	
	public void reportCard(String name,String emailID) {
		
		details[1]="Name: "+name;

		System.out.println("Student Name: "+name);
		System.out.println("---------------------------------------------------------");
		System.out.println("Enter marks for the subjects respectively\n");
		
		for(String subject:subjects) {
			System.out.print(subject+"\t");
		}
		System.out.println();
		
		int[] marks=new int[len];
		
		StringTokenizer token=new StringTokenizer(scan.nextLine(),"\t");
		while(token.hasMoreTokens())
			for(int i=0;i<len;i++)
				marks[i]=Integer.parseInt(token.nextElement().toString());
		
		excel(name,marks,emailID);
	}
	
	public void excel(String name,int[] marks,String emailID) {
		Sheet sheet=workbook.createSheet(name);
		
		CellStyle fontStyle=workbook.createCellStyle();
		
		XSSFFont font=((XSSFWorkbook)workbook).createFont();
		font.setFontName("Times");
		font.setFontHeightInPoints((short)12);
		fontStyle.setFont(font);
		fontStyle.setAlignment(HorizontalAlignment.LEFT);
		
		Row row;
		Cell cell;
		for(int i=0;i<4;i++) {
			row=sheet.createRow(i);
			cell=row.createCell(0);
			cell.setCellStyle(fontStyle);
			cell.setCellValue(details[i]);
			sheet.addMergedRegion(new CellRangeAddress(i, i, 0, len));
		}
		
		row=sheet.createRow(4);
		cell=row.createCell(0);
		cell.setCellValue("Subject");
		cell.setCellStyle(fontStyle);
		for(int i=0;i<len;i++) {
			cell=row.createCell(i+1);
			cell.setCellValue(subjects[i]);
			cell.setCellStyle(fontStyle);
		}
		
		row=sheet.createRow(5);
		cell=row.createCell(0);
		cell.setCellValue("Marks");
		cell.setCellStyle(fontStyle);
		boolean flag=true;
		int total=0;
		
		for(int i=0;i<len;i++) {
			cell=row.createCell(i+1);
			cell.setCellValue(marks[i]);
			cell.setCellStyle(fontStyle);
			flag&=(marks[i]>=35)?true:false;
			total+=marks[i];
		}
		
		row=sheet.createRow(6);
		int mid=len/2;
		
		cell=row.createCell(0);
		cell.setCellValue("Total:     "+total);
		cell.setCellStyle(fontStyle);
		cell=row.createCell(mid+1);
		cell.setCellValue("Result:    "+(flag?"Pass":"Fail"));
		cell.setCellStyle(fontStyle);
		sheet.addMergedRegion(new CellRangeAddress(6, 6, 0, mid));
		sheet.addMergedRegion(new CellRangeAddress(6,6,mid+1,len));
		
		row=sheet.getRow(0);
		cell=row.getCell(0);
		CellStyle headerCell=workbook.createCellStyle();
		headerCell.setAlignment(HorizontalAlignment.CENTER);
		XSSFFont headerfont=((XSSFWorkbook)workbook).createFont();
		headerfont.setFontName("Times");
		headerfont.setFontHeightInPoints((short)14);
		headerfont.setBold(true);
		headerCell.setFont(headerfont);
		cell.setCellStyle(headerCell);
		
		mail(name,emailID,workbook);
	}
	
	private void mail(String name, String emailID, Workbook workbook) {
		
		try {
			File file=File.createTempFile(name, "xlsx");
		FileOutputStream outfile=new FileOutputStream(file);
		workbook.write(outfile);

		BodyPart textPart=new MimeBodyPart();
		MimeBodyPart bodyPart=new MimeBodyPart();
		
		//Content of the mail composed with setText method
		textPart.setText("Dear "+name+"\nYour Report Card is released kindly check it."
				+ "mail project.\nWith Regards,\n"+System.getenv("USER_NAME"));
		DataSource source=new FileDataSource(file);
		bodyPart.setDataHandler(new DataHandler(source));
		bodyPart.setFileName(name+".xlsx");
		file.deleteOnExit();
		
		Multipart multipart=new MimeMultipart();
		multipart.addBodyPart(textPart);
		multipart.addBodyPart(bodyPart);
		
		FileOutputStream sample=new FileOutputStream("src/resource/Progress.xlsx");
		workbook.write(sample);
		SendMail send=new SendMail(name, emailID, multipart);
		
			send.mail();
		}catch(Exception ex){
			ex.printStackTrace();
			}
	}
}

class ReadFile{
	
	private File file=null;
	
	public ReadFile(String path) {
		this.file=new File(path);
	}
	
	public Properties readStudents() throws Exception{

		Properties students=new Properties();
		try {
			String student;
			Scanner scan=new Scanner(file);
			while(scan.hasNextLine()) {
				student=scan.nextLine();
				int index=student.lastIndexOf("-");
				students.put(student.substring(0,index), student.substring(index+1,student.length()));
			}
			scan.close();
		}
		catch(IOException ex) {
			ex.printStackTrace();
		}
		return students;
	}
}