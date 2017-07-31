package com.ql.upload;

import com.mashape.unirest.http.HttpResponse;
import com.mashape.unirest.http.Unirest;
import org.apache.http.conn.ssl.SSLConnectionSocketFactory;
import org.apache.http.conn.ssl.SSLContexts;
import org.apache.http.conn.ssl.TrustSelfSignedStrategy;
import org.apache.http.impl.client.CloseableHttpClient;
import org.apache.http.impl.client.HttpClients;
import org.apache.poi.ss.usermodel.*;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;



import javax.net.ssl.SSLContext;



/**
 * Created by ranjith on 017 Jul 17 2017.
 */
public class PromptUpload {

	public static void main(String[] args) {
		try {
			System.setProperty("tenant.admin.userName", "ssoadmin@wells.edu");
			System.setProperty("tenant.admin.password", "Admin@123");
			System.setProperty("file.path","PromptData.xlsx");
			System.setProperty("ql.url", "https://qlsso-staging.quicklaunchsso.com/admin/rest/savePromptValue");
			System.setProperty("outputFile.path","C:/Users/harry_prasad/Desktop/NewPrompt/promptresponse.xlsx" );
			System.out.println("Starting data processing.");
			System.out.println("Usage java -Dtenant.admin.userName=[userName] -Dtenant.admin.password=[password] -Dfile.path=[fullPathToExcel] -Dql.url=[urlToPostData] -jar [jarfileName]");
			new PromptUpload().processData();

		} catch (Exception e) {
			System.out.println("Could not upload due to "+e.getMessage());
		}
	}

	private void processData() throws Exception {
		String userName =  System.getProperty("tenant.admin.userName");
		String password = System.getProperty("tenant.admin.password");
		if (userName == null || userName.isEmpty() || password == null || password.isEmpty()) {
			throw new Exception("Tenant Admin username and password needs to be provided as java -Dtenant.admin.userName=[userName] -Dtenant.admin.password=[password] -Dfile.path=[fullPathToExcelFile] -jar upload.jar");
		}
		List<PromptData> promptDataList = readDataFromExcel();
		postDataToQL(promptDataList);
		System.out.println("Prompt Data is "+promptDataList);
	}

	/**
	 * Posts Data to ql.
	 *
	 * @param promptDataList
	 */
	private void postDataToQL(List<PromptData> promptDataList) throws Exception {
		String userName =  System.getProperty("tenant.admin.userName");
		String password = System.getProperty("tenant.admin.password");
		String url = System.getProperty("ql.url");
		String outputFilePath = System.getProperty("outputFile.path");
		if (userName == null || userName.isEmpty() || password == null || password.isEmpty() || url == null || url.isEmpty()) {
			throw new Exception("Tenant Admin usernamd and password needs to be provided as java -Dtenant.admin.userName=[userName] -Dtenant.admin.password=[password] -Dfile.path=[fullPathToExcelFile] -jar upload.jar");
		}
		if (promptDataList != null && !promptDataList.isEmpty()) {
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet("Prompt Data Response");
			int rowCount = 0;	     
			for (PromptData promptData : promptDataList) {	
				String body="{"+ "\"applicationName\":"+"\""+promptData.getApplicationName()+"\""+","+"\"userId\":"+"\""+promptData.getUserId()+"\""+","+"\"appUserName\":"+"\""+promptData.getUserName()+"\""+","+"\"appPassword\":"+"\""+promptData.getPassword()+"\""+"}";          	    
				SSLContext sslcontext = SSLContexts.custom()
						.loadTrustMaterial(null, new TrustSelfSignedStrategy())
						.build();
				SSLConnectionSocketFactory sslsf = new SSLConnectionSocketFactory(sslcontext);
				CloseableHttpClient httpclient = HttpClients.custom()
						.setSSLSocketFactory(sslsf)
						.build();
				Unirest.setHttpClient(httpclient);
				HttpResponse<String> bodyResponse=	Unirest.post(url)
						.basicAuth(userName, password)
						.header("accept", "application/json")
						.header("Content-Type", "application/json")
						.body(body)
						.asString();
				rowCount = createExcelSheetForResponse(sheet, rowCount, promptData, bodyResponse); 
			}
			try (FileOutputStream outputStream = new FileOutputStream(outputFilePath)) {
				workbook.write(outputStream);
			}
		}
	}

	private int createExcelSheetForResponse(XSSFSheet sheet, int rowCount, PromptData promptData,
			HttpResponse<String> bodyResponse) {
		Row row = sheet.createRow(rowCount++);
		int columnCount = 0;
		 Cell cell = row.createCell(columnCount++);
		    if (promptData.getApplicationName() instanceof String) {
		        cell.setCellValue((String) promptData.getApplicationName());
		    } 
		    cell = row.createCell(columnCount++);
		    if (promptData.getUserId() instanceof String) {
		        cell.setCellValue((String) promptData.getUserId());
		    } 
		    cell = row.createCell(columnCount++);
		    if (promptData.getUserName() instanceof String) {
		        cell.setCellValue((String) promptData.getUserName());
		    } 
		    cell = row.createCell(columnCount++);
		    if (promptData.getPassword() instanceof String) {
		        cell.setCellValue((String) promptData.getPassword());
		    } 
		 cell = row.createCell(columnCount++);
		 if (bodyResponse.getBody() instanceof String) {
		        cell.setCellValue((String) bodyResponse.getBody());
		    }
		return rowCount;
	}

	/**
	 * Reads data from excel file.
	 * @throws Exception
	 */
	private List<PromptData> readDataFromExcel() throws Exception {
		String excelFilePath = "PromptData.xlsx";
		List<PromptData> promptDataList = new ArrayList<PromptData>();
		InputStream inputStream = getClass().getClassLoader().getResourceAsStream("PromptData.xlsx");
		String filePath = System.getProperty("file.path");
		if (filePath != null) {
			inputStream = new FileInputStream(filePath);
		} else {
			if (inputStream == null)
				inputStream = new FileInputStream("./PromptData.xlsx");
		}
		Workbook workbook = new XSSFWorkbook(inputStream);
		Sheet firstSheet = workbook.getSheetAt(0);
		Iterator<Row> iterator = firstSheet.iterator();
		while (iterator.hasNext()) {
			Row nextRow = iterator.next();
			Iterator<Cell> cellIterator = nextRow.cellIterator();
			int count=0;
			PromptData pd = new PromptData();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				cell.setCellType(CellType.STRING);
				String value = cell.getStringCellValue();
				if (cell.getColumnIndex() == 0) {
					pd.setApplicationName(value);
				} else if (cell.getColumnIndex() == 1) {
					pd.setUserId(value);
				} else if (cell.getColumnIndex() == 2) {
					pd.setUserName(value);
				} else if (cell.getColumnIndex() == 3){
					pd.setPassword(value);
				}
			}	
			if (pd.getApplicationName() != null && !pd.getApplicationName().isEmpty() && pd.getUserId() != null && !pd.getUserId().isEmpty())
				promptDataList.add(pd);
		}
		workbook.close();
		inputStream.close();
		return  promptDataList;
	}
}
