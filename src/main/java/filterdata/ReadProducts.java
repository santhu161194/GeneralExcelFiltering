package filterdata;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.solr.client.solrj.SolrClient;
import org.apache.solr.client.solrj.SolrQuery;
import org.apache.solr.client.solrj.SolrServerException;
import org.apache.solr.client.solrj.impl.HttpSolrClient;
import org.apache.solr.client.solrj.response.QueryResponse;

public class ReadProducts {
	SolrClient solrClient=new HttpSolrClient.Builder("http://192.168.1.192:8983/solr/search_products").build();
	SolrQuery query=new SolrQuery();
	QueryResponse response=null;
	Map<String,String> companies=null;
	Map<String,Map<String,String>> searchTagCompaniesMap=new HashMap<>();
	File file=new File("/home/testingteam/Desktop/apollo.xlsx");
	Sheet sheet=null;
	void readExcelFile(File file) throws EncryptedDocumentException, InvalidFormatException, IOException {
		Workbook workbook=WorkbookFactory.create(file);
		System.out.println("The workbook has "+workbook.getNumberOfSheets());
		workbook.forEach(sheet->{
			System.out.println("The name of the current sheet is"+sheet.getSheetName());
			if(sheet.getSheetName().equals("apollo"))
			sheet.rowIterator().forEachRemaining(row->{
				row.cellIterator().forEachRemaining(cell->{
					if(cell.getColumnIndex()==0) {
						companies=new HashMap<>();
						String cellval;
						try {
						cellval=cell.getStringCellValue();
						}catch(Exception e) {
							cellval=String.valueOf(cell.getNumericCellValue());
						}
						query.setQuery(cellval);
						query.setFields("name","company");
						try {
							response=solrClient.query(query);
						} catch (SolrServerException | IOException e) {
							e.printStackTrace();
						}
						response.getResults().forEach(respo->companies.put(String.valueOf(respo.get("name")),String.valueOf(respo.get("company"))));
						searchTagCompaniesMap.put(cellval, companies);
					}
				});
				writeToExcel(sheet.getSheetName(),searchTagCompaniesMap);
			});
		//System.out.println(searchTagCompaniesMap);
		System.exit(0);
		});
	}
	
	private void writeToExcel(String sheetName, Map<String, Map<String, String>> searchTagCompaniesMap2) {
		try (FileOutputStream fos = new FileOutputStream(file);Workbook workbook=new XSSFWorkbook();){
		sheet = workbook.getSheet(sheetName);
		if(sheet==null)
				workbook.createSheet(sheetName);
		sheet = workbook.getSheet(sheetName);
		sheet.setColumnWidth(1, 11000);
		sheet.setColumnWidth(2, 11000);
		sheet.setColumnWidth(3, 11000);
		Row header = sheet.createRow(0);
		header.createCell(0).setCellValue("ProductName");
		header.createCell(1).setCellValue("Medplus");
		header.createCell(2).setCellValue("1mg");
		header.createCell(3).setCellValue("netMeds");
		searchTagCompaniesMap2.forEach((product,companyMap)->{
			Row row = sheet.createRow(sheet.getLastRowNum()+1);
			row.createCell(0).setCellValue(product);
			List<String> medplusList = companyMap.entrySet().stream().filter(entry->entry.getValue().equals("Medplus")).map(Map.Entry::getKey).collect(Collectors.toList());
			List<String> onemgList = companyMap.entrySet().stream().filter(entry->entry.getValue().equals("1mg")).map(Map.Entry::getKey).collect(Collectors.toList());
			List<String> netMedsList = companyMap.entrySet().stream().filter(entry->entry.getValue().equals("netmeds")).map(Map.Entry::getKey).collect(Collectors.toList());
			CellStyle style = workbook.createCellStyle();
			style.setWrapText(true);
			row.setRowStyle(style);
			row.setHeight((short)1000);
			Cell cell1 = row.createCell(1);
			cell1.setCellValue(StringUtils.join(medplusList,","));
			cell1.setCellStyle(style);
			Cell cell2 = row.createCell(2);
			cell2.setCellValue(StringUtils.join(onemgList,","));
			cell2.setCellStyle(style);
			Cell cell3 = row.createCell(3);
			cell3.setCellValue(StringUtils.join(netMedsList,","));
			cell3.setCellStyle(style);
		});
			workbook.write(fos);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		File file=new File("/home/testingteam/Desktop/demo.xlsx");
		WriteProductDetails readProducts=new WriteProductDetails();
		readProducts.readExcelFile(file,"medlife");
	}
}
