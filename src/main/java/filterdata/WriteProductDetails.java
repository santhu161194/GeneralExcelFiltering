package filterdata;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.LinkedHashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
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

public class WriteProductDetails {
	SolrClient solrClient = new HttpSolrClient.Builder("http://192.168.1.192:8983/solr/search_products").build();
	SolrQuery query = new SolrQuery();
	QueryResponse response = null;
	File file = null;
	Sheet sheet = null;
	List<String> companies = new ArrayList<>(Arrays.asList("Medplus", "1mg", "netmeds"));
	List<SolrSearchResponse> productsList = new ArrayList<>();
	int rowNumber=1;

	void readExcelFile(File file, String sheetName)
			throws EncryptedDocumentException, InvalidFormatException, IOException {
		Workbook workbook = WorkbookFactory.create(file);

		System.out.println("The workbook has " + workbook.getNumberOfSheets());

		final Set<SolrSearchResponse> keywords = new LinkedHashSet<>();
		Sheet sheet = workbook.getSheet(sheetName);

		System.out.println("The name of the current sheet is " + sheet.getSheetName());


		sheet.rowIterator().forEachRemaining(row -> {
			if (row.getRowNum() != 0) {
				String cellval = null;
				Cell cell;	
				SolrSearchResponse solrSearchResponse = new SolrSearchResponse();
				cell=row.getCell(3);
				try {	
					solrSearchResponse.setSearchHits(cell.getNumericCellValue());
				}catch(Exception e) {
					solrSearchResponse.setSearchHits(0d);
				}
				cell=row.getCell(0);
				try {
					cellval = cell.getStringCellValue();
				} catch (Exception e) {
					try {
						cellval = String.valueOf(cell.getNumericCellValue());
					}catch(Exception e1) {
						e1.printStackTrace();
					}
				}
				solrSearchResponse.setSearchTag(cellval);
				populateSolrData(solrSearchResponse);
				keywords.add(solrSearchResponse);
			}
		});
		workbook.close();
		System.out.println("total keyword for " + sheetName + " is " + keywords.size());
		writeToExcel(sheetName, keywords);
		//int tempSize=keywords.size()/1000;
		/*for(int i=0;i<tempSize;i++) {
			List<SolrSearchResponse> tempKeyWords = new LinkedList<SolrSearchResponse>();
			Iterator<SolrSearchResponse> itr2 = keywords.iterator();
			int i1=0;
			while(itr2.hasNext()&&i1<1000) {
				SolrSearchResponse solrres = itr2.next();
				populateSolrData(solrres);
				tempKeyWords.add(solrres);
				i1++;
			}
			writeToExcel(sheetName, tempKeyWords);
			keywords.removeAll(tempKeyWords);
			tempKeyWords = null;
			System.out.println("The keywords size is"+keywords.size());
			System.out.println((i*1000) + " are done ");
		}
		*/
		
		
		
	}

	void populateSolrData(SolrSearchResponse solrSearchResponse) {
		companies.forEach(company->{
			query.setQuery("name:" + solrSearchResponse.getSearchTag() + " AND company:"+company);
			query.setFields("name", "company");
			query.setRows(10);
			try {
				response = solrClient.query(query);
			} catch (SolrServerException | IOException e) {
				e.printStackTrace();
			}
			if(company.equals("Medplus"))
				solrSearchResponse.setMedplusProducts(response.getResults().stream().map(doc->doc.get("name").toString()).collect(Collectors.toList()));
			else if(company.equals("1mg"))
				solrSearchResponse.setOneMgProducts(response.getResults().stream().map(doc->doc.get("name").toString()).collect(Collectors.toList()));
			else if(company.equals("netmeds"))
				solrSearchResponse.setNetMedsProducts(response.getResults().stream().map(doc->doc.get("name").toString()).collect(Collectors.toList()));
		});
	}


	private void writeToExcel(String sheetName, Set<SolrSearchResponse> productsList2) throws IOException {
		file = new File("/home/testingteam/Desktop/KeywordMapping/" + sheetName + ".xlsx");
		Workbook workbook=null;
		if(!file.exists()) {
			//Workbook workbook = new XSSFWorkbook();
			workbook = new XSSFWorkbook();
			sheet = workbook.createSheet(sheetName);
			
			sheet.setColumnWidth(2, 11000);
			sheet.setColumnWidth(3, 11000);
			sheet.setColumnWidth(4, 11000);
			
			Row header = sheet.createRow(0);
			header.createCell(0).setCellValue("ProductName");
			header.createCell(1).setCellValue("Search Count");
			header.createCell(2).setCellValue("Medplus");
			header.createCell(3).setCellValue("1mg");
			header.createCell(4).setCellValue("netMeds");
			//FileOutputStream fos = new FileOutputStream(file);
			//workbook.write(fos);
			//workbook.close();
			//workbook = null;
		}
			
		try {
			//FileInputStream myxls = new FileInputStream(file);
			//Workbook workbook = WorkbookFactory.create(myxls);
			sheet = workbook.getSheet(sheetName);
			
			productsList2.forEach(product -> {
				int maxRowsCount = 1;
				if(product.getMedplusProducts().size() > maxRowsCount)
					maxRowsCount = product.getMedplusProducts().size();
				else if(product.getNetMedsProducts().size() > maxRowsCount)
					maxRowsCount = product.getNetMedsProducts().size();
				else if(product.getOneMgProducts().size()>maxRowsCount)
					maxRowsCount=product.getOneMgProducts().size();
				
				for(int i = 0; i < maxRowsCount; i++) {				
					Row row = sheet.createRow(sheet.getLastRowNum() + 1);
					row.createCell(0).setCellValue(product.getSearchTag());
					row.createCell(1).setCellValue(product.getSearchHits());
					
					Cell cell2 = row.createCell(2);
					if(product.getMedplusProducts().size()>i)
						cell2.setCellValue(product.getMedplusProducts().get(i));
					
					Cell cell3 = row.createCell(3);
					if(product.getOneMgProducts().size()>i)
						cell3.setCellValue(product.getOneMgProducts().get(i));
					
					Cell cell4 = row.createCell(4);
					if(product.getNetMedsProducts().size()>i)
						cell4.setCellValue(product.getNetMedsProducts().get(i));
				}
			});
			System.out.println("The keywords done are "+sheet.getLastRowNum());
			//myxls.close();
			FileOutputStream fos = new FileOutputStream(file);
			workbook.write(fos);
			fos.close();
			workbook.close();
			workbook =null;
		} catch (IOException e) {
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		
		//String[] sheetNames = new String[] {"medlife","apollo","netmeds","1mg","medplus"};
		
		String[] sheetNames = new String[] {"medplus"};
		WriteProductDetails readProducts = new WriteProductDetails();
		
		for (String sheet : sheetNames) {
			
			File file = new File("/home/testingteam/Desktop/demoNew.xlsx");			
			readProducts.readExcelFile(file, sheet);
		}
		
		System.out.println("all done");
		
		
	}
}
