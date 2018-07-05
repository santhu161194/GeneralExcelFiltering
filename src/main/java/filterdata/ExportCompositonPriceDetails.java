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
import org.apache.solr.client.solrj.SolrResponse;
import org.apache.solr.client.solrj.SolrServerException;
import org.apache.solr.client.solrj.impl.HttpSolrClient;
import org.apache.solr.client.solrj.response.QueryResponse;
import org.apache.solr.common.SolrDocument;

public class ExportCompositonPriceDetails {
	private SolrClient solrClient = new HttpSolrClient.Builder("http://192.168.1.192:8983/solr/search_products")
			.build();
	private SolrQuery query = new SolrQuery();
	private QueryResponse response;
	private List<String> companies = new ArrayList<>(Arrays.asList("Medplus", "1mg", "netmeds"));
	private Sheet sheet;

	void readExcelFile(File file, String sheetName)
			throws EncryptedDocumentException, InvalidFormatException, IOException {
		Workbook workbook = WorkbookFactory.create(file);

		System.out.println("The workbook has " + workbook.getNumberOfSheets());

		final Set<SolrSearchResponseForNAP> keywords = new LinkedHashSet<>();
		Sheet sheet = workbook.getSheet(sheetName);

		System.out.println("The name of the current sheet is " + sheet.getSheetName());

		sheet.rowIterator().forEachRemaining(row -> {
			if (row.getRowNum() != 0) {
				String cellval = null;
				Cell cell;
				SolrSearchResponseForNAP solrSearchResponse = new SolrSearchResponseForNAP();
				cell = row.getCell(1);
				try {
					solrSearchResponse.setQuantity(cell.getNumericCellValue());
				} catch (Exception e) {
					solrSearchResponse.setQuantity(0d);
				}
				cell = row.getCell(0);
				try {
					cellval = cell.getStringCellValue();
				} catch (Exception e) {
					try {
						cellval = String.valueOf(cell.getNumericCellValue());
					} catch (Exception e1) {
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
		// int tempSize=keywords.size()/1000;
		/*
		 * for(int i=0;i<tempSize;i++) { List<SolrSearchResponse> tempKeyWords = new
		 * LinkedList<SolrSearchResponse>(); Iterator<SolrSearchResponse> itr2 =
		 * keywords.iterator(); int i1=0; while(itr2.hasNext()&&i1<1000) {
		 * SolrSearchResponse solrres = itr2.next(); populateSolrData(solrres);
		 * tempKeyWords.add(solrres); i1++; } writeToExcel(sheetName, tempKeyWords);
		 * keywords.removeAll(tempKeyWords); tempKeyWords = null;
		 * System.out.println("The keywords size is"+keywords.size());
		 * System.out.println((i*1000) + " are done "); }
		 */

	}

	void populateSolrData(SolrSearchResponseForNAP solrSearchResponse) {
			query.setQuery("name:" + solrSearchResponse.getSearchTag().replaceAll("/", ""));
			query.setRows(100);
			try {
				response = solrClient.query(query);
			} catch (SolrServerException | IOException e) {
				e.printStackTrace();
			}
			List<SolrDocument> medplusProductsList=response.getResults().stream().filter(result->result.get("company").toString().equals("Medplus")).collect(Collectors.toList());
			solrSearchResponse.setMedplusProducts(extractSolrResponse(medplusProductsList.subList(0, medplusProductsList.size()>10?10:medplusProductsList.size())));
			List<SolrDocument> oneMgProductsList=response.getResults().stream().filter(result->result.get("company").toString().equals("1mg")).collect(Collectors.toList());
			solrSearchResponse.setOneMgProducts(extractSolrResponse(oneMgProductsList.subList(0, oneMgProductsList.size()>10?10:oneMgProductsList.size())));
			List<SolrDocument> netMedsProductsList=response.getResults().stream().filter(result->result.get("company").toString().equals("netmeds")).collect(Collectors.toList());
			solrSearchResponse.setNetMedsProducts(extractSolrResponse(netMedsProductsList.subList(0, netMedsProductsList.size()>10?10:netMedsProductsList.size())));
	}
	
	private List<Product> extractSolrResponse(List<SolrDocument> productsList) {
		List<Product> products=new ArrayList<>();
		productsList.forEach(result->{
			Product product=new Product();
			product.setProductId(result.get("name")==null?"":result.get("name").toString());
			product.setComposition(result.get("composition")==null?"":result.get("composition").toString());
			product.setManufacturerId(result.get("manufacturer")==null?"":result.get("manufacturer").toString());
			product.setPrice(result.get("price")==null?"":result.get("price").toString());
			product.setUrl(result.get("url")==null?"":result.get("url").toString());
			products.add(product);
		});
		return products;
	}

	private void writeToExcel(String sheetName, Set<SolrSearchResponseForNAP> productsList2) throws IOException {
		File file = new File("/home/testingteam/Desktop/notAvailableList/" + sheetName + ".xlsx");
		Workbook workbook = null;
		if (!file.exists()) {
			// Workbook workbook = new XSSFWorkbook();
			workbook = new XSSFWorkbook();
			sheet = workbook.createSheet(sheetName);
			for(int i=2;i<=12;i++)
				sheet.setColumnWidth(i, 11000);
			
			Row header = sheet.createRow(0);
			header.createCell(0).setCellValue("ProductName");
			header.createCell(1).setCellValue("Quantity");
			header.createCell(2).setCellValue("Medplus Product");
			header.createCell(3).setCellValue("1mg Product");
			header.createCell(4).setCellValue("1mg Manufacturer");
			header.createCell(5).setCellValue("1mg Composition");
			header.createCell(6).setCellValue("1mg Price");
			header.createCell(7).setCellValue("1mg Url");
			header.createCell(8).setCellValue("netMeds Product");
			header.createCell(9).setCellValue("netMeds Manufacturer");
			header.createCell(10).setCellValue("netMeds Composition");
			header.createCell(11).setCellValue("netMeds Price");
			header.createCell(12).setCellValue("netMeds Url");
			// FileOutputStream fos = new FileOutputStream(file);
			// workbook.write(fos);
			// workbook.close();
			// workbook = null;
		}

		try {
			// FileInputStream myxls = new FileInputStream(file);
			// Workbook workbook = WorkbookFactory.create(myxls);
			sheet = workbook.getSheet(sheetName);

			productsList2.forEach(product -> {
				int maxRowsCount = 1;
				if (product.getMedplusProducts().size() > maxRowsCount)
					maxRowsCount = product.getMedplusProducts().size();
				else if (product.getNetMedsProducts().size() > maxRowsCount)
					maxRowsCount = product.getNetMedsProducts().size();
				else if (product.getOneMgProducts().size() > maxRowsCount)
					maxRowsCount = product.getOneMgProducts().size();

				for (int i = 0; i < maxRowsCount; i++) {
					Row row = sheet.createRow(sheet.getLastRowNum() + 1);
					Cell[] cells = new Cell[13];
					for(int j=0;j<=12;j++)
						cells[j]=row.createCell(j);
					cells[0].setCellValue(product.getSearchTag());
					cells[1].setCellValue(product.getQuantity());
					if (product.getMedplusProducts().size() > i)
						cells[2].setCellValue(product.getMedplusProducts().get(i).getProductId());
					if (product.getOneMgProducts().size() > i) {
						cells[3].setCellValue(product.getOneMgProducts().get(i).getProductId());
						cells[4].setCellValue(product.getOneMgProducts().get(i).getManufacturerId());
						cells[5].setCellValue(product.getOneMgProducts().get(i).getComposition());
						cells[6].setCellValue(product.getOneMgProducts().get(i).getPrice());
						cells[7].setCellValue(product.getOneMgProducts().get(i).getUrl());
					}
					if (product.getNetMedsProducts().size() > i) {
						cells[8].setCellValue(product.getNetMedsProducts().get(i).getProductId());
						cells[9].setCellValue(product.getNetMedsProducts().get(i).getManufacturerId());
						cells[10].setCellValue(product.getNetMedsProducts().get(i).getComposition());
						cells[11].setCellValue(product.getNetMedsProducts().get(i).getPrice());
						cells[12].setCellValue(product.getNetMedsProducts().get(i).getUrl());
					}
				}
			});
			System.out.println("The keywords done are " + sheet.getLastRowNum());
			FileOutputStream fos = new FileOutputStream(file);
			workbook.write(fos);
			fos.close();
			workbook.close();
			workbook = null;
		} catch (IOException e) {
			e.printStackTrace();
		} catch (EncryptedDocumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {

		// String[] sheetNames = new String[]
		// {"medlife","apollo","netmeds","1mg","medplus"};

		String[] sheetNames = new String[] { "Sheet1" };
		ExportCompositonPriceDetails readProducts = new ExportCompositonPriceDetails();

		for (String sheet : sheetNames) {

			File file = new File("/home/testingteam/Desktop/notAvailableList/NOTAVAILABLELIST1.0.xlsx");
			readProducts.readExcelFile(file, sheet);
		}

		System.out.println("all done");

	}
}
