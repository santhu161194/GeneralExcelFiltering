package filterdata;

import lombok.Data;

@Data
public class Product {
	private String productId;
	private String manufacturerId;
	private String composition;
	private String price;
	private String url;
}
