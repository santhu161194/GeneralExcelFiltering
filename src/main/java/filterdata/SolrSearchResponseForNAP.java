
package filterdata;

import java.util.List;

import lombok.Data;
import lombok.EqualsAndHashCode;

@Data
@EqualsAndHashCode(exclude= {"medplusProducts", "netMedsProducts", "oneMgProducts", "searchHits"})
public class SolrSearchResponseForNAP {
	private String searchTag;
	private Double quantity;
	private List<Product> medplusProducts;
	private List<Product> netMedsProducts;
	private List<Product> oneMgProducts;
}
