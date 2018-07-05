package filterdata;

import java.util.List;

import lombok.Data;
import lombok.EqualsAndHashCode;

@Data
@EqualsAndHashCode(exclude= {"medplusProducts", "netMedsProducts", "oneMgProducts", "searchHits"})
public class SolrSearchResponse {
	private String searchTag;
	private List<String> medplusProducts;
	private List<String> netMedsProducts;
	private List<String> oneMgProducts;
	private Double searchHits;
}
