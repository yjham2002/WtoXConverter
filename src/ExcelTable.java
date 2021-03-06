import java.util.ArrayList;
import java.util.List;

/**
 * Created by a on 2017-02-03.
 */
public class ExcelTable {
    public boolean noParam = false;
    public boolean noReturn = false;
    public String api = "#API";
    public String exp = "#EXP";
    public String sam = "#SAM";
    public String ret = "#RET";
    public List<ExcelRow> listParams = new ArrayList<ExcelRow>();
    public List<ExcelRow> listReturns = new ArrayList<ExcelRow>();

    public ExcelTable(){}

    public ExcelTable(String api, String exp, String sam, String ret, List<ExcelRow> listParams, List<ExcelRow> listReturns, boolean noParam, boolean noReturn) {
        this.api = api;
        this.exp = exp;
        this.sam = sam;
        this.ret = ret;
        this.noParam = noParam;
        this.noReturn = noReturn;
        this.listParams = listParams;
        this.listReturns = listReturns;
    }
}
