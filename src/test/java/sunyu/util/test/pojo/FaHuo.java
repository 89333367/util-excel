package sunyu.util.test.pojo;

import org.ttzero.excel.annotation.ExcelColumn;

import java.time.LocalDateTime;

public class FaHuo {
    @ExcelColumn(value = "ICCID/")
    private String iccid;
    @ExcelColumn(value = "SIM卡号/显示器底板编号/ICCID/")
    private String simMsisdn;
    @ExcelColumn(value = "SIM卡到期年限")
    private LocalDateTime exprTime;

    public FaHuo() {
    }

    public String getIccid() {
        return iccid;
    }

    public void setIccid(String iccid) {
        this.iccid = iccid;
    }

    public String getSimMsisdn() {
        return simMsisdn;
    }

    public void setSimMsisdn(String simMsisdn) {
        this.simMsisdn = simMsisdn;
    }

    public LocalDateTime getExprTime() {
        return exprTime;
    }

    public void setExprTime(LocalDateTime exprTime) {
        this.exprTime = exprTime;
    }
}
