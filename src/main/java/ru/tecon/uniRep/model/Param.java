package ru.tecon.uniRep.model;

import java.io.Serializable;
import java.util.List;
import java.util.StringJoiner;

/**
 * Класс описывающий структуру параметров для каждого отдельного объекта в отчете
 * @author Aleksey Sergeev
 */
public class Param implements Serializable {

    private Integer parId;
    private String parName;
    private String tecProc;
    private String units;
    private String total;
    private String statAggr;
    private List<Value> curDateParam;

    public Param(Integer parId, String parName, String tecProc, String units, String total, String statAggr, List<Value> curDateParam) {
        this.parId = parId;
        this.parName = parName;
        this.tecProc = tecProc;
        this.units = units;
        this.total = total;
        this.statAggr = statAggr;
        this.curDateParam = curDateParam;
    }

    public Integer getParId() {
        return parId;
    }

    public void setParId(Integer parId) {
        this.parId = parId;
    }

    public String getParName() {
        return parName;
    }

    public void setParName(String parName) {
        this.parName = parName;
    }

    public String getTecProc() {
        return tecProc;
    }

    public void setTecProc(String tecProc) {
        this.tecProc = tecProc;
    }

    public String getUnits() {
        return units;
    }

    public void setUnits(String units) {
        this.units = units;
    }

    public String getTotal() {
        return total;
    }

    public void setTotal(String total) {
        this.total = total;
    }

    public List<Value> getCurDateParam() {
        return curDateParam;
    }

    public void setCurDateParam(List<Value> curDateParam) {
        this.curDateParam = curDateParam;
    }

    public String getStatAggr() {
        return statAggr;
    }

    public void setStatAggr(String statAggr) {
        this.statAggr = statAggr;
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", Param.class.getSimpleName() + "[", "]")
                .add("parId=" + parId)
                .add("parName='" + parName + "'")
                .add("tecProc='" + tecProc + "'")
                .add("units='" + units + "'")
                .add("total='" + total + "'")
                .add("statAggr='" + statAggr + "'")
                .add("curDateParam=" + curDateParam)
                .toString();
    }
}
