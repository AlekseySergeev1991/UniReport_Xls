package ru.tecon.uniRep.model;

import java.io.Serializable;
import java.util.List;
import java.util.StringJoiner;

/**
 * Класс описывающий структуру параметров для каждого отдельного объекта в отчете
 * @author Aleksey Sergeev
 */
public class Param implements Serializable {

    private int parId;
    private String parName;
    private String tecProc;
    private String units;
    private String total;
    private List<Value> curDateParam;

    public Param() {
    }

    public Param(int parId, String parName, String tecProc, String units, String total, List<Value> curDateParam) {
        this.parId = parId;
        this.parName = parName;
        this.tecProc = tecProc;
        this.units = units;
        this.total = total;
        this.curDateParam = curDateParam;
    }

    public int getParId() {
        return parId;
    }

    public void setParId(int parId) {
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

//    public String getParCateg() {
//        return parCateg;
//    }
//
//    public void setParCateg(String parCateg) {
//        this.parCateg = parCateg;
//    }
//
//    public String getDifInt() {
//        return difInt;
//    }
//
//    public void setDifInt(String difInt) {
//        this.difInt = difInt;
//    }

    public List<Value> getCurDateParam() {
        return curDateParam;
    }

    public void setCurDateParam(List<Value> curDateParam) {
        this.curDateParam = curDateParam;
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", Param.class.getSimpleName() + "[", "]")
                .add("parId=" + parId)
                .add("parName='" + parName + "'")
                .add("tecProc='" + tecProc + "'")
                .add("units='" + units + "'")
                .add("total='" + total + "'")
//                .add("parCateg='" + parCateg + "'")
//                .add("difInt='" + difInt + "'")
                .add("curDateParam=" + curDateParam)
                .toString();
    }
}
