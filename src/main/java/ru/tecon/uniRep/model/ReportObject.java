package ru.tecon.uniRep.model;

import java.io.Serializable;
import java.util.List;
import java.util.StringJoiner;

/**
 * Класс описывающий структуру объектов, для которых формируется отчет
 * @author Aleksey Sergeev
 */
public class ReportObject implements Serializable {
    private int objId;
    private String name;
    private String addres;
    private List<Param> paramList;

    public ReportObject(int objId, String name, String addres) {
        this.objId = objId;
        this.name = name;
        this.addres = addres;
    }

    public int getObjId() {
        return objId;
    }

    public void setObjId(int objId) {
        this.objId = objId;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getAddres() {
        return addres;
    }

    public void setAddres(String addres) {
        this.addres = addres;
    }

    public List<Param> getParamList() {
        return paramList;
    }

    public void setParamList(List<Param> paramList) {
        this.paramList = paramList;
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", ReportObject.class.getSimpleName() + "[", "]")
                .add("objId=" + objId)
                .add("name='" + name + "'")
                .add("addres='" + addres + "'")
                .add("paramList=" + paramList)
                .toString();
    }
}
