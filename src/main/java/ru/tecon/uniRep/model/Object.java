package ru.tecon.uniRep.model;

import java.io.Serializable;
import java.util.List;
import java.util.StringJoiner;

/**
 * Класс описывающий структуру объектов, для которых формируется отчет
 * @author Aleksey Sergeev
 */
public class Object implements Serializable {

    private int objId;
    private String name;
    private List<Param> paramList;

    public Object(int objId, String name) {
        this.objId = objId;
        this.name = name;
    }

    public int getObjId() {
        return objId;
    }

    public String getName() {
        return name;
    }

    public List<Param> getParamList() {
        return paramList;
    }

    public void setParamList(List<Param> paramList) {
        this.paramList = paramList;
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", Object.class.getSimpleName() + "[", "]")
                .add("objId=" + objId)
                .add("name='" + name + "'")
                .add("paramList=" + paramList)
                .toString();
    }
}
