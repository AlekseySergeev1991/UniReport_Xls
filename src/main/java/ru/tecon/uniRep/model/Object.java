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
    private boolean isInterrupt = false;

    public Object(int objId, String name) {
        this.objId = objId;
        this.name = name;
    }

    public Object(int objId, String name, List<Param> paramList, boolean isInterrupt) {
        this.objId = objId;
        this.name = name;
        this.paramList = paramList;
        this.isInterrupt = isInterrupt;
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

    public List<Param> getParamList() {
        return paramList;
    }

    public void setParamList(List<Param> paramList) {
        this.paramList = paramList;
    }

    public boolean isInterrupt() {
        return isInterrupt;
    }

    public void setInterrupt(boolean interrupt) {
        isInterrupt = interrupt;
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", Object.class.getSimpleName() + "[", "]")
                .add("objId=" + objId)
                .add("name='" + name + "'")
                .add("paramList=" + paramList)
                .add("isInterrupt=" + isInterrupt)
                .toString();
    }
}
