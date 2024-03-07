package ru.tecon.uniRep.model;

import org.apache.poi.ss.usermodel.CellStyle;

import java.util.StringJoiner;

public class CellStyleClass {
    private String colorHex;
    private CellStyle coloredCell;

    public CellStyleClass(String colorHex, CellStyle coloredCell) {
        this.colorHex = colorHex;
        this.coloredCell = coloredCell;
    }

    public String getColorHex() {
        return colorHex;
    }

    public void setColorHex(String colorHex) {
        this.colorHex = colorHex;
    }

    public CellStyle getColoredCell() {
        return coloredCell;
    }

    public void setColoredCell(CellStyle coloredCell) {
        this.coloredCell = coloredCell;
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", CellStyleClass.class.getSimpleName() + "[", "]")
                .add("colorHex='" + colorHex + "'")
                .add("coloredCell=" + coloredCell)
                .toString();
    }
}
