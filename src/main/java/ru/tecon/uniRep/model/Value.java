package ru.tecon.uniRep.model;

import java.io.Serializable;
import java.util.StringJoiner;

/**
 * Класс описывающий структуру значение параметра на выбранный перриод
 * @author Aleksey Sergeev
 */
public class Value implements Serializable {
    private String value;
    private String color;

    public Value(String value, String color) {
        this.value = value;
        this.color = color;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public String getColor() {
        return color;
    }

    public void setColor(String color) {
        this.color = color;
    }


    @Override
    public String toString() {
        return new StringJoiner(", ", Value.class.getSimpleName() + "[", "]")
                .add("value='" + value + "'")
                .add("color='" + color + "'")
                .toString();
    }
}
