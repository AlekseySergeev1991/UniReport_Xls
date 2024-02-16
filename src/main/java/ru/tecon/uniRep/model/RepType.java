package ru.tecon.uniRep.model;

import java.io.Serializable;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.StringJoiner;

/**
 * Класс описывающий тип отчета с указанием дат и интервалов
 * @author Aleksey Sergeev
 */
public class RepType implements Serializable {

    private int maskId;
    private LocalDateTime beg;
    private LocalDateTime end;
    private String interval;

    public RepType() {
    }

    public RepType(int maskId, LocalDateTime beg, LocalDateTime end, String interval) {
        this.maskId = maskId;
        this.beg = beg;
        this.end = end;
        this.interval = interval;
    }

    public int getMaskId() {
        return maskId;
    }

    public void setMaskId(int maskId) {
        this.maskId = maskId;
    }

    public LocalDateTime getBeg() {
        return beg;
    }

    public void setBeg(LocalDateTime beg) {
        this.beg = beg;
    }

    public LocalDateTime getEnd() {
        return end;
    }

    public void setEnd(LocalDateTime end) {
        this.end = end;
    }

    public String getInterval() {
        return interval;
    }

    public void setInterval(String interval) {
        this.interval = interval;
    }

    @Override
    public String toString() {
        return new StringJoiner(", ", RepType.class.getSimpleName() + "[", "]")
                .add("maskId=" + maskId)
                .add("beg=" + beg)
                .add("end=" + end)
                .add("interval='" + interval + "'")
                .toString();
    }
}
