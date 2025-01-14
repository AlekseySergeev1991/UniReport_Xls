package ru.tecon.uniRep.ejb;

import jakarta.annotation.Resource;
import jakarta.ejb.Stateless;
import ru.tecon.uniRep.UniRep;


import javax.sql.DataSource;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.logging.Logger;

@Stateless
public class UniRepBean {

    private static final Logger LOGGER = Logger.getLogger(UniRepBean.class.getName());

    @Resource(name = "jdbc/DataSourceR")
    private DataSource dsR;
    @Resource(name = "jdbc/DataSourceRW")
    private DataSource dsRW;


    public void createReport(int reportId) {
        ExecutorService executor = Executors.newSingleThreadExecutor();
        executor.execute(() -> {
            try {
                UniRep.makeReport(reportId, dsR, dsRW);
            } catch (InterruptedException e) {
                throw new RuntimeException(e);
            }
        });
    }
}
