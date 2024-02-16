package ru.tecon.uniRep;

import org.apache.commons.codec.DecoderException;
import org.apache.commons.codec.binary.Hex;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import ru.tecon.uniRep.model.Object;
import ru.tecon.uniRep.model.Param;
import ru.tecon.uniRep.model.RepType;
import ru.tecon.uniRep.model.Value;

import javax.ejb.TransactionAttribute;
import javax.ejb.TransactionAttributeType;
import javax.sql.DataSource;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.sql.*;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.Date;
import java.util.logging.Level;
import java.util.logging.Logger;

public class UniRep {

    private static final Logger LOGGER = Logger.getLogger(UniRep.class.getName());
    private static final String LOAD_REP_TYPE = "SELECT * FROM td_rep.get_rep_info (?)";
    private static final String LOAD_OBJECT = "select obj_id,  admin.get_obj_name (obj_id) from td_rep.get_objects_from_mask (?);";
//    private static final String LOAD_OBJECT_NAME = "SELECT OBJ_NAME FROM admin.obj_object WHERE obj_id = ?";
    private static final String LOAD_PARAMS = "SELECT * FROM td_rep.get_common_obj_par_values (?, ?, ?, ?, ?) order by 2, 4;";
    private static final String GET_PAR_ID = "SELECT par_id FROM td_rep.get_params_from_mask (?, ?)";
    private static final String LOAD_H_VALUE ="SELECT * FROM td_rep.get_hour_obj_par_values(?, ?, ?, ?, ?)";
    private static final String LOAD_D_VALUE ="SELECT * FROM dsp_0032t.get_data_param_m(?,?,?,?) where time_stamp = ?";
    private static final String LOAD_M_VALUE ="SELECT * FROM dsp_0032t.get_data_param_y(?,?,?,?) where time_stamp = ?";
    private static final String PERCENT = "call td_rep.set_run_procent(?, ?)";
    private static final String INTERRUPTED = "SELECT STATUS FROM td_rep.get_rep_info (?)";
    private static final String DELSQL = "delete from admin.REP_BLOB where REP_id = ?";
    private static final String SQL = "call td_rep.ins_rep_blob(?,?,?::integer)";
    private static final String FINSQL = "update admin.REP_REPORT set RUN_PROCENT = 100, STATUS = 'F' where id = ?";

    private DataSource dsR;
    public void setDsR(DataSource dsR) {
        this.dsR = dsR;
    }

    private DataSource dsRW;

    public void setDsRW(DataSource dsRW) {
        this.dsRW = dsRW;
    }


    private int Rows;


//    @EJB
//    private UniRepInterface bean;

//    public static void main(String[] args) throws IOException, SQLException, ParseException {
//        // Здесь вызов этой порнографии
//        UniRep mr = new UniRep();
//
//        SXSSFWorkbook w = mr.printReport(31956, mr.ds);
//        int j = mr.saveReportIntoFile(w,"C:\\abc\\TEST.xlsx");
////        int i = mr.saveReportIntoTable (w, 31956);
//    }

    public static void makeReport (int p_Rep_Id, DataSource dsR, DataSource dsRW) {
        long currentTime = System.nanoTime();
        LOGGER.log(Level.INFO, "start make report {0}", p_Rep_Id);

        UniRep mr = new UniRep();
        mr.setDsR(dsR);
        mr.setDsRW(dsRW);
        SXSSFWorkbook w = null;
        int i = 0;
        try {
            System.out.println("start");
//            w = mr.printReport(p_Rep_Id, dsR, dsRW);
            w = mr.printReportTemp(p_Rep_Id, dsR, dsRW);
            i = mr.saveBlobIntoTable (w, p_Rep_Id, dsRW);
//            mr.saveReportIntoFile(w, "C:\\abc\\TESTNewHeader64137.xlsx");
        } catch (IOException | SQLException | ParseException | DecoderException e) {
            LOGGER.log(Level.WARNING, "makeReport error", e);
//            e.printStackTrace();
        }
//        int j = saveReportIntoFile(w,"D:\\TEST.xls");

        LOGGER.log(Level.INFO, "report created {0} created time {1}", new java.lang.Object[]{p_Rep_Id, (System.nanoTime() - currentTime)});

    }

    /*
      Метод создает нужный воркбук. Параметры:
      p_Rep_Id - идентификатор отчета
      p_Rep_Type - тип репорта. Принимает на вход один символ, малая латинская буква: h - часовой, d - дневной, m - месячный
      p_Beg_Date - Начальная дата в ткстовом формате (тут пишу в ораклиной нотации) dd-mm-yyyy hh24:mi
      p_End_Date - Конечная дата в ткстовом формате (тут пишу в ораклиной нотации) dd-mm-yyyy hh24:mi. Прекрасно понимаю, что ее можно рассчитать
                   с помощью количества колонок, типа отчета и начальной даты
      p_Rows - количество строк в отчете
      p_Data_Cols - количество колонок - значений параметров. В шапке на 4 колонки больше

    */
    public SXSSFWorkbook printReport (int p_Rep_Id, DataSource dsR, DataSource dsRW) throws IOException, SQLException, ParseException, DecoderException {
//        SXSSFWorkbook wb = new SXSSFWorkbook();
        SXSSFWorkbook wb = new SXSSFWorkbook();
        SXSSFSheet sh = wb.createSheet("Отчет");
        final int begRow = 7;  // строка в екселе, с которой начинается собственно отчет.

        System.out.println(p_Rep_Id);
        System.out.println(dsR);
//        RepType repType = bean.loadRepType(p_Rep_Id, dsR);
        RepType repType = loadRepType(p_Rep_Id, dsR);
        System.out.println("repType "+repType);

        long cols = 0;
        List<LocalDateTime> dateList = new ArrayList<>();
        LocalDateTime localDateTemp = repType.getBeg();
//        dateList.add(localDateTemp);
        switch (repType.getInterval()) {
            case ("D"):
                // Считаем количество часов между датами; Для часов убираем последний час
//                cols = (diffDate + 1) * 24;
//                cols = ChronoUnit.HOURS.between((Temporal) repType.getEnd(), (Temporal) repType.getBeg());
//                cols = ((repType.getEnd().getTime() - repType.getBeg().getTime())/ (1000*60*60)) +1;
                for (;;) {
//                    if (cols == 0 || repType.getEnd().plusHours(1).isAfter(localDateTemp)) {

                   if (cols == 0 || repType.getEnd().plusDays(1).isAfter(localDateTemp)) {
                       cols++;
//                       if (cols != 0) {
                           localDateTemp = localDateTemp.plusHours(1);
                           dateList.add(localDateTemp);
//                       }
                   } else {
                       break;
                   }
                }

//                for (int i = 0; i < cols; i++) {
//                    long tempDate = repType.getBeg().getTime() + i * (1000*60*60);
//                    Date date = new Date(tempDate);
//                    dateList.add(i, date);
//                }
                break;

            case ("M"):
                dateList.add(localDateTemp);
                //-- Считаем количество дней между датами
//                cols = (diffDate + 1);
//                cols = ((repType.getEnd().getTime() - repType.getBeg().getTime())/ (1000*60*60*24)) +1;
//                for (int i = 0; i < cols; i++) {
//                    long tempDate = repType.getBeg().getTime() + i * (1000*60*60*24);
//                    Date date = new Date(tempDate);
//                    dateList.add(i, date);
//                }
                for (;;) {
                    if (cols == 0 || repType.getEnd().isAfter(localDateTemp)) {
                        if (cols != 0) {
                            localDateTemp = localDateTemp.plusDays(1);
                            dateList.add(localDateTemp);
                        }
                        cols++;
                    } else {
                        break;
                    }
                }
                break;
            case ("Y"):
                dateList.add(localDateTemp);
//                cols = (monthes + 1);
//                cols = ChronoUnit.MONTHS.between(YearMonth.from((TemporalAccessor) repType.getBeg()), YearMonth.from((TemporalAccessor) repType.getEnd())) +1;
//                for (int i = 0; i < cols; i++) {
//
//                    LocalDate localDate = (Date) repType.getBeg();
//
//                    dateList.add(i, date);
//                }
                for (;;) {
                    if (cols == 0 || repType.getEnd().isAfter(localDateTemp)) {
                        if (cols != 0) {
                            localDateTemp = localDateTemp.plusMonths(1);
                            dateList.add(localDateTemp);
                            System.out.println();
                        }
                        cols++;
                        if (repType.getEnd().getMonth().equals(localDateTemp.getMonth()) && repType.getEnd().getYear() == localDateTemp.getYear()) {
                            break;
                        }
                    } else {
                        break;
                    }
                }
                break;
        }

        // Пожалуй, наполню-ка я лист отдельным методом. И сначала заполняем его, чтобы узнать Rows. Cols

        // Колонок - 4 первых и колонка с датами

//        CellStyle usualStyle = setUsualStyle(wb, "LEFT");
        CellStyle headerStyle = setHeaderStyle(wb);
        CellStyle headerStyleNoBold = setHeaderStyleNoBold(wb);
//        CellStyle notBoldHeaderStyle = setNotBoldHeaderStyle(wb);
        CellStyle cappyStyle = setCappyStyle(wb, "UpDn");
        CellStyle nowStyle = setCellNow (wb);


        // Устанавливаем ширины колонок. В конце мероприятия
        sh.setColumnWidth(0, 3 * 256);
        sh.setColumnWidth(1, 32 * 256);
        for (int i = 2; i < cols; i++) {
            sh.setColumnWidth(i, 10 * 256);
        }

        SXSSFRow row_1 = sh.createRow(0);
        row_1.setHeight((short) 435);
        SXSSFCell cell_1_1 = row_1.createCell(0);
        cell_1_1.setCellValue("ПАО \"МОЭК\": АС \"ТЕКОН - Диспетчеризация\"");
        CellRangeAddress title = new CellRangeAddress(0, 0, 0, 17);
        sh.addMergedRegion(title);
        cell_1_1.setCellStyle(headerStyle);

        SXSSFRow row_2 = sh.createRow(1);
        row_2.setHeight((short) 435);
        SXSSFCell cell_2_1 = row_2.createCell(0);
        cell_2_1.setCellValue("Многофункциональный отчет");
        CellRangeAddress formName = new CellRangeAddress(1, 1, 0, 17);
        sh.addMergedRegion(formName);
        cell_2_1.setCellStyle(headerStyle);

        // Через строчку идет период дат
        SXSSFRow row_3 = sh.createRow(2);
        row_3.setHeight((short) 435);
        SXSSFCell cell_3_1 = row_3.createCell(0);
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd.MM.yyyy");
        LocalDateTime begFormatted = repType.getBeg();
        String stringBeg = begFormatted.format(formatter);
        LocalDateTime endFormatted;
        if (repType.getInterval().equals("D")) {
            endFormatted = repType.getEnd().plusDays(1);
        } else {
            endFormatted = repType.getEnd();
        }
        String stringEnd = endFormatted.format(formatter);
        cell_3_1.setCellValue("за период: " + stringBeg + " - " + stringEnd);
        CellRangeAddress datePer = new CellRangeAddress(2, 2, 0, 17);
        sh.addMergedRegion(datePer);
        cell_3_1.setCellStyle(headerStyleNoBold);

        SXSSFRow row_4 = sh.createRow(3);
        row_4.setHeight((short) 435);
        SXSSFCell cell_4_1 = row_4.createCell(0);
        String interval = "";
        switch (repType.getInterval()) {
            case ("D"):
                interval = "Часовой";
                break;
            case ("M"):
                interval = "Дневной";
                break;
            case ("Y"):
                interval = "Месячный";
                break;
        }
        cell_4_1.setCellValue("Интервал: " + interval);
        CellRangeAddress intervalR = new CellRangeAddress(3, 3, 0, 17);
        sh.addMergedRegion(intervalR);
        cell_4_1.setCellStyle(headerStyleNoBold);



        // Печатаем отчетов зад
        String now = new SimpleDateFormat("dd.MM.yyyy HH:mm").format(new Date());
        SXSSFRow row_5 = sh.createRow(4);
        row_5.setHeight((short) 435);
        SXSSFCell cell_5_1 = row_5.createCell(0);
        cell_5_1.setCellStyle(nowStyle);
        cell_5_1.setCellValue("Отчет сформирован  " + now);
        CellRangeAddress nowDone = new CellRangeAddress(4, 4, 0, 17);
        sh.addMergedRegion(nowDone);


        // Прекрасно, Заголовок сделали. Готовим шапку.
        SXSSFRow row_7 = sh.createRow(6);
        SXSSFCell cell_7_0 = row_7.createCell(0);
        cell_7_0.setCellStyle(cappyStyle);
        cell_7_0.setCellValue("№");

        SXSSFCell cell_7_1 = row_7.createCell(1);
        cell_7_1.setCellStyle(cappyStyle);
        cell_7_1.setCellValue("Объект/Параметр");

        SXSSFCell cell_7_2 = row_7.createCell(2);
        cell_7_2.setCellStyle(cappyStyle);
        cell_7_2.setCellValue("Тех.Проц.");

        SXSSFCell cell_7_3 = row_7.createCell(3);
        cell_7_3.setCellStyle(cappyStyle);
        cell_7_3.setCellValue("Ед.Изм.");

        SXSSFCell cell_7_4 = row_7.createCell(4);
        cell_7_4.setCellStyle(cappyStyle);
        cell_7_4.setCellValue("Итого:");
        // декларируем переменные для шапки
        SXSSFCell cell;

//        SimpleDateFormat format = new SimpleDateFormat();
//        format.applyPattern("dd-MM-yyyy HH:mm");
//
//        SimpleDateFormat formatH = new SimpleDateFormat();
//        formatH.applyPattern("dd-MM HH");
//
//        SimpleDateFormat formatD = new SimpleDateFormat();
//        formatD.applyPattern("dd-MM");
//
//        SimpleDateFormat formatM = new SimpleDateFormat();
//        formatM.applyPattern("MM-yyyy");

        int i = 0;
        for (LocalDateTime curDateLDT : dateList) {
            String curDateS = String.valueOf(curDateLDT);
            curDateS = curDateS.replace('T', ' ');
            cell = row_7.createCell(i + 5);
            cell.setCellStyle(cappyStyle);
            switch (repType.getInterval()) {
                case ("D"):
                    curDateS = curDateS.substring(8, 10)+ "." + curDateS.substring(5, 7)+ " " + curDateS.substring(11, 13);
                    cell.setCellValue(curDateS);
                    break;
                case ("M"):
                    curDateS = curDateS.substring(8, 10)+ "." + curDateS.substring(5, 7);
                    cell.setCellValue(curDateS);
                    break;
                case ("Y"):
                    curDateS = curDateS.substring(5, 7)+ " " + curDateS.substring(0, 4);
                    cell.setCellValue(curDateS);
                    break;
            }
            i++;
        }


//        LocalDateTime curDateLDT = dateList.get(0);
//        System.out.println("curDateLDT "+curDateLDT);

//        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd-MM-yyyy HH:mm");
//        LocalDateTime dateTime = LocalDateTime.from(curDateLDT);
//        String formattedDateTime = dateTime.format(formatter);
//
//        Date curDate = formattedDateTime;
//        Date curDate = java.sql.Date.valueOf(String.valueOf(curDateLDT));
//        LocalDate curDateLD = LocalDate.from(curDateLDT);
//        Date curDate = Date.from(curDateLD.atStartOfDay(ZoneId.systemDefault()).toInstant());
//
//
//
//        Calendar calendar = Calendar.getInstance();
//        calendar.setTime(curDate);
//        System.out.println("calendar.getTime() "+calendar.getTime());
//        System.out.println("cols "+cols);
//
//        for (int i = 0; i < cols; i++) {
//            cell = row_4.createCell(i + 5);
//            cell.setCellStyle(cappyStyle);
//
//
//            switch (repType.getStatus()) {
//                case ("D"):
//                    cell.setCellValue(formatH.format(calendar.getTime()));
//                    calendar.add(Calendar.HOUR_OF_DAY, 1);
//                    break;
//                case ("M"):
//                    cell.setCellValue(formatD.format(calendar.getTime()));
//                    calendar.add(Calendar.DAY_OF_YEAR, 1);
//                    break;
//                case ("Y"):
//                    cell.setCellValue(formatM.format(calendar.getTime()));
//                    calendar.add(Calendar.MONTH, 1);
//                    break;
//            }
//        }

        LOGGER.log(Level.INFO, "Report head created {0}", p_Rep_Id);
        // Отлично. Заголовок и шапка сделаны. Идем по таблице, создаем и заполняем ячейки.

//        fillSheet(wb, p_Rep_Id, begRow, repType.getMaskId(), repType.getBeg(), repType.getEnd(), repType.getStatus(), dateList);

        LOGGER.log(Level.INFO, "Report body created {0}", p_Rep_Id);

//        LOGGER.log(Level.INFO, "Report footer created {0}", p_Rep_Id);

        return wb;
    }

    private void fillSheet (SXSSFWorkbook wb, int p_Rep_Id, int begRow, int maskId, LocalDateTime begDate, LocalDateTime endDate, String status, List<LocalDateTime> dateList) throws SQLException, DecoderException {
        SXSSFSheet sh = wb.getSheetAt(0);
        CellStyle usualStyle = setUsualStyle(wb, "LEFT");
//        CellStyle usualYEllowStyle = setUsualYellowStyle(wb, "LEFT");
        CellStyle boldUpStyle = setBoldUpStyle(wb);


        // Заполняем лист значениями, взятыми из таблицы
        List<Object> objects = loadObjects(maskId, dsR);
        System.out.println("p_Rep_Id " + p_Rep_Id +" objects " + objects);

        Rows = begRow;
        int objNum = 1;
        BigDecimal percentage = new BigDecimal(0);
        double size = objects.size();
        double iterationPercent = 100/size;
        BigDecimal iterationPercentBD = new BigDecimal(iterationPercent);

        for (Object object : objects) {
            System.out.println("object " + object);
            System.out.println("interrupted " + interrupted(p_Rep_Id , dsR));
            if (!interrupted(p_Rep_Id , dsR).equals("Q")) {
//                System.out.println("попали в иф?");
                object.setParamList(loadParams(maskId, object.getObjId(), begDate, endDate, status, dsR));
                SXSSFRow row = sh.createRow(Rows);
                SXSSFCell objNumCell = row.createCell(0);
                objNumCell.setCellValue(objNum);
                objNumCell.setCellStyle(boldUpStyle);
                objNum++;
                SXSSFCell objNameCell = row.createCell(1);
                objNameCell.setCellValue(object.getName());
                objNameCell.setCellStyle(boldUpStyle);
                for (int i = 2; i < dateList.size()+5; i++) {
                    SXSSFCell cell = row.createCell(i);
                    cell.setCellStyle(boldUpStyle);
                }
                Rows++;
                for (Param param : object.getParamList()) {
                    if (!interrupted(p_Rep_Id,dsR).equals("Q")) {
                        SXSSFRow parRow = sh.createRow(Rows);
                        SXSSFCell parNameCell = parRow.createCell(1);
                        parNameCell.setCellValue(param.getParName());
                        parNameCell.setCellStyle(usualStyle);
                        SXSSFCell techProcCell = parRow.createCell(2);
                        techProcCell.setCellValue(param.getTecProc());
                        techProcCell.setCellStyle(usualStyle);
                        SXSSFCell unitCell = parRow.createCell(3);
                        unitCell.setCellValue(param.getUnits());
                        unitCell.setCellStyle(usualStyle);
                        SXSSFCell totalCell = parRow.createCell(4);
                        totalCell.setCellValue(param.getTotal());
                        totalCell.setCellStyle(usualStyle);
                        Rows++;
                        int colNum = 0;
//                        float total = 0;
                        System.out.println("List<Value> "+ param.getCurDateParam());
                        for (Value value : param.getCurDateParam()) {
                            if (!interrupted(p_Rep_Id,dsR).equals("Q")) {
                                SXSSFCell valueCell = parRow.createCell(5 + colNum);
                                CellStyle cellStyle = setUsualStyle(wb, "LEFT");

                                valueCell.setCellValue(value.getValue());
//                                System.out.println("тут ломаемся? "+value.getColor());
//                                System.out.println(value);

                                if (value.getColor() != null) {

                                    System.out.println("if (value.getColor() != null)");
                                        String rgbS = value.getColor();
    //                                    rgbS = "A020F0"; //todo не забудь убрать
                                        byte [] rgbB = Hex.decodeHex(rgbS);
    //                                byte [] rgbB = new byte[]{0, (byte) 100, 0};
                                        XSSFColor color = new XSSFColor(rgbB, null);
    //                                XSSFCellStyle cellStyle = (XSSFCellStyle) wb.createCellStyle();
                                        cellStyle.setFillForegroundColor(color);
                                        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                                    }

                                    valueCell.setCellStyle(cellStyle);


                                colNum++;
//                                total = total + value.getValue();
//                                totalCell.setCellValue(total);
                            } else {
                                break;
                            }
                        }
                    } else {
                        break;
                    }
                }
                percentage = percentage.add(iterationPercentBD).setScale(3, RoundingMode.HALF_UP);
                System.out.println("percentage " + percentage);
//                percent(p_Rep_Id, percentage, dsRW); //todo не забудь вернуть
            } else {
                break;
            }
        }
//            Rows = Rows + 10;
    }







    private int saveReportIntoFile (SXSSFWorkbook workbook, String file) {
        FileOutputStream fos;
        System.out.println("saveReportIntoFile");
        int res = 0;

        try {
            fos = new FileOutputStream(file);
            workbook.write(fos);

        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("Cant create file");
            System.out.println(e.getMessage());
            res = 1;
        }

        System.out.println("ok report");
        return res;
    }

    private int saveBlobIntoTable (SXSSFWorkbook p_wb, int Rep_Id, DataSource dsRW) throws IOException, SQLException {

        int res = 0;
//        if (!interrupted(Rep_Id , dsR).equals("Q")) {
            saveReportIntoTable(p_wb, Rep_Id, dsRW);
//        }
        return res;
    }

    public RepType loadRepType(int repId, DataSource ds) {
//        System.out.println("Сюда вообще попадаем?");
        RepType result = new RepType();
        try (Connection connect = ds.getConnection();
             PreparedStatement stm = connect.prepareStatement(LOAD_REP_TYPE)) {
            stm.setInt(1, repId);
            ResultSet res = stm.executeQuery();
            if (res.next()) {
                result.setMaskId(res.getInt("mask_id"));
                result.setBeg (res.getTimestamp("beg_date").toLocalDateTime());
                result.setEnd(res.getTimestamp("end_date").toLocalDateTime());
                result.setInterval(res.getString("rep_interval"));
//                result.setEnd(result.getEnd().plusDays(1));
                System.out.println(result.getBeg());
                return result;
            }

        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load Rep Type", e);
        }
        return result;
//        RepType repType = new RepType();
//        repType.setMaskId(97485);
//        repType.setBeg(LocalDateTime.of(2023, Month.OCTOBER, 8, 0,0,0, 0));
//        repType.setEnd(LocalDateTime.of(2023, Month.OCTOBER, 11, 0,0,0, 0));
//        repType.setStatus("D");
//
//        return repType;
    }

    public List<Object> loadObjects(int maskId, DataSource ds) {
        List<Object> result = new ArrayList<>();
        try (Connection connect = ds.getConnection();
             PreparedStatement stm = connect.prepareStatement(LOAD_OBJECT)) {
            stm.setInt(1, maskId);
            System.out.println("maskId " + maskId);
            ResultSet res = stm.executeQuery();
            while (res.next()) {
                Object item = new Object(res.getInt("obj_id"), res.getString("name"));
                System.out.println("item "+item);
//                item.setName(loadObjName(item.getObjId(), ds));
                result.add(item);
            }
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load Object", e);
        }
//        Object object1 = new Object(1);
//        object1.setName("01-04-0926/003");
//        Object object2 = new Object(2);
//        object2.setName("01-04-0926/030");
//        result.add(object1);
//        result.add(object2);
        return result;
    }

//    public String loadObjName (int objId, DataSource ds) {
//        try (Connection connection = ds.getConnection();
//             PreparedStatement stm = connection.prepareStatement(LOAD_OBJECT_NAME)) {
//            stm.setInt(1, objId);
//
//            ResultSet res = stm.executeQuery();
//            if (res.next() && (res.getString(1) != null)) {
//                return res.getString(1);
//            }
//        } catch (SQLException e) {
//            LOGGER.log(Level.WARNING, "error load name ", e);
//        }
//        return null;
//    }


    public List<Param> loadParams(int maskId, int objId, LocalDateTime begDate, LocalDateTime endDate, String status, DataSource ds) {
        List<Param> result = new ArrayList<>();
//        Param item = new Param();
//        int parId = loadParId(maskId, objId, ds);
        try (Connection connect = ds.getConnection();
             PreparedStatement stm = connect.prepareStatement(LOAD_PARAMS)) {
            stm.setInt(1, maskId);
            stm.setInt(2, objId);
            stm.setString(3, status);
            stm.setTimestamp(4, Timestamp.valueOf(begDate));
            stm.setTimestamp(5, Timestamp.valueOf(endDate));
            ResultSet res = stm.executeQuery();
            while (res.next()) {
                if (result.get(result.size()-1).getParId() == res.getInt("par_id")) {
                    result.get(result.size()-1).getCurDateParam().add(new Value(res.getString("par_value"),
                            res.getString("color")));
//                    item.getCurDateParam().add(new Value(res.getString("par_value"), res.getString("color")));
                } else {
                    List<Value> curDateParam = new ArrayList<>();
                    curDateParam.add(new Value(res.getString("par_value"), res.getString("color")));
                    result.add(new Param(res.getInt("par_id"), res.getString("par_memo"),
                            res.getString("techproc_type_char"), res.getString("short_measure"), "total",
                            curDateParam));
                }
//                Param item = new Param(res.getInt("par_id"), res.getString("par_memo"),
//                        res.getString("techproc_type_char"), res.getString("short_measure"),
//                        res.getString("par_categ"), res.getString("dif_int"));
//                List<Value> valueList = new ArrayList<>();
//                System.out.println("dateList " + dateList);
//                for (LocalDateTime localDateTime : dateList) {
//                    Value value = loadValue(objId, item.getParId(), localDateTime, status, ds);
//                    valueList.add(value);
//                    if (value.getValue() != null) {
//                        try {
//                            float tempFloat = Float.parseFloat(value.getValue());
//                            total = total + tempFloat;
//                        } catch (NumberFormatException e) {
//                            LOGGER.log(Level.WARNING, "NumberFormatException for value");
//                        }
//                    }

//                }
//                String total = loadTotal(objId, item.getParId(), ds);
//                List <Value> curDateParam = loadValue(objId, item.getParId(), item.getParCateg(), item.getDifInt(), begDate, endDate, status, ds);
//                String total = curDateParam.get(curDateParam.size()-1).getValue();
//                curDateParam.remove(curDateParam.size()-1);
//                item.setCurDateParam(curDateParam);
//                item.setTotal(total);
//                result.add(item);
            }
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load Params", e);
        }

//        Param param1 = new Param(1,"Par1", "Techproc1", "KG");
//        List<Value> valueList = new ArrayList<>();
//        float total = 0;
//        for (LocalDateTime localDateTime : dateList) {
//            Value value = loadValue(objId, param1.getParId(), localDateTime, ds);
//            System.out.println(value);
//            valueList.add(value);
//            total = total + value.getValue();
//        }
//        param1.setTotal(total);
//        param1.setCurDateParam(valueList);
//        Param param2 = new Param(2,"Par2", "Techproc2", "M");
//        List<Value> valueList2 = new ArrayList<>();
//        float tota2 = 0;
//        for (LocalDateTime localDateTime : dateList) {
//            Value value = loadValue(objId, param2.getParId(), localDateTime, ds);
//            valueList2.add(value);
//            tota2 = tota2 + value.getValue();
//        }
//        param2.setTotal(tota2);
//        param2.setCurDateParam(valueList2);
//        Param param3 = new Param(3,"Par3", "Techproc3", "MgW");
//        List<Value> valueList3 = new ArrayList<>();
//        float tota3 = 0;
//        for (LocalDateTime localDateTime : dateList) {
//            Value value = loadValue(objId, param3.getParId(), localDateTime, ds);
//            valueList3.add(value);
//            tota3 = tota3 + value.getValue();
//        }
//        param3.setTotal(tota3);
//        param3.setCurDateParam(valueList3);
//
//
//        result.add(param1);
//        result.add(param2);
//        result.add(param3);

        return result;
    }

    public int loadParId(int maskId, int objId, DataSource ds) {
        try (Connection connection = ds.getConnection();
             PreparedStatement stm = connection.prepareStatement(GET_PAR_ID)) {
            stm.setInt(1, maskId);
            stm.setInt(2, objId);

            ResultSet res = stm.executeQuery();

            if (res.next()) {
                int result = res.getInt("par_id");
                System.out.println("parId result "+result);
                return result;
            }

        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load parId", e);
        }
        return Integer.valueOf(null);
    }
    public List<Value> loadValue(int objId, int parId, String parCateg, String difInt, LocalDateTime begDate, LocalDateTime endDate, String status, DataSource ds) {
        String loadValue = "";
        List <Value> result = new ArrayList<>();

        System.out.println("loadValue status" + status);
//        LocalDateTime dateForSQL = date;
        switch (status) {
            case ("D"):
                loadValue = LOAD_H_VALUE;
//                dateForSQL = dateForSQL.minusHours(1);
                break;
            case ("M"):
                loadValue = LOAD_D_VALUE;

                break;
            case ("Y"):
                loadValue = LOAD_M_VALUE;
                break;
        }
        try (Connection connection = ds.getConnection();
             PreparedStatement stm = connection.prepareStatement(loadValue)) {
            stm.setInt(1, objId);
            stm.setInt(2, parId);
            stm.setString(3, parCateg);
            stm.setString(4, difInt);
            stm.setTimestamp(4, Timestamp.valueOf(begDate));
            stm.setTimestamp(5, Timestamp.valueOf(endDate));
            System.out.println("Что передаем loadValue " + "objId " + objId + " parId "+parId+ " parCateg " + parCateg+ " difInt " + difInt + " begDate "+ begDate + " endDate " + endDate);

            ResultSet res = stm.executeQuery();

            while (res.next()) {
                Value value = new Value(res.getString("par_value"), res.getString("color"));
                if (value.getColor() != null) {
                    if (value.getColor().isEmpty()) {
                        value.setColor(null);
                    }
                }
                result.add(value);
            }
                return result;

        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load Value", e);
        }
        return result;
//        float random = (float) (Math.random()*100);
//        if (test < 10) {
//            test++;
//            return new Value("value", null);
//
//        }
//        return new Value("value", "A020F0");
    }

    public String loadTotal(int objId, int parId, DataSource ds) {
//        try (Connection connection = ds.getConnection();
//             PreparedStatement stm = connection.prepareStatement(LOAD_TOTAL)) {
//            stm.setInt(1, objId);
//            stm.setInt(2, parId);
//            ResultSet res = stm.executeQuery();
//            if (res.next() && (res.getString(1) != null)) {
//                return res.getString(1);
//            }
//        } catch (SQLException e) {
//            LOGGER.log(Level.WARNING, "error load total ", e);
//        }
//        return null;
        return "Total";
    }


    public void percent(int repId, BigDecimal percent, DataSource ds) {

        try (Connection connection = ds.getConnection();
             PreparedStatement stm = connection.prepareStatement(PERCENT)) {
            stm.setInt(1, repId);
            stm.setBigDecimal(2, percent);
            stm.executeUpdate();
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "update percent error ", e);
        }
    }

    public String interrupted(int repId, DataSource ds) {
        try (Connection connection = ds.getConnection();
             PreparedStatement stm = connection.prepareStatement(INTERRUPTED)) {
            stm.setInt(1, repId);
            ResultSet res = stm.executeQuery();
            if (res.next() && (res.getString(1) != null)) {
                return res.getString(1);
            }
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load interrupt status ", e);
        }
        return null;
//        return "Q";
    }

    public void saveReportIntoTable(SXSSFWorkbook p_wb, int repId, DataSource dsRW) {
        try (Connection connection = dsRW.getConnection();
             PreparedStatement delStmt = connection.prepareStatement(DELSQL);
             PreparedStatement finStmt = connection.prepareStatement(FINSQL)) {

// Блоб удаляем в любом случае.
            delStmt.setLong(1, repId);
            delStmt.executeUpdate();
            delStmt.close();

//            if  (!interrupted(repId, dsRW).equals("Q")) {
// Меняем статус и записываем блоб только если отчет не был прерван

                addBlob(p_wb, repId, dsRW);
                finStmt.setLong(1, repId);
                // выполняем запись. Ура, товарищи.
                finStmt.executeUpdate();
//            }

        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "SQL saving blob error ", e);
        }
    }

    public void addBlob(SXSSFWorkbook wb, int repId, DataSource ds) {
        System.out.println("addBlob");
        try (Connection connection = ds.getConnection();
             CallableStatement stm = connection.prepareCall(SQL)) {


            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            try {
                wb.write(bos);
            } catch (IOException e) {
                LOGGER.log(Level.WARNING, "saving blob error ", e);
            }

            byte[] bytes = bos.toByteArray();

            stm.setLong(1, repId);
            stm.setBytes(2, bytes);
            stm.registerOutParameter(3, Types.INTEGER);

            stm.executeUpdate();
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "add blob error ", e);
        }
    }

    public SXSSFWorkbook printReportTemp (int p_Rep_Id, DataSource dsR, DataSource dsRW) throws IOException, SQLException, ParseException, DecoderException {
        SXSSFWorkbook wb = new SXSSFWorkbook();
        SXSSFSheet sh = wb.createSheet("Отчет находится в стадии разработки");

        SXSSFRow row_0 = sh.createRow(0);
        SXSSFCell cell = row_0.createCell(4);
        cell.setCellValue("Отчет находится в стадии разработки");

        LOGGER.log(Level.INFO, "Report footer created {0}", p_Rep_Id);

        return wb;
    }

    ///////////////////////////////////////////  Определение стилей тут
//  Стиль ячеек с данными. "LEFT" - выравнивает по левому краю, "RIGHT" - по правому

    private  CellStyle setCellNow(SXSSFWorkbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);


        Font headerFont = wb.createFont();
        headerFont.setBold(false);
        headerFont.setFontName("Times New Roman");
        headerFont.setFontHeightInPoints((short) 12);

        style.setFont(headerFont);

        return style;
    }

    private CellStyle setUsualStyle(SXSSFWorkbook p_wb, String p_LeftRight) {
        CellStyle style = p_wb.createCellStyle();

        if (p_LeftRight.equals("LEFT")) style.setAlignment(HorizontalAlignment.LEFT);
        if (p_LeftRight.equals("RIGHT")) style.setAlignment(HorizontalAlignment.RIGHT);
        if (p_LeftRight.equals("CENTER")) style.setAlignment(HorizontalAlignment.CENTER);

        style.setWrapText(false);

        style.setBorderTop(BorderStyle.NONE);
        style.setBorderLeft(BorderStyle.NONE);
        style.setBorderRight(BorderStyle.NONE);
        style.setBorderBottom(BorderStyle.NONE);

        Font headerFont = p_wb.createFont();
        headerFont.setBold(false);
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short) 10);

        style.setFont(headerFont);


        return style;
    }

//    private  CellStyle setUsualYellowStyle(SXSSFWorkbook p_wb, String p_LeftRight) throws DecoderException {
//        CellStyle style = p_wb.createCellStyle();
//
////        SSFPalette palette = p_wb.getCustomPalette();
////        XSSFColor color = (XSSFColor) palette.findColor((byte) 0, (byte) 255, (byte) 255);
////        XSSFColor color = getColor(XSSFColor.YELLOW.index);
//
////        style.setFillForegroundColor(IndexedColors.YELLOW.index);
////        style.setFillForegroundColor((Color) java.awt.Color.decode("00FF00"));//decode hex to rgb
////        style.setFillForegroundColor(new XSSFColor(new java.awt.Color(ebeb0f)));
////        String rgbS = "00FF00";
////        byte [] rgbB = Hex.decodeHex(rgbS);
////        System.out.println("rgbB "+rgbB);
////        XSSFColor color = new XSSFColor(rgbB);
////        System.out.println("color "+color);
//
////        byte [] rgbB = new byte[]{0, (byte) 128,0};
////        XSSFColor color = new XSSFColor(rgbB);
//
////
////        byte[] rgb = new byte[3];
////        rgb[0] = (byte) Integer.valueOf(rgbS.substring(0,2),16) 0,128,0
//
////        style.setFillForegroundColor(new XSSFColor((IndexedColorMap) java.awt.Color.decode("00FF00")));
//
//
////        style.setFillForegroundColor(color);
//
//
//////
////        style.setFillForegroundColor(color);
////        java.awt.Color green = new java.awt.Color(Integer.valueOf(rgbS.substring(1, 3),16),
////                Integer.valueOf(rgbS.substring(3, 5), 16), Integer.valueOf(rgbS.substring(5, 7), 16));
//
////        XSSFColor color = new XSSFColor(green);
//
////        style.setFillForegroundColor(color.getIndex());
////        style.setFillForegroundColor(new XSSFColor(new java.awt.Color (Integer.valueOf(rgbS.substring(1, 3),16),
////                Integer.valueOf(rgbS.substring(3, 5), 16), Integer.valueOf(rgbS.substring(5, 7), 16))), new DefaultIndexedColorMap());
//
//        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//
//
//        if (p_LeftRight.equals("LEFT")) style.setAlignment(HorizontalAlignment.LEFT);
//        if (p_LeftRight.equals("RIGHT")) style.setAlignment(HorizontalAlignment.RIGHT);
//        if (p_LeftRight.equals("CENTER")) style.setAlignment(HorizontalAlignment.CENTER);
//
//        style.setWrapText(false);
//
//        style.setBorderTop(BorderStyle.NONE);
//        style.setBorderLeft(BorderStyle.NONE);
//        style.setBorderRight(BorderStyle.NONE);
//        style.setBorderBottom(BorderStyle.NONE);
//
//        Font headerFont = p_wb.createFont();
//        headerFont.setBold(false);
//        headerFont.setFontName("Arial");
//        headerFont.setFontHeightInPoints((short) 10);
//
//        style.setFont(headerFont);
//
//        return style;
//    }



    //  Стиль заголовка
    private  CellStyle setHeaderStyle(SXSSFWorkbook p_wb) {

        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(false);

        Font headerFont = p_wb.createFont();
        headerFont.setBold(true);
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short) 16);

        style.setFont(headerFont);

        return style;

    }

    private  CellStyle setHeaderStyleNoBold(SXSSFWorkbook p_wb) {

        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(false);

        Font headerFont = p_wb.createFont();
        headerFont.setBold(false);
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short) 16);

        style.setFont(headerFont);

        return style;

    }

//    private  CellStyle setNotBoldHeaderStyle(SXSSFWorkbook p_wb) {
//
//        CellStyle style = p_wb.createCellStyle();
//        style.setAlignment(HorizontalAlignment.LEFT);
//        style.setWrapText(false);
//
//        style.setBorderTop(BorderStyle.NONE);
//        style.setBorderLeft(BorderStyle.NONE);
//        style.setBorderRight(BorderStyle.NONE);
//        style.setBorderBottom(BorderStyle.NONE);
//
//        Font headerFont = p_wb.createFont();
//        headerFont.setBold(false);
//        headerFont.setFontName("Arial");
//        headerFont.setFontHeightInPoints((short) 14);
//
//        style.setFont(headerFont);
//
//        return style;
//
//    }

    private  CellStyle setBoldUpStyle(SXSSFWorkbook p_wb) {

        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setWrapText(false);

        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.NONE);
        style.setBorderRight(BorderStyle.NONE);
        style.setBorderBottom(BorderStyle.NONE);

        Font headerFont = p_wb.createFont();
        headerFont.setBold(true);
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short) 10);

        style.setFont(headerFont);

        return style;
    }



    //  Стиль шапки
    private  CellStyle setCappyStyle(SXSSFWorkbook p_wb, String p_BorderType) {
        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        if ("Up".equals(p_BorderType)) {
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            style.setBorderBottom(BorderStyle.NONE);
        } else if ("Dn".equals(p_BorderType) ) {
            style.setBorderTop(BorderStyle.NONE);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            style.setBorderBottom(BorderStyle.THIN);
        } else if ("UpDn".equals(p_BorderType) ) {
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            style.setBorderBottom(BorderStyle.THIN);
        } else if ("None".equals(p_BorderType)) {
            style.setBorderTop(BorderStyle.NONE);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
            style.setBorderBottom(BorderStyle.NONE);
        }

        Font cappyFont = p_wb.createFont();

        cappyFont.setBold(true);
        cappyFont.setFontName("Arial");
        cappyFont.setFontHeightInPoints((short) 10);

        style.setFont(cappyFont);

        return style;
    }
}