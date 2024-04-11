package ru.tecon.uniRep;

import org.apache.commons.codec.DecoderException;
import org.apache.commons.codec.binary.Hex;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import ru.tecon.uniRep.model.*;

import javax.sql.DataSource;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.sql.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

public class UniRep {

    private static final Logger LOGGER = Logger.getLogger(UniRep.class.getName());
    private static final String LOAD_REP_TYPE = "SELECT * FROM td_rep.get_rep_info (?)";
    private static final String LOAD_OBJECT = "SELECT obj_id, obj_name, obj_address FROM td_rep.get_full_objects_from_mask(?)";
    private static final String LOAD_PARAMS = "SELECT * FROM td_rep.get_common_obj_par_values (?, ?, ?, ?, ?)";
    private static final String PERCENT = "call td_rep.set_run_procent(?, ?)";
    private static final String INTERRUPTED = "SELECT STATUS FROM td_rep.get_rep_info (?)";
    private static final String DELSQL = "delete from admin.REP_BLOB where REP_id = ?";
    private static final String SQL = "call td_rep.ins_rep_blob(?,?,?::integer)";
    private static final String FINSQL = "update admin.REP_REPORT set STATUS = 'F' where id = ?";

    private DataSource dsR;
    public void setDsR(DataSource dsR) {
        this.dsR = dsR;
    }

    private DataSource dsRW;

    public void setDsRW(DataSource dsRW) {
        this.dsRW = dsRW;
    }


    private int Rows;

    //init метод при создании отчета
    public static void makeReport (int p_Rep_Id, DataSource dsR, DataSource dsRW) throws InterruptedException {
        long currentTime = System.nanoTime();
        LOGGER.log(Level.INFO, "start make report {0}", p_Rep_Id);

        Thread.sleep(1000);

        UniRep mr = new UniRep();
        mr.setDsR(dsR);
        mr.setDsRW(dsRW);
        SXSSFWorkbook w;
        try {
            w = mr.printReport(p_Rep_Id);
            mr.saveReportIntoTable (w, p_Rep_Id, dsRW);
//            mr.saveReportIntoFile(w, "C:\\abc\\MOEK_64636.xlsx");
        } catch (IOException | SQLException | ParseException | DecoderException e) {
            LOGGER.log(Level.WARNING, "makeReport error", e);
            e.printStackTrace();
        }
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
    public SXSSFWorkbook printReport (int p_Rep_Id) throws IOException, SQLException, ParseException, DecoderException {
        SXSSFWorkbook wb = new SXSSFWorkbook();
        SXSSFSheet sh = wb.createSheet("Отчет");
        int begRow = 7;  // строка в екселе, с которой начинается собственно отчет.

        RepType repType = loadRepType(p_Rep_Id, dsR);

        long cols = 0;
        List<LocalDateTime> dateList = new ArrayList<>();
        LocalDateTime localDateTemp = repType.getBeg();

        dateList.add(localDateTemp);
        //-- Считаем количество дней между датами
        if (repType.getEnd().getYear() == repType.getBeg().getYear() && repType.getEnd().getMonth() == repType.getBeg().getMonth()
                && repType.getEnd().getDayOfMonth() == repType.getBeg().getDayOfMonth()) {

            cols++;
        } else {
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
        }
        if (repType.getInterval().equals("H")) {
            cols = cols*24;
        }
        // Пожалуй, наполню-ка я лист отдельным методом. И сначала заполняем его, чтобы узнать Rows. Cols

        // Колонок - 4 первых и колонка с датами

        CellStyle headerStyle = setHeaderStyle(wb);
        CellStyle headerStyleNoBold = setHeaderStyleNoBold(wb);
        CellStyle nowStyle = setCellNow (wb);
        CellStyle tableHeaderStyle = setTableHeaderStyle(wb);


        // Устанавливаем ширины колонок. В конце мероприятия
        sh.setColumnWidth(0, 6 * 256);
        sh.setColumnWidth(1, 39 * 256);
        sh.setColumnWidth(2, 14 * 256);
        sh.setColumnWidth(3, 14 * 256);
        sh.setColumnWidth(4, 14 * 256);

        for (int i = 5; i < cols+5; i++) {
            sh.setColumnWidth(i, 10 * 256);
        }

        SXSSFRow row_1 = sh.createRow(0);
        row_1.setHeight((short) 435);
        SXSSFCell cell_1_1 = row_1.createCell(0);
        cell_1_1.setCellValue("ПАО \"МОЭК\": АС \"ТЕКОН - Диспетчеризация\"");
        CellRangeAddress title = new CellRangeAddress(0, 0, 0, 4);
        sh.addMergedRegion(title);
        cell_1_1.setCellStyle(headerStyle);

        SXSSFRow row_2 = sh.createRow(1);
        row_2.setHeight((short) 435);
        SXSSFCell cell_2_1 = row_2.createCell(0);
        cell_2_1.setCellValue("Многофункциональный отчет");
        CellRangeAddress formName = new CellRangeAddress(1, 1, 0, 4);
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
        endFormatted = repType.getEnd();
        String stringEnd = endFormatted.format(formatter);
        cell_3_1.setCellValue("за период: " + stringBeg + " - " + stringEnd);
        CellRangeAddress datePer = new CellRangeAddress(2, 2, 0, 4);
        sh.addMergedRegion(datePer);
        cell_3_1.setCellStyle(headerStyleNoBold);

        //указываем интервал
        SXSSFRow row_4 = sh.createRow(3);
        row_4.setHeight((short) 435);
        SXSSFCell cell_4_1 = row_4.createCell(0);
        String interval = "";
        switch (repType.getInterval()) {
            case ("H"):
                interval = "Часовой";
                break;
            case ("D"):
                interval = "Дневной";
                break;
        }
        cell_4_1.setCellValue("Интервал: " + interval);
        CellRangeAddress intervalR = new CellRangeAddress(3, 3, 0, 4);
        sh.addMergedRegion(intervalR);
        cell_4_1.setCellStyle(headerStyleNoBold);

        // Печатаем время формирования отчета
        String now = new SimpleDateFormat("dd.MM.yyyy HH:mm").format(new Date());
        SXSSFRow row_5 = sh.createRow(4);
        row_5.setHeight((short) 435);
        SXSSFCell cell_5_1 = row_5.createCell(0);
        cell_5_1.setCellStyle(nowStyle);
        cell_5_1.setCellValue("Отчет сформирован  " + now);
        CellRangeAddress nowDone = new CellRangeAddress(4, 4, 0, 4);
        sh.addMergedRegion(nowDone);

        // Прекрасно, Заголовок сделали. Готовим шапку.
        SXSSFRow row_7 = sh.createRow(6);
        SXSSFCell cell_7_0 = row_7.createCell(0);
        cell_7_0.setCellStyle(tableHeaderStyle);
        cell_7_0.setCellValue("№");

        SXSSFCell cell_7_1 = row_7.createCell(1);
        cell_7_1.setCellStyle(tableHeaderStyle);
        cell_7_1.setCellValue("Объект/Параметр");

        SXSSFCell cell_7_2 = row_7.createCell(2);
        cell_7_2.setCellStyle(tableHeaderStyle);
        cell_7_2.setCellValue("Тех.Проц.");

        SXSSFCell cell_7_3 = row_7.createCell(3);
        cell_7_3.setCellStyle(tableHeaderStyle);
        cell_7_3.setCellValue("Ед.Изм.");

        SXSSFCell cell_7_4 = row_7.createCell(4);
        cell_7_4.setCellStyle(tableHeaderStyle);
        cell_7_4.setCellValue("Итого:");

        if (repType.getInterval().equals("H")) {
            SXSSFRow row_8 = sh.createRow(7);
            begRow++;
            CellRangeAddress num = new CellRangeAddress(6, 7, 0, 0);
            sh.addMergedRegion(num);
            CellRangeAddress borderForNum = new CellRangeAddress(6, 7, 0, 0);
            RegionUtil.setBorderBottom(BorderStyle.THICK, borderForNum, sh);
            RegionUtil.setBorderTop(BorderStyle.THICK, borderForNum, sh);
            RegionUtil.setBorderLeft(BorderStyle.THICK, borderForNum, sh);
            RegionUtil.setBorderRight(BorderStyle.THICK, borderForNum, sh);

            CellRangeAddress objPar = new CellRangeAddress(6, 7, 1, 1);
            sh.addMergedRegion(objPar);
            CellRangeAddress borderForObjPar = new CellRangeAddress(6, 7, 1, 1);
            RegionUtil.setBorderBottom(BorderStyle.THICK, borderForObjPar, sh);
            RegionUtil.setBorderTop(BorderStyle.THICK, borderForObjPar, sh);
            RegionUtil.setBorderLeft(BorderStyle.THICK, borderForObjPar, sh);
            RegionUtil.setBorderRight(BorderStyle.THICK, borderForObjPar, sh);

            CellRangeAddress techProc = new CellRangeAddress(6, 7, 2, 2);
            sh.addMergedRegion(techProc);
            CellRangeAddress borderForTechProc = new CellRangeAddress(6, 7, 2, 2);
            RegionUtil.setBorderBottom(BorderStyle.THICK, borderForTechProc, sh);
            RegionUtil.setBorderTop(BorderStyle.THICK, borderForTechProc, sh);
            RegionUtil.setBorderLeft(BorderStyle.THICK, borderForTechProc, sh);
            RegionUtil.setBorderRight(BorderStyle.THICK, borderForTechProc, sh);

            CellRangeAddress unit = new CellRangeAddress(6, 7, 3, 3);
            sh.addMergedRegion(unit);
            CellRangeAddress borderForUnit = new CellRangeAddress(6, 7, 3, 3);
            RegionUtil.setBorderBottom(BorderStyle.THICK, borderForUnit, sh);
            RegionUtil.setBorderTop(BorderStyle.THICK, borderForUnit, sh);
            RegionUtil.setBorderLeft(BorderStyle.THICK, borderForUnit, sh);
            RegionUtil.setBorderRight(BorderStyle.THICK, borderForUnit, sh);

            CellRangeAddress total = new CellRangeAddress(6, 7, 4, 4);
            sh.addMergedRegion(total);
            CellRangeAddress borderForTotal = new CellRangeAddress(6, 7, 4, 4);
            RegionUtil.setBorderBottom(BorderStyle.THICK, borderForTotal, sh);
            RegionUtil.setBorderTop(BorderStyle.THICK, borderForTotal, sh);
            RegionUtil.setBorderLeft(BorderStyle.THICK, borderForTotal, sh);
            RegionUtil.setBorderRight(BorderStyle.THICK, borderForTotal, sh);
            for (int i = 0; i < dateList.size(); i++) {
                SXSSFCell dateCell = row_7.createCell(i*24 + 5);
                dateCell.setCellStyle(tableHeaderStyle);
                String curDateString = dateList.get(i).format(formatter);
                dateCell.setCellValue(curDateString);
                CellRangeAddress date = new CellRangeAddress(6, 6, i*24+5, i*24+28);
                sh.addMergedRegion(date);
                CellRangeAddress borderForDate = new CellRangeAddress(6, 6, i*24+5, i*24+28);
                RegionUtil.setBorderBottom(BorderStyle.THICK, borderForDate, sh);
                RegionUtil.setBorderTop(BorderStyle.THICK, borderForDate, sh);
                RegionUtil.setBorderLeft(BorderStyle.THICK, borderForDate, sh);
                RegionUtil.setBorderRight(BorderStyle.THICK, borderForDate, sh);

                for (int j = 0; j < 24; j++) {
                    SXSSFCell hourCell = row_8.createCell(i*24 + 5 + j);
                    hourCell.setCellStyle(tableHeaderStyle);
                    if (j<9) {
                        hourCell.setCellValue("0" + (j + 1) + " ч.");
                    } else if (j == 23) {
                        hourCell.setCellValue("00 ч.");
                    } else {
                        hourCell.setCellValue((j+1) + " ч.");
                    }
                }
            }
            sh.createFreezePane(5, 8);

        } else {
            sh.createFreezePane(5, 7);
        }

        // декларируем переменные для шапки
        SXSSFCell cell;

        int i = 0;
        for (LocalDateTime curDateLDT : dateList) {
            String curDateS = String.valueOf(curDateLDT);

            if ("D".equals(repType.getInterval())) {
                curDateS = curDateS.replace('T', ' ');
                cell = row_7.createCell(i + 5);
                cell.setCellStyle(tableHeaderStyle);
                curDateS = curDateS.substring(8, 10) + "." + curDateS.substring(5, 7);
                cell.setCellValue(curDateS);
            }
            i++;
        }

        LOGGER.log(Level.INFO, "Report head created {0}", p_Rep_Id);
        // Отлично. Заголовок и шапка сделаны. Идем по таблице, создаем и заполняем ячейки.

        fillSheet(wb, p_Rep_Id, begRow, repType.getMaskId(), repType.getBeg(), repType.getEnd(), repType.getInterval(), dateList);


        LOGGER.log(Level.INFO, "Report body created {0}", p_Rep_Id);

        return wb;
    }

    private void fillSheet (SXSSFWorkbook wb, int p_Rep_Id, int begRow, int maskId, LocalDateTime begDate, LocalDateTime endDate, String interval, List<LocalDateTime> dateList) throws DecoderException {
        SXSSFSheet sh = wb.getSheetAt(0);
        CellStyle cellBoldStyle = setCellBoldStyle(wb);
        CellStyle cellNoBoldStyle = setCellNoBoldStyle(wb);

        // Заполняем лист значениями, взятыми из таблицы
        List<ReportObject> objects = loadObjects(maskId, dsR);
        if (!objects.isEmpty()) {
            Rows = begRow;
            int objNum = 1;
            BigDecimal percentage = new BigDecimal(0);
            double size = objects.size();
            double iterationPercent = 100/size;
            BigDecimal iterationPercentBD = new BigDecimal(iterationPercent);
            HashMap<String, CellStyle> colors = new HashMap<>();

            for (ReportObject object : objects) {
                if (!interrupted(p_Rep_Id , dsR).equals("Q")) {
                    object.setParamList(loadParams(maskId, object.getObjId(), begDate, endDate, interval, dsR));
                    SXSSFRow row = sh.createRow(Rows);
                    SXSSFCell objNumCell = row.createCell(0);
                    objNumCell.setCellValue(objNum);
                    objNumCell.setCellStyle(cellBoldStyle);
                    objNum++;
                    SXSSFCell objNameCell = row.createCell(1);
                    objNameCell.setCellValue(object.getName());
                    objNameCell.setCellStyle(cellBoldStyle);

                    SXSSFCell objAddrCell = row.createCell(2);
                    objAddrCell.setCellValue(object.getAddres());
                    objAddrCell.setCellStyle(cellBoldStyle);
                    CellRangeAddress address = new CellRangeAddress(Rows, Rows, 2, 4);
                    sh.addMergedRegion(address);
                    CellRangeAddress borderForTotal = new CellRangeAddress(Rows, Rows, 2, 4);
                    RegionUtil.setBorderBottom(BorderStyle.THIN, borderForTotal, sh);
                    RegionUtil.setBorderTop(BorderStyle.THICK, borderForTotal, sh);
                    RegionUtil.setBorderLeft(BorderStyle.THIN, borderForTotal, sh);
                    RegionUtil.setBorderRight(BorderStyle.THIN, borderForTotal, sh);

                    if ("D".equals(interval)) {
                        for (int i = 5; i < dateList.size()+5; i++) {
                            SXSSFCell cell = row.createCell(i);
                            cell.setCellStyle(cellBoldStyle);
                        }
                    } else {
                        for (int i = 5; i < (dateList.size()*24)+5; i++) {
                            SXSSFCell cell = row.createCell(i);
                            cell.setCellStyle(cellBoldStyle);
                        }
                    }

                    Rows++;
                    for (Param param : object.getParamList()) {
                        SXSSFRow parRow = sh.createRow(Rows);
                        SXSSFCell emptyNumCell = parRow.createCell(0);
                        emptyNumCell.setCellStyle(cellNoBoldStyle);
                        SXSSFCell parNameCell = parRow.createCell(1);
                        parNameCell.setCellValue(param.getParName());
                        parNameCell.setCellStyle(cellNoBoldStyle);
                        SXSSFCell techProcCell = parRow.createCell(2);
                        techProcCell.setCellValue(param.getTecProc());
                        techProcCell.setCellStyle(cellNoBoldStyle);
                        SXSSFCell unitCell = parRow.createCell(3);
                        unitCell.setCellValue(param.getUnits());
                        unitCell.setCellStyle(cellNoBoldStyle);
                        SXSSFCell totalCell = parRow.createCell(4);
                        totalCell.setCellValue(param.getTotal());
                        totalCell.setCellStyle(cellNoBoldStyle);
                        Rows++;
                        int colNum = 0;
                        for (Value value : param.getCurDateParam()) {
                            SXSSFCell valueCell = parRow.createCell(5 + colNum);
                            valueCell.setCellValue(value.getValue());

                            if (value.getColor() != null) {
                                if (colors.containsKey(value.getColor())) {
                                    valueCell.setCellStyle(colors.get(value.getColor()));
                                } else {
                                    CellStyle cellColoredStyle = setCellNoBoldStyle(wb);
                                    String rgbS = value.getColor();
                                    byte [] rgbB = Hex.decodeHex(rgbS);
                                    XSSFColor color = new XSSFColor(rgbB, null);
                                    cellColoredStyle.setFillForegroundColor(color);
                                    cellColoredStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                                    colors.put(value.getColor(), cellColoredStyle);
                                    valueCell.setCellStyle(cellColoredStyle);
                                }
                            } else {
                                valueCell.setCellStyle(cellNoBoldStyle);
                            }
                            colNum++;
                        }
                    }
                    percentage = percentage.add(iterationPercentBD).setScale(3, RoundingMode.DOWN);
                percent(p_Rep_Id, percentage, dsRW);
//                    System.out.println(percentage);
                } else {
                    break;
                }
            }
        } else {
            if ("H".equals(interval)) {
                SXSSFRow row_9 = sh.createRow(8);
                SXSSFCell cell_9_1 = row_9.createCell(1);
                cell_9_1.setCellValue("Не выбран ни один объект");
            } else {
                SXSSFRow row_8 = sh.createRow(7);
                SXSSFCell cell_8_1 = row_8.createCell(1);
                cell_8_1.setCellValue("Не выбран ни один объект");
            }

        }
    }

//    private int saveReportIntoFile (SXSSFWorkbook workbook, String file) {
//        FileOutputStream fos;
//        System.out.println("saveReportIntoFile");
//        int res = 0;
//
//        try {
//            fos = new FileOutputStream(file);
//            workbook.write(fos);
//
//        } catch (IOException e) {
//            e.printStackTrace();
//            System.out.println("Cant create file");
//            System.out.println(e.getMessage());
//            res = 1;
//        }
//
//        System.out.println("ok report");
//        return res;
//    }

    public RepType loadRepType(int repId, DataSource ds) {
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
                return result;
            }

        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load Rep Type", e);
        }
        return result;
    }

    public List<ReportObject> loadObjects(int maskId, DataSource ds) {
        List<ReportObject> result = new ArrayList<>();
        try (Connection connect = ds.getConnection();
             PreparedStatement stm = connect.prepareStatement(LOAD_OBJECT)) {
            stm.setInt(1, maskId);
            ResultSet res = stm.executeQuery();
            while (res.next()) {
                ReportObject item = new ReportObject(res.getInt("obj_id"), res.getString("obj_name"),
                        res.getString("obj_address"));


                result.add(item);
            }
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load Object", e);
        }
        return result;
    }


    public List<Param> loadParams(int maskId, int objId, LocalDateTime begDate, LocalDateTime endDate, String status, DataSource ds) {
        List<Param> result = new ArrayList<>();

        try (Connection connect = ds.getConnection();
             PreparedStatement stm = connect.prepareStatement(LOAD_PARAMS)) {

            stm.setInt(1, maskId);
            stm.setInt(2, objId);
            stm.setString(3, status);
            stm.setTimestamp(4, Timestamp.valueOf(begDate));
            stm.setTimestamp(5, Timestamp.valueOf(endDate));
            ResultSet res = stm.executeQuery();
            while (res.next()) {
                if (!result.isEmpty() && !result.get(result.size()-1).getStatAggr().equals(res.getString("stataggr"))) {
                    List<Value> curDateParam = new ArrayList<>();
                    Value value = new Value(res.getString("par_value"), res.getString("color"));
                    if (value.getColor()!=null) {
                        if (value.getColor().isEmpty()) {
                            value.setColor(null);
                        }
                    }

                    curDateParam.add(value);

                    Param param = new Param(null, null, null, null, null, null, null);
                    param.setParId(res.getInt("par_id"));
                    param.setParName(res.getString("par_memo"));
                    param.setTecProc(res.getString("techproc_type_char"));
                    param.setUnits(res.getString("short_measure"));
                    param.setStatAggr(res.getString("stataggr"));
                    param.setTotal(res.getString("itog"));
                    param.setCurDateParam(curDateParam);

                    result.add(param);
                } else if (!result.isEmpty() && result.get(result.size()-1).getParId() == res.getInt("par_id")) {
                    Value value = new Value(res.getString("par_value"), res.getString("color"));
                    if (value.getColor() != null) {
                        if (value.getColor().isEmpty()) {
                            value.setColor(null);
                        }
                    }
                    result.get(result.size() - 1).getCurDateParam().add(value);
                } else {
                    List<Value> curDateParam = new ArrayList<>();
                    Value value = new Value(res.getString("par_value"), res.getString("color"));
                    if (value.getColor()!=null) {
                        if (value.getColor().isEmpty()) {
                            value.setColor(null);
                        }
                    }

                    curDateParam.add(value);

                    Param param = new Param(null, null, null, null, null, null, null);
                    param.setParId(res.getInt("par_id"));
                    param.setParName(res.getString("par_memo"));
                    param.setTecProc(res.getString("techproc_type_char"));
                    param.setUnits(res.getString("short_measure"));
                    param.setStatAggr(res.getString("stataggr"));
                    param.setTotal(res.getString("itog"));
                    param.setCurDateParam(curDateParam);

                    result.add(param);
                }
            }
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "error load Params", e);
        }
        return result;
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
    }

    public void saveReportIntoTable(SXSSFWorkbook p_wb, int repId, DataSource dsRW) {
        try (Connection connection = dsRW.getConnection();
             PreparedStatement delStmt = connection.prepareStatement(DELSQL);
             PreparedStatement finStmt = connection.prepareStatement(FINSQL)) {

// Блоб удаляем в любом случае.
            delStmt.setLong(1, repId);
            delStmt.executeUpdate();
            delStmt.close();

            if  (!interrupted(repId, dsRW).equals("Q")) {
// Меняем статус и записываем блоб только если отчет не был прерван

                addBlob(p_wb, repId, dsRW);
                finStmt.setLong(1, repId);
                // выполняем запись. Ура, товарищи.
                percent(repId, BigDecimal.valueOf(100),dsRW);
                finStmt.executeUpdate();
            }

        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "SQL saving blob error ", e);
        }
    }

    public void addBlob(SXSSFWorkbook wb, int repId, DataSource ds) {
        try (Connection connection = ds.getConnection();
             CallableStatement stm = connection.prepareCall(SQL);
             ByteArrayOutputStream bos = new ByteArrayOutputStream()) {

                wb.write(bos);

            stm.setLong(1, repId);
            stm.setBytes(2, bos.toByteArray());
            stm.registerOutParameter(3, Types.INTEGER);

            stm.executeUpdate();
        } catch (IOException e) {
            LOGGER.log(Level.WARNING, "saving blob error ", e);
        } catch (SQLException e) {
            LOGGER.log(Level.WARNING, "add blob error ", e);
        } finally {
            try {
                wb.close();
            } catch (IOException e) {
                LOGGER.log(Level.WARNING, "wb close error ", e);
            }
        }
    }

    ///////////////////////////////////////////  Определение стилей тут

    //  Стиль заголовка жирный
    private  CellStyle setHeaderStyle(SXSSFWorkbook p_wb) {

        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(false);

        Font headerFont = p_wb.createFont();
        headerFont.setBold(true);
        headerFont.setFontName("Times New Roman");
        headerFont.setFontHeightInPoints((short) 16);

        style.setFont(headerFont);

        return style;
    }

    //  Стиль заголовка не жирный
    private  CellStyle setHeaderStyleNoBold(SXSSFWorkbook p_wb) {

        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(false);

        Font headerFontNoBold = p_wb.createFont();
        headerFontNoBold.setBold(false);
        headerFontNoBold.setFontName("Times New Roman");
        headerFontNoBold.setFontHeightInPoints((short) 16);

        style.setFont(headerFontNoBold);

        return style;
    }

    //стиль для даты создания отчета
    private  CellStyle setCellNow(SXSSFWorkbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);


        Font nowFont = wb.createFont();
        nowFont.setBold(false);
        nowFont.setFontName("Times New Roman");
        nowFont.setFontHeightInPoints((short) 12);

        style.setFont(nowFont);

        return style;
    }

    //  Стиль шапки таблицы
    private  CellStyle setTableHeaderStyle(SXSSFWorkbook p_wb) {
        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(false);

        style.setBorderTop(BorderStyle.THICK);
        style.setBorderLeft(BorderStyle.THICK);
        style.setBorderRight(BorderStyle.THICK);
        style.setBorderBottom(BorderStyle.THICK);

        Font tableHeaderFont = p_wb.createFont();

        tableHeaderFont.setBold(true);
        tableHeaderFont.setFontName("Times New Roman");
        tableHeaderFont.setFontHeightInPoints((short) 12);

        style.setFont(tableHeaderFont);

        return style;
    }

    private  CellStyle setCellBoldStyle(SXSSFWorkbook p_wb) {
        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);

        style.setBorderTop(BorderStyle.THICK);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);

        Font cellBoldFont = p_wb.createFont();

        cellBoldFont.setBold(true);
        cellBoldFont.setFontName("Times New Roman");
        cellBoldFont.setFontHeightInPoints((short) 11);

        style.setFont(cellBoldFont);

        return style;
    }

    private  CellStyle setCellNoBoldStyle(SXSSFWorkbook p_wb) {
        CellStyle style = p_wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);

        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);

        Font cellNoBoldFont = p_wb.createFont();

        cellNoBoldFont.setBold(false);
        cellNoBoldFont.setFontName("Times New Roman");
        cellNoBoldFont.setFontHeightInPoints((short) 11);

        style.setFont(cellNoBoldFont);

        return style;
    }
}