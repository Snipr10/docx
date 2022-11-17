import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.*;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xddf.usermodel.text.XDDFTextBody;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import org.openxmlformats.schemas.drawingml.x2006.chart.CTDLbls;
import org.openxmlformats.schemas.drawingml.x2006.chart.STDLblPos;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.IOException;
import java.math.BigInteger;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static java.lang.Math.min;


class NewReport {
    private static  int entityOnPage = 0;
    private static int commentsLenght= 100;
    private static String format = "Times New Roman";
    static CellReference setTitleInDataSheet(XWPFChart chart, String title, int column) throws Exception {
        XSSFWorkbook workbook = chart.getWorkbook();
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(0);
        if (row == null)
            row = sheet.createRow(0);
        XSSFCell cell = row.getCell(column);
        if (cell == null)
            cell = row.createCell(column);
        cell.setCellValue(title);
        return new CellReference(sheet.getSheetName(), 0, column, true, true);
    }

    public static <T> T[] append(T[] arr, T element) {
        final int N = arr.length;
        arr = Arrays.copyOf(arr, N + 1);
        arr[N] = element;
        return arr;
    }


    public static XWPFDocument createDoc(String type, String name, String date,
                                         DataForDocx data, JSONObject jsonPosts, JSONObject jsonComments, JSONObject stat,
                                         JSONObject sex, JSONObject age, JSONObject usersJson, JSONArray jsonCity, JSONArray posts,
                                         JSONArray postsContent,JSONArray commentContent, int first_month, int first_year
    ) {
        int users = Integer.parseInt(usersJson.get("count").toString());
        NumberFormat f = NumberFormat.getInstance();

        try {
            XWPFDocument docxModel = new XWPFDocument();
            XWPFParagraph bodyParagraph = docxModel.createParagraph();
            bodyParagraph.setAlignment(ParagraphAlignment.RIGHT);
            XWPFRun paragraphConfig = bodyParagraph.createRun();
            paragraphConfig.setFontSize(22);
            paragraphConfig.setBold(true);
            paragraphConfig.setFontFamily(format);
            paragraphConfig.setText(
                    " "
            );
            paragraphConfig.addBreak();
            paragraphConfig.addBreak();
            paragraphConfig.addBreak();
            paragraphConfig.addBreak();
            paragraphConfig.addBreak();
            paragraphConfig.addBreak();
            paragraphConfig.addBreak();
            paragraphConfig.addBreak();


            XWPFParagraph bodyParagraphLenta = docxModel.createParagraph();
            bodyParagraphLenta.setAlignment(ParagraphAlignment.LEFT);
            XWPFRun paragraphConfigLenta = bodyParagraphLenta.createRun();
            paragraphConfigLenta.setFontSize(22);
            paragraphConfigLenta.setBold(true);
            paragraphConfigLenta.setFontFamily(format);
            paragraphConfigLenta.setText(
                    "Лента: "
            );

            XWPFParagraph bodyParagraphName = docxModel.createParagraph();
            bodyParagraphName.setAlignment(ParagraphAlignment.LEFT);
            XWPFRun paragraphConfigName = bodyParagraphName.createRun();
            paragraphConfigName.setFontSize(26);
            paragraphConfigName.setBold(true);
            paragraphConfigName.setFontFamily(format);
            paragraphConfigName.setText(
                    name

            );
            paragraphConfigName.addBreak();

            XWPFParagraph bodyParagraphAnalyze = docxModel.createParagraph();
            bodyParagraphAnalyze.setAlignment(ParagraphAlignment.LEFT);
            XWPFRun paragraphConfigAnalyze = bodyParagraphAnalyze.createRun();
            paragraphConfigAnalyze.setFontSize(14);
            paragraphConfigAnalyze.setFontFamily(format);
            paragraphConfigAnalyze.setText("Аналитический отчет по упоминаниям в онлайн-СМИ и соцмедиа");
            paragraphConfigAnalyze.addBreak();


            XWPFParagraph bodyParagraphDate = docxModel.createParagraph();
            bodyParagraphDate.setAlignment(ParagraphAlignment.LEFT);
            XWPFRun paragraphConfigDate = bodyParagraphDate.createRun();
            paragraphConfigDate.setFontSize(14);
            paragraphConfigDate.setFontFamily(format);
            paragraphConfigDate.setText(
                    "Период анализа: " + date
            );
            paragraphConfigDate.addBreak();
            paragraphConfigDate.addBreak();
            paragraphConfigDate.addBreak();
            paragraphConfigDate.addBreak();
            paragraphConfigDate.addBreak();
            paragraphConfigDate.addBreak();
            paragraphConfigDate.addBreak();
            paragraphConfigDate.addBreak();
            paragraphConfigDate.addBreak();
            paragraphConfigDate.addBreak();
            paragraphConfigDate.addBreak();
            paragraphConfigDate.addBreak();


            XWPFParagraph paragraph = docxModel.createParagraph();
            paragraph.setPageBreak(true);
            XWPFRun run;
            run = paragraph.createRun();
            run.setText("Оглавление");

            run.setFontSize(14);
            run.setBold(true);
            run.setFontFamily(format);

            CTP ctP = paragraph.getCTP();
            CTSimpleField toc = ctP.addNewFldSimple();
            toc.setInstr("TOC \\h");
            toc.setDirty(STOnOff.TRUE);
            addCustomHeadingStyle(docxModel, "Heading1", 1, false);
            addCustomHeadingStyle(docxModel, "Heading2", 2, true);

            XWPFParagraph bodyParagraphStatistic = docxModel.createParagraph();
            bodyParagraphStatistic.setStyle("Heading1");
            bodyParagraphStatistic.setAlignment(ParagraphAlignment.LEFT);

            XWPFRun paragraphConfigStatic = bodyParagraphStatistic.createRun();
            paragraphConfigStatic.setFontSize(14);
            paragraphConfigStatic.setBold(true);
            paragraphConfigStatic.setFontFamily(format);
            paragraphConfigStatic.setText(
                    "Базовые статистики и количество реакций пользователей на публикации"
            );

            XWPFTable table = docxModel.createTable();
            deleteBoarder(table);

            XWPFTableRow tableRowOne = table.getRow(0);
            run = tableRowOne.getCell(0).getParagraphs().get(0).createRun();
            run.setText("Совокупная аудитория, чел.");
            run.setFontSize(12);
            run.setFontFamily(format);
            XWPFTableCell cell_one = tableRowOne.addNewTableCell();
            cell_one.removeParagraph(0);
            XWPFParagraph addParagraph_one = cell_one.addParagraph();
            XWPFRun run_one = addParagraph_one.createRun();
            run_one.setFontFamily(format);
            run_one.setFontSize(12);
            run_one.setText(get_format_stng(users));

            XWPFTableRow tableRowTwo = table.createRow();
            XWPFRun run1 = tableRowTwo.getCell(0).getParagraphs().get(0).createRun();
            run1.setText("Количество источников публикаций, шт.");
            run1.setFontSize(12);
            run1.setFontFamily(format);

            XWPFTableCell cell_two = tableRowTwo.getCell(1);
            cell_two.removeParagraph(0);
            XWPFParagraph addParagraph_two = cell_two.addParagraph();
            XWPFRun run_two = addParagraph_two.createRun();
            run_two.setFontFamily(format);
            run_two.setFontSize(12);
            run_two.setText(get_format_stng(data.total_sources));


            XWPFTableRow tableRowThree = table.createRow();
            XWPFRun run2 = tableRowThree.getCell(0).getParagraphs().get(0).createRun();
            run2.setText("Количество публикаций, шт.");
            run2.setFontSize(12);
            run2.setFontFamily(format);


            XWPFTableCell cell__2 = tableRowThree.getCell(1);
            cell__2.removeParagraph(0);
            XWPFParagraph addParagraph___2 = cell__2.addParagraph();
            XWPFRun run___2 = addParagraph___2.createRun();
            run___2.setFontFamily(format);
            run___2.setFontSize(12);
            run___2.setText(get_format_stng(data.total_publication));

//            XWPFTableRow tableRowFour = table.createRow();
//            XWPFRun run3 = tableRowFour.getCell(0).getParagraphs().get(0).createRun();
//            run3.setText("Количество комментариев к публикациям, шт.");
//            tableRowFour.getCell(1).setText(String.valueOf(data.total_comment));

            XWPFTableRow tableRow4 = table.createRow();
            XWPFRun run4_1 = tableRow4.getCell(0).getParagraphs().get(0).createRun();
            run4_1.setFontSize(12);
            run4_1.setFontFamily(format);

            run4_1.setText("Количество реакций пользователей на публикации, шт.");

            XWPFTableCell cell__4 = tableRow4.getCell(1);
            cell__4.removeParagraph(0);
            XWPFParagraph addParagraph___4 = cell__4.addParagraph();
            XWPFRun run___4 = addParagraph___4.createRun();
            run___4.setFontFamily(format);
            run___4.setFontSize(12);
            run___4.setText(get_format_stng(data.total_views));


            for (int x = 0; x < table.getNumberOfRows(); x++) {
                XWPFTableRow row = table.getRow(x);
                XWPFTableCell cell0 = row.getCell(0);
                cell0.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(8000));
                XWPFTableCell cell1 = row.getCell(1);
                cell1.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1000));
                cell1.getParagraphs().get(0).setAlignment(ParagraphAlignment.RIGHT);
            }
            entityOnPage =1;
            int diagramCount = 1;
            int tableCount = 1;
            ParseData postData = getWeekData(type, (JSONArray) ((JSONObject) jsonPosts.get("total")).get("total"), first_month, first_year);
//            diagramCount = addChats(docxModel, postData.categories, postData.valuesA, String.format("Диаграмма %s Динамика количества публикаций", diagramCount),  diagramCount);

            ParseData comments = getWeekData(type, (JSONArray) (jsonComments).get("total"), first_month, first_year);
//            diagramCount= addChats(docxModel, comments.categories, comments.valuesA, String.format("Диаграмма %s Динамика количества комментариев к публикациям", diagramCount), diagramCount);
            String postDate;

            // one function
            JSONArray jsonArray;
            String[] categoriesPost = postData.categories;
            Double[] valuesAPost = postData.valuesA;
            Double[] postCommentData = new Double[]{};
            String[] categoriesComments = comments.categories;
            Double[] valuesAComments = comments.valuesA;

            double postCommentD;
            for (int i = 0; i < categoriesPost.length; i++) {
                postCommentD = 0;
                postDate = categoriesPost[i];
                if (valuesAPost[i] != 0) {
                    for (int j = 0; j < categoriesComments.length; j++) {
                        if (postDate.equals(categoriesComments[j])) {
                            if (valuesAComments[i] == 0){
                                postCommentD = 0.0;
                            } else {
                                postCommentD = Math.round(new Double(valuesAPost[i].toString()) / new Double(valuesAComments[i].toString()) * 100.0) / 100.0;
                            }
                            break;
                        }
                    }

                }
                postCommentData = append(postCommentData, postCommentD);
            }
            diagramCount = addChats(docxModel, categoriesPost, postCommentData, String.format("Диаграмма %s Динамика количества комментариев на 1 публикацию", diagramCount), diagramCount);


            JSONObject jsonPostTotal = ((JSONObject) jsonPosts.get("total"));
            Double[] variableDouble  = new Double[]{(double) getComment(jsonPostTotal, "negative"), (double) getComment(jsonPostTotal, "positive"), (double) getComment(jsonPostTotal, "netural")};
            Double allComments = variableDouble[0] + variableDouble[1] + variableDouble[2];
            for ( int i = 0; i < variableDouble.length; i++) {
                variableDouble[i] = variableDouble[i]/allComments * 100;
            }
//
//            diagramCount = addPieFormat(docxModel, new String[]{"Негативная", "Позитивная", "Нейтральная"},variableDouble,
//                    String.format("Диаграмма %s Cтатистика эмоционального окраса в публикациях", diagramCount), diagramCount, true);


            JSONArray positive = (JSONArray) (jsonPostTotal).get("positive");
            JSONArray netural = (JSONArray) (jsonPostTotal).get("netural");
            JSONArray negative = (JSONArray) (jsonPostTotal).get("negative");
            JSONArray totalComments = ((JSONArray) (jsonPostTotal).get("total"));


            double commnetsCount = 0;
//            for (Double dV: variableDouble) {
//                commnetsCount +=dV;
//            }
//            if (commnetsCount >0) {
//                DataForArea d = new DataForArea(type, totalComments, positive, netural,
//                        negative, first_month, first_year);
////                diagramCount = addArea(docxModel, d.categoriesPostType,
////                        d.valuesNegative,
////                        d.valuesPositive,
////                        d.valuesNetural,
////                        String.format("Диаграмма %s Динамика распределения публикаций по тональности", diagramCount), diagramCount);
//            }
            addParagraph(docxModel, "Статистика эмоционального окраса в публикациях");

            XWPFTable table_r = docxModel.createTable();

            int total_count =  positive.length() + netural.length() + negative.length();
            int positive_t = (Integer) ((positive.length() *100) / total_count) ;
            int netural_t= (Integer) ((netural.length() *100) / total_count) ;
            int negative_t = 100 - positive_t - netural_t ;
            // totalComments


            XWPFTableRow tableRowTwo_n = table_r.getRow(0);
            XWPFRun run1_n = tableRowTwo_n.getCell(0).getParagraphs().get(0).createRun();
            run1_n.setText("Тональность публикаций");
            run1_n.setBold(true);
            run1_n.setFontSize(12);
            run1_n.setFontFamily(format);
            XWPFTableCell cell_two_n = tableRowTwo_n.addNewTableCell();
            cell_two_n.removeParagraph(0);
            XWPFParagraph addParagraph_two_n = cell_two_n.addParagraph();
            XWPFRun run_two_n = addParagraph_two_n.createRun();
            run_two_n.setFontFamily(format);
            run_two_n.setFontSize(12);
            run_two_n.setBold(true);
            run_two_n.setText("Доля публикаций");

            XWPFTableRow tableRowThree_r = table_r.createRow();
            XWPFRun run2_r = tableRowThree_r.getCell(0).getParagraphs().get(0).createRun();
            run2_r.setText("Негативная");
            run2_r.setFontSize(12);
            run2_r.setFontFamily(format);
            XWPFTableCell cell__2_r = tableRowThree_r.getCell(1);
            cell__2_r.removeParagraph(0);
            XWPFParagraph addParagraph___2_r = cell__2_r.addParagraph();
            XWPFRun run___2_r = addParagraph___2_r.createRun();
            run___2_r.setFontFamily(format);
            run___2_r.setFontSize(12);
            run___2_r.setText(String.valueOf(netural_t) + "%");

            XWPFTableRow tableRowThree_r_1 = table_r.createRow();
            XWPFRun run2_r_1 = tableRowThree_r_1.getCell(0).getParagraphs().get(0).createRun();
            run2_r_1.setText("Позитивная");
            run2_r_1.setFontSize(12);
            run2_r_1.setFontFamily(format);
            XWPFTableCell cell__2_r_1 =tableRowThree_r_1.getCell(1);
            cell__2_r_1.removeParagraph(0);
            XWPFParagraph addParagraph___2_r_1 = cell__2_r_1.addParagraph();
            XWPFRun run___2_r_1 = addParagraph___2_r_1.createRun();
            run___2_r_1.setFontFamily(format);
            run___2_r_1.setFontSize(12);
            run___2_r_1.setText(String.valueOf(positive_t)+ "%");


            XWPFTableRow tableRowThree_r_2 = table_r.createRow();
            XWPFRun run2_r_2 = tableRowThree_r_2.getCell(0).getParagraphs().get(0).createRun();
            run2_r_2.setText("Нейтральная");
            run2_r_2.setFontSize(12);
            run2_r_2.setFontFamily(format);
            XWPFTableCell cell__2_r_2 =tableRowThree_r_2.getCell(1);
            cell__2_r_2.removeParagraph(0);
            XWPFParagraph addParagraph___2_r_2 = cell__2_r_2.addParagraph();
            XWPFRun run___2_r_2 = addParagraph___2_r_2.createRun();
            run___2_r_2.setFontFamily(format);
            run___2_r_2.setFontSize(12);
            run___2_r_2.setText(String.valueOf(negative_t)+ "%");

            XWPFTableRow tableRowThree_r_3 = table_r.createRow();
            XWPFRun run2_r_3 = tableRowThree_r_3.getCell(0).getParagraphs().get(0).createRun();
            run2_r_3.setText("Всего");
            run2_r_3.setFontSize(12);
            run2_r_3.setFontFamily(format);
            XWPFTableCell cell__2_r_3 =tableRowThree_r_3.getCell(1);
            cell__2_r_3.removeParagraph(0);
            XWPFParagraph addParagraph___2_r_3 = cell__2_r_3.addParagraph();
            XWPFRun run___2_r_3 = addParagraph___2_r_3.createRun();
            run___2_r_3.setFontFamily(format);
            run___2_r_3.setFontSize(12);
            run___2_r_3.setText("100%");

            for (int x = 0; x < table_r.getNumberOfRows(); x++) {
                XWPFTableRow row = table_r.getRow(x);
                XWPFTableCell cell0 = row.getCell(0);
                cell0.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(5000));
                row.getCell(1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4500));

//                XWPFTableCell cell1 = row.getCell(1);
//                cell1.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4500));
//                cell1.getParagraphs().get(0).setAlignment(ParagraphAlignment.RIGHT);
            }

//
//
//            XWPFTableCell cell__2_r_1 = tableRowThree_r_1.getCell(1);
//            cell__2_r.removeParagraph(0);
//            XWPFParagraph addParagraph___2_r_1 = cell__2_r_1.addParagraph();
//            XWPFRun run___2_r_1 = addParagraph___2_r_1.createRun();
//            run___2_r_1.setFontFamily(format);
//            run___2_r_1.setFontSize(12);
//            run___2_r_1.setText("12");
//            for (int x = 0; x < table_r.getNumberOfRows(); x++) {
//                XWPFTableRow row = table_r.getRow(x);
//                XWPFTableCell cell0 = row.getCell(0);
//                cell0.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4500));
//                XWPFTableCell cell1 = row.getCell(1);
//                cell1.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4500));
//                cell1.getParagraphs().get(0).setAlignment(ParagraphAlignment.RIGHT);
//            }




            int total_vk = getTotalMedia(jsonPosts, "vk");
            int total_tw = getTotalMedia(jsonPosts, "tw");
            int total_fb = getTotalMedia(jsonPosts, "fb");
            int total_yt = getTotalMedia(jsonPosts, "yt");
            int total_gs = getTotalMedia(jsonPosts, "gs");
            int total_tg = getTotalMedia(jsonPosts, "tg");
            int total_ig = getTotalMedia(jsonPosts, "ig");
            int all = total_vk + total_tw + total_fb + total_gs + total_tg + total_ig + total_yt;


            ParseData soData = getWeekDataMedia(type, jsonPosts, first_month, first_year);
            double val = 0;
            for (Double d:soData.valuesA){
                val += d;
            }
            for (Double d:soData.valuesB){
                val += d;
            }

            if ((all != 0) || (val != 0) || jsonCity.length() != 0) {

                XWPFParagraph bodyParagraphIst = docxModel.createParagraph();
                bodyParagraphIst.setPageBreak(true);
                bodyParagraphIst.setStyle("Heading1");
                bodyParagraphIst.setAlignment(ParagraphAlignment.LEFT);
                XWPFRun paragraphConfigIst = bodyParagraphIst.createRun();
                paragraphConfigIst.setFontSize(22);
                paragraphConfigIst.setBold(true);
                paragraphConfigIst.setFontFamily(format);
                paragraphConfigIst.setText(
                        "Источники"
                );
                entityOnPage = 0;
                if (all == 0) {
                    dataLost(docxModel);
                } else {
                    addParagraph(docxModel, String.format("Таблица %s Общее количество публикаций по каждой социальной сети",
                            tableCount));
                    tableCount += 1;
                    XWPFTable tableIst = docxModel.createTable();
                    XWPFTableRow tableRowOneIst = tableIst.getRow(0);
                    XWPFRun run4 = tableRowOneIst.getCell(0).getParagraphs().get(0).createRun();
                    run4.setText("Площадка");
                    run4.setBold(true);
                    run4.setFontSize(12);
                    run4.setFontFamily(format);

                    tableRowOneIst.addNewTableCell();
                    XWPFRun run5 = tableRowOneIst.getCell(1).getParagraphs().get(0).createRun();
                    run5.setText("Количество публикаций, шт.");
                    run5.setBold(true);
                    run5.setFontSize(12);
                    run5.setFontFamily(format);

                    tableRowOneIst.addNewTableCell();
                    XWPFRun run6 = tableRowOneIst.getCell(2).getParagraphs().get(0).createRun();
                    run6.setText("   %     ");
                    run6.setBold(true);
                    run6.setFontSize(12);
                    run6.setFontFamily(format);

                    XWPFTableRow tableRowTwoIst = tableIst.createRow();
                    setText(tableRowTwoIst, "Вконтакте", 0);
                    setText(tableRowTwoIst,  get_format_stng(total_vk), 1);
                    setText(tableRowTwoIst, String.valueOf(Math.round((float) total_vk * 100.00 / (float) all * 100.00) / 100.0), 2);

                    XWPFTableRow tableRowThreeIst = tableIst.createRow();
                    setText(tableRowThreeIst, "Facebook", 0);
                    setText(tableRowThreeIst, get_format_stng(total_fb), 1);
                    setText(tableRowThreeIst, String.valueOf(Math.round((float) total_fb * 100.00 / (float) all * 100.00) / 100.0), 2);

                    XWPFTableRow tableRowThIst = tableIst.createRow();
                    setText(tableRowThIst, "Twitter", 0);
                    setText(tableRowThIst, get_format_stng(total_tw), 1);
                    setText(tableRowThIst, String.valueOf(Math.round((float) total_tw * 100.00 / (float) all * 100.00) / 100.0), 2);

                    XWPFTableRow tableRowFIst = tableIst.createRow();
                    setText(tableRowFIst, "Инстаграм", 0);
                    setText(tableRowFIst, get_format_stng(total_ig), 1);
                    setText(tableRowFIst, String.valueOf(Math.round((float) total_ig * 100.00 / (float) all * 100.00) / 100.0), 2);

                    XWPFTableRow tableRowSixIst = tableIst.createRow();
                    setText(tableRowSixIst, "Telegram", 0);
                    setText(tableRowSixIst, get_format_stng(total_tg), 1);
                    setText(tableRowSixIst, String.valueOf(Math.round((float) total_tg * 100.00 / (float) all * 100.00) / 100.0), 2);

                    XWPFTableRow tableRowSevenIst = tableIst.createRow();
                    setText(tableRowSevenIst, "YouTube", 0);
                    setText(tableRowSevenIst, get_format_stng(total_yt), 1);
                    setText(tableRowSevenIst, String.valueOf(Math.round((float) total_yt * 100.00 / (float) all * 100.00) / 100.0), 2);

                    XWPFTableRow tableRowSevIst = tableIst.createRow();
                    setText(tableRowSevIst, "СМИ", 0);
                    setText(tableRowSevIst, get_format_stng(total_gs), 1);
                    setText(tableRowSevIst, String.valueOf(Math.round((float) total_gs * 100.00 / (float) all * 100.00) / 100.0), 2);

                    XWPFTableRow tableRowSevAll = tableIst.createRow();
                    setText(tableRowSevAll, "Итог", 0);
                    setText(tableRowSevAll, get_format_stng(all), 1);
                    setText(tableRowSevAll, "100", 2);


                    for (int x = 0; x < tableIst.getNumberOfRows(); x++) {
                        XWPFTableRow row = tableIst.getRow(x);
                        XWPFTableCell cell0 = row.getCell(0);
                        cell0.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(3000));
                        XWPFTableCell cell1 = row.getCell(1);
                        cell1.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(5000));
                        cell1.getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                        XWPFTableCell cell2 = row.getCell(2);
                        cell2.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1500));
                        cell2.getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                        AligmentCenter(cell0);
                        AligmentCenter(cell1);
                        AligmentCenter(cell2);
                    }
                }


                diagramCount = addDоubleChats(docxModel, soData.categories, soData.valuesA, soData.valuesB,
                        String.format("Диаграмма %s Динамика публикаций по предоставленным источникам", diagramCount), diagramCount);

//                if (posts.length() == 0) {
//                    dataLost(docxModel);
//                } else {
//                    addParagraph(docxModel, String.format("Таблица %s Топ-%s источников по количеству публикаций", tableCount, posts.length()));
//                    tableCount += 1;
//                    XWPFTable tableTop10Own = docxModel.createTable();
//                    XWPFTableRow tableTop10OwnRow = tableTop10Own.getRow(0);
//
//                    XWPFRun run12 = tableTop10OwnRow.getCell(0).getParagraphs().get(0).createRun();
//                    run12.setText("Название источника");
//                    run12.setBold(true);
//
//                    tableTop10OwnRow.addNewTableCell();
//                    XWPFRun run11 = tableTop10OwnRow.getCell(1).getParagraphs().get(0).createRun();
//                    run11.setText("URL");
//                    run11.setBold(true);
//                    tableTop10OwnRow.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//
//                    tableTop10OwnRow.addNewTableCell();
//                    XWPFRun run13 = tableTop10OwnRow.getCell(2).getParagraphs().get(0).createRun();
//                    run13.setText("Количество публикаций");
//                    run13.setBold(true);
//
//                    JSONObject jsonObject;
//                    for (Object o : posts) {
//                        jsonObject = (JSONObject) o;
//                        getRow(tableTop10Own, jsonObject.get("username").toString(), jsonObject.get("url").toString(),
//                                jsonObject.get("coefficient").toString());
//                    }
//
//
//                    for (int x = 0; x < tableTop10Own.getNumberOfRows(); x++) {
//                        XWPFTableRow row = tableTop10Own.getRow(x);
//                        XWPFTableCell cell0 = row.getCell(0);
//                        cell0.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(5000));
//                        XWPFTableCell cell1 = row.getCell(1);
//                        cell1.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(3000));
//                        cell1.getParagraphs().get(0).setAlignment(ParagraphAlignment.LEFT);
//                        XWPFTableCell cell2 = row.getCell(2);
//                        cell2.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1500));
//                        cell2.getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
//                    }
//                }
            }

            ParseData auditData = getWeekData(type, (JSONArray) (stat).get("graph_data"), first_month, first_year);
            JSONObject sexJson= (JSONObject) ((JSONObject)sex.get("additional_data")).get("sex");
            String[] categoriesCity = new String[]{};
            Double[] valuesACity  = new Double[]{};
            double valueCity;
            int i = 0;
            int count10 = 0;
            JSONObject jsonObject;
            for (Object o :  jsonCity) {
                jsonObject = (JSONObject) o;
                if (i == 10) {
                    break;
                }
                count10 += Integer.parseInt(jsonObject.get("users").toString());
                i ++;
            }
            for (Object o :  jsonCity) {
                if (categoriesCity.length >= 10) {
                    break;
                }

                jsonObject = (JSONObject) o;
                valueCity= Math.round(Double.parseDouble( jsonObject.get("users").toString())*100.00/count10 * 100.00) / 100.00;
                if (valueCity < 1) {
                    break;
                }
                categoriesCity = (String[]) append(categoriesCity, jsonObject.get("city"));
                valuesACity = append(valuesACity, valueCity );
            }

            double valAudit = 0;
            double valSex = 0;
            double valAge= 0;
            double valCity = 0;



            Double[] masSex = new Double[]{new  Double(sexJson.get("u").toString()),
                    new  Double(sexJson.get("m").toString()), new  Double(sexJson.get("w").toString())};


            Double [] masAge = new Double[]{new Double(((JSONObject)age.get("group1")).get("graph").toString()),
                    new Double(((JSONObject)age.get("group2")).get("graph").toString()),
                    new Double(((JSONObject)age.get("group3")).get("graph").toString()),
                    new Double(((JSONObject)age.get("group4")).get("graph").toString())
            };

            for (Double d:masSex){
                valSex += d;
            }
            for (Double d:auditData.valuesA){
                valAudit += d;
            }
            for (Double d:masAge){
                valAge += d;
            }
            for (Double d:masAge){
                valCity += d;
            }


//            if ((valSex != 0) || (valAudit != 0) || (valAge != 0) || (valCity != 0 ) || ( jsonCity.length() != 0)
//                    || (jsonCity.length() != 0) || ((JSONArray) usersJson.get("users")).length() !=0) {
//                XWPFParagraph bodyParagraphAudit = docxModel.createParagraph();
//                bodyParagraphAudit.setAlignment(ParagraphAlignment.LEFT);
//                bodyParagraphAudit.setPageBreak(true);
//                bodyParagraphAudit.setStyle("Heading1");
//                XWPFRun paragraphConfigAudit = bodyParagraphAudit.createRun();
//                paragraphConfigAudit.setFontSize(22);
//                paragraphConfigAudit.setBold(true);
//                paragraphConfigAudit.setFontFamily("Arial");
//                paragraphConfigAudit.setText(
//                        "Аудитория"
//                );
//                entityOnPage = 0;
//                diagramCount = addChats(docxModel, auditData.categories, auditData.valuesA, String.format("Диаграмма %s Динамика объема аудитории", diagramCount), diagramCount);
//
//
//                diagramCount = addPie(docxModel, new String[]{"Не указан", "Мужчины", "Женщины"}, masSex, String.format("Диаграмма %s Распределение аудитории по полу", diagramCount), diagramCount);
//
//
//                diagramCount = addPie(docxModel, new String[]{"18-25 лет", "26-40 лет", "40 лет и старше", "не указан"},
//                        masAge, String.format("Диаграмма %s Распределение аудитории по возрасту", diagramCount), diagramCount);
//
//
//                diagramCount = addPieFormat(docxModel, categoriesCity, valuesACity, String.format("Диаграмма %s Распределение аудитории по геолокации", diagramCount), diagramCount, false);
//
//                if (jsonCity.length() == 0) {
//                    dataLost(docxModel);
//                } else {
//                    addParagraph(docxModel, String.format("Таблица %s Топ-%s городов", tableCount, jsonCity.length()));
//                    tableCount += 1;
//                    XWPFTable tableTop10OCity = docxModel.createTable();
//                    XWPFTableRow tableTop10OCityRow = tableTop10OCity.getRow(0);
//                    XWPFRun runCity = tableTop10OCityRow.getCell(0).getParagraphs().get(0).createRun();
//                    tableTop10OCityRow.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//                    runCity.setText("Город");
//                    runCity.setBold(true);
//
//                    tableTop10OCityRow.addNewTableCell();
//                    XWPFRun r9 = tableTop10OCityRow.getCell(1).getParagraphs().get(0).createRun();
//                    r9.setText("Количество");
//                    r9.setBold(true);
//                    tableTop10OCityRow.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//
//
//                    tableTop10OCityRow.addNewTableCell();
//                    XWPFRun run10a = tableTop10OCityRow.getCell(2).getParagraphs().get(0).createRun();
//                    run10a.setText("%");
//                    run10a.setBold(true);
//                    tableTop10OCityRow.getCell(2).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//
//                    try {
//                        for (Object o : jsonCity) {
//                            jsonObject = (JSONObject) o;
//                            getRow(tableTop10OCity, jsonObject.get("city").toString(), jsonObject.get("users").toString(),
//                                    String.format("%.1f", Double.parseDouble(jsonObject.get("users").toString()) * 100.0 / Double.valueOf(count10)));
//                        }
//                    } catch (Exception e) {
//                        System.out.println("S");
//                    }
//
//                    for (int x = 0; x < tableTop10OCity.getNumberOfRows(); x++) {
//                        XWPFTableRow row000 = tableTop10OCity.getRow(x);
//                        XWPFTableCell cell0000 = row000.getCell(0);
//                        cell0000.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4000));
//                        XWPFTableCell cell1000 = row000.getCell(1);
//                        cell1000.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4000));
//                        cell1000.getParagraphs().get(0).setAlignment(ParagraphAlignment.LEFT);
//                        XWPFTableCell cell2000 = row000.getCell(2);
//                        cell2000.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1500));
//                        cell2000.getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
//                    }
//                }
//
//                if (((JSONArray) usersJson.get("users")).length() == 0) {
//                    dataLost(docxModel);
//                } else {
//                    addParagraph(docxModel, String.format("Таблица %s Топ-%s активных пользователей по сумме реакции (лайков, комментариев, репостов)", tableCount, ((JSONArray) usersJson.get("users")).length()));
//                    tableCount += 1;
//                    XWPFTable tableTop10OUser = docxModel.createTable();
//                    XWPFTableRow tableTop10OUserRow = tableTop10OUser.getRow(0);
//                    XWPFRun run8 = tableTop10OUserRow.getCell(0).getParagraphs().get(0).createRun();
//                    tableTop10OUserRow.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//                    run8.setText("Пользователь");
//                    run8.setBold(true);
//
//                    tableTop10OUserRow.addNewTableCell();
//                    XWPFRun run9 = tableTop10OUserRow.getCell(1).getParagraphs().get(0).createRun();
//                    run9.setText("URL");
//                    run9.setBold(true);
//                    tableTop10OUserRow.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//
//                    tableTop10OUserRow.addNewTableCell();
//                    XWPFRun run10 = tableTop10OUserRow.getCell(2).getParagraphs().get(0).createRun();
//                    run10.setText("Сумма реакции");
//                    run10.setBold(true);
//                    tableTop10OUserRow.getCell(2).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//
//                    for (Object o : (JSONArray) usersJson.get("users")) {
//                        jsonObject = (JSONObject) o;
//                        getRow(tableTop10OUser, jsonObject.get("name").toString(), jsonObject.get("url").toString(),
//                                jsonObject.get("coefficient").toString());
//                    }
//
//                    for (int x = 0; x < tableTop10OUser.getNumberOfRows(); x++) {
//                        XWPFTableRow row = tableTop10OUser.getRow(x);
//                        XWPFTableCell cell0 = row.getCell(0);
//                        cell0.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4000));
//                        XWPFTableCell cell1 = row.getCell(1);
//                        cell1.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4000));
//                        cell1.getParagraphs().get(0).setAlignment(ParagraphAlignment.LEFT);
//                        XWPFTableCell cell2 = row.getCell(2);
//                        cell2.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1500));
//                        cell2.getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
//                    }
//                }
//            }
            int likesPosts = 0;
            int likesComment = 0;

            for (Object o : postsContent) {
                likesPosts+=1;
//
//                jsonObject = (JSONObject) o;
//                if (Integer.parseInt(jsonObject.get("viewed").toString()) + Integer.parseInt(jsonObject.get("reposts").toString()) +
//                        Integer.parseInt(jsonObject.get("likes").toString()) + Integer.parseInt(jsonObject.get("comments").toString()) +
//                        Integer.parseInt(jsonObject.get("attendance").toString()) > 0) {
//                    likesPosts+=1;
//                }
            }
            for (Object o : commentContent) {
//                if (Integer.parseInt(((JSONObject) o).get("likes").toString())  > 0) {
                    likesComment+=1;
//                }
            }
            if((likesComment !=0) || (likesPosts !=0 )) {
                XWPFParagraph bodyParagrapKeysP = docxModel.createParagraph();
                bodyParagrapKeysP.setAlignment(ParagraphAlignment.LEFT);
                bodyParagrapKeysP.setPageBreak(true);
                bodyParagrapKeysP.setStyle("Heading1");
                XWPFRun paragraphConfigKeysP = bodyParagrapKeysP.createRun();
                paragraphConfigKeysP.setFontSize(22);
                paragraphConfigKeysP.setBold(true);
                paragraphConfigKeysP.setFontFamily(format);
//                paragraphConfigKeysP.setText(
//                        "Ключевые публикации и комментарии"
//                );
                entityOnPage =0;
                if (likesPosts == 0) {
                    dataLost(docxModel);
                } else {
                    addParagraph_new(docxModel, String.format("Таблица %s Топ-%s публикаций за сутки", tableCount, postsContent.length()), tableCount==1);
                    tableCount += 1;
                    XWPFTable tableTop10Post = docxModel.createTable();
                    XWPFTableRow tableTop10PostRow = tableTop10Post.getRow(0);


                    XWPFRun run15 = tableTop10PostRow.getCell(0).getParagraphs().get(0).createRun();
                    run15.setText("Публикация");
                    run15.setBold(true);
                    run15.setFontSize(12);
                    run15.setFontFamily(format);
//                    tableTop10PostRow.addNewTableCell();
//                    XWPFRun run16 = tableTop10PostRow.getCell(1).getParagraphs().get(0).createRun();
//                    run16.setText("URL");
//                    run16.setBold(true);
//                    tableTop10PostRow.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                    tableTop10PostRow.addNewTableCell();
                    XWPFRun run17 = tableTop10PostRow.getCell(1).getParagraphs().get(0).createRun();
                    run17.setText("Охват");
                    run17.setBold(true);
                    run17.setFontSize(12);
                    run17.setFontFamily(format);

                    String text;

                    for (Object o : postsContent) {
                        jsonObject = (JSONObject) o;
                        text = updateText(jsonObject.get("text").toString());
//                        getRow(tableTop10Post, text, jsonObject.get("uri").toString(), res(jsonObject)
//                                );
                        XWPFTableRow tableRowTwoIst = tableTop10Post.createRow();
                        setText(tableRowTwoIst, text.replaceAll("\\<[^>]*>",""), 0);
                        setText(tableRowTwoIst, get_format_stng((Integer) jsonObject.get("attendance")), 1);

                    }
                    for (int x = 0; x < tableTop10Post.getNumberOfRows(); x++) {
                        XWPFTableRow row = tableTop10Post.getRow(x);
                        XWPFTableCell cell0 = row.getCell(0);


                        cell0.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(7500));
                        cell0.getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                        AligmentCenter(cell0);

//                        XWPFTableCell cell1 = row.getCell(1);
//                        cell1.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(3000));
//                        cell1.getParagraphs().get(0).setAlignment(ParagraphAlignment.LEFT);
                        XWPFTableCell cell2 = row.getCell(1);
                        cell2.getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                        cell2.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(2000));
                        AligmentCenter(cell2);
                    }
                }

//                if (likesComment== 0) {
                if (likesPosts== 0) {
                    dataLost(docxModel);
                } else {
                    addParagraph_new(docxModel, String.format("Таблица %s Перечень основных информационных поводов публикаций", tableCount, min(likesPosts, 3)), true);
                    tableCount += 1;
                    String text;
                    XWPFTable tableTop10Comment = docxModel.createTable();
                    XWPFTableRow tableTop10CommentRow = tableTop10Comment.getRow(0);
                    XWPFRun run19 = tableTop10CommentRow.getCell(0).getParagraphs().get(0).createRun();
                    run19.setText("Текст");
                    run19.setBold(true);
                    run19.setFontSize(12);
                    run19.setFontFamily(format);


                    tableTop10CommentRow.addNewTableCell();
                    XWPFRun run20 = tableTop10CommentRow.getCell(1).getParagraphs().get(0).createRun();
                    run20.setText("URL");
                    run20.setBold(true);
                    run20.setFontSize(12);
                    run20.setFontFamily(format);

                    tableTop10CommentRow.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                    tableTop10CommentRow.addNewTableCell();
                    XWPFRun run21 = tableTop10CommentRow.getCell(2).getParagraphs().get(0).createRun();
                    run21.setText("Охват");
                    run21.setBold(true);
                    run21.setFontSize(12);
                    run21.setFontFamily(format);

//                    for (Object o : commentContent) {
                    int i_z = 0;
                    String Uri;
                    for (Object o : postsContent) {
                        i_z += 1;
                        if (i_z > 3) {
                            break;
                        }
                        jsonObject = (JSONObject) o;
                        try {
                            text = updateText(jsonObject.get("title").toString());
                        } catch (Exception e) {
                            text = updateText(jsonObject.get("text").toString());

                        }
                        if (text.equals(""))
                            text = updateText(jsonObject.get("text").toString());
                        try {
                            Uri = jsonObject.get("post_url").toString();
                        } catch (Exception e) {
                            Uri =  jsonObject.get("uri").toString();

                        }
                        getRow(tableTop10Comment, text.replaceAll("\\<[^>]*>",""), Uri,
                                get_format_stng((Integer) jsonObject.get("attendance")));
                    }

                    for (int x = 0; x < tableTop10Comment.getNumberOfRows(); x++) {
                        XWPFTableRow row = tableTop10Comment.getRow(x);
                        XWPFTableCell cell0 = row.getCell(0);
                        cell0.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(6500));
                        cell0.getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                        cell0.getParagraphs().get(0).setAlignment(ParagraphAlignment.LEFT);
                        AligmentCenter(cell0);

                        XWPFTableCell cell1 = row.getCell(1);
                        cell1.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(2000));
                        cell1.getParagraphs().get(0).setAlignment(ParagraphAlignment.LEFT);
                        cell1.getParagraphs().get(0).setVerticalAlignment(TextAlignment.CENTER);
                        AligmentCenter(cell1);

                        XWPFTableCell cell2 = row.getCell(2);
                        cell2.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1000));
                        cell2.getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                        AligmentCenter(cell2);

                    }
                }
                if (likesPosts == 0) {
                    dataLost(docxModel);
                } else {
                    addParagraph_new(docxModel, String.format("Таблица %s Ссылки на публикации", tableCount, likesPosts), true);
                    tableCount += 1;
                    XWPFTable tableTop10Post = docxModel.createTable();
                    XWPFTableRow tableTop10PostRow = tableTop10Post.getRow(0);


                    XWPFRun run15 = tableTop10PostRow.getCell(0).getParagraphs().get(0).createRun();
                    run15.setText("Публикация");
                    run15.setBold(true);
                    run15.setFontSize(12);
                    run15.setFontFamily(format);

                    tableTop10PostRow.addNewTableCell();
                    XWPFRun run16 = tableTop10PostRow.getCell(1).getParagraphs().get(0).createRun();
                    run16.setText("URL");
                    run16.setBold(true);
                    run16.setFontSize(12);
                    run16.setFontFamily(format);

                    tableTop10PostRow.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                    //                tableTop10PostRow.addNewTableCell();
                    //                XWPFRun run17 = tableTop10PostRow.getCell(1).getParagraphs().get(0).createRun();
                    //                run17.setText("Резонанс");
                    //                run17.setBold(true);


                    String text;

                    for (Object o : postsContent) {
                        jsonObject = (JSONObject) o;
                        text = updateText(jsonObject.get("text").toString());
                        //                        getRow(tableTop10Post, text, jsonObject.get("uri").toString(), res(jsonObject)
                        //                                );
                        XWPFTableRow tableRowTwoIst = tableTop10Post.createRow();
                        setText(tableRowTwoIst, text.replaceAll("\\<[^>]*>",""), 0);
                        setText(tableRowTwoIst, jsonObject.get("uri").toString(), 1, true);

                    }
                    for (int x = 0; x < tableTop10Post.getNumberOfRows(); x++) {
                        XWPFTableRow row = tableTop10Post.getRow(x);
                        XWPFTableCell cell0 = row.getCell(0);
                        cell0.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(7500));
                        cell0.getParagraphs().get(0).setAlignment(ParagraphAlignment.LEFT);
                        AligmentCenter(cell0);

                        //                        XWPFTableCell cell1 = row.getCell(1);
                        //                        cell1.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(3000));
                        //                        cell1.getParagraphs().get(0).setAlignment(ParagraphAlignment.LEFT);
                        XWPFTableCell cell2 = row.getCell(1);
                        cell2.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(2000));
                        cell2.getParagraphs().get(0).setAlignment(ParagraphAlignment.LEFT);
                        AligmentCenter(cell2);

                    }

                }
            }
            CTP ctp = CTP.Factory.newInstance();

            ctp.addNewR().addNewPgNum();

            XWPFParagraph codePara = new XWPFParagraph(ctp, docxModel);
            XWPFParagraph[] paragraphs = new XWPFParagraph[1];
            paragraphs[0] = codePara;

            codePara.setAlignment(ParagraphAlignment.CENTER);

            CTSectPr sectPr = docxModel.getDocument().getBody().addNewSectPr();

            XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(docxModel, sectPr);
            headerFooterPolicy.createFooter(STHdrFtr.DEFAULT, paragraphs);
            CTSectPr sect = docxModel.getDocument().getBody().getSectPr();
            sect.addNewTitlePg();
            return docxModel;

        } catch (Exception e) {
            e.printStackTrace();
        }
        return new XWPFDocument();

    }
    private static void AligmentCenter(XWPFTableCell cell) {
        cell.getParagraphArray(0).setSpacingAfter(0);
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
    }

    public static String res(JSONObject jsonObject){
        return String.valueOf(Integer.parseInt(jsonObject.get("viewed").toString()) + Integer.parseInt(jsonObject.get("reposts").toString()) +
                Integer.parseInt(jsonObject.get("likes").toString()) + Integer.parseInt(jsonObject.get("comments").toString()));
    }
    public static int getTotalMedia(JSONObject jsonPosts, String key){
        int total = 0;
        JSONArray jsonArray;
        for(Object o: (JSONArray)((JSONObject)(jsonPosts).get(key)).get("total")){
            jsonArray = (JSONArray) o;
            total += (int) jsonArray.get(1);
        }
        return total;
    }

    private static void getRow(XWPFTable table, String str1, String str2, String str3){
        XWPFTableRow tableRowTwoIst = table.createRow();
        setText(tableRowTwoIst, str1, 0);
        setText(tableRowTwoIst, str2, 1, true);
        setText(tableRowTwoIst, str3, 2);

    }

    public static int getComment(JSONObject jsonComments, String key) {
        int res = 0;
        JSONArray jsonArray;
        for (Object o : (JSONArray) (jsonComments).get(key)) {
            jsonArray = (JSONArray) o;
            res += (int) jsonArray.get(1);
        }
        return res;
    }

    private static int addChats(XWPFDocument docxModel, String[] categories, Double[] valuesA, String name, Integer dia ) throws Exception {
        // create the data
        double val = 0;
        for (Double d:valuesA){
            val += d;
        }
        if (val <= 0) {
            return dia;
        }
        addParagraph(docxModel, name);
        dia +=1;

        // create the chart
        XWPFChart chart = docxModel.createChart(17 * Units.EMU_PER_CENTIMETER,  6 * Units.EMU_PER_CENTIMETER);

        // create data sources
        int numOfPoints = categories.length;
        String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
        String valuesDataRangeA = chart.formatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
        XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
        XDDFNumericalDataSource<Double> valuesDataA = XDDFDataSourcesFactory.fromArray(valuesA, valuesDataRangeA, 1);


        XDDFSolidFillProperties WHITE_SMOKE = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.GRAY));
        XDDFLineProperties line = new XDDFLineProperties();
        line.setFillProperties(WHITE_SMOKE);
        XDDFSolidFillProperties WHITE = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.WHITE));
        XDDFLineProperties lineWhite = new XDDFLineProperties();
        lineWhite.setFillProperties(WHITE);

        // create axis
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFShapeProperties s = bottomAxis.getOrAddShapeProperties();
        s.setFillProperties(new XDDFSolidFillProperties(XDDFColor.from(PresetColor.WHITE)));

        bottomAxis.setVisible(true);
        bottomAxis.setMajorTickMark(AxisTickMark.CROSS);
        bottomAxis.getOrAddShapeProperties().setLineProperties(lineWhite);
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        leftAxis.getOrAddMajorGridProperties();


        leftAxis.getOrAddMajorGridProperties().setLineProperties(line);
        leftAxis.setVisible(true);
        // Set AxisCrossBetween, so the left axis crosses the category axis between the categories.
        // Else first and last category is exactly on cross points and the bars are only half visible.
        leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
        leftAxis.getOrAddShapeProperties().setLineProperties(lineWhite);

        // create chart data
        XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
        ((XDDFBarChartData) data).setBarDirection(BarDirection.COL);

        // create series
        // if only one series do not vary colors for each bar
        ((XDDFBarChartData) data).setVaryColors(false);
        XDDFChartData.Series series = data.addSeries(categoriesData, valuesDataA);
        series.setTitle("", setTitleInDataSheet(chart, "a", 1));


        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.CORNFLOWER_BLUE));

        XDDFShapeProperties properties = series.getShapeProperties();
        if (properties == null) {
            properties = new XDDFShapeProperties();
        }
        properties.setFillProperties(fill);
        series.setShapeProperties(properties);

        // add data labels
        chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).addNewDLbls();
        chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowVal().setVal(true);
        chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowLegendKey().setVal(false);
        chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowCatName().setVal(false);
        chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowSerName().setVal(false);
        chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewDLblPos().setVal(org.openxmlformats.schemas.drawingml.x2006.chart.STDLblPos.CTR);

        chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewTxPr()
                .addNewBodyPr().setRot((int)(-90.00 * 60000));
        chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().getTxPr()
                .addNewP().addNewPPr().addNewDefRPr();

//        chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowPercent()

        chart.plot(data);
        return dia;
    }

    private static XWPFDocument dataLost(XWPFDocument docxModel) {
//        XWPFParagraph paragraphDinColPubCom = docxModel.createParagraph();
//        paragraphDinColPubCom.setAlignment(ParagraphAlignment.CENTER);
//        XWPFRun paragraphConfigDinColPubCom = paragraphDinColPubCom.createRun();
//        paragraphConfigDinColPubCom.setFontSize(10);
//        paragraphConfigDinColPubCom.setFontFamily("Arial");
//        paragraphConfigDinColPubCom.setBold(true);
//        paragraphConfigDinColPubCom.addBreak();
//        paragraphConfigDinColPubCom.setText(
//                "данные отсутвуют"
//        );

        return docxModel;
    }

    private static int addDоubleChats(XWPFDocument docxModel, String[] categories, Double[] valuesA, Double[] valuesB, String name, int dia) throws Exception {
        int i_i =0;
        for (String s : categories) {
            String [] splyt_d = s.split("-");
            categories[i_i] = splyt_d[2] + "-" + splyt_d[1] +  "-" + splyt_d[0];
            i_i += 1;
        }
        double val = 0;
        for (Double d:valuesA){
            val += d;
        }
        for (Double d:valuesB){
            val += d;
        }
        if (val <= 0) {
            return dia;
        }
        addParagraph(docxModel, name);
        dia +=1;
        XWPFChart chart = docxModel.createChart(17 * Units.EMU_PER_CENTIMETER,  6 * Units.EMU_PER_CENTIMETER);

        int numOfPoints = categories.length;
        String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
        String valuesDataRangeA = chart.formatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
        String valuesDataRangeB = chart.formatRange(new CellRangeAddress(1, numOfPoints, 2, 2));

        XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
        XDDFNumericalDataSource<Double> valuesDataA = XDDFDataSourcesFactory.fromArray(valuesA, valuesDataRangeA, 1);
        XDDFNumericalDataSource<Double> valuesDataB = XDDFDataSourcesFactory.fromArray(valuesB, valuesDataRangeB, 2);

        XDDFSolidFillProperties WHITE = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.WHITE));
        XDDFLineProperties lineWhite = new XDDFLineProperties();
        lineWhite.setFillProperties(WHITE);

        // create axis
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.getOrAddShapeProperties().setLineProperties(lineWhite);

        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
        leftAxis.getOrAddShapeProperties().setLineProperties(lineWhite);
        XDDFSolidFillProperties WHITE_SMOKE = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.LIGHT_GRAY));
        XDDFLineProperties line = new XDDFLineProperties();
        line.setFillProperties(WHITE_SMOKE);
        leftAxis.getOrAddMajorGridProperties().setLineProperties(line);
        // create chart data
        XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
        ((XDDFBarChartData) data).setBarDirection(BarDirection.COL);

        // create series
        // if only one series do not vary colors for each bar
        ((XDDFBarChartData) data).setVaryColors(false);
        XDDFChartData.Series series = data.addSeries(categoriesData, valuesDataA);
        XDDFChartData.Series series2 = data.addSeries(categoriesData, valuesDataB);

        series.setTitle("Сми", null);
        series2.setTitle("Соцмедиа", null);

        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.YELLOW));
        transp(fill);
        XDDFShapeProperties properties = series.getShapeProperties();
        if (properties == null) {
            properties = new XDDFShapeProperties();
        }
        properties.setFillProperties(fill);
        series.setShapeProperties(properties);

        fill = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.CORNFLOWER_BLUE));
        transp(fill);

        properties = series2.getShapeProperties();
        if (properties == null) {
            properties = new XDDFShapeProperties();
        }
        properties.setFillProperties(fill);
        series2.setShapeProperties(properties);

        // add data labels
        for (int i =0; i < 2; i ++) {
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).addNewDLbls();
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowVal().setVal(true);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowLegendKey().setVal(false);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowCatName().setVal(false);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowSerName().setVal(false);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewDLblPos().setVal(org.openxmlformats.schemas.drawingml.x2006.chart.STDLblPos.OUT_END);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewTxPr()
                    .addNewBodyPr().setRot((int)(-90.00 * 60000));
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().getTxPr()
                    .addNewP().addNewPPr().addNewDefRPr();
        }

        chart.plot(data);
        chart.getCTChart().getPlotArea().getBarChartArray(0).addNewOverlap().setVal((byte)-25);
        chart.getCTChart().getPlotArea().getBarChartArray(0).addNewGapWidth().setVal(500);
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP);
//        legend.setOverlay(true);
        XDDFTextBody legendTextBody = new XDDFTextBody(legend);
        legendTextBody.getXmlObject().addNewBodyPr();
        legendTextBody.addNewParagraph().addDefaultRunProperties().setFontSize(8d);
        legend.setTextBody(legendTextBody);
        return dia;
    }
    private static void transp(XDDFSolidFillProperties fill ) {
        org.openxmlformats.schemas.drawingml.x2006.main.CTSolidColorFillProperties ctSolidColorFillProperties =
                (org.openxmlformats.schemas.drawingml.x2006.main.CTSolidColorFillProperties) fill.getXmlObject();
        org.openxmlformats.schemas.drawingml.x2006.main.CTPresetColor ctPresetColor = ctSolidColorFillProperties.getPrstClr();
        org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveFixedPercentage ctPositiveFixedPercentage = ctPresetColor.addNewAlpha();
        ctPositiveFixedPercentage.setVal(54000);
    }

    private static int addPie(XWPFDocument document,  String[] categories, Double[] valuesA, String name, int dia) throws IOException, InvalidFormatException {
        return addPie( document, categories,  valuesA, name, dia, false);
    }

    private static int addPie(XWPFDocument document,  String[] categories, Double[] valuesA, String name, int dia, boolean RGB) throws IOException, InvalidFormatException {
        double val = 0;
        for (Double d:valuesA){
            val += d;
        }
        if (val <= 0) {
            return dia;
        }
        addParagraph(document,  name);
        dia+=1;
        XWPFChart chart = document.createChart(17 * Units.EMU_PER_CENTIMETER, 5 * Units.EMU_PER_CENTIMETER);

        int numOfPoints = categories.length;
        String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
        String valuesDataRangeA = chart.formatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
        XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);


        XDDFNumericalDataSource<Double> valuesDataA = XDDFDataSourcesFactory.fromArray(valuesA, valuesDataRangeA, 1);

        XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
        XDDFChartData.Series series = data.addSeries(categoriesData, valuesDataA);
        data.setVaryColors(true);
//        series.setShowLeaderLines(false);
        series.setTitle("", chart.setSheetTitle("", 1));
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).addNewDLbls();
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowVal().setVal(false);
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowLeaderLines().setVal(true);
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowLegendKey().setVal(true);
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowCatName().setVal(true);
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowSerName().setVal(false);
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowBubbleSize().setVal(false);
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowPercent().setVal(true);

        if (RGB) {

            int pointCount = series.getCategoryData().getPointCount();
            for (int p = 0; p < pointCount; p++) {
                chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).addNewDPt().addNewIdx().setVal(p);
                if (p == 1) {
                    chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDPtArray(p)
                            .addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(DefaultIndexedColorMap.getDefaultRGB(23));
                }
                if (p == 2) {
                    chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDPtArray(p)
                            .addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(DefaultIndexedColorMap.getDefaultRGB(10));
                }
                if (p == 0) {
                    chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDPtArray(p)
                            .addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(DefaultIndexedColorMap.getDefaultRGB(11));
                }
            }
        }

        chart.plot(data);
        return dia;
    }

    private static int addPieFormat(XWPFDocument document, String[] categories, Double[] valuesA, String name,
                                    int dia, boolean RGB) throws IOException, InvalidFormatException {
        // create the data


        for (int i=0; i < categories.length; i++) {
            categories[i] += " " + String.format("%.2f", valuesA[i]) + "% ";
        }
        // create data sources

        int numOfPoints = categories.length;
        if (numOfPoints == 0){
            return dia;
        }
        double val = 0;
        for (Double d:valuesA){
            val += d;
        }
        if (val <= 0) {
            return dia;
        }
        addParagraph(document,  name);
        dia+=1;
        XWPFChart chart = document.createChart(17 * Units.EMU_PER_CENTIMETER, 5 * Units.EMU_PER_CENTIMETER);

        String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
        String valuesDataRangeA = chart.formatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
        XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);

        XDDFNumericalDataSource<Double> valuesDataA = XDDFDataSourcesFactory.fromArray(valuesA, valuesDataRangeA, 1);

        XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
        XDDFChartData.Series series = data.addSeries(categoriesData, valuesDataA);

        data.setVaryColors(true);
//        series.setShowLeaderLines(false);
        series.setTitle("", chart.setSheetTitle("", 1));

        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).addNewDLbls();
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowVal().setVal(false);
        if (RGB) {
            chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowLegendKey().setVal(false);
            chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowCatName().setVal(false);
        } else {
            chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowLegendKey().setVal(true);
            chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowCatName().setVal(true);
        }
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowSerName().setVal(false);
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowBubbleSize().setVal(false);
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowPercent().setVal(false);
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowLeaderLines().setVal(false);
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().setSeparator("\n");
//        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewDLblPos().setVal(STDLblPos.Enum.forString("inEnd"));


        if (RGB) {

            int pointCount = series.getCategoryData().getPointCount();
            for (int p = 0; p < pointCount; p++) {
                chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).addNewDPt().addNewIdx().setVal(p);
                if (p == 1) {
                    chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDPtArray(p)
                            .addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(DefaultIndexedColorMap.getDefaultRGB(11));
                }
                if (p == 2) {
                    chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDPtArray(p)
                            .addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(DefaultIndexedColorMap.getDefaultRGB(23));
                }
                if (p == 0) {
                    chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDPtArray(p)
                            .addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(DefaultIndexedColorMap.getDefaultRGB(10));
                }
            }
        }

        chart.plot(data);
        if (RGB) {
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.RIGHT);
        }
        return dia;
    }

    private static int addArea(XWPFDocument document, String[] categories,
                               Double[] valuesNegative,
                               Double[] valuesPositive,
                               Double[] valuesNetural, String name, int dia) throws IOException, InvalidFormatException {
        double val = 0;
        for (Double d:valuesNegative){
            val += d;
        }
        for (Double d:valuesPositive){
            val += d;
        }
        for (Double d:valuesNetural){
            val += d;
        }
        if (val <= 0) {
            return dia;
        }
        addParagraph(document,  name);
        dia+=1;
        XWPFChart chart = document.createChart(17 * Units.EMU_PER_CENTIMETER, 6 * Units.EMU_PER_CENTIMETER);
        // create data sources
        int numOfPoints = categories.length;
        String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
        String valuesDataRangeA = chart.formatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
        String valuesDataRangeB = chart.formatRange(new CellRangeAddress(1, numOfPoints, 2, 2));
        String valuesDataRangeC = chart.formatRange(new CellRangeAddress(1, numOfPoints, 3, 3));

        XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
        XDDFNumericalDataSource<Double> valuesDataA = XDDFDataSourcesFactory.fromArray(valuesNegative, valuesDataRangeA, 1);
        XDDFNumericalDataSource<Double> valuesDataB = XDDFDataSourcesFactory.fromArray(valuesNetural, valuesDataRangeB, 2);
        XDDFNumericalDataSource<Double> valuesDataC = XDDFDataSourcesFactory.fromArray(valuesPositive, valuesDataRangeC, 3);

        XDDFSolidFillProperties WHITE = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.WHITE));
        XDDFLineProperties lineWhite = new XDDFLineProperties();
        lineWhite.setFillProperties(WHITE);

        // create axis
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.getOrAddShapeProperties().setLineProperties(lineWhite);
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        // Set AxisCrossBetween, so the left axis crosses the category axis between the categories.
        // Else first and last category is exactly on cross points and the bars are only half visible.
        leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
        leftAxis.getOrAddShapeProperties().setLineProperties(lineWhite);
        XDDFSolidFillProperties WHITE_SMOKE = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.GRAY));
        XDDFLineProperties line = new XDDFLineProperties();
        line.setFillProperties(WHITE_SMOKE);
        leftAxis.getOrAddMajorGridProperties().setLineProperties(line);
        // create chart data
        XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
        ((XDDFBarChartData) data).setBarDirection(BarDirection.COL);
        ((XDDFBarChartData) data).setBarGrouping(BarGrouping.STACKED);
        XDDFBarChartData bar = (XDDFBarChartData) data;
        bar.setBarGrouping(BarGrouping.PERCENT_STACKED);

        chart.getCTChart().getPlotArea().getBarChartArray(0).addNewOverlap().setVal((byte)100);


        XDDFChartData.Series series1 = data.addSeries(categoriesData, valuesDataA);
        series1.setTitle("Негативная тональность %", null);
        XDDFChartData.Series series2 = data.addSeries(categoriesData, valuesDataB);
        series2.setTitle("Нейтральная тональность %", null);
        XDDFChartData.Series series3 = data.addSeries(categoriesData, valuesDataC);
        series3.setTitle("Позитивная тональность %", null);


        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.RED));

        XDDFShapeProperties properties = series1.getShapeProperties();
        if (properties == null) {
            properties = new XDDFShapeProperties();
        }
        properties.setFillProperties(fill);
        series1.setShapeProperties(properties);
        fill = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.GRAY));

        properties = series2.getShapeProperties();
        if (properties == null) {
            properties = new XDDFShapeProperties();
        }
        properties.setFillProperties(fill);
        series2.setShapeProperties(properties);
        fill = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.GREEN));

        properties = series3.getShapeProperties();
        if (properties == null) {
            properties = new XDDFShapeProperties();
        }
        properties.setFillProperties(fill);
        series3.setShapeProperties(properties);

        for (int i =0; i < 3; i ++) {
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).addNewDLbls();
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowPercent().setVal(true);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowLegendKey().setVal(false);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowCatName().setVal(false);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowSerName().setVal(false);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowBubbleSize().setVal(true);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowVal().setVal(true);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewTxPr()
                    .addNewBodyPr().setRot((int)(-90.00 * 60000));
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().getTxPr()
                    .addNewP().addNewPPr().addNewDefRPr();
        }


        data.setVaryColors(true);
        chart.plot(data);

        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP);
        return dia;
    }


    public static ParseData getWeekData(String type, JSONArray jsonTotal, int first_month, int first_year) throws ParseException {
        String[] categoriesPost = new String[]{};
        Double[] valuesAPost = new Double[]{};
        JSONArray jsonArray;
        if  (type.equals("day")) {
            for (Object o : jsonTotal) {
                jsonArray = (JSONArray) o;
                categoriesPost = (String[]) append(categoriesPost, jsonArray.get(0));
                valuesAPost = append(valuesAPost, new Double(jsonArray.get(1).toString()));
            }}
        else {
            boolean isContain;
            long circe = 1;
            int circeM;
            int lastDate = 0;
            if (type.equals("week")) {
                circeM = 100;
            } else {
                if (type.equals("month")) {
                    circeM = 100;
                } else{
                    circeM = 10; }
            }

            for (Object o : jsonTotal) {
                jsonArray = (JSONArray) o;

                int dateInt = getDate((String) jsonArray.get(0), type);
                if (lastDate != 1 && dateInt == 1 && categoriesPost.length > 0) {
                    circe = circe * circeM;
                }
                String date = String.valueOf(dateInt * circe);

                isContain = false;
                for (int j = 0; j < categoriesPost.length; j++) {
                    if (categoriesPost[j].equals(date)) {
                        valuesAPost[j] +=  new Double(jsonArray.get(1).toString());
                        isContain = true;
                        break;
                    }
                }
                if (!isContain) {
                    categoriesPost = append(categoriesPost, date);
                    valuesAPost = append(valuesAPost, new Double(jsonArray.get(1).toString()));

                }
                lastDate = dateInt;
            }
            changeWeekString(categoriesPost, type, first_month, first_year);


        }
        return new ParseData(categoriesPost, valuesAPost);
    }

    public static ParseData getWeekDataMedia(String type, JSONObject jsonPosts, int first_month, int first_year) throws ParseException {
        String[] categoriesSoMedia = new String[]{};
        Double[] valuesSo = new Double[]{};
        Double[] valuesMedia = new Double[]{};
        JSONArray Gs = ((JSONArray) ((JSONObject)jsonPosts.get("gs")).get("total"));
        JSONArray Vk = ((JSONArray) ((JSONObject)jsonPosts.get("vk")).get("total"));
        JSONArray Tw = ((JSONArray) ((JSONObject)jsonPosts.get("tw")).get("total"));
        JSONArray Fb = ((JSONArray) ((JSONObject)jsonPosts.get("fb")).get("total"));
        JSONArray Tg = ((JSONArray) ((JSONObject)jsonPosts.get("tg")).get("total"));
        JSONArray Ig = ((JSONArray) ((JSONObject)jsonPosts.get("ig")).get("total"));
        if  (type.equals("day")) {

            for (int i =0; i<Gs.length(); i++) {
                categoriesSoMedia = append(categoriesSoMedia, (String)((JSONArray)Gs.get(i)).get(0));

                valuesSo = append(valuesSo,             new Double (((JSONArray)Vk.get(i)).get(1).toString()) +
                        new Double (((JSONArray)Tw.get(i)).get(1).toString())+
                        new Double (((JSONArray)Fb.get(i)).get(1).toString())+
                        +new Double (((JSONArray)Tg.get(i)).get(1).toString())+
                        +new Double (((JSONArray)Ig.get(i)).get(1).toString()));
                valuesMedia = append(valuesMedia, new Double (((JSONArray)Gs.get(i)).get(1).toString()));

            }
        } else {
            boolean isContain;
            long circe = 1;
            int circeM;
            int lastDate = 0;
            if (type.equals("week")) {
                circeM = 100;
            } else {
                if (type.equals("month")) {
                    circeM = 100;
                } else{
                    circeM = 10; }
            }


            for (int i = 0; i < Gs.length(); i++) {
                int dateInt = getDate((String) ((JSONArray) Gs.get(i)).get(0), type);
                if (lastDate != 1 && dateInt == 1 && categoriesSoMedia.length > 0) {
                    circe = circe * circeM;
                }
                String date = String.valueOf(dateInt * circe);
                isContain = false;
                for (int j = 0; j < categoriesSoMedia.length; j++) {
                    if (categoriesSoMedia[j].equals(date)) {
                        valuesSo[j] += new Double(((JSONArray) Vk.get(i)).get(1).toString()) +
                                new Double(((JSONArray) Tw.get(i)).get(1).toString()) +
                                new Double(((JSONArray) Fb.get(i)).get(1).toString()) +
                                +new Double(((JSONArray) Tg.get(i)).get(1).toString()) +
                                +new Double(((JSONArray) Ig.get(i)).get(1).toString());
                        valuesMedia[j] += new Double(((JSONArray) Gs.get(i)).get(1).toString());
                        isContain = true;
                        break;
                    }
                }
                if (!isContain) {
                    categoriesSoMedia = append(categoriesSoMedia, date);
                    valuesSo = append(valuesSo, new Double(((JSONArray) Vk.get(i)).get(1).toString()) +
                            new Double(((JSONArray) Tw.get(i)).get(1).toString()) +
                            new Double(((JSONArray) Fb.get(i)).get(1).toString()) +
                            +new Double(((JSONArray) Tg.get(i)).get(1).toString()) +
                            +new Double(((JSONArray) Ig.get(i)).get(1).toString()));
                    valuesMedia = append(valuesMedia, new Double(((JSONArray) Gs.get(i)).get(1).toString()));

                }
                lastDate = dateInt;
            }
            changeWeekString(categoriesSoMedia, type, first_month, first_year);


        }
        return new ParseData(categoriesSoMedia, valuesMedia,valuesSo);
    }
    public static Integer getDate(String date, String type) throws ParseException {
        int res;

        Date dateDate = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").parse(date  + " 23:59:59");
        Calendar cal = Calendar.getInstance(TimeZone.getTimeZone("Europe/Paris"));
        cal.setTime(dateDate);
        if (type.equals("week")){
            res = cal.get(Calendar.WEEK_OF_YEAR);
        } else {
            if (type.equals("month")){
                res = cal.get(Calendar.MONTH) + 1;
            } else {
                res = cal.get(Calendar.MONTH) / 3 + 1;
            }
        }
        return res;
    }
    public static String[] changeWeekString(String[] categories, String type, int first_month, int first_year){
        String dateType;
        if (type.equals("week")) {
            dateType = " неделя";
        }
        else
        {
            if (type.equals("month")) {
                dateType = "месяц";
            }
            else {
                dateType = " квартал";
            }
        }

        if (dateType.equals("месяц")) {
            int i = 0;
            String[] monthNames = { "январь ", "февраль ", "март ", "апрель ", "май ", "июнь ", "июль ", "август ", "сентябрь ", "октябрь ", "ноябрь ", "декабрь " };
            for (int j = first_month; j < categories.length + first_month ; j++) {
                categories[i] = String.valueOf(monthNames[j%12]) + String.valueOf(first_year + j/12);
                i += 1;
            }
        } else {
            for (int j = 0; j < categories.length; j++) {
                categories[j] = String.valueOf(j + 1) + dateType;
            }}
        return categories;
    }

    private static void deleteBoarder(XWPFTable table) {
        table.setRightBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "");
        table.setLeftBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "");
        table.setTopBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "");
        table.setBottomBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "");
        table.setInsideVBorder(XWPFTable.XWPFBorderType.NONE, 0, 0, "");
    }


    private static void addCustomHeadingStyle(XWPFDocument docxDocument, String strStyleId, int headingLevel, boolean bold) {

        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(strStyleId);

        CTString styleName = CTString.Factory.newInstance();
        styleName.setVal(strStyleId);
        ctStyle.setName(styleName);

        CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
        indentNumber.setVal(BigInteger.valueOf(headingLevel));

        // lower number > style is more prominent in the formats bar
        ctStyle.setUiPriority(indentNumber);

        CTOnOff onoffnull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onoffnull);

        // style shows up in the formats bar
        ctStyle.setQFormat(onoffnull);

        // style defines a heading of the given level
        CTPPr ppr = CTPPr.Factory.newInstance();
        ppr.setOutlineLvl(indentNumber);
        ctStyle.setPPr(ppr);
        CTRPr rPr = ctStyle.addNewRPr();
        if (bold) {
            rPr.addNewB();
            rPr.addNewBCs();
        }

        XWPFStyle style = new XWPFStyle(ctStyle);

        XWPFStyles styles = docxDocument.createStyles();

        style.setType(STStyleType.PARAGRAPH);
        styles.addStyle(style);
    }

    private static String getOrNone(JSONObject j, String key) {
        try {
            String res = j.get(key).toString();
            if (res.equals("null")) {
                return "0";
            }
            return res;
        } catch (JSONException e){
            return "0";
        }
    }
    private static XWPFDocument addParagraph(XWPFDocument docxModel, String name) {
        return addParagraph_new(docxModel, name, false);
    }
    private static XWPFDocument addParagraph_new(XWPFDocument docxModel, String name, Boolean new_page) {

        XWPFParagraph paragraph = docxModel.createParagraph();
        paragraph.setStyle("Heading2");
        if (new_page) {
            paragraph.setPageBreak(true);
        }
        else{
            if ((entityOnPage != 0 && entityOnPage % 3 ==0)) {
                paragraph.setPageBreak(true);
            }}
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun paragraphConfig= paragraph.createRun();
        paragraphConfig.setFontSize(12);
        paragraphConfig.setFontFamily(format);
        paragraphConfig.setBold(true);
        paragraphConfig.addBreak();
        paragraphConfig.setText(
                name
        );
        entityOnPage+=1;
        return docxModel;
    }
    public static String updateText(String text){
        StringBuilder sb = new StringBuilder(text);
        deleteDataInString(sb, "<span.*?>");
        deleteDataInString(sb, "</span.*?>");
        deleteDataInString(sb, "<div.*?>");
        deleteDataInString(sb, "</div.*?>");
        String res = sb.toString();
        if (res.length() > commentsLenght) {
            res = res.substring(0, commentsLenght) + "...";
        }
        return res;
    }

    private static StringBuilder deleteDataInString(StringBuilder sb, String reg){
        Pattern p = Pattern.compile(reg, Pattern.CASE_INSENSITIVE);
        boolean stop = false;
        while (!stop)
        {
            Matcher m = p.matcher(sb.toString());
            if (m.find()) {
                sb.delete(m.start(), m.end());
            }
            else
                stop = true;
        }
        return sb;
    }
    private static void setText(XWPFTableRow tableRow, String data, int celNum) {
        setText(tableRow, data, celNum, false);
    }
    private static void setText(XWPFTableRow tableRow, String data, int celNum, boolean is_link){
        XWPFTableCell cell_two = tableRow.getCell(celNum);
        cell_two.removeParagraph(0);
        XWPFParagraph addParagraph_two = cell_two.addParagraph();
        if (is_link) {
            addParagraph_two.setAlignment(ParagraphAlignment.CENTER);
            addHyperlink(addParagraph_two, data, data);
        }
        else {
            XWPFRun run_two = addParagraph_two.createRun();
            run_two.setFontFamily(format);
            run_two.setFontSize(12);
            run_two.setText(data);
        }
    }

    private static String get_format_stng(int i) {
        return String.format(Locale.CANADA_FRENCH, "%,d", i);

    }
    private static void addHyperlink(XWPFParagraph para, String text, String bookmark) {
        //Create hyperlink in paragraph
        CTHyperlink cLink=para.getCTP().addNewHyperlink();
        cLink.setAnchor(bookmark);
        //Create the linked text
        CTText ctText=CTText.Factory.newInstance();
        if (text.length() > 30) {
            text = text.substring(0, 27) + "...";
        }
        ctText.setStringValue(text);
        CTR ctr=CTR.Factory.newInstance();
        ctr.setTArray(new CTText[]{ctText});

        //Create the formatting
        CTFonts fonts = CTFonts.Factory.newInstance();
        fonts.setAscii("Times New Roman" );
        CTRPr rpr = ctr.addNewRPr();
        CTColor colour = CTColor.Factory.newInstance();
        colour.setVal("0000FF");
        rpr.setColor(colour);
        rpr.setRFonts(fonts);
        CTHpsMeasure size = CTHpsMeasure.Factory.newInstance();
        size.setVal(new BigInteger("24"));
        rpr.setSz(size);
        CTRPr rpr1 = ctr.addNewRPr();
        rpr1.addNewU().setVal(STUnderline.SINGLE);


        //Insert the linked text into the link
        cLink.setRArray(new CTR[]{ctr});
    }

}
