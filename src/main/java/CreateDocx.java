import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.*;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.IOException;
import java.math.BigInteger;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

class WordWorker {
    private static  int entityOnPage = 0;
    private static int commentsLenght= 100;
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

    static <T> T[] append(T[] arr, T element) {
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

        try {
            XWPFDocument docxModel = new XWPFDocument();
            XWPFParagraph bodyParagraph = docxModel.createParagraph();
            bodyParagraph.setAlignment(ParagraphAlignment.RIGHT);
            XWPFRun paragraphConfig = bodyParagraph.createRun();
            paragraphConfig.setFontSize(22);
            paragraphConfig.setBold(true);
            paragraphConfig.setFontFamily("Arial");
            paragraphConfig.setText(
                    "Базовый отчет"
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
            paragraphConfigLenta.setFontFamily("Arial");
            paragraphConfigLenta.setText(
                    "Лента: "
            );

            XWPFParagraph bodyParagraphName = docxModel.createParagraph();
            bodyParagraphName.setAlignment(ParagraphAlignment.LEFT);
            XWPFRun paragraphConfigName = bodyParagraphName.createRun();
            paragraphConfigName.setFontSize(26);
            paragraphConfigName.setBold(true);
            paragraphConfigName.setFontFamily("Arial");
            paragraphConfigName.setText(
                    name

            );
            paragraphConfigName.addBreak();

            XWPFParagraph bodyParagraphAnalyze = docxModel.createParagraph();
            bodyParagraphAnalyze.setAlignment(ParagraphAlignment.LEFT);
            XWPFRun paragraphConfigAnalyze = bodyParagraphAnalyze.createRun();
            paragraphConfigAnalyze.setFontSize(14);
            paragraphConfigAnalyze.setFontFamily("Arial");
            paragraphConfigAnalyze.setText("Аналитический отчет по упоминаниям в онлайн-СМИ и соцмедиа");
            paragraphConfigAnalyze.addBreak();


            XWPFParagraph bodyParagraphDate = docxModel.createParagraph();
            bodyParagraphDate.setAlignment(ParagraphAlignment.LEFT);
            XWPFRun paragraphConfigDate = bodyParagraphDate.createRun();
            paragraphConfigDate.setFontSize(14);
            paragraphConfigDate.setFontFamily("Arial");
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
            run.setFontFamily("Arial");

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
            paragraphConfigStatic.setFontFamily("Arial");
            paragraphConfigStatic.setText(
                    "Базовые статистики"
            );

            XWPFTable table = docxModel.createTable();
            deleteBoarder(table);

            XWPFTableRow tableRowOne = table.getRow(0);
            run = tableRowOne.getCell(0).getParagraphs().get(0).createRun();
            run.setText("Совокупная аудитория, чел.");
            tableRowOne.addNewTableCell().setText(String.valueOf(users));

            XWPFTableRow tableRowTwo = table.createRow();
            XWPFRun run1 = tableRowTwo.getCell(0).getParagraphs().get(0).createRun();
            run1.setText("Количество источников публикаций, шт.");
            tableRowTwo.getCell(1).setText(String.valueOf(data.total_sources));

            XWPFTableRow tableRowThree = table.createRow();
            XWPFRun run2 = tableRowThree.getCell(0).getParagraphs().get(0).createRun();
            run2.setText("Количество публикаций, шт.");
            tableRowThree.getCell(1).setText(String.valueOf(data.total_publication));

            XWPFTableRow tableRowFour = table.createRow();
            XWPFRun run3 = tableRowFour.getCell(0).getParagraphs().get(0).createRun();
            run3.setText("Количество комментариев к публикациям, шт.");
            tableRowFour.getCell(1).setText(String.valueOf(data.total_comment));

            XWPFTableRow tableRow4 = table.createRow();
            XWPFRun run4_1 = tableRow4.getCell(0).getParagraphs().get(0).createRun();
            run4_1.setText("Количество просмотров, шт.");
            tableRow4.getCell(1).setText(String.valueOf(data.total_views));

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
            diagramCount = addChats(docxModel, postData.categories, postData.valuesA, String.format("Диаграмма %s Динамика количества публикаций", diagramCount),  diagramCount);


            ParseData comments = getWeekData(type, (JSONArray) (jsonComments).get("total"), first_month, first_year);
            diagramCount= addChats(docxModel, comments.categories, comments.valuesA, String.format("Диаграмма %s Динамика количества комментариев к публикациям", diagramCount), diagramCount);
            String postDate;

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
            diagramCount = addPie(docxModel, new String[]{"Нейтральная", "Позитивная", "Негативная"}, new Double[]{(double) getComment(jsonPostTotal, "netural"),
                    (double) getComment(jsonPostTotal, "positive"), (double) getComment(jsonPostTotal, "negative")}, String.format("Диаграмма %s Тональность публикаций", diagramCount), diagramCount, true);


            String[] categoriesPostType = new String[]{};
            Double[] valuesNegative = new Double[]{};
            Double[] valuesPositive = new Double[]{};
            Double[] valuesNetural = new Double[]{};

            JSONArray positive = (JSONArray) (jsonPostTotal).get("positive");
            JSONArray netural = (JSONArray) (jsonPostTotal).get("netural");
            JSONArray negative = (JSONArray) (jsonPostTotal).get("negative");
            JSONArray totalComments = ((JSONArray) (jsonPostTotal).get("total"));

            double positiveInt;
            double neturalInt;
            double negativeInt;
            double sum;
            if (type.equals("day")) {
                for (int i = 0; i < totalComments.length(); i++) {
                    //            for (int i =0; i<31; i++) {

                    negativeInt = new Double(((JSONArray) negative.get(i)).get(1).toString());
                    positiveInt = new Double(((JSONArray) positive.get(i)).get(1).toString());
                    neturalInt = new Double(((JSONArray) netural.get(i)).get(1).toString());
                    sum = negativeInt + neturalInt + positiveInt;
                    categoriesPostType = append(categoriesPostType, (String) ((JSONArray) negative.get(i)).get(0));
                    if (sum == 0) {
                        valuesNegative = append(valuesNegative, 33d);
                        // 033
                        valuesPositive = append(valuesPositive, 33d);
                        // 033
                        valuesNetural = append(valuesNetural, 33d);

                    } else {
                        valuesNegative = append(valuesNegative, (double) Math.round(negativeInt / sum * 100));
                        // 033
                        valuesPositive = append(valuesPositive, (double) Math.round(positiveInt / sum * 100));
                        // 033
                        valuesNetural = append(valuesNetural, (double) Math.round(neturalInt / sum * 100));
                    }

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
                    } else {
                        circeM = 10;
                    }
                }
                for (int i = 0; i < totalComments.length(); i++) {

                    negativeInt = new Double(((JSONArray) negative.get(i)).get(1).toString());
                    positiveInt = new Double(((JSONArray) positive.get(i)).get(1).toString());
                    neturalInt = new Double(((JSONArray) netural.get(i)).get(1).toString());
                    int dateInt = getDate((String) ((JSONArray) negative.get(i)).get(0), type);
                    if (lastDate != 1 && dateInt == 1 && categoriesPostType.length > 0) {
                        circe = circe * circeM;
                    }
                    String dateSo = String.valueOf(dateInt * circe);
                    isContain = false;
                    for (int j = 0; j < categoriesPostType.length; j++) {
                        if (categoriesPostType[j].equals(dateSo)) {
                            valuesNegative[j] += negativeInt;
                            valuesPositive[j] += positiveInt;
                            valuesNetural[j] += neturalInt;
                            isContain = true;
                            break;
                        }
                    }
                    if (!isContain) {
                        categoriesPostType = append(categoriesPostType, dateSo);
                        valuesNegative = append(valuesNegative, 100.00 * negativeInt);
                        // 033
                        valuesPositive = append(valuesPositive, 100.00 * positiveInt);
                        // 033
                        valuesNetural = append(valuesNetural, 100.00 * neturalInt);
                    }
                    lastDate = dateInt;

                }
                for (int j = 0; j < categoriesPostType.length; j++) {
                    if (valuesNegative[j] == 0 && valuesPositive[j] == 0 && valuesNetural[j] == 0) {
                        valuesNegative[j] = 33d;
                        valuesPositive[j] = 33d;
                        valuesNetural[j] = 33d;
                    } else {
                        sum = valuesNegative[j] + valuesPositive[j] + valuesNetural[j];
                        valuesNegative[j] =
                                (double) Math.round((valuesNegative[j] / sum) * 100.0);
                        valuesPositive[j] =
                                (double) Math.round((valuesPositive[j] / sum) * 100.0);
                        valuesNetural[j] =
                                (double) Math.round((valuesNetural[j] / sum) * 100.0);
                    }
                }
                changeWeekString(categoriesPostType, type, first_month, first_year);
            }

            diagramCount= addArea(docxModel, categoriesPostType,
                    valuesNegative,
                    valuesPositive,
                    valuesNetural,
                    String.format("Диаграмма %s Динамика распределения публикаций по тональности", diagramCount), diagramCount);

            int total_vk = getTotalMedia(jsonPosts, "vk");
            int total_tw = getTotalMedia(jsonPosts, "tw");
            int total_fb = getTotalMedia(jsonPosts, "fb");
            int total_gs = getTotalMedia(jsonPosts, "gs");
            int total_tg = getTotalMedia(jsonPosts, "tg");
            int total_ig = getTotalMedia(jsonPosts, "ig");
            int all = total_vk + total_tw + total_fb + total_gs + total_tg + total_ig;


            ParseData soData = getWeekDataMedia(type, jsonPosts, first_month, first_year);
            double val = 0;
            for (Double d:soData.valuesA){
                val += d;
            }
            for (Double d:soData.valuesB){
                val += d;
            }

            if ((all != 0) || (val != 0) || jsonCity.length() == 0) {

                XWPFParagraph bodyParagraphIst = docxModel.createParagraph();
                bodyParagraphIst.setPageBreak(true);
                bodyParagraphIst.setStyle("Heading1");
                bodyParagraphIst.setAlignment(ParagraphAlignment.LEFT);
                XWPFRun paragraphConfigIst = bodyParagraphIst.createRun();
                paragraphConfigIst.setFontSize(22);
                paragraphConfigIst.setBold(true);
                paragraphConfigIst.setFontFamily("Arial");
                paragraphConfigIst.setText(
                        "Источники"
                );
                entityOnPage = 0;
                if (all == 0) {
                    dataLost(docxModel);
                } else {
                    addParagraph(docxModel, String.format("Таблица %s Ключевые площадки", tableCount));
                    tableCount += 1;
                    XWPFTable tableIst = docxModel.createTable();
                    XWPFTableRow tableRowOneIst = tableIst.getRow(0);
                    XWPFRun run4 = tableRowOneIst.getCell(0).getParagraphs().get(0).createRun();
                    run4.setText("Площадка");
                    run4.setBold(true);
                    tableRowOneIst.addNewTableCell();
                    XWPFRun run5 = tableRowOneIst.getCell(1).getParagraphs().get(0).createRun();
                    run5.setText("Количество публикаций, шт.");
                    run5.setBold(true);
                    tableRowOneIst.addNewTableCell();
                    XWPFRun run6 = tableRowOneIst.getCell(2).getParagraphs().get(0).createRun();
                    run6.setText("   %     ");
                    run6.setBold(true);

                    XWPFTableRow tableRowTwoIst = tableIst.createRow();
                    tableRowTwoIst.getCell(0).setText("Вконтакте");
                    tableRowTwoIst.getCell(1).setText(String.valueOf(total_vk));
                    tableRowTwoIst.getCell(2).setText(String.valueOf(Math.round((float) total_vk * 100.00 / (float) all * 100.00) / 100.0));

                    XWPFTableRow tableRowThreeIst = tableIst.createRow();
                    tableRowThreeIst.getCell(0).setText("Facebook");
                    tableRowThreeIst.getCell(1).setText(String.valueOf(total_fb));
                    tableRowThreeIst.getCell(2).setText(String.valueOf(Math.round((float) total_fb * 100 / (float) all * 100.00) / 100.0));

                    XWPFTableRow tableRowThIst = tableIst.createRow();
                    tableRowThIst.getCell(0).setText("Twitter");
                    tableRowThIst.getCell(1).setText(String.valueOf(total_tw));
                    tableRowThIst.getCell(2).setText(String.valueOf(Math.round((float) total_tw * 100 / (float) all * 100.00) / 100.0));

                    XWPFTableRow tableRowFIst = tableIst.createRow();
                    tableRowFIst.getCell(0).setText("Инстаграм");
                    tableRowFIst.getCell(1).setText(String.valueOf(total_ig));
                    tableRowFIst.getCell(2).setText(String.valueOf(Math.round((float) total_ig * 100 / (float) all * 100.00) / 100.0));

                    XWPFTableRow tableRowSixIst = tableIst.createRow();
                    tableRowSixIst.getCell(0).setText("Telegram");
                    tableRowSixIst.getCell(1).setText(String.valueOf(total_tg));
                    tableRowSixIst.getCell(2).setText(String.valueOf(Math.round((float) total_tg * 100 / (float) all * 100.00) / 100.0));

                    XWPFTableRow tableRowSevIst = tableIst.createRow();
                    tableRowSevIst.getCell(0).setText("СМИ");
                    tableRowSevIst.getCell(1).setText(String.valueOf(total_gs));
                    tableRowSevIst.getCell(2).setText(String.valueOf(Math.round((float) total_gs * 100 / (float) all * 100.00) / 100.0));

                    XWPFTableRow tableRowSevAll = tableIst.createRow();
                    tableRowSevAll.getCell(0).setText("Итог");
                    tableRowSevAll.getCell(1).setText(String.valueOf(all));
                    tableRowSevAll.getCell(2).setText("100");


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
                    }
                }


                diagramCount = addDоubleChats(docxModel, soData.categories, soData.valuesA, soData.valuesB,
                        String.format("Диаграмма %s Динамика количества публикаций на отдельных площадках", diagramCount), diagramCount);

                if (posts.length() == 0) {
                    dataLost(docxModel);
                } else {
                    addParagraph(docxModel, String.format("Таблица %s Топ-%s источников по количеству публикаций", tableCount, posts.length()));
                    tableCount += 1;
                    XWPFTable tableTop10Own = docxModel.createTable();
                    XWPFTableRow tableTop10OwnRow = tableTop10Own.getRow(0);

                    XWPFRun run12 = tableTop10OwnRow.getCell(0).getParagraphs().get(0).createRun();
                    run12.setText("Название источника");
                    run12.setBold(true);

                    tableTop10OwnRow.addNewTableCell();
                    XWPFRun run11 = tableTop10OwnRow.getCell(1).getParagraphs().get(0).createRun();
                    run11.setText("URL");
                    run11.setBold(true);
                    tableTop10OwnRow.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                    tableTop10OwnRow.addNewTableCell();
                    XWPFRun run13 = tableTop10OwnRow.getCell(2).getParagraphs().get(0).createRun();
                    run13.setText("Количество публикаций");
                    run13.setBold(true);

                    JSONObject jsonObject;
                    for (Object o : posts) {
                        jsonObject = (JSONObject) o;
                        getRow(tableTop10Own, jsonObject.get("username").toString(), jsonObject.get("url").toString(),
                                jsonObject.get("coefficient").toString());
                    }


                    for (int x = 0; x < tableTop10Own.getNumberOfRows(); x++) {
                        XWPFTableRow row = tableTop10Own.getRow(x);
                        XWPFTableCell cell0 = row.getCell(0);
                        cell0.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(5000));
                        XWPFTableCell cell1 = row.getCell(1);
                        cell1.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(3000));
                        cell1.getParagraphs().get(0).setAlignment(ParagraphAlignment.LEFT);
                        XWPFTableCell cell2 = row.getCell(2);
                        cell2.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1500));
                        cell2.getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                    }
                }
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


            if ((valSex != 0) || (valAudit != 0) || (valAge != 0) || (valCity != 0 ) || ( jsonCity.length() != 0)
                    || (jsonCity.length() != 0) || ((JSONArray) usersJson.get("users")).length() !=0) {
                XWPFParagraph bodyParagraphAudit = docxModel.createParagraph();
                bodyParagraphAudit.setAlignment(ParagraphAlignment.LEFT);
                bodyParagraphAudit.setPageBreak(true);
                bodyParagraphAudit.setStyle("Heading1");
                XWPFRun paragraphConfigAudit = bodyParagraphAudit.createRun();
                paragraphConfigAudit.setFontSize(22);
                paragraphConfigAudit.setBold(true);
                paragraphConfigAudit.setFontFamily("Arial");
                paragraphConfigAudit.setText(
                        "Аудитория"
                );
                entityOnPage = 0;
                diagramCount = addChats(docxModel, auditData.categories, auditData.valuesA, String.format("Диаграмма %s Динамика объема аудитории", diagramCount), diagramCount);


                diagramCount = addPie(docxModel, new String[]{"Не указан", "Мужчины", "Женщины"}, masSex, String.format("Диаграмма %s Распределение аудитории по полу", diagramCount), diagramCount);


                diagramCount = addPie(docxModel, new String[]{"18-25 лет", "26-40 лет", "40 лет и старше", "не указан"},
                        masAge, String.format("Диаграмма %s Распределение аудитории по возрасту", diagramCount), diagramCount);


                diagramCount = addPieCity(docxModel, categoriesCity, valuesACity, String.format("Диаграмма %s Распределение аудитории по геолокации", diagramCount), diagramCount);

                if (jsonCity.length() == 0) {
                    dataLost(docxModel);
                } else {
                    addParagraph(docxModel, String.format("Таблица %s Топ-%s городов", tableCount, jsonCity.length()));
                    tableCount += 1;
                    XWPFTable tableTop10OCity = docxModel.createTable();
                    XWPFTableRow tableTop10OCityRow = tableTop10OCity.getRow(0);
                    XWPFRun runCity = tableTop10OCityRow.getCell(0).getParagraphs().get(0).createRun();
                    tableTop10OCityRow.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    runCity.setText("Город");
                    runCity.setBold(true);

                    tableTop10OCityRow.addNewTableCell();
                    XWPFRun r9 = tableTop10OCityRow.getCell(1).getParagraphs().get(0).createRun();
                    r9.setText("Количество");
                    r9.setBold(true);
                    tableTop10OCityRow.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);


                    tableTop10OCityRow.addNewTableCell();
                    XWPFRun run10a = tableTop10OCityRow.getCell(2).getParagraphs().get(0).createRun();
                    run10a.setText("%");
                    run10a.setBold(true);
                    tableTop10OCityRow.getCell(2).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                    try {
                        for (Object o : jsonCity) {
                            jsonObject = (JSONObject) o;
                            getRow(tableTop10OCity, jsonObject.get("city").toString(), jsonObject.get("users").toString(),
                                    String.format("%.1f", Double.parseDouble(jsonObject.get("users").toString()) * 100.0 / Double.valueOf(count10)));
                        }
                    } catch (Exception e) {
                        System.out.println("S");
                    }

                    for (int x = 0; x < tableTop10OCity.getNumberOfRows(); x++) {
                        XWPFTableRow row000 = tableTop10OCity.getRow(x);
                        XWPFTableCell cell0000 = row000.getCell(0);
                        cell0000.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4000));
                        XWPFTableCell cell1000 = row000.getCell(1);
                        cell1000.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4000));
                        cell1000.getParagraphs().get(0).setAlignment(ParagraphAlignment.LEFT);
                        XWPFTableCell cell2000 = row000.getCell(2);
                        cell2000.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1500));
                        cell2000.getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                    }
                }

                if (((JSONArray) usersJson.get("users")).length() == 0) {
                    dataLost(docxModel);
                } else {
                    addParagraph(docxModel, String.format("Таблица %s Топ-%s активных пользователей по сумме реакции (лайков, комментариев, репостов)", tableCount, ((JSONArray) usersJson.get("users")).length()));
                    tableCount += 1;
                    XWPFTable tableTop10OUser = docxModel.createTable();
                    XWPFTableRow tableTop10OUserRow = tableTop10OUser.getRow(0);
                    XWPFRun run8 = tableTop10OUserRow.getCell(0).getParagraphs().get(0).createRun();
                    tableTop10OUserRow.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    run8.setText("Пользователь");
                    run8.setBold(true);

                    tableTop10OUserRow.addNewTableCell();
                    XWPFRun run9 = tableTop10OUserRow.getCell(1).getParagraphs().get(0).createRun();
                    run9.setText("URL");
                    run9.setBold(true);
                    tableTop10OUserRow.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                    tableTop10OUserRow.addNewTableCell();
                    XWPFRun run10 = tableTop10OUserRow.getCell(2).getParagraphs().get(0).createRun();
                    run10.setText("Сумма реакции");
                    run10.setBold(true);
                    tableTop10OUserRow.getCell(2).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                    for (Object o : (JSONArray) usersJson.get("users")) {
                        jsonObject = (JSONObject) o;
                        getRow(tableTop10OUser, jsonObject.get("name").toString(), jsonObject.get("url").toString(),
                                jsonObject.get("coefficient").toString());
                    }

                    for (int x = 0; x < tableTop10OUser.getNumberOfRows(); x++) {
                        XWPFTableRow row = tableTop10OUser.getRow(x);
                        XWPFTableCell cell0 = row.getCell(0);
                        cell0.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4000));
                        XWPFTableCell cell1 = row.getCell(1);
                        cell1.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4000));
                        cell1.getParagraphs().get(0).setAlignment(ParagraphAlignment.LEFT);
                        XWPFTableCell cell2 = row.getCell(2);
                        cell2.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1500));
                        cell2.getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                    }
                }
            }
            int likesPosts = 0;
            int likesComment = 0;

            for (Object o : postsContent) {
                jsonObject = (JSONObject) o;
                if (Integer.parseInt(jsonObject.get("likes").toString()) > 0) {
                    likesPosts+=1;
                }
            }
            for (Object o : postsContent) {
                jsonObject = (JSONObject) o;
                if (Integer.parseInt(((JSONObject) o).get("likes").toString()) > 0) {
                    likesComment+=1;
                }
            }
            for (Object o : commentContent) {
                if (Integer.parseInt(((JSONObject) o).get("likes").toString()) > 0) {
                    likesComment+=1;
                }
            }
            if((likesComment !=0) || (likesPosts !=0 )) {
                XWPFParagraph bodyParagrapKeysP = docxModel.createParagraph();
                bodyParagrapKeysP.setAlignment(ParagraphAlignment.LEFT);
                bodyParagrapKeysP.setPageBreak(true);
                bodyParagrapKeysP.setStyle("Heading1");
                XWPFRun paragraphConfigKeysP = bodyParagrapKeysP.createRun();
                paragraphConfigKeysP.setFontSize(22);
                paragraphConfigKeysP.setBold(true);
                paragraphConfigKeysP.setFontFamily("Arial");
                paragraphConfigKeysP.setText(
                        "Ключевые публикации и комментарии"
                );
                entityOnPage =0;
                if (likesPosts == 0) {
                    dataLost(docxModel);
                } else {
                    addParagraph(docxModel, String.format("Таблица %s Топ-%s публикаций по сумме резонанса", tableCount, likesPosts));
                    tableCount += 1;
                    XWPFTable tableTop10Post = docxModel.createTable();
                    XWPFTableRow tableTop10PostRow = tableTop10Post.getRow(0);


                    XWPFRun run15 = tableTop10PostRow.getCell(0).getParagraphs().get(0).createRun();
                    run15.setText("Публикация");
                    run15.setBold(true);

                    tableTop10PostRow.addNewTableCell();
                    XWPFRun run16 = tableTop10PostRow.getCell(1).getParagraphs().get(0).createRun();
                    run16.setText("URL");
                    run16.setBold(true);
                    tableTop10PostRow.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                    tableTop10PostRow.addNewTableCell();
                    XWPFRun run17 = tableTop10PostRow.getCell(2).getParagraphs().get(0).createRun();
                    run17.setText("Резонанс");
                    run17.setBold(true);


                    String text;

                    for (Object o : postsContent) {
                        jsonObject = (JSONObject) o;
                        text = jsonObject.get("text").toString();
                        if (text.length() > commentsLenght) {
                            text = jsonObject.get("text").toString().substring(0, commentsLenght);
                        }
                        getRow(tableTop10Post, text, jsonObject.get("uri").toString(),
                                String.valueOf(Integer.parseInt(jsonObject.get("viewed").toString()) + Integer.parseInt(jsonObject.get("reposts").toString()) +
                                        Integer.parseInt(jsonObject.get("likes").toString()) + Integer.parseInt(jsonObject.get("comments").toString())));

                    }
                    for (int x = 0; x < tableTop10Post.getNumberOfRows(); x++) {
                        XWPFTableRow row = tableTop10Post.getRow(x);
                        XWPFTableCell cell0 = row.getCell(0);
                        cell0.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(5000));
                        XWPFTableCell cell1 = row.getCell(1);
                        cell1.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(3000));
                        cell1.getParagraphs().get(0).setAlignment(ParagraphAlignment.LEFT);
                        XWPFTableCell cell2 = row.getCell(2);
                        cell2.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1500));
                        cell2.getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);

                    }
                }

                if (likesComment== 0) {
                    dataLost(docxModel);
                } else {
                    addParagraph(docxModel, String.format("Таблица %s Топ-%s комментариев к публикациям по сумме лайков", tableCount, likesComment));
                    tableCount += 1;
                    String text;
                    XWPFTable tableTop10Comment = docxModel.createTable();
                    XWPFTableRow tableTop10CommentRow = tableTop10Comment.getRow(0);
                    XWPFRun run19 = tableTop10CommentRow.getCell(0).getParagraphs().get(0).createRun();
                    run19.setText("Комментарий");
                    run19.setBold(true);

                    tableTop10CommentRow.addNewTableCell();
                    XWPFRun run20 = tableTop10CommentRow.getCell(1).getParagraphs().get(0).createRun();
                    run20.setText("URL");
                    run20.setBold(true);
                    tableTop10CommentRow.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                    tableTop10CommentRow.addNewTableCell();
                    XWPFRun run21 = tableTop10CommentRow.getCell(2).getParagraphs().get(0).createRun();
                    run21.setText("Резонанс");
                    run21.setBold(true);

                    for (Object o : commentContent) {
                        jsonObject = (JSONObject) o;
                        text = jsonObject.get("text").toString();
                        if (text.length() > commentsLenght) {
                            text = jsonObject.get("text").toString().substring(0, commentsLenght);
                        }

                        getRow(tableTop10Comment, text, jsonObject.get("post_url").toString(),
                                jsonObject.get("likes").toString());
                    }

                    for (int x = 0; x < tableTop10Comment.getNumberOfRows(); x++) {
                        XWPFTableRow row = tableTop10Comment.getRow(x);
                        XWPFTableCell cell0 = row.getCell(0);
                        cell0.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(5000));
                        XWPFTableCell cell1 = row.getCell(1);
                        cell1.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(3000));
                        cell1.getParagraphs().get(0).setAlignment(ParagraphAlignment.LEFT);
                        XWPFTableCell cell2 = row.getCell(2);
                        cell2.getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1500));
                        cell2.getParagraphs().get(0).setAlignment(ParagraphAlignment.CENTER);
                    }
                }
            }
            CTP ctp = CTP.Factory.newInstance();
//this add page number incremental
            ctp.addNewR().addNewPgNum();

            XWPFParagraph codePara = new XWPFParagraph(ctp, docxModel);
            XWPFParagraph[] paragraphs = new XWPFParagraph[1];
            paragraphs[0] = codePara;
//position of number
            codePara.setAlignment(ParagraphAlignment.CENTER);

            CTSectPr sectPr = docxModel.getDocument().getBody().addNewSectPr();

            XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(docxModel, sectPr);
            headerFooterPolicy.createFooter(STHdrFtr.DEFAULT, paragraphs);
            CTSectPr sect = docxModel.getDocument().getBody().getSectPr();
            sect.addNewTitlePg();
            // сохраняем модель docx документа в файл
//            try (FileOutputStream fileOut = new FileOutputStream("/home/oleg/Documents/test1t.docx")) {
//                docxModel.write(fileOut);
//            }

            return docxModel;

        } catch (Exception e) {
            e.printStackTrace();
        }
        return new XWPFDocument();

    }

    private static int getTotalMedia(JSONObject jsonPosts, String key){
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
        tableRowTwoIst.getCell(0).setText(str1);
        tableRowTwoIst.getCell(1).setText(str2);
        tableRowTwoIst.getCell(2).setText(str3);
    }

    private static int getComment(JSONObject jsonComments, String key) {
        int res = 0;
        JSONArray jsonArray;
        for (Object o : (JSONArray) (jsonComments).get(key)) {
            jsonArray = (JSONArray) o;
            res += (int) jsonArray.get(1);
        }
        return res;
    }

    private static CTP createFooterModel(String footerContent) {
        // создаем футер или нижний колонтитул
        CTP ctpFooterModel = CTP.Factory.newInstance();
        CTR ctrFooterModel = ctpFooterModel.addNewR();
        CTText cttFooter = ctrFooterModel.addNewT();

        cttFooter.setStringValue(footerContent);
        return ctpFooterModel;
    }

    private static CTP createHeaderModel(String headerContent) {
        // создаем хедер или верхний колонтитул
        CTP ctpHeaderModel = CTP.Factory.newInstance();
        CTR ctrHeaderModel = ctpHeaderModel.addNewR();
        CTText cttHeader = ctrHeaderModel.addNewT();

        cttHeader.setStringValue(headerContent);
        return ctpHeaderModel;
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


//        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(255, 209, 48));
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


        //PERCENTE
//        chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(0).getDLbls().addNewShowPercent()

        chart.plot(data);

//        // create legend
//        XDDFChartLegend legend = chart.getOrAddLegend();
//        legend.setPosition(LegendPosition.LEFT);
//        legend.setOverlay(false);

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
//        leftAxis.getOrAddMajorGridProperties();
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

        // create series
        // if only one series do not vary colors for each bar
        ((XDDFBarChartData) data).setVaryColors(false);
        XDDFChartData.Series series = data.addSeries(categoriesData, valuesDataA);
        XDDFChartData.Series series2 = data.addSeries(categoriesData, valuesDataB);

        series.setTitle("Сми", null);
        series2.setTitle("Соцмедиа", null);


//        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(255, 209, 48));
        XDDFSolidFillProperties fill = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.BLUE_VIOLET));

        XDDFShapeProperties properties = series.getShapeProperties();
        if (properties == null) {
            properties = new XDDFShapeProperties();
        }
        properties.setFillProperties(fill);
        series.setShapeProperties(properties);

        fill = new XDDFSolidFillProperties(XDDFColor.from(PresetColor.CORNFLOWER_BLUE));

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
        }

        chart.plot(data);

        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.TOP);
//        legend.setOverlay(true);

        return dia;
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

        // Set AxisCrossBetween, so the left axis crosses the category axis between the categories.
        // Else first and last category is exactly on cross points and the bars are only half visible.
//        Method andPrivateMethod
//                = XDDFDoughnutChartData.class.XDDFDoughnutChartData(
//                "privateAnd", boolean.class, boolean.class);
//        new XDDFDoughnutChartData(1, null);
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
                            .addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(DefaultIndexedColorMap.getDefaultRGB(11));
                }
                if (p == 2) {
                    chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDPtArray(p)
                            .addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(DefaultIndexedColorMap.getDefaultRGB(10));
                }
                if (p == 0) {
                    chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDPtArray(p)
                            .addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(DefaultIndexedColorMap.getDefaultRGB(23));
                }
            }
        }
//        if (chart.getCTChart().getAutoTitleDeleted() == null) chart.getCTChart().addNewAutoTitleDeleted();
//        chart.getCTChart().getAutoTitleDeleted().setVal(false);
        chart.plot(data);



//        XDDFChartLegend legend = chart.getOrAddLegend();
//        legend.setPosition(LegPosition.RIGHT);
//        legend.setOverlay(true);
        return dia;
    }

    private static int addPieCity(XWPFDocument document,  String[] categories, Double[] valuesA, String name, int dia) throws IOException, InvalidFormatException {
        // create the data

        // create the chart

        for (int i=0; i < categories.length; i++) {
            categories[i] += "; " + String.format("%.1f", valuesA[i]) + "%";
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

        // Set AxisCrossBetween, so the left axis crosses the category axis between the categories.
        // Else first and last category is exactly on cross points and the bars are only half visible.
//        Method andPrivateMethod
//                = XDDFDoughnutChartData.class.XDDFDoughnutChartData(
//                "privateAnd", boolean.class, boolean.class);
//        new XDDFDoughnutChartData(1, null);
        XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
        XDDFChartData.Series series = data.addSeries(categoriesData, valuesDataA);
        data.setVaryColors(true);
//        series.setShowLeaderLines(false);
        series.setTitle("", chart.setSheetTitle("", 1));

        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).addNewDLbls();
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowVal().setVal(false);
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowLegendKey().setVal(true);
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowCatName().setVal(true);
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowSerName().setVal(false);
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowBubbleSize().setVal(false);
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowPercent().setVal(false);
        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDLbls().addNewShowLeaderLines().setVal(true);
                //            int pointCount = series.getCategoryData().getPointCount();
//            for (int p = 0; p < pointCount; p++) {
//                chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).addNewDPt().addNewIdx().setVal(p);
//                if (p == 1) {
//                    chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDPtArray(p)
//                            .addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(DefaultIndexedColorMap.getDefaultRGB(11));
//                }
//                if (p == 2) {
//                    chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDPtArray(p)
//                            .addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(DefaultIndexedColorMap.getDefaultRGB(10));
//                }
//                if (p == 0) {
//                    chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).getDPtArray(p)
//                            .addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(DefaultIndexedColorMap.getDefaultRGB(23));
//                }
//            }

//       chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).addNewIdx().setVal(1);
//        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).(p)
////                    .addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(DefaultIndexedColorMap.getDefaultRGB(p+10));
//        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).addNewIdx()  addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(DefaultIndexedColorMap.getDefaultRGB(2));
//                chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(DefaultIndexedColorMap.getDefaultRGB(3));
//                        chart.getCTChart().getPlotArea().getPieChartArray(0).getSerArray(0).addNewSpPr().addNewSolidFill().addNewSrgbClr().setVal(DefaultIndexedColorMap.getDefaultRGB(4));
        chart.plot(data);



//        XDDFChartLegend legend = chart.getOrAddLegend();
//        legend.setPosition(LegPosition.RIGHT);
//        legend.setOverlay(true);
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
//            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowVal().setVal(true);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowPercent().setVal(true);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowLegendKey().setVal(false);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowCatName().setVal(false);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowSerName().setVal(false);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowBubbleSize().setVal(true);
            chart.getCTChart().getPlotArea().getBarChartArray(0).getSerArray(i).getDLbls().addNewShowVal().setVal(true);
        }

        data.setVaryColors(true);
        chart.plot(data);

        XDDFChartLegend legend = chart.getOrAddLegend();
//        legend.setPosition(LegendPosition.TOP);
        legend.setPosition(LegendPosition.TOP);
//        legend.set
//        legend.setOverlay(true);
        return dia;
    }


    private static ParseData getWeekData(String type, JSONArray jsonTotal, int first_month, int first_year) throws ParseException {
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

    private static ParseData getWeekDataMedia(String type, JSONObject jsonPosts, int first_month, int first_year) throws ParseException {
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
    private static Integer getDate(String date, String type) throws ParseException {
        int res;
        Date dateDate = new SimpleDateFormat("yyyy-MM-dd").parse(date);
        Calendar cal = Calendar.getInstance(TimeZone.getTimeZone("Europe/Paris"));
        cal.setTime(dateDate);
        if (type.equals("week")){
            res = cal.get(Calendar.WEEK_OF_YEAR);
        } else {
            if (type.equals("month")){
                res = cal.get(Calendar.MONTH);
            } else {
                res = cal.get(Calendar.MONTH) / 3 + 1;
            }
        }
        return res;
    }
    private static String[] changeWeekString(String[] categories, String type, int first_month, int first_year){
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

     XWPFParagraph paragraph = docxModel.createParagraph();
     paragraph.setStyle("Heading2");
     if (entityOnPage != 0 && entityOnPage % 3 ==0) {
        paragraph.setPageBreak(true);
     }
     paragraph.setAlignment(ParagraphAlignment.LEFT);
     XWPFRun paragraphConfig= paragraph.createRun();
     paragraphConfig.setFontSize(12);
     paragraphConfig.setFontFamily("Arial");
     paragraphConfig.setBold(true);
     paragraphConfig.addBreak();
     paragraphConfig.setText(
             name
     );
     entityOnPage+=1;
     return docxModel;
 }
}
