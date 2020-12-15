import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.*;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTBoolean;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.CTPlotArea;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

class WordWorker {

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


    public static void createDoc(String name, String date, DataForDocx data, JSONObject jsonPosts, JSONObject jsonComments) {
        try {
            XWPFDocument docxModel = new XWPFDocument();
            CTSectPr ctSectPr = docxModel.getDocument().getBody().addNewSectPr();
            // получаем экземпляр XWPFHeaderFooterPolicy для работы с колонтитулами
            XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(docxModel, ctSectPr);

            // создаем верхний колонтитул Word файла
            CTP ctpHeaderModel = createHeaderModel(
                    "Верхний колонтитул - создано с помощью Apache POI на Java :)"
            );
            // устанавливаем сформированный верхний
            // колонтитул в модель документа Word
            XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeaderModel, docxModel);
            headerFooterPolicy.createHeader(
                    XWPFHeaderFooterPolicy.DEFAULT,
                    new XWPFParagraph[]{headerParagraph}
            );

//             создаем нижний колонтитул docx файла
            CTP ctpFooterModel = createFooterModel(
                    "На базе данных программного обеспечения собственной разработки ООО «SNIPR»");
//             устанавливаем сформированый нижний
//             колонтитул в модель документа Word
            XWPFParagraph footerParagraph = new XWPFParagraph(ctpFooterModel, docxModel);
            headerFooterPolicy.createFooter(
                    XWPFHeaderFooterPolicy.DEFAULT,
                    new XWPFParagraph[]{footerParagraph}
            );


            // создаем обычный параграф, который будет расположен слева,
            // будет синим курсивом со шрифтом 25 размера
            XWPFParagraph bodyParagraph = docxModel.createParagraph();
            bodyParagraph.setAlignment(ParagraphAlignment.RIGHT);
            XWPFRun paragraphConfig = bodyParagraph.createRun();
            paragraphConfig.setFontSize(22);
            paragraphConfig.setBold(true);
            paragraphConfig.setFontFamily("Century Gothic");
            paragraphConfig.setText(
                    "Базовый ответ"
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
            paragraphConfigLenta.setFontFamily("Century Gothic");
            paragraphConfigLenta.setText(
                    "Лента: "
            );

            XWPFParagraph bodyParagraphName = docxModel.createParagraph();
            bodyParagraphName.setAlignment(ParagraphAlignment.LEFT);
            XWPFRun paragraphConfigName = bodyParagraphName.createRun();
            paragraphConfigName.setFontSize(26);
            paragraphConfigName.setBold(true);
            paragraphConfigName.setFontFamily("Century Gothic");
            paragraphConfigName.setText(
                    name

            );
            paragraphConfigName.addBreak();

            XWPFParagraph bodyParagraphAnalyze = docxModel.createParagraph();
            bodyParagraphAnalyze.setAlignment(ParagraphAlignment.LEFT);
            XWPFRun paragraphConfigAnalyze = bodyParagraphAnalyze.createRun();
            paragraphConfigAnalyze.setFontSize(14);
            paragraphConfigAnalyze.setFontFamily("Yu Gothic UI");
            paragraphConfigAnalyze.setText("Аналитический отчет по упоминаниям в онлайн-СМИ и соцмедиа");
            paragraphConfigAnalyze.addBreak();


            XWPFParagraph bodyParagraphDate = docxModel.createParagraph();
            bodyParagraphDate.setAlignment(ParagraphAlignment.LEFT);
            XWPFRun paragraphConfigDate = bodyParagraphDate.createRun();
            paragraphConfigDate.setFontSize(14);
            paragraphConfigDate.setFontFamily("Yu Gothic UI");
            paragraphConfigDate.setText(
                    "Период анализа: " + date
            );
            XWPFParagraph bodyParagraphToc = docxModel.createParagraph();
            bodyParagraphToc.setAlignment(ParagraphAlignment.LEFT);
            bodyParagraphToc.setPageBreak(true);
            docxModel.createTOC();

            XWPFParagraph bodyParagraphStatistic = docxModel.createParagraph();
            bodyParagraphStatistic.setStyle("Heading1");
            bodyParagraphStatistic.setAlignment(ParagraphAlignment.LEFT);
            bodyParagraphStatistic.setPageBreak(true);
            XWPFRun paragraphConfigStatic = bodyParagraphStatistic.createRun();
            paragraphConfigStatic.setFontSize(14);
            paragraphConfigStatic.setFontFamily("Yu Gothic UI");
            paragraphConfigStatic.setText(
                    "Базовые статистики"
            );

            //create table
            XWPFTable table = docxModel.createTable();

            //create first row
            XWPFTableRow tableRowOne = table.getRow(0);
            tableRowOne.getCell(0).setText("Совокупная аудитория1, чел.");
            tableRowOne.addNewTableCell().setText("0");
            //create second row

            XWPFTableRow tableRowTwo = table.createRow();

            tableRowTwo.getCell(0).setText("Количество источников публикаций, шт.");
            tableRowTwo.getCell(1).setText(String.valueOf(data.total_sources));

            XWPFTableRow tableRowThree = table.createRow();
            tableRowThree.getCell(0).setText("Количество публикаций, шт.");
            tableRowThree.getCell(1).setText(String.valueOf(data.total_publication));

            XWPFTableRow tableRowFour = table.createRow();
            tableRowFour.getCell(0).setText("Количество комментариев к публикациям, шт.");
            tableRowFour.getCell(1).setText(String.valueOf(data.total_comment));

//           HASH MAP
            String[] categoriesPost = new String[]{};
            Double[] valuesAPost = new Double[]{};
            JSONArray jsonArray;
            for (Object o : (JSONArray) (jsonPosts).get("total")) {
                jsonArray = (JSONArray) o;
                categoriesPost = (String[]) append(categoriesPost, jsonArray.get(0));
                valuesAPost = append(valuesAPost, new Double(jsonArray.get(1).toString()));
            }
            docxModel = addChats(docxModel, categoriesPost, valuesAPost);


            String[] categoriesComments = new String[]{};
            Double[] valuesAComments = new Double[]{};
            for (Object o : (JSONArray) (jsonComments).get("total")) {
                jsonArray = (JSONArray) o;
                categoriesComments = (String[]) append(categoriesComments, jsonArray.get(0));
                valuesAComments = append(valuesAComments, new Double(jsonArray.get(1).toString()));
            }
            docxModel = addChats(docxModel, categoriesComments, valuesAComments);
            String postDate;

            Double[] postCommentData = new Double[]{};
            double postCommentD;
            for (int i =0; i < categoriesPost.length; i ++){
                postCommentD = 0;
                postDate = categoriesPost[i];
                if (valuesAPost[i] != 0) {
                for (int j =0; j < categoriesComments.length; j ++) {
                    if (postDate.equals(categoriesComments[j])) {
                        postCommentD = new Double( valuesAPost[i].toString()) /new Double( valuesAComments[i].toString());
                        break;
                    }
                }

                }
                postCommentData= append(postCommentData, postCommentD);
            }
            docxModel = addChats(docxModel, categoriesPost, postCommentData);


            addPie(docxModel,  new Double[]{(double) getComment(jsonComments, "netural"),
                    (double) getComment(jsonComments, "positive"), (double) getComment(jsonComments, "negative")});

            // сохраняем модель docx документа в файл
            try (FileOutputStream fileOut = new FileOutputStream("/home/oleg/Documents/test1t.docx")) {
                docxModel.write(fileOut);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println("Успешно записан в файл");
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


    private static XWPFDocument addChats(XWPFDocument docxModel, String[] categories, Double[] valuesA ) throws Exception {

        // create the data


        // create the chart
        XWPFChart chart = docxModel.createChart(15 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);

        // create data sources
        int numOfPoints = categories.length;
        String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
        String valuesDataRangeA = chart.formatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
        String valuesDataRangeB = chart.formatRange(new CellRangeAddress(1, numOfPoints, 2, 2));
        XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
        XDDFNumericalDataSource<Double> valuesDataA = XDDFDataSourcesFactory.fromArray(valuesA, valuesDataRangeA, 1);
//            XDDFNumericalDataSource<Double> valuesDataB = XDDFDataSourcesFactory.fromArray(valuesB, valuesDataRangeB, 2);

        // create axis
        XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
        XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        // Set AxisCrossBetween, so the left axis crosses the category axis between the categories.
        // Else first and last category is exactly on cross points and the bars are only half visible.
        leftAxis.setCrossBetween(AxisCrossBetween.BETWEEN);

        // create chart data
        XDDFChartData data = chart.createData(ChartTypes.BAR, bottomAxis, leftAxis);
        ((XDDFBarChartData) data).setBarDirection(BarDirection.COL);

        // create series
        // if only one series do not vary colors for each bar
        ((XDDFBarChartData) data).setVaryColors(false);
        XDDFChartData.Series series = data.addSeries(categoriesData, valuesDataA);
        // XDDFChart.setSheetTitle is buggy. It creates a Table but only half way and incomplete.
        // Excel cannot opening the workbook after creatingg that incomplete Table.
        // So updating the chart data in Word is not possible.
//        series.setTitle("a", chart.setSheetTitle("a", 1));
            series.setTitle("", setTitleInDataSheet(chart, "a", 1));
			/*
			   // if more than one series do vary colors of the series
			   ((XDDFBarChartData)data).setVaryColors(true);
			   series = data.addSeries(categoriesData, valuesDataB);
			   //series.setTitle("b", chart.setSheetTitle("b", 2));
			   series.setTitle("b", setTitleInDataSheet(chart, "b", 2));
			*/
        // plot chart data
        chart.plot(data);

        // create legend
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.LEFT);
        legend.setOverlay(false);

        return docxModel;
    }
    private static XWPFDocument addPie(XWPFDocument document, Double[] valuesA) throws IOException, InvalidFormatException {
        // create the data
        String[] categories = new String[]{"Нейтральность","Позитив","Негатив"};

        // create the chart

        XWPFChart chart = document.createChart(15 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);

        // create data sources
        int numOfPoints = categories.length;
        String categoryDataRange = chart.formatRange(new CellRangeAddress(1, numOfPoints, 0, 0));
        String valuesDataRangeA = chart.formatRange(new CellRangeAddress(1, numOfPoints, 1, 1));
        XDDFDataSource<String> categoriesData = XDDFDataSourcesFactory.fromArray(categories, categoryDataRange, 0);
        XDDFNumericalDataSource<Double> valuesDataA = XDDFDataSourcesFactory.fromArray(valuesA, valuesDataRangeA, 1);


        XDDFChartData data = chart.createData(ChartTypes.PIE, null, null);
        XDDFChartData.Series series = data.addSeries(categoriesData, valuesDataA);
        data.setVaryColors(true);
        series.setShowLeaderLines(false);
        series.setTitle("", chart.setSheetTitle("", 1));




//            CTDLbls dLbls = chart.getCTChart().getPlotArea().addNewPieChart().addNewDLbls();
//            dLbls.addNewShowBubbleSize().setVal(true);
//            dLbls.addNewShowLegendKey().setVal(true);
//            dLbls.addNewShowCatName().setVal(true);
//            dLbls.addNewShowSerName().setVal(true);
//            dLbls.addNewShowPercent().setVal(true);
//            dLbls.addNewShowVal().setVal(true);
//            dLbls.addNewDLblPos(Dlc)
        chart.plot(data);

        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.RIGHT);
        legend.setOverlay(true);
        return document;
    }
}