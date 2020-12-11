import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileOutputStream;
import java.math.BigInteger;

class WordWorker {
    public static void createDoc(String name, String date, DataForDocx data) {
        try {
            // создаем модель docx документа, 
            // к которой будем прикручивать наполнение (колонтитулы, текст)
            XWPFDocument docxModel = new XWPFDocument();
            CTSectPr ctSectPr = docxModel.getDocument().getBody().addNewSectPr();
            // получаем экземпляр XWPFHeaderFooterPolicy для работы с колонтитулами
            XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(docxModel, ctSectPr);

            // создаем верхний колонтитул Word файла
//            CTP ctpHeaderModel = createHeaderModel(
//                    "Верхний колонтитул - создано с помощью Apache POI на Java :)"
//            );
//            // устанавливаем сформированный верхний
//            // колонтитул в модель документа Word
//            XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeaderModel, docxModel);
//            headerFooterPolicy.createHeader(
//                    XWPFHeaderFooterPolicy.DEFAULT,
//                    new XWPFParagraph[]{headerParagraph}
//            );

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
            XWPFRun  paragraphConfigName = bodyParagraphName.createRun();
            paragraphConfigName.setFontSize(26);
            paragraphConfigName.setBold(true);
            paragraphConfigName.setFontFamily("Century Gothic");
            paragraphConfigName.setText(
                    name

            );
            paragraphConfigName.addBreak();

            XWPFParagraph bodyParagraphAnalyze = docxModel.createParagraph();
            bodyParagraphAnalyze.setAlignment(ParagraphAlignment.LEFT);
            XWPFRun  paragraphConfigAnalyze = bodyParagraphAnalyze.createRun();
            paragraphConfigAnalyze.setFontSize(14);
            paragraphConfigAnalyze.setFontFamily("Yu Gothic UI");
            paragraphConfigAnalyze.setText("Аналитический отчет по упоминаниям в онлайн-СМИ и соцмедиа");
            paragraphConfigAnalyze.addBreak();


            XWPFParagraph bodyParagraphDate = docxModel.createParagraph();
            bodyParagraphDate.setAlignment(ParagraphAlignment.LEFT);
            XWPFRun  paragraphConfigDate = bodyParagraphDate.createRun();
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
            XWPFRun  paragraphConfigStatic = bodyParagraphStatistic.createRun();
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
            tableRowOne.getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(200000));
            tableRowOne.addNewTableCell().setText("0");
            //create second row

            XWPFTableRow tableRowTwo = table.createRow();

            tableRowTwo.getCell(0).setText("Количество источников публикаций, шт.");
            tableRowTwo.getCell(1).setText(String.valueOf(data.total_sources));
            tableRowTwo.getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(200000));

            XWPFTableRow tableRowThree = table.createRow();
            tableRowThree.getCell(0).setText("Количество публикаций, шт.");
            tableRowThree.getCell(1).setText(String.valueOf(data.total_publication));
            tableRowThree.getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(200000));

            XWPFTableRow tableRowFour = table.createRow();
            tableRowFour.getCell(0).setText("Количество комментариев к публикациям, шт.");
            tableRowFour.getCell(1).setText(String.valueOf(data.total_comment));
            tableRowFour.getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(200000));

//            CTTblWidth tblWidth = table.getRow(0).getCell(0).getCTTc().addNewTcPr().addNewTcW();
//            tblWidth.setW(BigInteger.valueOf(10*1440));
//            XWPFParagraph bodyParagraphTOC = docxModel.createParagraph();
//            bodyParagraphTOC.setAlignment(ParagraphAlignment.LEFT);
//            CTP ctP = bodyParagraphTOC.getCTP();
//            CTSimpleField toc = ctP.addNewFldSimple();
//            toc.setInstr("TOC \\h");
//            toc.setDirty(STOnOff.TRUE);
//
//            XWPFParagraph bodyParagraphTest = docxModel.createParagraph();
//            bodyParagraphTest.setStyle("Heading1");
//            bodyParagraphTest.setAlignment(ParagraphAlignment.LEFT);
//            bodyParagraphTest.setPageBreak(true);
//            XWPFRun  paragraphConfigTest = bodyParagraphTest.createRun();
//            paragraphConfigTest.setFontSize(14);
//            paragraphConfigTest.setFontFamily("Yu Gothic UI");
//            paragraphConfigTest.setText(
//                    "test 1"
//            );
//
//            XWPFParagraph bodyParagraphTest2 = docxModel.createParagraph();
//            bodyParagraphTest2.setStyle("Heading1");
//            bodyParagraphTest2.setAlignment(ParagraphAlignment.LEFT);
//            bodyParagraphTest2.setPageBreak(true);
//            XWPFRun  paragraphConfigTest2 = bodyParagraphTest2.createRun();
//            paragraphConfigTest2.setFontSize(14);
//            paragraphConfigTest2.setFontFamily("Yu Gothic UI");
//            paragraphConfigTest2.setText(
//                    "test 2"
//            );
//

            // сохраняем модель docx документа в файл
            FileOutputStream outputStream = new FileOutputStream("/home/oleg/Documents/1.docx");
            docxModel.write(outputStream);
            outputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        System.out.println("Успешно записан в файл");
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
}