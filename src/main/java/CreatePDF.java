import com.itextpdf.awt.DefaultFontMapper;
import com.itextpdf.awt.DefaultFontMapper.BaseFontParameters;
import com.itextpdf.awt.FontMapper;
import com.itextpdf.text.Font;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.*;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.*;
import org.jfree.chart.labels.ItemLabelAnchor;
import org.jfree.chart.labels.ItemLabelPosition;
import org.jfree.chart.labels.StandardCategoryItemLabelGenerator;
import org.jfree.chart.labels.StandardPieSectionLabelGenerator;
import org.jfree.chart.plot.*;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.renderer.category.GroupedStackedBarRenderer;
import org.jfree.chart.renderer.category.StandardBarPainter;
import org.jfree.chart.title.LegendTitle;
import org.jfree.data.KeyToGroupMap;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.ui.GradientPaintTransformType;
import org.jfree.ui.RectangleInsets;
import org.jfree.ui.StandardGradientPaintTransformer;
import org.jfree.ui.TextAnchor;
import org.json.JSONArray;
import org.json.JSONObject;
import com.itextpdf.text.Chunk;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfPageEventHelper;
import com.itextpdf.text.pdf.PdfWriter;

import java.io.*;
import java.util.AbstractMap.SimpleEntry;
import java.util.ArrayList;
import java.util.List;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.ParseException;

import java.util.Iterator;
import java.util.UUID;


public class CreatePDF {
    private static String fontUrl = "arial.ttf";
    private static BaseFont fontRegular;
    private static FontMapper fontMapper;
    private static String fontUrlBold;
    private static String fontUrlRus;
    private static String encoding;
    private static Font fontFrazeBOLD;
    private static Font font;
    private static Font fontFraze;
    private static Color GrayColor = new Color(224, 225, 225);
    private static Color YellowColor = new Color(255, 209, 48);
    private static Color BlueColor = new Color(0, 0, 139);
    private static Color WhiteColor = new Color(240, 240, 240);
    private static Color SeaColor = new Color(0, 206, 209);
    private static Color[] cityColor = new Color[] {YellowColor,
            BlueColor,
            new Color(0, 206, 209),
            new Color(255, 160, 122),
            new Color(119, 136, 153),
            new Color(147, 112, 219),
            new Color(255, 192, 203),
            new Color(255, 239, 213),
            SeaColor,
            new Color(152, 251, 152),
    };
    private static Color NegativeColor =  new Color(255, 0, 0);
    private static Color NeutralColor =  new Color(82, 82, 82, 50);
    private static Color PositiveColor = new Color(144, 238, 144);

    public CreatePDF() throws IOException, DocumentException {
    }

    public static String createPDF(String docName, String type, String name, String date, DataForDocx data, JSONObject jsonPosts, JSONObject jsonComments, JSONObject stat, JSONObject sex, JSONObject age, JSONObject usersJson, JSONArray jsonCity, JSONArray posts, JSONArray postsContent, JSONArray commentContent, int first_month, int first_year, JSONArray postsContentLikes, int cityAll) throws DocumentException, IOException, ParseException, FontFormatException {
        DefaultFontMapper mapper = new DefaultFontMapper();
        mapper.insertDirectory(fontUrl);
        BaseFontParameters pp = mapper.getBaseFontParameters("Arial Unicode MS");
        if (pp != null) {
            pp.encoding = "Identity-H";
        }

        String encoding = "cp1251";
        int with = 500;
        int users = Integer.parseInt(usersJson.get("count").toString());
        int diagramCount = 1;
        Paragraph paragraphEnter = new Paragraph("\n", FontFactory.getFont(fontUrl, encoding, true, 10.0F));
        Document document = new Document(PageSize.A4, 50.0F, 50.0F, 50.0F, 50.0F);
        String paring_name = UUID.randomUUID().toString() + ".pdf";
        PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(paring_name));

        writer.createXmpMetadata();
        writer.setTagged();
        writer.setPageEvent(new Header());
        TOCEvent event = new TOCEvent();
        writer.setPageEvent(event);
        document.open();
        String title = "Базовый отчет";
        Chunk c;
        Paragraph paragraphTitle = new Paragraph(title, FontFactory.getFont(fontUrlBold, encoding, true, 22.0F));        paragraphTitle.setAlignment(2);
        document.add(paragraphTitle);
        document.add(new Phrase("\n\n\n\n\n\n\n\n"));
        document.add(new Phrase("\n\n\n\n"));
        paragraphTitle = new Paragraph("Лента:", FontFactory.getFont(fontUrlBold, encoding, true, 22.0F));
        document.add(paragraphTitle);
        paragraphTitle = new Paragraph(name, FontFactory.getFont(fontUrlBold, encoding, true, 28.0F));
        document.add(paragraphTitle);
        document.add(new Phrase("\n"));
        paragraphTitle = new Paragraph("Аналитический отчет по упоминаниям в онлайн-СМИ и соцмедиа", font);
        document.add(paragraphTitle);
        document.add(new Phrase("\n"));
        paragraphTitle = new Paragraph("Период анализа: " + date, font);
        document.add(paragraphTitle);
        document.add(new Phrase("\n"));
        document.newPage();
        Paragraph paragraphTOB = new Paragraph("Оглавление", FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
        document.add(paragraphTOB);
        document.newPage();
        title = "Базовые статистики";
        c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
        c.setGenericTag(title);
        Paragraph paragraphBaseStatistic = new Paragraph(c);
        document.add(paragraphBaseStatistic);
        document.add(paragraphEnter);
        PdfPTable table = new PdfPTable(2);
        table.setTotalWidth((float)with);
        table.setLockedWidth(true);

        addCell("Совокупная аудитория, чел.", table, 0);
        addCell(String.valueOf(users), table, 2);

        addCell("Количество источников публикаций, шт.", table, 0);
        addCell(String.valueOf(data.total_sources), table, 2);

        addCell("Количество публикаций, шт.", table, 0);
        addCell(String.valueOf(data.total_publication), table, 2);

        addCell("Количество комментариев к публикациям, шт.", table, 0);
        addCell(String.valueOf(data.total_comment), table, 2);

        addCell("Количество просмотров, шт.", table, 0);
        addCell(String.valueOf(data.total_views), table, 2);


        document.add(table);

        document.add(new Phrase("\n"));
        ParseData postData = WordWorker.getWeekData(type, (JSONArray)((JSONObject)jsonPosts.get("total")).get("total"), first_month, first_year);
        int diagramY = 400;
        int tableCount = 1;

        double val_test = 0;
        for (Double d:postData.valuesA){
            val_test += d;
        }
        if (val_test > 0) {
            title = String.format("Диаграмма %s Динамика количества публикаций", diagramCount);
            c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
            c.setGenericTag(title);
            paragraphBaseStatistic = new Paragraph(c);
            document.add(paragraphBaseStatistic);
            AddBar(postData.categories, postData.valuesA, writer, diagramY);
            diagramY = ChangeY(diagramY, document, false);
            diagramCount = diagramCount + 1;
        }
        ParseData comments = WordWorker.getWeekData(type, (JSONArray)jsonComments.get("total"), first_month, first_year);
        val_test = 0;
        for (Double d:comments.valuesA){
            val_test += d;
        }
        if (val_test > 0) {
            title = String.format("Диаграмма %s Динамика количества комментариев к публикациям", diagramCount);
            c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
            c.setGenericTag(title);
            paragraphBaseStatistic = new Paragraph(c);
            document.add(paragraphBaseStatistic);
            AddBar(comments.categories, comments.valuesA, writer, diagramY);
            diagramY = ChangeY(diagramY, document, false);
            ++diagramCount;
        }
        String[] categoriesPost = postData.categories;
        Double[] valuesAPost = postData.valuesA;
        Double[] postCommentData = new Double[0];
        String[] categoriesComments = comments.categories;
        Double[] valuesAComments = comments.valuesA;

        for(int i = 0; i < categoriesPost.length; ++i) {
            double postCommentD = 0.0D;
            String postDate = categoriesPost[i];
            if (valuesAPost[i] != 0.0D) {
                for(int j = 0; j < categoriesComments.length; ++j) {
                    if (postDate.equals(categoriesComments[j])) {
                        if (valuesAComments[i] == 0.0D) {
                            postCommentD = 0.0D;
                        } else {
                            postCommentD = (double)Math.round(new Double(valuesAPost[i].toString()) / new Double(valuesAComments[i].toString()) * 100.0D) / 100.0D;
                        }
                        break;
                    }
                }
            }

            postCommentData = (Double[])WordWorker.append(postCommentData, postCommentD);
        }
        val_test = 0;
        for (Double d:postCommentData){
            val_test += d;
        }
        if (val_test > 0) {
            title = String.format("Диаграмма %s Динамика количества комментариев на 1 публикацию", diagramCount);
            c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
            c.setGenericTag(title);
            paragraphBaseStatistic = new Paragraph(c);
            document.add(paragraphBaseStatistic);
            AddBar(categoriesPost, postCommentData, writer, diagramY, true);
            diagramY = ChangeY(diagramY, document, false);
            ++diagramCount;
            document.add(paragraphEnter);
        }

        JSONObject jsonPostTotal = (JSONObject)jsonPosts.get("total");
        Double[] variableDouble = new Double[]{(double)WordWorker.getComment(jsonPostTotal, "netural"), (double)WordWorker.getComment(jsonPostTotal, "positive"), (double)WordWorker.getComment(jsonPostTotal, "negative")};
        val_test = 0;
        for (Double d:variableDouble){
            val_test += d;
        }
        if (val_test > 0) {
            title = String.format("Диаграмма %s Тональность публикаций", diagramCount);
            c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
            c.setGenericTag(title);
            paragraphBaseStatistic = new Paragraph(c);
            document.add(paragraphBaseStatistic);

            addPie(new String[]{"Негативная", "Нейтральная", "Позитивная"}, new Double[]{variableDouble[2],
                    variableDouble[0], variableDouble[1]}, writer, diagramY, false, true);
            diagramY = ChangeY(diagramY, document, false);
            ++diagramCount;
        }
        for (Double d:variableDouble){
            val_test += d;
        }
        if (val_test > 0) {
            document.newPage();
            diagramY = 550;
            title = String.format("Диаграмма %s Тональность публикаци", diagramCount);
            c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
            c.setGenericTag(title);
            paragraphBaseStatistic = new Paragraph(c);
            document.add(paragraphBaseStatistic);
            JSONArray positive = (JSONArray) jsonPostTotal.get("positive");
            JSONArray netural = (JSONArray) jsonPostTotal.get("netural");
            JSONArray negative = (JSONArray) jsonPostTotal.get("negative");
            JSONArray totalComments = (JSONArray) jsonPostTotal.get("total");
            DataForArea d = new DataForArea(type, totalComments, positive, netural, negative, first_month, first_year);
            addArea(d, writer, diagramY);
            ChangeY(diagramY, document, false);
            ++diagramCount;
        }
        int total_vk = WordWorker.getTotalMedia(jsonPosts, "vk");
        int total_tw = WordWorker.getTotalMedia(jsonPosts, "tw");
        int total_fb = WordWorker.getTotalMedia(jsonPosts, "fb");
        int total_gs = WordWorker.getTotalMedia(jsonPosts, "gs");
        int total_tg = WordWorker.getTotalMedia(jsonPosts, "tg");
        int total_ig = WordWorker.getTotalMedia(jsonPosts, "ig");
        int total_yt = WordWorker.getTotalMedia(jsonPosts, "yt");

        int all = total_vk + total_tw + total_fb + total_gs + total_tg + total_ig + total_yt;
        ParseData soData = WordWorker.getWeekDataMedia(type, jsonPosts, first_month, first_year);
        double val = 0;
        for (Double d:soData.valuesA){
            val += d;
        }
        for (Double d:soData.valuesB){
            val += d;
        }
        Paragraph paragraphSources;
        PdfPTable tableSource;
        JSONObject sexJson;
        if ((all != 0) || (val != 0)) {

            document.newPage();

            title = "Источники";
            c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 22.0F));
            c.setGenericTag(title);
            paragraphSources = new Paragraph(c);
            document.add(paragraphSources);


            diagramY = 0;
            title = String.format("Таблица %s Ключевые площадки", tableCount);
            c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
            c.setGenericTag(title);
            paragraphSources = new Paragraph(c);
            document.add(paragraphSources);
            document.add(paragraphEnter);
            tableCount = tableCount + 1;
            tableSource = new PdfPTable(3);
            tableSource.setTotalWidth((float) with);
            tableSource.setLockedWidth(true);
            addToTable3(tableSource, "Площадка", "Количество публикаций, шт.", "   %     ", fontFrazeBOLD);
            addToTable3(tableSource, "Вконтакте", String.valueOf(total_vk), String.valueOf((double) Math.round((double) ((float) total_vk) * 100.0D / (double) ((float) all) * 100.0D) / 100.0D), fontFraze);
            addToTable3(tableSource, "Facebook", String.valueOf(total_fb), String.valueOf((double) Math.round((double) ((float) total_fb) * 100.0D / (double) ((float) all) * 100.0D) / 100.0D), fontFraze);
            addToTable3(tableSource, "Twitter", String.valueOf(total_tw), String.valueOf((double) Math.round((double) ((float) total_tw) * 100.0D / (double) ((float) all) * 100.0D) / 100.0D), fontFraze);
            addToTable3(tableSource, "Инстаграм", String.valueOf(total_ig), String.valueOf((double) Math.round((double) ((float) total_ig) * 100.0D / (double) ((float) all) * 100.0D) / 100.0D), fontFraze);
            addToTable3(tableSource, "Telegram", String.valueOf(total_tg), String.valueOf((double) Math.round((double) ((float) total_tg) * 100.0D / (double) ((float) all) * 100.0D) / 100.0D), fontFraze);
            addToTable3(tableSource, "YouTube", String.valueOf(total_yt), String.valueOf((double) Math.round((double) ((float) total_yt) * 100.0D / (double) ((float) all) * 100.0D) / 100.0D), fontFraze);
            addToTable3(tableSource, "СМИ", String.valueOf(total_gs), String.valueOf((double) Math.round((double) ((float) total_gs) * 100.0D / (double) ((float) all) * 100.0D) / 100.0D), fontFraze);
            addToTable3(tableSource, "Итог", String.valueOf(all), "100", font);
            document.add(tableSource);
            diagramY = 280;

            title = String.format("Диаграмма %s Динамика количества публикаций на отдельных площадках", diagramCount);
            c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
            c.setGenericTag(title);
            paragraphSources = new Paragraph(c);
            document.add(paragraphSources);
            addDouble(soData.categories, soData.valuesA, soData.valuesB, writer, diagramY);
            ++diagramCount;
            ChangeY(diagramY, document, false);
            document.newPage();
        }
        if (posts.length() != 0) {
            title = String.format("Таблица %s Топ-%s источников по количеству публикаций", tableCount, posts.length());
            c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
            c.setGenericTag(title);
            paragraphSources = new Paragraph(c);
            document.add(paragraphSources);
            document.add(paragraphEnter);
            tableSource = new PdfPTable(3);
            tableSource.setTotalWidth((float) with);
            tableSource.setLockedWidth(true);
            document.add(new Phrase(""));
            ++tableCount;
            addToTable3(tableSource, "Название источника", "    URL   ", "Количество публикаций", fontFrazeBOLD);

            Iterator var97 = posts.iterator();


            while (var97.hasNext()) {
                Object o = var97.next();
                sexJson = (JSONObject) o;
                addToTable3(tableSource, sexJson.get("username").toString(), sexJson.get("url").toString(), sexJson.get("coefficient").toString(), fontFraze);
            }

            document.add(tableSource);
        }
        ParseData auditData = WordWorker.getWeekData(type, (JSONArray)stat.get("graph_data"), first_month, first_year);
        sexJson = (JSONObject)((JSONObject)sex.get("additional_data")).get("sex");
        String[] categoriesCity = new String[0];
        Double[] valuesACity = new Double[0];
        JSONArray JsoncityUpdate = new JSONArray(jsonCity.toString());
        int res = 0;
        for ( Object o1 :JsoncityUpdate){
            res += Integer.parseInt( (String) (((JSONObject)o1).get("users")));
        }
        JSONObject json1_city = new JSONObject();
        json1_city.put("city", "Другие города");
        json1_city.put("users", Integer.toString(cityAll-res));
        JsoncityUpdate.put(json1_city);

        int i = 0;
        int count10 = 0;

        JSONObject jsonObject;
        Iterator var65;
        Object o;
        for(var65 = JsoncityUpdate.iterator(); var65.hasNext(); ++i) {
            o = var65.next();
            jsonObject = (JSONObject)o;

            count10 += Integer.parseInt(jsonObject.get("users").toString());
        }

        double valueCity;
        for(var65 = JsoncityUpdate.iterator();
            var65.hasNext();
            ) {
            o = var65.next();
            if (categoriesCity.length >= 11) {
                break;
            }

            jsonObject = (JSONObject)o;
            valueCity = (double)Math.round(Double.parseDouble(jsonObject.get("users").toString()) * 100.0D / (double)count10 * 100.0D) / 100.0D;
            if (valueCity < 1.0D) {
                continue;
            }
            valuesACity = (Double[])WordWorker.append(valuesACity, valueCity);
            categoriesCity = (String[])((String[])WordWorker.append(categoriesCity, jsonObject.get("city")));
        }

        double valAudit = 0.0D;
        double valSex = 0.0D;
        double valAge = 0.0D;
        double valCity = 0.0D;
        Double[] masSex = new Double[]{new Double(sexJson.get("u").toString()), new Double(sexJson.get("m").toString()), new Double(sexJson.get("w").toString())};
        Double[] masAge = new Double[]{new Double(((JSONObject)age.get("group1")).get("graph").toString()), new Double(((JSONObject)age.get("group2")).get("graph").toString()), new Double(((JSONObject)age.get("group3")).get("graph").toString()), new Double(((JSONObject)age.get("group4")).get("graph").toString())};
        int likesComment = masSex.length;
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
            Iterator var105;
            document.newPage();
            title = "Аудитория";
            c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 22.0F));
            c.setGenericTag(title);
            Paragraph paragraphAudit = new Paragraph(c);
            document.add(paragraphAudit);
            diagramY = 520;

            val_test = 0;
            for (Double d:auditData.valuesA){
                val_test += d;
            }
            if (val_test > 0) {
                title = String.format("Диаграмма %s Динамика объема аудитории", diagramCount);
                c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
                c.setGenericTag(title);
                paragraphAudit = new Paragraph(c);
                document.add(paragraphAudit);
                AddBar(auditData.categories, auditData.valuesA, writer, diagramY);
                diagramY = ChangeY(diagramY, document, false);
                ++diagramCount;
            }

            val_test = 0;
            for (Double d:masSex){
                val_test += d;
            }
            if (val_test > 0) {
                title = String.format("Диаграмма %s Распределение аудитории по полу", diagramCount);
                c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
                c.setGenericTag(title);
                paragraphAudit = new Paragraph(c);

                document.add(paragraphAudit);
                addPie(new String[]{"Не указан", "Мужчины", "Женщины"}, masSex, writer, diagramY);
                diagramY = ChangeY(diagramY, document, false);
                ++diagramCount;
            }
            val_test = 0;
            for (Double d:masAge){
                val_test += d;
            }
            if (val_test > 0) {
                title = String.format("Диаграмма %s Распределение аудитории по возрасту", diagramCount);
                c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
                c.setGenericTag(title);
                paragraphAudit = new Paragraph(c);
                document.add(paragraphAudit);
                addPie(new String[]{"18-25 лет", "26-40 лет", "40 лет и старше", "не указан"}, masAge, writer, diagramY);
                diagramY = ChangeY(diagramY, document, false);
                ++diagramCount;
            }
            val_test = 0;
            for (Double d:valuesACity){
                val_test += d;
            }
            if (val_test > 0) {
                title = String.format("Диаграмма %s Распределение аудитории по геолокаци", diagramCount);
                c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
                c.setGenericTag(title);
                paragraphAudit = new Paragraph(c);
                document.add(paragraphAudit);
                addPie(categoriesCity, valuesACity, writer, diagramY, true);
                ChangeY(diagramY, document, false);
                ++diagramCount;
            }
            PdfPTable tableAudit;
            if (jsonCity.length() > 0) {
                title = String.format("Таблица %s Топ-%s городов", tableCount, jsonCity.length());
                c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
                c.setGenericTag(title);
                paragraphAudit = new Paragraph(c);
                document.add(paragraphAudit);
                document.add(paragraphEnter);
                tableAudit = new PdfPTable(3);
                tableAudit.setTotalWidth((float) with);
                tableAudit.setLockedWidth(true);
                document.add(new Phrase(""));
                ++tableCount;
                addToTable3(tableAudit, "Город", "Количество", "%", fontFrazeBOLD);
                var105 = jsonCity.iterator();

                while (var105.hasNext()) {
                    o = var105.next();
                    jsonObject = (JSONObject) o;
                    addToTable3(tableAudit, jsonObject.get("city").toString(), jsonObject.get("users").toString(), String.format("%.1f", Double.parseDouble(jsonObject.get("users").toString()) * 100.0D / Double.valueOf((double) count10)), fontFraze);
                }

                document.add(tableAudit);
            }
            if (((JSONArray) usersJson.get("users")).length() > 0) {
                title = String.format("Таблица %s Топ-%s активных пользователей по сумме реакции (лайков, комментариев, репостов)", tableCount, ((JSONArray) usersJson.get("users")).length());
                c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
                c.setGenericTag(title);
                paragraphAudit = new Paragraph(c);

                document.add(paragraphAudit);
                document.add(paragraphEnter);
                tableAudit = new PdfPTable(3);
                tableAudit.setTotalWidth((float) with);
                tableAudit.setLockedWidth(true);
                document.add(new Phrase(""));
                ++tableCount;
                addToTable3(tableAudit, "Пользователь", "URL", "Сумма реакции", fontFrazeBOLD);
                var105 = ((JSONArray) usersJson.get("users")).iterator();

                while (var105.hasNext()) {
                    o = var105.next();
                    jsonObject = (JSONObject) o;
                    addToTable3(tableAudit, jsonObject.get("name").toString(), jsonObject.get("url").toString(), jsonObject.get("coefficient").toString(), fontFraze);
                }

                document.add(tableAudit);
            }


        }
        int likesPosts = 0;
        likesComment = 0;
        for (Object o1 : postsContent) {
            jsonObject = (JSONObject) o1;
            if (Integer.parseInt(jsonObject.get("viewed").toString()) + Integer.parseInt(jsonObject.get("reposts").toString()) +
                    Integer.parseInt(jsonObject.get("likes").toString()) + Integer.parseInt(jsonObject.get("comments").toString()) > 0) {
                likesPosts+=1;
            }
        }
        for (Object o2 : commentContent) {
            if (Integer.parseInt(((JSONObject) o2).get("likes").toString()) > 0) {
                likesComment+=1;
            }
        }
        if((likesComment !=0) || (likesPosts !=0 )) {
            document.newPage();
            title = "Ключевые публикации";
            c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 22.0F));
            c.setGenericTag(title);
            Paragraph paragraphPublication = new Paragraph(c);


            document.add(paragraphPublication);
            diagramY = 0;
            PdfPTable tablePublication;
            Iterator var80;
            String text;
            if (likesPosts >= 0) {
                title = String.format("Таблица %s Топ-%s публикаций по сумме резонанса", tableCount, likesPosts);
                c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
                c.setGenericTag(title);
                paragraphPublication = new Paragraph(c);

                document.add(paragraphPublication);
                document.add(paragraphEnter);
                tablePublication = new PdfPTable(3);
                tablePublication.setTotalWidth((float) with);
                tablePublication.setLockedWidth(true);
                document.add(new Phrase(""));
                ++tableCount;
                addToTable3(tablePublication, "Публикация", "URL", "Резонанс", fontFrazeBOLD);
                var80 = postsContent.iterator();


                while (var80.hasNext()) {
                    try {
                        o = var80.next();
                        jsonObject = (JSONObject) o;
                        text = WordWorker.updateText(jsonObject.get("text").toString());
                        addToTable3(tablePublication, text, jsonObject.get("uri").toString(), WordWorker.res(jsonObject), fontFraze, false);
                }  catch(Exception e){
                    System.out.println("can not add comment");
                }
            }
                document.add(tablePublication);
            }
            if (postsContentLikes.length()> 0) {
                title = String.format("Таблица %s Топ-%s публикаций по количеству лайков", tableCount, postsContentLikes.length());
                c = new Chunk(title, FontFactory.getFont(fontUrlBold, encoding, true, 14.0F));
                c.setGenericTag(title);
                paragraphPublication = new Paragraph(c);

                document.add(paragraphPublication);
                document.add(paragraphEnter);
                tablePublication = new PdfPTable(3);
                tablePublication.setTotalWidth((float) with);
                tablePublication.setLockedWidth(true);
                document.add(new Phrase(""));
                ++tableCount;
                addToTable3(tablePublication, "Публикация", "URL", "Лайки", fontFrazeBOLD);
                var80 = postsContentLikes.iterator();

                while (var80.hasNext()) {
                    o = var80.next();
                    jsonObject = (JSONObject) o;
                    try {
                        if (jsonObject.get("title").equals("")){
                            text = WordWorker.updateText(jsonObject.get("text").toString());

                        }
                        else {
                            text = WordWorker.updateText(jsonObject.get("title") + "\n" + jsonObject.get("text").toString());

                        }
                    }catch (Exception e){
                        text = WordWorker.updateText(jsonObject.get("text").toString());

                    }

                    addToTable3(tablePublication, text, jsonObject.get("uri").toString(), jsonObject.get("likes").toString(), fontFraze, false);
                }

                document.add(tablePublication);
            }
        }
        document.newPage();
        document.add(new Paragraph("Оглавление", FontFactory.getFont(fontUrlBold, encoding, true, 22.0F)));
        document.add(paragraphEnter);

        List<SimpleEntry<String, Integer>> entries = event.getTOC();
        Paragraph p;
        String nameTOB;
        String page;

        PdfPTable toc = new PdfPTable(new float[] {10, 1});
        toc.setTotalWidth((float)with);
        toc.setLockedWidth(true);
        toc.deleteBodyRows();
        String last = null;
        for (SimpleEntry<String, Integer> entry : entries) {
            nameTOB = entry.getKey();
            if ((nameTOB.toLowerCase().contains("диаграмма") || nameTOB.toLowerCase().contains("таблица")) ) {
                nameTOB = "    " + entry.getKey();
            }
            if (!nameTOB.equals(last)) {
                PdfPCell cellOneSource = new PdfPCell(new Phrase(nameTOB, FontFactory.getFont(fontUrl, encoding, true, 9.0F)));
                cellOneSource.setHorizontalAlignment(3);
                cellOneSource.setVerticalAlignment(5);
                cellOneSource.setBorder(Rectangle.NO_BORDER);
                toc.addCell(cellOneSource);
//
                page = entry.getValue().toString();

                cellOneSource = new PdfPCell(new Phrase(page, FontFactory.getFont(fontUrl, encoding, true, 9.0F)));
                cellOneSource.setHorizontalAlignment(2);
                cellOneSource.setVerticalAlignment(5);
                cellOneSource.setBorder(Rectangle.NO_BORDER);
                toc.addCell(cellOneSource);
            }
            last = nameTOB;

        }
        document.setPageCount(2);
        document.add(toc);
        document.close();
        File parsingFile = new File(paring_name);
        PdfReader reader = new PdfReader(new FileInputStream(parsingFile));
        int n = reader.getNumberOfPages();
        reader.selectPages(String.format("1, %d, 3-%d", n, n-1));
        String fullName= docName + ".pdf";
        PdfStamper stamper = new PdfStamper(reader, new FileOutputStream(fullName));
        stamper.close();
        try {
            parsingFile.delete();
        } catch (Exception e){
            System.out.println("can not delete file");
        }
        return fullName;
    }

    private static void addCell(String data, PdfPTable table, int alignment){
        PdfPCell cellOne;
        cellOne = new PdfPCell(new Phrase(data, fontFraze));
        cellOne.setFixedHeight(20.0F);
        cellOne.setHorizontalAlignment(alignment);
        cellOne.setBorder(2);
        table.addCell(cellOne);
    }

    private static void addToTable3(PdfPTable tableSource, String d1, String d2, String d3, Font font) {
        addToTable3(tableSource, d1, d2, d3, font, true);
    }

    private static void addToTable3(PdfPTable tableSource, String d1, String d2, String d3, Font font, boolean isFixed_size) {

        PdfPCell cellOneSource = new PdfPCell(new Phrase(d1, font));
        cellOneSource.setHorizontalAlignment(0);
        cellOneSource.setVerticalAlignment(5);
        if (isFixed_size) {
            cellOneSource.setFixedHeight(20.0F);
        }

        tableSource.addCell(cellOneSource);
        cellOneSource = new PdfPCell(new Phrase(d2, font));
        if (isFixed_size) {
            cellOneSource.setFixedHeight(20.0F);
        }

        cellOneSource.setHorizontalAlignment(1);
        cellOneSource.setVerticalAlignment(5);
        tableSource.addCell(cellOneSource);
        cellOneSource = new PdfPCell(new Phrase(d3, font));
        if (isFixed_size) {
            cellOneSource.setFixedHeight(20.0F);
        }

        cellOneSource.setHorizontalAlignment(2);
        cellOneSource.setVerticalAlignment(5);
        tableSource.addCell(cellOneSource);
    }

    private static int ChangeY(int diagramY, Document document, boolean isLast) throws DocumentException {
        diagramY -= 260;
        if (diagramY < 0) {
            if (!isLast) {
                document.newPage();
            }

            diagramY = 550;
        } else {
            for(int i = 0; i <= 12; ++i) {
                document.add(new Phrase("\n"));
            }
        }

        return diagramY;
    }
    private static void AddBar(String[] categories, Double[] valuesA, PdfWriter writer, int diagramY){
        AddBar(categories, valuesA, writer, diagramY, false);
    }
    private static void AddBar(String[] categories, Double[] valuesA, PdfWriter writer, int diagramY, boolean decimal) {
        DefaultCategoryDataset defaultCategoryDataset = new DefaultCategoryDataset();

        for(int i = 0; i < categories.length; ++i) {
            defaultCategoryDataset.setValue(valuesA[i], "", categories[i]);
        }

        PdfContentByte pdfContentByte = writer.getDirectContent();
        int width = 500;
        int height = 208;
        PdfTemplate pdfTemplate = pdfContentByte.createTemplate((float)width, (float)height);
        Graphics2D graphics2d = pdfTemplate.createGraphics((float)width, (float)height, new DefaultFontMapper());
        graphics2d.setColor(Color.BLACK);
        Rectangle2D rectangle2d = new java.awt.geom.Rectangle2D.Double(0.0D, 0.0D, (double)width, (double)height);
        JFreeChart jFreeChart = ChartFactory.createBarChart("", "", "", defaultCategoryDataset, PlotOrientation.VERTICAL, false, false, false);
        jFreeChart.getPlot().setBackgroundPaint(Color.WHITE);
        CategoryPlot plot = jFreeChart.getCategoryPlot();
        plot.setOutlinePaint((Paint)null);
        plot.setRangeGridlinePaint(WhiteColor);
//        plot.setRangeGridlinePaint(Color.BLACK);
        plot.setRangeGridlineStroke(new BasicStroke(0.01f,BasicStroke.CAP_SQUARE, BasicStroke.JOIN_BEVEL, 0.0f, null, 1.0f));
//        plot.setRangeGridlineStroke(new BasicStroke(0.1f));
        plot.setDomainGridlinesVisible(true);
        plot.setDomainGridlinePaint(WhiteColor);
        plot.setDomainGridlineStroke(new BasicStroke(0.01f,BasicStroke.CAP_SQUARE, BasicStroke.JOIN_BEVEL, 0.0f, null, 0.0f));

        CategoryAxis categoryAxis = plot.getDomainAxis();
//        categoryAxis.setCategoryLabelPositions(CategoryLabelPositions.UP_45);
        categoryAxis.setCategoryLabelPositions(CategoryLabelPositions.UP_90);

        categoryAxis.setTickLabelFont(new java.awt.Font("Arial", 0, 5));
//        categoryAxis.setAxisLinePaint(Color.WHITE);
        categoryAxis.setAxisLinePaint(Color.BLACK);

        categoryAxis.setTickLabelPaint(Color.BLACK);
        ValueAxis valueAxis = plot.getRangeAxis();
        valueAxis.setTickLabelFont(new java.awt.Font("Arial", 0, 5));
        valueAxis.setTickLabelPaint(Color.BLACK);
//        valueAxis.setAxisLinePaint(Color.WHITE);
        valueAxis.setAxisLinePaint(Color.BLACK);

        BarRenderer render = (BarRenderer)plot.getRenderer();
        render.setSeriesPaint(0, YellowColor);
//        render.setMaximumBarWidth(0.05D);
        render.setMaximumBarWidth(0.09D);
        render.setBarPainter(new StandardBarPainter());
        if (decimal) {
            render.setSeriesItemLabelGenerator(0, new StandardCategoryItemLabelGenerator());
        } else {
            DecimalFormat labelFormat = new DecimalFormat("##");
            plot.getRenderer().setSeriesItemLabelGenerator(0, new StandardCategoryItemLabelGenerator("{2}", labelFormat));
        }
        render.setSeriesItemLabelsVisible(1, true);
        render.setBaseItemLabelsVisible(true);
//        render.setBaseSeriesVisible(true);
        render.setBaseItemLabelFont(new java.awt.Font("Arial", 0, 5));
//        ItemLabelPosition position = new ItemLabelPosition(ItemLabelAnchor.CENTER, TextAnchor.CENTER, TextAnchor.TOP_CENTER, -1.57D);
//        render.setBasePositiveItemLabelPosition(position);
        jFreeChart.draw(graphics2d, rectangle2d);
        graphics2d.dispose();
        pdfContentByte.addTemplate(pdfTemplate, 40.0F, (float)diagramY);
    }

    private static void addPie(String[] categories, Double[] valuesA, PdfWriter writer, int diagramY) throws FontFormatException, DocumentException, IOException {
        addPie(categories, valuesA, writer, diagramY, false, false);
    }

    private static void addPie(String[] categories, Double[] valuesA, PdfWriter writer, int diagramY, boolean is_city) throws FontFormatException, DocumentException, IOException {
        addPie(categories, valuesA, writer, diagramY, is_city, false);
    }

    private static void addPie(String[] categories, Double[] valuesA, PdfWriter writer, int diagramY, boolean is_city, boolean is_tonal) throws IOException, FontFormatException, DocumentException {
        DefaultPieDataset dataset = new DefaultPieDataset();
        String [] categoriesUpdate =  new String[categories.length];
        String updateCategory;
        for(int i = 0; i < categories.length; ++i) {
            if (is_city) {
                updateCategory = categories[i] + " " + String.format("%.1f", valuesA[i]) + "%";
                categoriesUpdate[i]= updateCategory;
                dataset.setValue(updateCategory, valuesA[i]);
            } else {
                dataset.setValue(categories[i], valuesA[i]);
            }
        }

        PdfContentByte pdfContentByte = writer.getDirectContent();
        int width = 500;
        int height = 208;
        PdfTemplate pdfTemplate = pdfContentByte.createTemplate((float)width, (float)height);
        JFreeChart chart = ChartFactory.createRingChart("", dataset, false, false, false);
        chart.setBackgroundPaint(java.awt.Color.white);
        chart.setPadding(new RectangleInsets(4, 8, 2, 2));

        RingPlot plot = (RingPlot)chart.getPlot();
        if (!is_city) {
            StandardPieSectionLabelGenerator generator;
            if (is_tonal) {
                generator = new StandardPieSectionLabelGenerator("{0}; {2}", new DecimalFormat("0"), new DecimalFormat("0.00%"));
            } else {
                generator = new StandardPieSectionLabelGenerator("{0}; {2}", new DecimalFormat("0"), new DecimalFormat("0%"));
            }

            plot.setLabelGenerator(generator);
        }
        if (is_city) {
            plot.setSectionDepth(1.0);
            int s = 0;
            for (String value : categoriesUpdate) {
                if (value.contains("Не указан")) {
                    plot.setSectionPaint(value, GrayColor);
                } else {
                    plot.setSectionPaint(value, cityColor[s]);
                    s++;
                }
            }
        } else {
            plot.setSectionDepth(0.50);
            plot.setSectionPaint("Негативная", NegativeColor);
            plot.setSectionPaint("Нейтральная", NeutralColor);
            plot.setSectionPaint("Не указан", GrayColor);
            plot.setSectionPaint("Позитивная", PositiveColor);
            plot.setSectionPaint("Женщины", YellowColor);
            plot.setSectionPaint("Мужчины", BlueColor);
            plot.setSectionPaint("18-25 лет", YellowColor);
            plot.setSectionPaint("26-40 лет", BlueColor);
            plot.setSectionPaint("40 лет и старше", SeaColor);
            plot.setSectionPaint("не указан", GrayColor);

        }

        plot.setOutlinePaint((Paint)null);
        plot.setBackgroundPaint(Color.WHITE);
        plot.setLabelOutlinePaint(Color.WHITE);


        plot.setLabelFont(new java.awt.Font("Arial", 0, 5));
        plot.setLabelBackgroundPaint(Color.WHITE);
        plot.setNoDataMessage("No data available");
//        plot.setBackgroundPaint(null);
//        plot.setOutlineVisible(false);
//        plot.setLabelGenerator(null);
//        plot.setSectionDepth(0.35);

//        plot.setSectionOutlinesVisible(false);
//        plot.setSimpleLabels(true);


        plot.setLabelShadowPaint(null);
        plot.setShadowXOffset(0.0D);
        plot.setShadowYOffset(0.0D);
        plot.setLabelLinkStyle(PieLabelLinkStyle.STANDARD);
        plot.setLabelLinkPaint(new Color(50, 50, 50));
        plot.setShadowPaint(null);
        plot.setOuterSeparatorExtension(0);
        plot.setInnerSeparatorExtension(0);
        plot.setLabelBackgroundPaint(null);
        plot.setLabelOutlinePaint(null);
//        plot.setLabelGap(0.02D);
        Graphics2D graphics2d = pdfTemplate.createGraphics((float)width, (float)height, fontMapper);
        Rectangle2D rectangle2d = new java.awt.geom.Rectangle2D.Double(0.0D, 0.0D, (double)width, (double)height);
        chart.draw(graphics2d, rectangle2d);
        graphics2d.dispose();
        pdfContentByte.addTemplate(pdfTemplate, 40.0F, (float)diagramY);
    }
    private static Double roundValue(double d, boolean check, int n) {
        if (check){
            if (n  == 10) {
                if (d > 10){
                    return Math.round(d * 10.0) / 10.0;
                }
            }
            if (n == 1) {
                if (d > 1){
                    return Math.round(d)/1.0;
                }
                else {
                    return Math.round(d * 10.0) / 10.0;

                }
            }
         }
        return d;
    }

    private static void addArea(DataForArea d, PdfWriter writer, int diagramY) throws IOException, FontFormatException, DocumentException {
        DefaultCategoryDataset result = new DefaultCategoryDataset();
        int n = 1;
        boolean check = false;
        if (d.categoriesPostType.length > 49) {
            check = true;
            n = 1;

        }
        else {
            if (d.categoriesPostType.length > 29) {
                check = true;
                n = 10;
            }
        }
        for(int i = 0; i < d.categoriesPostType.length; ++i) {
            result.addValue(roundValue(d.valuesNegative[i], check, n), "Негативная тональность %", d.categoriesPostType[i]);
            result.addValue(roundValue(d.valuesNetural[i], check, n), "Нейтральная тональность %", d.categoriesPostType[i]);
            result.addValue(roundValue(d.valuesPositive[i], check, n), "Позитивная тональность %", d.categoriesPostType[i]);
        }

        JFreeChart chart = ChartFactory.createStackedBarChart("", "%", "", result, PlotOrientation.VERTICAL, true, false, false);
        LegendTitle legend = chart.getLegend();
        legend.setItemFont(new java.awt.Font("Arial", 0, 5));
        legend.setBorder(0.0D, 0.0D, 0.0D, 0.0D);
        GroupedStackedBarRenderer renderer = new GroupedStackedBarRenderer();
        KeyToGroupMap map = new KeyToGroupMap("G1");
        map.mapKeyToGroup("Негативная тональность %", "G1");
        map.mapKeyToGroup("Нейтральная тональность %", "G1");
        map.mapKeyToGroup("Позитивная тональность %", "G1");
        renderer.setSeriesToGroupMap(map);
        renderer.setItemMargin(0.0D);
        renderer.setBarPainter(new StandardBarPainter());
        Color p1 = NegativeColor;
        renderer.setSeriesPaint(0, p1);
        renderer.setSeriesPaint(4, p1);
        renderer.setSeriesPaint(8, p1);
        Color p2 =  NeutralColor;
        renderer.setSeriesPaint(1, p2);
        renderer.setSeriesPaint(5, p2);
        renderer.setSeriesPaint(9, p2);
        Color p3 =PositiveColor;
        renderer.setSeriesPaint(2, p3);
        renderer.setSeriesPaint(6, p3);
        renderer.setSeriesPaint(10, p3);
        renderer.setGradientPaintTransformer(new StandardGradientPaintTransformer(GradientPaintTransformType.HORIZONTAL));




        renderer.setSeriesPositiveItemLabelPosition(0,
                new ItemLabelPosition(ItemLabelAnchor.OUTSIDE12, TextAnchor.HALF_ASCENT_CENTER, TextAnchor.HALF_ASCENT_CENTER, 0.0));

        renderer.setSeriesPositiveItemLabelPosition(1,
                new ItemLabelPosition(ItemLabelAnchor.CENTER, TextAnchor.TOP_CENTER, TextAnchor.TOP_CENTER, 0.0));

        renderer.setSeriesPositiveItemLabelPosition(2,
                new ItemLabelPosition(ItemLabelAnchor.OUTSIDE6, TextAnchor.CENTER, TextAnchor.CENTER, 0.0));


        SubCategoryAxis domainAxis = new SubCategoryAxis("");
        domainAxis.setCategoryMargin(0.05D);
        chart.getPlot().setBackgroundPaint(Color.WHITE);
        CategoryPlot plot = (CategoryPlot)chart.getPlot();
        plot.setDomainAxis(domainAxis);
        plot.setRenderer(renderer);
        BarRenderer render = (BarRenderer)plot.getRenderer();
        render.setMaximumBarWidth(0.09D);
        render.setSeriesItemLabelGenerator(0, new StandardCategoryItemLabelGenerator());
        render.setSeriesItemLabelGenerator(1, new StandardCategoryItemLabelGenerator());
        render.setSeriesItemLabelGenerator(2, new StandardCategoryItemLabelGenerator());
        render.setSeriesItemLabelsVisible(1, true);
        render.setBaseItemLabelsVisible(true);
        render.setBaseSeriesVisible(true);

        render.setBaseItemLabelFont(new java.awt.Font("Arial", 0, 5));
        plot.setOutlinePaint((Paint)null);
        plot.setRangeGridlinePaint(WhiteColor);
        plot.setRangeGridlineStroke(new BasicStroke(0.01f,BasicStroke.CAP_SQUARE, BasicStroke.JOIN_BEVEL, 0.0f, null, 1.0f));

        CategoryAxis categoryAxis = plot.getDomainAxis();
        categoryAxis.setTickLabelFont(new java.awt.Font("Arial", 0, 5));
        categoryAxis.setAxisLinePaint(Color.BLACK);
        categoryAxis.setTickLabelPaint(Color.BLACK);
        categoryAxis.setCategoryLabelPositions(CategoryLabelPositions.UP_90);

        ValueAxis valueAxis = plot.getRangeAxis();
        valueAxis.setTickLabelFont(new java.awt.Font("Arial", 0, 5));
        valueAxis.setTickLabelPaint(Color.BLACK);
        valueAxis.setAxisLinePaint(Color.BLACK);
        DecimalFormat pctFormat = new DecimalFormat("#%");
        pctFormat.setMultiplier(1);
        NumberAxis rangeAxis = (NumberAxis)plot.getRangeAxis();
        rangeAxis.setNumberFormatOverride(pctFormat);
        PdfContentByte pdfContentByte = writer.getDirectContent();
        int width = 500;
        int height = 210;
        PdfTemplate pdfTemplate = pdfContentByte.createTemplate((float)width, (float)height);
        Graphics2D graphics2d = pdfTemplate.createGraphics((float)width, (float)height, fontMapper);
        Rectangle2D rectangle2d = new java.awt.geom.Rectangle2D.Double(0.0D, 0.0D, (double)width, (double)height);
        chart.draw(graphics2d, rectangle2d);
        graphics2d.dispose();
        pdfContentByte.addTemplate(pdfTemplate, 40.0F, (float)diagramY);
    }

    private static void addDouble(String[] categories, Double[] valuesA, Double[] valuesB, PdfWriter writer, int diagramY) throws IOException, FontFormatException, DocumentException {
        DefaultCategoryDataset result = new DefaultCategoryDataset();

        for(int i = 0; i < categories.length; ++i) {
            result.addValue(valuesA[i], "Сми", categories[i]);
            result.addValue(valuesB[i], "СоцМедиа", categories[i]);
        }

        JFreeChart chart = ChartFactory.createStackedBarChart("", "%", "", result, PlotOrientation.VERTICAL, true, false, false);
        LegendTitle legend = chart.getLegend();
        legend.setItemFont(new java.awt.Font("Arial", 0, 5));
        legend.setBorder(0.0D, 0.0D, 0.0D, 0.0D);
        GroupedStackedBarRenderer renderer = new GroupedStackedBarRenderer();
        KeyToGroupMap map = new KeyToGroupMap("G1");
        map.mapKeyToGroup("Сми", "G1");
        map.mapKeyToGroup("СоцМедиа", "G2");
        renderer.setSeriesToGroupMap(map);
        renderer.setItemMargin(0.0D);
//        Paint p1 = new GradientPaint(0.0F, 0.0F, new Color(129, 22, 244), 0.0F, 0.0F, new Color(129, 22, 244));
        renderer.setSeriesPaint(0, YellowColor);
        renderer.setMaximumBarWidth(0.09D);
        renderer.setBarPainter(new StandardBarPainter());

//        renderer.setSeriesPaint(0, p1);
        renderer.setSeriesPaint(1, BlueColor);
        renderer.setGradientPaintTransformer(new StandardGradientPaintTransformer(GradientPaintTransformType.HORIZONTAL));
        SubCategoryAxis domainAxis = new SubCategoryAxis("");
        domainAxis.setCategoryMargin(0.05D);
        chart.getPlot().setBackgroundPaint(Color.WHITE);
        CategoryPlot plot = (CategoryPlot)chart.getPlot();
        plot.setDomainAxis(domainAxis);
        plot.setRenderer(renderer);
        BarRenderer render = (BarRenderer)plot.getRenderer();
        render.setMaximumBarWidth(0.05D);

//        render.setMaximumBarWidth(0.05D);


        render.setSeriesItemLabelGenerator(0, new StandardCategoryItemLabelGenerator());
        render.setSeriesItemLabelGenerator(1, new StandardCategoryItemLabelGenerator());
        render.setSeriesItemLabelGenerator(2, new StandardCategoryItemLabelGenerator());
        render.setSeriesItemLabelsVisible(1, true);
        render.setBaseItemLabelsVisible(true);
        render.setBaseSeriesVisible(true);
        render.setBaseItemLabelFont(new java.awt.Font("Arial", 0, 4));
//        ItemLabelPosition position = new ItemLabelPosition(ItemLabelAnchor.INSIDE, TextAnchor.CENTER, TextAnchor.TOP_CENTER, 0);
        render.setBasePositiveItemLabelPosition(new ItemLabelPosition(ItemLabelAnchor.OUTSIDE12, TextAnchor.BOTTOM_CENTER, TextAnchor.CENTER, 0.0));
//        render.setBasePositiveItemLabelPosition(position);
        plot.setOutlinePaint((Paint)null);

        plot.setRangeGridlinePaint(WhiteColor);
//        plot.setRangeGridlinePaint(Color.BLACK);
        plot.setRangeGridlineStroke(new BasicStroke(0.01f,BasicStroke.CAP_SQUARE, BasicStroke.JOIN_BEVEL, 0.0f, null, 1.0f));
//        plot.setRangeGridlineStroke(new BasicStroke(0.1f));
        plot.setDomainGridlinesVisible(true);
        plot.setDomainGridlinePaint(WhiteColor);
        plot.setDomainGridlineStroke(new BasicStroke(0.01f,BasicStroke.CAP_SQUARE, BasicStroke.JOIN_BEVEL, 0.0f, null, 0.0f));

        CategoryAxis categoryAxis = plot.getDomainAxis();
        categoryAxis.setTickLabelFont(new java.awt.Font("Arial", 0, 5));
        categoryAxis.setCategoryLabelPositions(CategoryLabelPositions.UP_90);

        categoryAxis.setAxisLinePaint(Color.BLACK);
        categoryAxis.setTickLabelPaint(Color.BLACK);
        ValueAxis valueAxis = plot.getRangeAxis();
        valueAxis.setTickLabelFont(new java.awt.Font("Arial", 0, 5));
        valueAxis.setTickLabelPaint(Color.BLACK);
        valueAxis.setAxisLinePaint(Color.BLACK);
        PdfContentByte pdfContentByte = writer.getDirectContent();
        int width = 500;
        int height = 208;
        PdfTemplate pdfTemplate = pdfContentByte.createTemplate((float)width, (float)height);
        Graphics2D graphics2d = pdfTemplate.createGraphics((float)width, (float)height, fontMapper);
        Rectangle2D rectangle2d = new java.awt.geom.Rectangle2D.Double(0.0D, 0.0D, (double)width, (double)height);
        chart.draw(graphics2d, rectangle2d);
        graphics2d.dispose();
        pdfContentByte.addTemplate(pdfTemplate, 40.0F, (float)diagramY);
    }

    static {
        try {
            fontRegular = BaseFont.createFont(fontUrl, "Cp1251", true);
        } catch (DocumentException var1) {
            var1.printStackTrace();
        } catch (IOException var2) {
            var2.printStackTrace();
        }

        fontMapper = new FontMapper() {
            public java.awt.Font pdfToAwt(BaseFont arg0, int arg1) {
                return null;
            }

            public BaseFont awtToPdf(java.awt.Font font) {
                return CreatePDF.fontRegular;
            }
        };
        fontUrlBold = "arialbd.ttf";
        fontUrlRus = "ofont.ru_Arial Cyr.ttf";
        encoding = "cp1251";
        fontFrazeBOLD = FontFactory.getFont(fontUrlBold, encoding, true, 10.0F);
        font = FontFactory.getFont(fontUrl, encoding, true, 14.0F);
        fontFraze = FontFactory.getFont(fontUrl, encoding, true, 10.0F);
    }


    static class Header extends PdfPageEventHelper {
        PdfTemplate t;
        Image total;
        Font f = FontFactory.getFont(fontUrl, encoding, true, 5.0F);

        @Override
        public void onOpenDocument(PdfWriter writer, Document document) {
            t = writer.getDirectContent().createTemplate(30, 16);
            try {
                total = Image.getInstance(t);
                total.setRole(PdfName.ARTIFACT);
            } catch (DocumentException de) {
                throw new ExceptionConverter(de);
            }
        }

        @Override
        public void onEndPage(PdfWriter writer, Document document) {
            PdfPTable table = new PdfPTable(3);

            try {
                if (writer.getPageNumber() > 1) {
                    table.setWidths(new int[]{24, 24, 2});
                    table.setTotalWidth(770);
                    table.getDefaultCell().setFixedHeight(20);
                    table.getDefaultCell().setBorder(Rectangle.NO_BORDER);

                    table.addCell(new Phrase(String.format("Страница: %d", writer.getPageNumber()), f));
                    table.getDefaultCell().setHorizontalAlignment(Element.ALIGN_RIGHT);
                    table.addCell(new Phrase(String.format("", writer.getPageNumber()), f));
                    PdfPCell cell = new PdfPCell(total);
                    cell.setBorder(Rectangle.NO_BORDER);
                    table.addCell(cell);
                    PdfContentByte canvas = writer.getDirectContent();
                    canvas.beginMarkedContentSequence(PdfName.ARTIFACT);
                    table.writeSelectedRows(0, -1, 36, 30, canvas);
                    canvas.endMarkedContentSequence();
                }
            } catch (DocumentException de) {
                throw new ExceptionConverter(de);
            }
        }

        @Override
        public void onCloseDocument(PdfWriter writer, Document document) {
//            PdfPTable table = new PdfPTable(3);
//
//            try {
//                table.setWidths(new int[]{24, 24, 2});
//            } catch (DocumentException e) {
//                e.printStackTrace();
//            }
//            table.setTotalWidth(770);
//            table.getDefaultCell().setFixedHeight(20);
//            table.getDefaultCell().setBorder(Rectangle.NO_BORDER);
//
//            table.addCell(new Phrase(String.format("Страница: %d", 123), f));
//            table.getDefaultCell().setHorizontalAlignment(Element.ALIGN_RIGHT);
//            table.addCell(new Phrase(String.format("", writer.getPageNumber()), f));
//            PdfPCell cell = new PdfPCell(total);
//            cell.setBorder(Rectangle.NO_BORDER);
//            table.addCell(cell);
//            PdfContentByte canvas = writer.getDirectContent();
//            canvas.beginMarkedContentSequence(PdfName.ARTIFACT);
//            table.writeSelectedRows(0, -1, 36, 30, canvas);
//            canvas.endMarkedContentSequence();


//            ColumnText.showTextAligned(t, Element.ALIGN_LEFT,
//                    new Phrase(String.valueOf(writer.getPageNumber()), font),
//                    2, 4, 0);
        }
    }
    public static class TOCEvent extends PdfPageEventHelper {

        protected List<SimpleEntry<String, Integer>> toc = new ArrayList<>();

        @Override
        public void onGenericTag(PdfWriter writer, Document document, Rectangle rect, String text) {
            toc.add(new SimpleEntry(text, writer.getPageNumber()));
        }

        public List getTOC() {
            return toc;
        }
    }

}
