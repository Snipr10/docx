import org.json.JSONArray;

import java.text.ParseException;
import java.util.Arrays;

public class DataForArea {
    String[] categoriesPostType = new String[]{};
    Double[] valuesNegative = new Double[]{};
    Double[] valuesPositive = new Double[]{};
    Double[] valuesNetural = new Double[]{};


    public DataForArea(String type, JSONArray totalComments, JSONArray positive, JSONArray netural,
                                JSONArray negative, int first_month, int first_year) throws ParseException {
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
                categoriesPostType = WordWorker.append(categoriesPostType, (String) ((JSONArray) negative.get(i)).get(0));
                if (sum == 0) {
                    valuesNegative = WordWorker.append(valuesNegative, 33.00d);
                    // 033
                    valuesPositive = WordWorker.append(valuesPositive, 33.00d);
                    // 033
                    valuesNetural = WordWorker.append(valuesNetural, 33.00d);

                } else {
                    valuesNegative = WordWorker.append(valuesNegative, (double) Math.round((negativeInt / sum*100) * 100.0) / 100.0);
                    // 033
                    valuesPositive = WordWorker.append(valuesPositive, (double) Math.round((positiveInt / sum*100) * 100.0) / 100.0);
                    // 033
                    valuesNetural = WordWorker.append(valuesNetural, (double) Math.round((neturalInt / sum*100) * 100.0) / 100.0);
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
                int dateInt = WordWorker.getDate((String) ((JSONArray) negative.get(i)).get(0), type);
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
                    categoriesPostType = WordWorker.append(categoriesPostType, dateSo);
                    valuesNegative = WordWorker.append(valuesNegative, 100.00 * negativeInt);
                    // 033
                    valuesPositive = WordWorker.append(valuesPositive, 100.00 * positiveInt);
                    // 033
                    valuesNetural = WordWorker.append(valuesNetural, 100.00 * neturalInt);
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
            WordWorker.changeWeekString(categoriesPostType, type, first_month, first_year);
        }

    }

}
