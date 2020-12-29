public class ParseData {
    String[] categories;
    Double[] valuesA;
    Double[] valuesB;

    public ParseData( String[] categories, Double[] valuesA) {
        this.categories = categories;
        this.valuesA =valuesA;
    }
    public ParseData( String[] categories, Double[] valuesA, Double[] valuesB) {
        this.categories = categories;
        this.valuesA =valuesA;
        this.valuesB =valuesB;

    }
}
