public class Main {

    public static void main(String[] args) {
        Log.setDebug(true);
        ParseUtils2 parseUtils = new ParseUtils2();
//        parseUtils.parse("values.xlsx");
        parseUtils.parse("拟被许可企业及专利清单2022.11.9（深大-南方）.xlsx");
//        parseUtils.convertToJson();
    }
}
