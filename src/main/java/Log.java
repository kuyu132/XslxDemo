public class Log {
    private static boolean isDebug = true;

    public static void setDebug(boolean isDebug) {
        Log.isDebug = isDebug;
    }

    public static void i(String value) {
        if (isDebug) {
            System.out.println(value);
        }
    }
}
