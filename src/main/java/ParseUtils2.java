import com.google.gson.Gson;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

public class ParseUtils2 {
    private List<BaseBean> beanList = new ArrayList<>();
    private List<String> firstColValues = new ArrayList<>();
    private List<String> firstRowValues = new ArrayList<>();
    private Set<String> valueSet = new HashSet<>();

    private int numIndex = -1;

    public void parse(String path) {
        File file = new File(path);
        if (!file.exists()) {
            Log.i("文件不存在");
            return;
        }
        try {
            FileInputStream fis = new FileInputStream(file);
            Workbook workbook = new XSSFWorkbook(fis);
            int sheetCount = workbook.getNumberOfSheets();
            for (int sIndex = 0; sIndex < sheetCount; sIndex++) {
                Sheet sheet = workbook.getSheetAt(sIndex);
                String sheetName = sheet.getSheetName();

                int firstRow = sheet.getFirstRowNum();
                int lastRow = sheet.getLastRowNum();
                for (int rIndex = firstRow; rIndex <= lastRow; rIndex++) {
                    Row row = sheet.getRow(rIndex);
                    if (row == null) {
                        continue;
                    }
                    int firstCellIndex = row.getFirstCellNum();
                    int lastCellIndex = row.getLastCellNum();
                    String firstName = null;
                    for (int cIndex = firstCellIndex; cIndex < lastCellIndex; cIndex++) {
                        Cell cell = row.getCell(cIndex);
                        BaseBean baseBean = new BaseBean();
                        baseBean.setSheetName(sheetName);
                        if (cell != null) {
                            String value = cell.toString();
                            baseBean.setValue(cell.toString());
                            Log.i(cell.toString());
                            if (cIndex == firstCellIndex) {
                                firstName = cell.toString();
                            }
                            baseBean.setRowName(firstName);
                            if ("专利号".equals(value)) {
                                numIndex = cIndex;
                            } else {
                                if (numIndex == cIndex) {
                                    if (valueSet.contains(value)) {
                                        CellStyle style = cell.getCellStyle();
                                        //自定义颜色对象
                                        XSSFColor color = new XSSFColor();
                                        //根据你需要的rgb值获取byte数组
                                        color.setRGB(intToByteArray(getIntFromColor(0, 255, 255)));
                                        //自定义颜色
//                                        style.setFillForegroundColor(color);
                                        style.setFillBackgroundColor(color);
                                    } else {
                                        valueSet.add(value);
                                    }
                                }
                            }
                        }
                        beanList.add(baseBean);
                    }
                }
            }

            // 创建文件对象，作为Excel文件内容的输出文件
            File f = new File("test.xlsx");
            // 输出时通过流的形式对外输出，包装对应的目标文件
            OutputStream os = new FileOutputStream(f);
            // 将内存中的workbook数据写入到流中
            workbook.write(os);
            workbook.close();
            os.flush();
            os.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void convertToJson() {
        Gson gson = new Gson();
        String str = gson.toJson(beanList);
        Log.i(str);
    }

    /**
     * rgb转int
     */
    private static int getIntFromColor(int Red, int Green, int Blue){
        Red = (Red << 16) & 0x00FF0000;
        Green = (Green << 8) & 0x0000FF00;
        Blue = Blue & 0x000000FF;
        return 0xFF000000 | Red | Green | Blue;
    }

    /**
     * int转byte[]
     */
    public static byte[] intToByteArray(int i) {
        byte[] result = new byte[4];
        result[0] = (byte)((i >> 24) & 0xFF);
        result[1] = (byte)((i >> 16) & 0xFF);
        result[2] = (byte)((i >> 8) & 0xFF);
        result[3] = (byte)(i & 0xFF);
        return result;
    }
}
