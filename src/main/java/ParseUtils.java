import com.google.gson.Gson;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;

public class ParseUtils {
    private List<BaseBean> beanList = new ArrayList<>();
    private List<String> firstColValues = new ArrayList<>();
    private List<String> firstRowValues = new ArrayList<>();

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
                    String firstName;
                    for (int cIndex = firstCellIndex; cIndex < lastCellIndex; cIndex++) {
                        Cell cell = row.getCell(cIndex);
                        BaseBean baseBean = new BaseBean();
                        baseBean.setSheetName(sheetName);
                        if (cell != null) {
                            baseBean.setValue(cell.toString());
                            Log.i(cell.toString());
                            if (cIndex == firstCellIndex) {
                                firstName = cell.toString()
                            }
                            baseBean.setRowName(firstName);
                        }
                        beanList.add(baseBean);
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public void convertToJson(){
        Gson gson = new Gson();
        String str = gson.toJson(beanList);
        Log.i(str);
    }
}
