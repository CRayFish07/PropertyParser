package propertiy;


import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by hds on 17-5-24.
 */
public class ReadProperties {

    private static String inputFile;
    private static String outputFile;

    public static void main(String[] args) {
        try {
            System.out.println("版本:1.0.0");
            if (args == null || args.length == 0) {
                System.out.println("路径出错,eg:java -jar xxx.jar xxx/xxx.xml xxx/res");
                System.exit(0);
            }
            String xmlPath = args[0];
            outputFile = args[1];

            File file = new File(xmlPath);
            // 创建输入流，读取Excel
            InputStream is = new FileInputStream(file.getAbsolutePath());
            //根据上述创建的输入流 创建工作簿对象
            Workbook wb = WorkbookFactory.create(is);
            Sheet sheet = wb.getSheetAt(0);
            flatAllProperties(sheet);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 由于所有设备属性不整齐.
     *
     * @param sheet
     */
    private static void flatAllProperties(Sheet sheet) {
        Row rowFirst = sheet.getRow(0);
        //最后一列
        final int lastColumn = rowFirst.getLastCellNum();
        //最后一行
        final int lastRow = sheet.getLastRowNum();
        P p = new P();
        p.pList = new ArrayList<>();
        p.version = versionCode();
        for (int rowIndex = 1; rowIndex < lastRow; rowIndex++) {
            //读出一行
            Row rowContent = sheet.getRow(rowIndex);
            if (rowContent == null) continue;
            TreeMap<String, String> map = initMap();
            for (int columnIndex = 1; columnIndex < lastColumn; columnIndex++) {
                Cell startCell = rowContent.getCell(columnIndex);
                if (startCell == null || startCell.toString().isEmpty()) continue;
                //获取属性名字
                Cell cell = rowFirst.getCell(columnIndex);
                if (cell == null || cell.toString() == null || !cell.toString().contains("(")) continue;
                final String propertyName = getPropertyName(cell.toString());
                //获取属性的值
                cell = rowContent.getCell(columnIndex);
                final String content = getFinalContent(columnIndex, cell == null ? "" : cell.toString());
                map.put(propertyName, content);
            }
            final String mapString = new Gson().toJson(map);
            if (mapString.equals("{}")) continue;
            p.pList.add(map);
        }
        filterCam(p, "cam");
        filterCam(p, "bell");
        Gson gson = new GsonBuilder().setPrettyPrinting().create();
        String json = gson.toJson(p);
        Utils.write2File(json, outputFile);
    }

    private static TreeMap<String, String> initMap() {
        return new TreeMap<>((Comparator<String>) String::compareTo);
    }

    private static void filterCam(P p, final String tag) {
        List<TreeMap<String, String>> pList = p.pList;
        final int count = pList.size();
        List<Integer> pidList = new ArrayList<>();
        for (int i = 0; i < count; i++) {
            Map<String, String> map = pList.get(i);
            final String value = map.get("value");
            if (value != null && value.contains(tag)) {
                final String pid = map.get("Pid");
                final String os = map.get("os");
                if (pid != null && pid.length() > 0 && isDigitsOnly(pid)) {
                    if (!pidList.contains(Integer.parseInt(pid)))
                        pidList.add(Integer.parseInt(pid));
                }
                if (os != null && os.length() > 0 && isDigitsOnly(os)) {
                    if (!pidList.contains(Integer.parseInt(os)))
                        pidList.add(Integer.parseInt(os));
                }
            }
        }
        Collections.sort(pidList);
        if (p.typePidMap == null) p.typePidMap = new TreeMap<>((Comparator<String>) String::compareTo);
        p.typePidMap.put(tag, pidList);
    }

    private static class P {
        private List<TreeMap<String, String>> pList;
        private TreeMap<String, List<Integer>> typePidMap;
        String version;

    }

    static final String bingo = "√";
    static final String opposite = "×";

    private static String getFinalContent(int index, String content) {
        if (index > 7) {
            if (content != null && content.equals(bingo))
                return "1";
            else return "0";
        }
        if (content.contains("."))
            content = content.substring(0, content.lastIndexOf("."));
        return content;
    }

    private static String versionCode() {
        SimpleDateFormat format = new SimpleDateFormat("yyyyMMdd");
        return format.format(new Date(System.currentTimeMillis()));
    }

    private static String getPropertyName(String content) {
        return content.substring(content.indexOf("(") + 1, content.lastIndexOf(")"));
    }


    /**
     * Returns whether the given CharSequence contains only digits.
     */
    public static boolean isDigitsOnly(CharSequence str) {
        final int len = str.length();
        for (int cp, i = 0; i < len; i += Character.charCount(cp)) {
            cp = Character.codePointAt(str, i);
            if (!Character.isDigit(cp)) {
                return false;
            }
        }
        return true;
    }

}
