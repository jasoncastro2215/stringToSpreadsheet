import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

public class Main {
    public static void main(String[] args) {
        Map<String, Object[]> data = new TreeMap<>();
        String[] strArr = ("ARG-4-OPzV-280 4 2 290 286 281 258 232 2300 0.88 103 206 382 - 19\n" +
                "ARG-5-OPzV-350 5 2 363 358 352 323 290 2860 0.71 124 206 382 - 24\n" +
                "ARG-6-OPzV-420 6 2 435 429 423 388 347 3380 0.60 145 206 382 - 26\n" +
                "ARG-5-OPzV-490 5 2 536 527 517 472 420 3380 0.60 124 206 498 - 31\n" +
                "ARG-6-OPzV-588 6 2 644 633 622 567 504 3980 0.51 145 206 498 - 36\n" +
                "ARG-7-OPzV-686 7 2 753 740 727 662 588 4520 0.45 166 206 498 - 42\n" +
                "ARG-6-OPzV-840 6 2 937 920 892 809 716 4360 0.47 145 206 673 - 48\n" +
                "ARG-8-OPzV-1120 8 4 1247 1224 1187 1077 954 5980 0.34 191 210 673 80 65\n" +
                "ARG-10-OPzV-1400 10 4 1560 1532 1485 1347 1193 7380 0.28 233 210 673 110 82\n" +
                "ARG-12-OPzV-1680 12 4 1877 1842 1786 1618 1432 8640 0.24 275 210 673 140 96\n" +
                "ARG-12-OPzV-2100 12 4 2120 2086 2050 1878 1678 9440 0.22 275 210 824 140 110\n" +
                "ARG-16-OPzV-2800 16 6 2824 2780 2731 2503 2237 12680 0.16 399 214 799 110 159\n" +
                "ARG-20-OPzV-3500 20 8 3523 3468 3412 3127 2796 16240 0.13 487 212 799 110 202\n" +
                "ARG-24-OPzV-4200 24 8 4248 4181 4106 3760 3357 18460 0.11 576 212 799 140 227\n" +
                "ARG-6V-4-OPzV-240 4 2 253 250 243 233 212 2260 2.70 272 205 371 - 48\n" +
                "ARG-6V-5-OPzV-300 5 2 317 313 304 292 265 2740 2.22 380 205 371 - 63\n" +
                "ARG-6V-6-OPzV-360 6 2 381 377 365 350 318 3220 1.89 380 205 371 - 70\n" +
                "ARG-12V-1-OPzV-60 1 2 63 62 60 58 52 620 19.80 272 205 371 - 43\n" +
                "ARG-12V-2-OPzV-120 2 2 126 125 120 115 105 1240 9.90 272 205 371 - 52\n" +
                "ARG-12V-3-OPzV-180 3 2 188 186 180 173 158 1720 7.08 380 205 371 - 74").split("\n");
        int i = 1;
        for (String str : strArr) {
            String[] items = str.split(" ");
            System.out.println(str);
            data.put(Integer.toString(i), new Object[] {items[0],items[1],items[2],items[3],items[4],items[5],items[6],items[7],items[8],items[9],items[10],items[11],
                    items[12],items[13],items[14]});
            i++;
        }

        XSSFWorkbook workbook = new XSSFWorkbook();

        XSSFSheet sheet = workbook.createSheet("Carousell Data");

        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset)
        {
            Row row = sheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
                Cell cell = row.createCell(cellnum++);
                if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
        try
        {
            FileOutputStream out = new FileOutputStream("CarousellData.xlsx");
            workbook.write(out);
            out.close();
            System.out.println("CarousellData.xlsx written successfully on disk.");
        }
        catch (Exception e) {
            e.printStackTrace();
        }
    }
}
