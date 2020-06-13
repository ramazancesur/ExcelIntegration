package tr.com.ramazan.fatura_entegrasyon.utils;

import org.apache.commons.collections4.IteratorUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;

import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Created by ramazancesur on 13/06/2020.
 */
public class Utils {
    private final static DataFormatter DATA_FORMATTER = new DataFormatter();

    private static Stream createIteratorToStream(Iterator iterator) {
        return StreamSupport.stream(Spliterators.spliteratorUnknownSize(iterator, Spliterator.ORDERED), false);
    }

    public static int getFilteredColumnNumber(Row headerRow, String filteredValue) {
        Iterator<Cell> cells = headerRow.cellIterator();
        while (cells.hasNext()) {
            Cell currentCell = cells.next();
            String value = DATA_FORMATTER.formatCellValue(currentCell);
            if (value.equals(filteredValue)) {
                return currentCell.getColumnIndex();
            }
        }
        return -1;
    }

    public static List<Map> excelDataToMap(List<Row> rows){
        List<Map> excelRowDataList =
                        rows.stream()
                        .map(row -> {
                            Iterator<Cell> cells = row.cellIterator();
                            Map<Integer, String> rowDataMap = ((Stream<Cell>) Utils.createIteratorToStream(cells))
                                    .collect(Collectors.toMap(Cell::getColumnIndex,
                                            currentCell -> Utils.DATA_FORMATTER.formatCellValue(currentCell)));
                            return rowDataMap;
                        }).collect(Collectors.toList());

        return  excelRowDataList;
    }


    public static List filteredExcelColumn(List<Row> rows,int columnNumber) {
        List<Map> excelRowDataList= excelDataToMap(rows);
        return excelRowDataList.stream()
                .map(x -> x.get(columnNumber))
                .collect(Collectors.toList());
    }

    public static Map<Integer, List> filteredExcelColumn(List<Row> rows, int... columnNumbers){
        Map<Integer,List> filteredMap= new HashMap<>();
        for(int colomnNumber: columnNumbers){
            List filteredData= filteredExcelColumn(rows,colomnNumber);
            filteredMap.put(colomnNumber,filteredData);
        }
        return filteredMap;
    }
}
