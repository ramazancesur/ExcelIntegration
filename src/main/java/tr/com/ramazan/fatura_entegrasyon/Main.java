package tr.com.ramazan.fatura_entegrasyon;

import org.apache.commons.collections4.IteratorUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import tr.com.ramazan.fatura_entegrasyon.excel.operations.ExcelOperations;
import tr.com.ramazan.fatura_entegrasyon.interfaces.IExcelOperations;
import tr.com.ramazan.fatura_entegrasyon.utils.Utils;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

/**
 * Created by ramazancesur on 13/06/2020.
 */
public class Main {
    public static Properties readProps(){
        Properties prop = new Properties();
        try (InputStream inputStream = Main.class.getResourceAsStream("/application.properties")) {
            prop.load(inputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace(System.out);
        } catch (IOException e) {
            e.printStackTrace(System.out);
        }
        return prop;

    }
    public static void main(String[] args) throws IOException {
        IExcelOperations operations= new ExcelOperations();
        Properties props= readProps();

        InputStream ExcelFileToRead = new FileInputStream(props.getProperty("excelPath"));
        XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
        XSSFSheet sheet = wb.getSheetAt(0);

        Iterator<Row> rows = sheet.rowIterator();
        Row headerRow = rows.next();
        int noOfColumns = sheet.getRow(0).getPhysicalNumberOfCells();



        int specode3ColomnNumber = Utils.getFilteredColumnNumber(headerRow, "SPECODE3");
        int birimCarpanColumnNumber = Utils.getFilteredColumnNumber(headerRow, "BIRIM_2_CARPAN");
        int hamFiyatColumnNumber = Utils.getFilteredColumnNumber(headerRow, "FIYAT_HAM");

        // specode3 == 2 SATIŞ FİYATI=BIRIM_2_CARPAN*FIYAT_HAM
        // specode3 == 1 satış fiyatı =FIYAT_HAM

        List arrayOfList = IteratorUtils.toList(rows);
        List<String> satisFiyatList = new ArrayList<>();
        satisFiyatList.add("SATIS FIYATLARI");
        HashMap<Integer, List> filteredValue = (HashMap<Integer, List>) Utils.filteredExcelColumn(arrayOfList,
                specode3ColomnNumber, birimCarpanColumnNumber, hamFiyatColumnNumber);
        List speCodeList = filteredValue.get(specode3ColomnNumber);
        List<String> birimCarpanList = filteredValue.get(birimCarpanColumnNumber);
        List<String> hamFiyatList = filteredValue.get(hamFiyatColumnNumber);

        for (int i = 0; i < speCodeList.size(); i++) {
            if (speCodeList.get(i).equals("1") && !birimCarpanList.get(i).equals("")) {
                Double satisFiyat = Double.parseDouble(birimCarpanList.get(i).replace(",", "."))
                        * Double.parseDouble(hamFiyatList.get(i).replace(",", "."));
                satisFiyatList.add(satisFiyat.toString());
            } else {
                satisFiyatList.add(hamFiyatList.get(i).replace(",", "."));
            }
        }
        ExcelFileToRead.close();


        operations.updateXLSXFile(props.getProperty("excelPath"),satisFiyatList,noOfColumns);
        System.out.println("başarıyla excel oluşturuldu ve güncellendi");

    }
}
