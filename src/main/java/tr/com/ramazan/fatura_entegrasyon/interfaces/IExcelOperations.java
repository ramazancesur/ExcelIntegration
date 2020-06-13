package tr.com.ramazan.fatura_entegrasyon.interfaces;

import java.io.IOException;
import java.util.List;

/**
 * Created by ramazancesur on 13/06/2020.
 */
public interface IExcelOperations {
    void  readXLSFile(String excelPath) throws IOException;
    void writeXLSFile(String excelFilePath) throws IOException;
    void writeXLSXFile(String excelFilePath) throws IOException;
    void updateXLSXFile( String excelFileName, List<String> satisValues, int colomnNumber) throws IOException;

}
