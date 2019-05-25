package com.epion_t3.excel.command.runner;

import com.epion_t3.excel.command.model.ExcelBindVariables;
import com.epion_t3.core.command.bean.CommandResult;
import com.epion_t3.core.command.runner.impl.AbstractCommandRunner;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Iterator;

/**
 *
 */
@Slf4j
public class ExcelBindVariablesRunner extends AbstractCommandRunner<ExcelBindVariables> {

    /**
     * {@inheritDoc}
     */
    @Override
    public CommandResult execute(ExcelBindVariables command, Logger logger) throws Exception {

        // 対象のExcelファイルパス
        Path targetFilePath = Paths.get(getCommandBelongScenarioDirectory(), command.getTarget()).normalize();

        // 保存先のファイルパス
        Path saveFilePath = getEvidencePath("BindVariables_" + targetFilePath.getFileName().toString());

        try (Workbook wb = WorkbookFactory.create(targetFilePath.toFile());) {
            // 全シート、全セルを走査し、文字列セルに対してバインド実施後の文字列を設定
            Iterator<Sheet> sheetIte = wb.sheetIterator();
            while (sheetIte.hasNext()) {
                Sheet sheet = sheetIte.next();
                log.debug("Read Sheet. Name : {}", sheet.getSheetName());
                Iterator<Row> rowIte = sheet.rowIterator();
                while (rowIte.hasNext()) {
                    Row row = rowIte.next();
                    Iterator<Cell> cellIte = row.cellIterator();
                    while (cellIte.hasNext()) {
                        Cell cell = cellIte.next();
                        if (cell != null) {
                            switch (cell.getCellType()) {
                                case STRING:
                                    String value = cell.getStringCellValue();
                                    cell.setCellValue(bind(value));
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }

            // 書き込み
            try (OutputStream fileOut = new FileOutputStream(saveFilePath.toFile())) {
                wb.write(fileOut);
            }

            // エビデンスとして保存
            registrationFileEvidence(saveFilePath);
        }
        return CommandResult.getSuccess();
    }

}
