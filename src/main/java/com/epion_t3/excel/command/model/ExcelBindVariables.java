package com.epion_t3.excel.command.model;

import com.epion_t3.excel.command.runner.ExcelBindVariablesRunner;
import com.epion_t3.core.common.annotation.CommandDefinition;
import com.epion_t3.core.common.bean.scenario.Command;
import lombok.Getter;
import lombok.Setter;

/**
 * エクセルの全シート、全セルの文字列に対するバインド処理.
 *
 * @author takashno
 */
@Getter
@Setter
@CommandDefinition(id = "ExcelBindVariables", runner = ExcelBindVariablesRunner.class)
public class ExcelBindVariables extends Command {
}
