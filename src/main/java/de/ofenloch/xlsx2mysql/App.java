package de.ofenloch.xlsx2mysql;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.util.List;

import org.apache.commons.io.FilenameUtils;

import de.ofenloch.xlsx2mysql.*;
import org.apache.commons.cli.*;

/**
 * Hello world!
 *
 */
public class App {
    public static void main(String[] args) {
        System.out.println("Hello, World!");
        // set up command line parsing with org.apache.commons.cli
        Options options = new Options();
        CommandLineParser parser = new DefaultParser();
        HelpFormatter formatter = new HelpFormatter();
        CommandLine cmd = null;
        // add optins (command line arguments)
        Option optInput = new Option("i", "input", true, "XLS(X) input file path");
        String sInput = optInput.getLongOpt();
        optInput.setRequired(true);
        options.addOption(optInput);

        Option optOutput = new Option("o", "output", true, "SQL output file");
        String sOutput = optOutput.getLongOpt();
        optOutput.setRequired(true);
        options.addOption(optOutput);

        Option optSheet = new Option("s", "sheet", true,
                "params for the sheets: SheetIndex ColumnNameRowIndex FirstDataRowIndex ");
        String sSheet = optSheet.getLongOpt();
        optSheet.setRequired(false);
        optSheet.setArgs(3);
        // TODO: allow multiple --sheet arguments, e.g. --sheet 0 2 3 --sheet 1 1 4
        options.addOption(optSheet);

        // parse command line
        try {
            cmd = parser.parse(options, args);
        } catch (ParseException e) {
            System.out.println(e.getMessage());
            formatter.printHelp("xlsx2mysql", options);
            System.exit(1);
        } catch (Exception ex) {
            ex.printStackTrace();
        }

        // If we get here, we have all required command line arguments.
        String inputFileName = cmd.getOptionValue(sInput);
        String outputFileName = cmd.getOptionValue(sOutput);

        try {
            System.out.println("Reading data from file '" + inputFileName + "' ...");
            xlsx2mysql converter = new xlsx2mysql(inputFileName, outputFileName);
            // Depending on the structure of your input, you may have to
            // do something like this:
            // converter.setColumnNamesRowIndex(0, 4); // row #5
            // converter.setFirstDataRowIndex(0, 6); // row #7
            if (cmd.hasOption(sSheet) == true) {
                String sheetArg[] = cmd.getOptionValues(sSheet);
                int sheetIdx = Integer.parseInt(sheetArg[0]);
                int nameRowIdx = Integer.parseInt(sheetArg[1]);
                int dataRowIdx = Integer.parseInt(sheetArg[2]);
                converter.setColumnNamesRowIndex(sheetIdx, nameRowIdx);
                converter.setFirstDataRowIndex(sheetIdx, dataRowIdx);
            }
            /*
             * // These are the settings for data/formatted-test-with-formulae.xlsx // first
             * sheet (index 0) converter.setColumnNamesRowIndex(0, 0); // row #1
             * converter.setFirstDataRowIndex(0, 1); // row #2 // second sheet (index 1)
             * converter.setColumnNamesRowIndex(1, 2); // row #2
             * converter.setFirstDataRowIndex(1, 4); // row #5 // third sheet (index 2)
             * converter.setColumnNamesRowIndex(2, 0); // row #1
             * converter.setFirstDataRowIndex(2, 1); // row #2
             *
             * // This is for 2019-11-05-ETM_DE_42033_Wiegand_20191105.xlsx
             * converter.setColumnNamesRowIndex(0, 4); // row #5
             * converter.setFirstDataRowIndex(0, 6); // row #7
             */
            System.out.println("Creating SQL script ...");
            BufferedWriter writer = new BufferedWriter(new FileWriter(outputFileName));
            System.out.println("Saving SQL script to file '" + outputFileName + "' ...");
            String theScript = converter.generateSqlScript();
            writer.write(theScript);
            writer.close();
            System.out.println("All done. Bye!");
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
}
