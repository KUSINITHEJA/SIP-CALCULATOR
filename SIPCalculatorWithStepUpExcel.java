package com.example.sipCalulator.SIP;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Scanner;

public class SIPCalculatorWithStepUpExcel {

    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);

        // User inputs
        System.out.print("Enter initial monthly SIP amount (₹): ");
        double monthlySIP = scanner.nextDouble();

        System.out.print("Enter annual step-up percentage (%): ");
        double stepUpPercent = scanner.nextDouble();

        System.out.print("Enter investment duration (in years): ");
        int years = scanner.nextInt();

        System.out.print("Enter expected annual rate of return (%): ");
        double annualRate = scanner.nextDouble();

        double monthlyRate = annualRate / 12 / 100;
        int totalMonths = years * 12;
        long futureValue = 0;

        // Create Excel workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("SIP Calculation");

        // Header row
        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Year");
        header.createCell(1).setCellValue("Monthly SIP (₹)");
        header.createCell(2).setCellValue("Total Contribution (₹)");
        header.createCell(3).setCellValue("Future Value (₹)");
        header.createCell(4).setCellValue("Step-Up (%)");

        // Step-up info row
        Row stepUpRow = sheet.createRow(1);
        stepUpRow.createCell(4).setCellValue(Math.round(stepUpPercent));

        // Loop through each year
        for (int year = 0; year < years; year++) {
            double currentMonthlySIP = monthlySIP * Math.pow(1 + stepUpPercent / 100, year);
            double yearFutureValue = 0.0;

            for (int month = 0; month < 12; month++) {
                int monthsRemaining = totalMonths - (year * 12 + month);
                yearFutureValue += currentMonthlySIP * Math.pow(1 + monthlyRate, monthsRemaining);
            }

            long roundedMonthlySIP = Math.round(currentMonthlySIP);
            long roundedContribution = roundedMonthlySIP * 12;
            long roundedYearFutureValue = Math.round(yearFutureValue);
            futureValue += roundedYearFutureValue;

            // Write to Excel
            Row row = sheet.createRow(year + 2); // +2 to account for header and step-up row
            row.createCell(0).setCellValue(year + 1);
            row.createCell(1).setCellValue(roundedMonthlySIP);
            row.createCell(2).setCellValue(roundedContribution);
            row.createCell(3).setCellValue(roundedYearFutureValue);
        }

        // Final summary row
        Row summary = sheet.createRow(years + 3);
        summary.createCell(2).setCellValue("Total Future Value:");
        summary.createCell(3).setCellValue(futureValue);

        // Auto-size columns
        for (int i = 0; i <= 4; i++) {
            sheet.autoSizeColumn(i);
        }

        // Generate unique filename with timestamp
        String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss").format(new Date());
        String fileName = "SIP_StepUp_Report_" + timestamp + ".xlsx";

        // Save Excel file
        try (FileOutputStream fileOut = new FileOutputStream(fileName)) {
            workbook.write(fileOut);
            workbook.close();
            System.out.println("Excel file '" + fileName + "' created successfully.");
        } catch (IOException e) {
            System.out.println("Error writing Excel file: " + e.getMessage());
        }
    }
}
