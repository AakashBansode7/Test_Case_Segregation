package com.fedex.shipping;

import com.spire.data.table.DataTable;
import com.spire.data.table.common.JdbcAdapter;
import com.spire.xls.ExcelVersion;
import com.spire.xls.LineStyleType;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import javax.swing.*;
import java.awt.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.Properties;


public class Test_Segregation {
    public static String CYCLE = ReadPropertiesFile("CYCLE").trim();
    public static String VERSION = ReadPropertiesFile("VERSION").trim();
    public static String Device = ReadPropertiesFile("Device").trim();
    static String BP = ReadPropertiesFile("BP").trim();

    public static String ReadPropertiesFile(String KeyName) {


        Properties prop = new Properties();
        FileInputStream input;
        String KeyValue = null;

        try {
            input = new FileInputStream("E:\\Ashwini\\Aakash\\TestCase_Segregation\\input.properties");
            prop.load(input);

            KeyValue = prop.getProperty(KeyName);

        } catch (IOException ex) {
            ex.printStackTrace();
        }
        return KeyValue;
    }

    public static void main(String[] args) throws Exception {
        Connection connection;
        System.out.println("Data loading");
        System.out.println("data loaded");
        // Extract data from result set
        Class.forName("oracle.jdbc.driver.OracleDriver");
        System.out.println("DRIVER_LOAD OK");
        connection = DriverManager.getConnection("jdbc:oracle:thin:@T000305-scan.ute.iaas.fedex.com:1526/CMMSHP_PERM_01_CSHIP_S1.ute.iaas.fedex.com", "CSHIP_DBO", "fT6xEFj9wcTtdogXlY87CSHIPDBO");
        System.out.println("CONNECTION  OK");

        //Create Excel sheet
        Workbook wb = new Workbook();
        Workbook wb1 = new Workbook();
        //Get the first worksheet
        Worksheet sheet = wb.getWorksheets().get(0);

        System.out.println("Excel sheet For GRT creation= Ok");


        if (BP.equalsIgnoreCase("B")) {
            ResultSet rs;
            ResultSet rs1;
            ResultSet rs2;
            ResultSet rs3;

            DataTable dataTable = new DataTable();
            DataTable dataTable1 = new DataTable();
            DataTable dataTable2 = new DataTable();
            DataTable dataTable3 = new DataTable();
            //Fetch ILEP Tins
            PreparedStatement stmt = connection.prepareStatement("SELECT  Distinct Test_ID,METERNBR,CUSTNBR,GRPIND,NATBIND,SVCTYPCD,TEAM_ID FROM " + CYCLE + "_L3" + BP + Device + VERSION + "ILEP_BL " + " where   (TEST_ID NOT IN(select distinct test_id from " + CYCLE + "_L3" + BP + Device + VERSION + "ILEP_RSLT )) AND (Team_ID='GRT') ");
            System.out.println(stmt);
            rs = stmt.executeQuery();
            //insert data into excel file
            JdbcAdapter jdbcAdapter = new JdbcAdapter();
            jdbcAdapter.fillDataTable(dataTable, rs);
            //Write datatable to the worksheet
            sheet.insertDataTable(dataTable, true, 1, 1);

            //Auto fit column width
            sheet.getAllocatedRange().autoFitColumns();
            sheet.getAllocatedRange().borderInside(LineStyleType.Thin, Color.BLACK);
            sheet.getAllocatedRange().borderAround(LineStyleType.Thin, Color.BLACK);
            sheet.getCellRange("A1:Q1").getCellStyle().getExcelFont().isBold(true);

            //Save to an Excel TAB
            sheet.setName("ILEP");
            sheet.setTabColor(Color.red);

            //Fetch next tab
            sheet = wb.getWorksheets().get(1);
            System.out.println("opened sheet 1");
            //fetch DMEP
            PreparedStatement stmt1 = connection.prepareStatement("SELECT DISTINCT TEST_ID,SHP_SEQ_NBR,METERNBR,CUSTNBR,DANGGOODSFLG,TEAM_ID FROM " + CYCLE + "_L3" + BP + Device + VERSION + "DMEP_BL " + " where (TEST_ID NOT IN(select distinct test_id from " + CYCLE + "_L3" + BP + Device + VERSION + "DMEP_RSLT )) AND (Team_ID='GRT') ");
            System.out.println(stmt1);
            rs1 = stmt1.executeQuery();
            //insert data into excel file
            JdbcAdapter jdbcAdapter1 = new JdbcAdapter();
            jdbcAdapter1.fillDataTable(dataTable1, rs1);

            //Write datatable to the worksheet
            sheet.insertDataTable(dataTable1, true, 1, 1);

            //Auto fit column width
            sheet.getAllocatedRange().autoFitColumns();
            sheet.getAllocatedRange().borderInside(LineStyleType.Thin, Color.BLACK);
            sheet.getAllocatedRange().borderAround(LineStyleType.Thin, Color.BLACK);
            sheet.getCellRange("A1:Q1").getCellStyle().getExcelFont().isBold(true);
            sheet.setName("DMEP");
            sheet.setTabColor(Color.red);


            //Fetch DMGD
            sheet = wb.getWorksheets().get(2);
            System.out.println("opened sheet 2");
            PreparedStatement stmt2 = connection.prepareStatement("SELECT DISTINCT TEST_ID,METERNBR,CUSTNBR,SP_FLG,RTRNSHPIND,team_id FROM " + CYCLE + "_L3" + BP + Device + VERSION + "DMGD_BL " + " where (TEST_ID NOT IN(select distinct test_id from " + CYCLE + "_L3" + BP + Device + VERSION + "DMGD_RSLT )) AND (Team_ID='GRT') ");
            System.out.println(stmt2);
            rs2 = stmt2.executeQuery();
            //insert data into excel file
            JdbcAdapter jdbcAdapter2 = new JdbcAdapter();
            jdbcAdapter2.fillDataTable(dataTable2, rs2);

            //Write datatable to the worksheet
            sheet.insertDataTable(dataTable2, true, 1, 1);

            //Auto fit column width
            sheet.getAllocatedRange().autoFitColumns();
            sheet.getAllocatedRange().borderInside(LineStyleType.Thin, Color.BLACK);
            sheet.getAllocatedRange().borderAround(LineStyleType.Thin, Color.BLACK);
            sheet.getCellRange("A1:Q1").getCellStyle().getExcelFont().isBold(true);
            sheet.setName("DMGD");
            sheet.setTabColor(Color.green);


            //Fetch ILGD

            sheet = wb.getWorksheets().add("ILGD");
            System.out.println("Sheet created");

            sheet = wb.getWorksheets().get(3);
            System.out.println("opened sheet 3");
            PreparedStatement stmt3 = connection.prepareStatement("SELECT DISTINCT TEST_ID,METERNBR,CUSTNBR,GRPIND,NATBIND,RTRNSHPIND,ITGNBR,team_id FROM " + CYCLE + "_L3" + BP + Device + VERSION + "ILGD_BL " + " where (TEST_ID NOT IN(select distinct test_id from " + CYCLE + "_L3" + BP + Device + VERSION + "ILGD_RSLT )) AND (Team_ID='GRT') ");
            System.out.println(stmt3);
            rs3 = stmt3.executeQuery();
            //insert data into excel file
            JdbcAdapter jdbcAdapter3 = new JdbcAdapter();
            jdbcAdapter3.fillDataTable(dataTable3, rs3);

            //Write datatable to the worksheet
            sheet.insertDataTable(dataTable3, true, 1, 1);

            //Auto fit column width
            sheet.getAllocatedRange().autoFitColumns();
            sheet.getAllocatedRange().borderInside(LineStyleType.Thin, Color.BLACK);
            sheet.getAllocatedRange().borderAround(LineStyleType.Thin, Color.BLACK);
            sheet.getCellRange("A1:Q1").getCellStyle().getExcelFont().isBold(true);
            sheet.setName("ILGD");
            sheet.setTabColor(Color.green);
            sheet = wb.getWorksheets().add("Rough");
            wb.saveToFile("E:\\Ashwini\\Aakash\\TestCase_Segregation\\GRT_Test_Data.xlsx", ExcelVersion.Version2016);
            System.out.println(rs);
            System.out.println(dataTable1);
            rs = null;
            rs1 = null;
            rs2 = null;
            rs3 = null;
            DataTable dataTable4 = new DataTable();
            DataTable dataTable5 = new DataTable();
            DataTable dataTable6 = new DataTable();
            DataTable dataTable7 = new DataTable();

            Worksheet sheet1 = wb1.getWorksheets().get(0);

            System.out.println(dataTable);
            PreparedStatement stmt4 = connection.prepareStatement("SELECT  Distinct Test_ID,METERNBR,CUSTNBR,GRPIND,NATBIND,SVCTYPCD,TEAM_ID FROM " + CYCLE + "_L3" + BP + Device + VERSION + "ILEP_BL " + " where   (TEST_ID NOT IN(select distinct test_id from " + CYCLE + "_L3" + BP + Device + VERSION + "ILEP_RSLT )) AND (Team_ID!='GRT') ");
            System.out.println(stmt4);
            rs = stmt4.executeQuery();
            //insert data into excel file
            JdbcAdapter jdbcAdapter4 = new JdbcAdapter();
            jdbcAdapter4.fillDataTable(dataTable4, rs);
            //Write datatable to the worksheet
            sheet1.insertDataTable(dataTable4, true, 1, 1);

            //Auto fit column width
            sheet1.getAllocatedRange().autoFitColumns();
            sheet1.getAllocatedRange().borderInside(LineStyleType.Thin, Color.BLACK);
            sheet1.getAllocatedRange().borderAround(LineStyleType.Thin, Color.BLACK);
            sheet1.getCellRange("A1:Q1").getCellStyle().getExcelFont().isBold(true);

            //Save to an Excel TAB
            sheet1.setName("ILEP");
            sheet1.setTabColor(Color.red);


            sheet1 = wb1.getWorksheets().get(1);

            System.out.println(dataTable);
            PreparedStatement stmt5 = connection.prepareStatement("SELECT  Distinct TEST_ID,SHP_SEQ_NBR,METERNBR,CUSTNBR,DANGGOODSFLG,TEAM_ID FROM " + CYCLE + "_L3" + BP + Device + VERSION + "DMEP_BL " + " where   (TEST_ID NOT IN(select distinct test_id from " + CYCLE + "_L3" + BP + Device + VERSION + "DMEP_RSLT )) AND (Team_ID!='GRT') ");
            System.out.println(stmt5);
            rs1 = stmt5.executeQuery();
            //insert data into excel file
            JdbcAdapter jdbcAdapter5 = new JdbcAdapter();
            jdbcAdapter5.fillDataTable(dataTable5, rs1);
            //Write datatable to the worksheet
            sheet1.insertDataTable(dataTable5, true, 1, 1);

            //Auto fit column width
            sheet1.getAllocatedRange().autoFitColumns();
            sheet1.getAllocatedRange().borderInside(LineStyleType.Thin, Color.BLACK);
            sheet1.getAllocatedRange().borderAround(LineStyleType.Thin, Color.BLACK);
            sheet1.getCellRange("A1:Q1").getCellStyle().getExcelFont().isBold(true);
            sheet1.setName("DMEP");
            sheet1.setTabColor(Color.red);

            //Fetch DMGD
            sheet1 = wb1.getWorksheets().get(2);
            System.out.println("opened sheet 2");
            PreparedStatement stmt6 = connection.prepareStatement("SELECT DISTINCT TEST_ID,METERNBR,CUSTNBR,SP_FLG,RTRNSHPIND,team_id FROM " + CYCLE + "_L3" + BP + Device + VERSION + "DMGD_BL " + " where (TEST_ID NOT IN(select distinct test_id from " + CYCLE + "_L3" + BP + Device + VERSION + "DMGD_RSLT )) AND (Team_ID!='GRT') ");
            System.out.println(stmt6);
            rs2 = stmt6.executeQuery();
            //insert data into excel file
            JdbcAdapter jdbcAdapter6 = new JdbcAdapter();
            jdbcAdapter2.fillDataTable(dataTable6, rs2);

            //Write datatable to the worksheet
            sheet1.insertDataTable(dataTable6, true, 1, 1);

            //Auto fit column width
            sheet1.getAllocatedRange().autoFitColumns();
            sheet1.getAllocatedRange().borderInside(LineStyleType.Thin, Color.BLACK);
            sheet1.getAllocatedRange().borderAround(LineStyleType.Thin, Color.BLACK);
            sheet1.getCellRange("A1:Q1").getCellStyle().getExcelFont().isBold(true);
            sheet1.setName("DMGD");
            sheet1.setTabColor(Color.green);


            sheet1 = wb1.getWorksheets().add("ILGD");
            System.out.println("Sheet created");
            //Fetch ILGD
            sheet1 = wb1.getWorksheets().get(3);
            System.out.println("opened sheet 3");
            PreparedStatement stmt7 = connection.prepareStatement("SELECT DISTINCT TEST_ID,METERNBR,CUSTNBR,GRPIND,NATBIND,RTRNSHPIND,ITGNBR,team_id FROM " + CYCLE + "_L3" + BP + Device + VERSION + "ILGD_BL " + " where (TEST_ID NOT IN(select distinct test_id from " + CYCLE + "_L3" + BP + Device + VERSION + "ILGD_RSLT )) AND (Team_ID!='GRT') ");
            System.out.println(stmt7);
            rs3 = stmt7.executeQuery();
            //insert data into excel file
            JdbcAdapter jdbcAdapter7 = new JdbcAdapter();
            jdbcAdapter7.fillDataTable(dataTable7, rs3);

            //Write datatable to the worksheet
            sheet1.insertDataTable(dataTable7, true, 1, 1);

            //Auto fit column width
            sheet1.getAllocatedRange().autoFitColumns();
            sheet1.getAllocatedRange().borderInside(LineStyleType.Thin, Color.BLACK);
            sheet1.getAllocatedRange().borderAround(LineStyleType.Thin, Color.BLACK);
            sheet1.getCellRange("A1:Q1").getCellStyle().getExcelFont().isBold(true);
            sheet1.setName("ILGD");
            sheet1.setTabColor(Color.green);
            sheet1 = wb1.getWorksheets().add("Rough");
            wb1.saveToFile("E:\\Ashwini\\Aakash\\TestCase_Segregation\\GTM_Test_Data.xlsx", ExcelVersion.Version2016);

        } else {

            ResultSet rs;
            ResultSet rs1;
            ResultSet rs2;
            ResultSet rs3;


            DataTable dataTable = new DataTable();
            DataTable dataTable1 = new DataTable();
            DataTable dataTable2 = new DataTable();
            DataTable dataTable3 = new DataTable();

            //Fetch ILEP Tins
            PreparedStatement stmt = connection.prepareStatement("SELECT  Distinct Test_ID,METERNBR,CUSTNBR,GRPIND,NATBIND,SVCTYPCD,USMCACERTITY FROM " + CYCLE + "_L3" + BP + Device + VERSION + "ILEP_BL " + " where   (TEST_ID NOT IN(select distinct test_id from " + CYCLE + "_L3" + BP + Device + VERSION + "ILEP_RSLT ))");
            System.out.println(stmt);
            rs = stmt.executeQuery();
            //insert data into excel file
            JdbcAdapter jdbcAdapter = new JdbcAdapter();
            jdbcAdapter.fillDataTable(dataTable, rs);
            //Write datatable to the worksheet
            sheet.insertDataTable(dataTable, true, 1, 1);

            //Auto fit column width
            sheet.getAllocatedRange().autoFitColumns();
            sheet.getAllocatedRange().borderInside(LineStyleType.Thin, Color.BLACK);
            sheet.getAllocatedRange().borderAround(LineStyleType.Thin, Color.BLACK);
            sheet.getCellRange("A1:Q1").getCellStyle().getExcelFont().isBold(true);

            //Save to an Excel TAB
            sheet.setName("ILEP");
            sheet.setTabColor(Color.red);

            //Fetch next tab
            sheet = wb.getWorksheets().get(1);
            System.out.println("opened sheet 1");
            //fetch DMEP
            PreparedStatement stmt1 = connection.prepareStatement("SELECT DISTINCT TEST_ID,SHP_SEQ_NBR,METERNBR,CUSTNBR,DANGGOODSFLG FROM " + CYCLE + "_L3" + BP + Device + VERSION + "DMEP_BL " + " where (TEST_ID NOT IN(select distinct test_id from " + CYCLE + "_L3" + BP + Device + VERSION + "DMEP_RSLT ))");
            System.out.println(stmt1);
            rs1 = stmt1.executeQuery();
            //insert data into excel file
            JdbcAdapter jdbcAdapter1 = new JdbcAdapter();
            jdbcAdapter1.fillDataTable(dataTable1, rs1);

            //Write datatable to the worksheet
            sheet.insertDataTable(dataTable1, true, 1, 1);

            //Auto fit column width
            sheet.getAllocatedRange().autoFitColumns();
            sheet.getAllocatedRange().borderInside(LineStyleType.Thin, Color.BLACK);
            sheet.getAllocatedRange().borderAround(LineStyleType.Thin, Color.BLACK);
            sheet.getCellRange("A1:Q1").getCellStyle().getExcelFont().isBold(true);
            sheet.setName("DMEP");
            sheet.setTabColor(Color.red);

            //Fetch DMGD
            sheet = wb.getWorksheets().get(2);
            System.out.println("opened sheet 2");
            PreparedStatement stmt2 = connection.prepareStatement("SELECT DISTINCT TEST_ID,METERNBR,CUSTNBR,RTRNSHPIND FROM " + CYCLE + "_L3" + BP + Device + VERSION + "DMGD_BL " + " where (TEST_ID NOT IN(select distinct test_id from " + CYCLE + "_L3" + BP + Device + VERSION + "DMGD_RSLT))");
            System.out.println(stmt2);
            rs2 = stmt2.executeQuery();
            //insert data into excel file
            JdbcAdapter jdbcAdapter2 = new JdbcAdapter();
            jdbcAdapter2.fillDataTable(dataTable2, rs2);

            //Write datatable to the worksheet
            sheet.insertDataTable(dataTable2, true, 1, 1);

            //Auto fit column width
            sheet.getAllocatedRange().autoFitColumns();
            sheet.getAllocatedRange().borderInside(LineStyleType.Thin, Color.BLACK);
            sheet.getAllocatedRange().borderAround(LineStyleType.Thin, Color.BLACK);
            sheet.getCellRange("A1:Q1").getCellStyle().getExcelFont().isBold(true);
            sheet.setName("DMGD");
            sheet.setTabColor(Color.green);
            int count = wb.getWorksheets().getCapacity();


            //Fetch ILGD

            sheet = wb.getWorksheets().add("ILGD");
            System.out.println("Sheet created");

            sheet = wb.getWorksheets().get(3);
            System.out.println("opened sheet 3");
            PreparedStatement stmt3 = connection.prepareStatement("SELECT DISTINCT TEST_ID,METERNBR,CUSTNBR,GRPIND,NATBIND,RTRNSHPIND,ITGNBR,USMCACERTITY FROM " + CYCLE + "_L3" + BP + Device + VERSION + "ILGD_BL " + " where (TEST_ID NOT IN(select distinct test_id from " + CYCLE + "_L3" + BP + Device + VERSION + "ILGD_RSLT))");
            System.out.println(stmt3);
            rs3 = stmt3.executeQuery();
            //insert data into excel file
            JdbcAdapter jdbcAdapter3 = new JdbcAdapter();
            jdbcAdapter3.fillDataTable(dataTable3, rs3);

            //Write datatable to the worksheet
            sheet.insertDataTable(dataTable3, true, 1, 1);

            //Auto fit column width
            sheet.getAllocatedRange().autoFitColumns();
            sheet.getAllocatedRange().borderInside(LineStyleType.Thin, Color.BLACK);
            sheet.getAllocatedRange().borderAround(LineStyleType.Thin, Color.BLACK);
            sheet.getCellRange("A1:Q1").getCellStyle().getExcelFont().isBold(true);
            sheet.setName("ILGD");
            sheet.setTabColor(Color.green);
            sheet = wb.getWorksheets().add("Rough");

            sheet = wb.getWorksheets().add("ILGD");
            System.out.println("Sheet created");
            sheet = wb.getWorksheets().add("Rough");

            wb.saveToFile("E:\\Ashwini\\Aakash\\TestCase_Segregation\\CA_Test_Data.xlsx", ExcelVersion.Version2016);

        }


        //Dialogue Box
        ImageIcon icon = new ImageIcon("src\\test\\Fedex-logo-500x281.png");
        Image image = icon.getImage(); // transform it
        Image newimg = image.getScaledInstance(80, 60, java.awt.Image.SCALE_SMOOTH);
        icon = new ImageIcon(newimg);
        JPanel panel = new JPanel();
        panel.setBackground(new Color(102, 205, 170));
        panel.setSize(new Dimension(200, 64));
        panel.setLayout(null);
        JLabel label = new JLabel("Completed");
        label.setBounds(0, 0, 200, 80);
        label.setFont(new Font("Arial", Font.BOLD, 15));
        label.setHorizontalAlignment(SwingConstants.LEFT);
        label.setVerticalAlignment(SwingConstants.TOP);
        panel.add(label);
        JOptionPane.showMessageDialog(null, panel, "Script status", JOptionPane.PLAIN_MESSAGE, icon);
    }
}