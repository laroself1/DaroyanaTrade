import Beans.Truck;
import org.apache.poi.POIXMLProperties;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.sql.Time;
import java.util.*;
//0670652727
/**
 * Created by LaroSelf on 20.05.2016.
 */
public class FileOperations {
    public static List<String> getManagersNames(String fileOut) throws IOException,InvalidFormatException,OpenXML4JException {
        List<Truck> list = new ArrayList<Truck>();
        DataFormatter objDefaultFormat = new DataFormatter();
        FileInputStream input_document = new FileInputStream(new File(fileOut));
        XSSFWorkbook my_xlsx_workbook = new XSSFWorkbook(input_document);
        XSSFSheet sheet = my_xlsx_workbook.getSheet("Сводка");
        Iterator<Row> objIterator = sheet.rowIterator();

        List<String> consNames = new ArrayList<String>();
        List<String> consNamesForComb = new ArrayList<String>();
        while (objIterator.hasNext()) {

            Row row = objIterator.next();
            if (row.getRowNum()>=2) {
                String manager = objDefaultFormat.formatCellValue(row.getCell(24)).trim();
                if(!manager.equals("")){
                    if(!manager.equals(0)) {

                        consNames.add(manager);
                    }}}}
                    for (String a:consNames ) {
                        if(!consNamesForComb.contains(a))
                        {
                            consNamesForComb.add(a);
                            System.out.println(a);}

                    }




                    System.out.println("----------");
                    input_document.close();
              return consNamesForComb;}

    public static List<Truck> getManagersTrucks(String fileOut, String chosenMan) throws IOException,InvalidFormatException,OpenXML4JException {
        List<Truck> list = new ArrayList<Truck>();
        DataFormatter objDefaultFormat = new DataFormatter();
        FileInputStream input_document = new FileInputStream(new File(fileOut));
        XSSFWorkbook my_xlsx_workbook = new XSSFWorkbook(input_document);
        XSSFSheet sheet = my_xlsx_workbook.getSheet("Сводка");
        Iterator<Row> objIterator = sheet.rowIterator();
        while (objIterator.hasNext()) {

            Row row = objIterator.next();
            if (row.getRowNum()>=2) {
                String manager = objDefaultFormat.formatCellValue(row.getCell(24)).trim();
                if(!manager.equals("")){
                    if(!manager.equals(0)){
                        if(manager.equals(chosenMan)) {
                            String position = objDefaultFormat.formatCellValue(row.getCell(0)).trim();
                            String goodClass = objDefaultFormat.formatCellValue(row.getCell(6)).trim();
                            String number = objDefaultFormat.formatCellValue(row.getCell(3)).trim();
                            if (position.equals("") & goodClass.equals("")) {
                                System.out.println("----no truck position-");
                                break;
                            }
                            String date = objDefaultFormat.formatCellValue(row.getCell(1)).trim();
                            if (date.contains("/")) {
                                String dateTemp = date.replace("/", ".").trim();
                                StringTokenizer st = new StringTokenizer(dateTemp);
                                String arr[] = dateTemp.split("\\.");

                                if (arr[1].length() < 2) {
                                    arr[1] = "0" + arr[1];
                                }
                                if (arr[0].length() < 2) {
                                    arr[0] = "0" + arr[0];
                                }
                                if (arr[2].length() == 2) {
                                    arr[2] = "20" + arr[2];
                                }
                                date = arr[1] + "." + arr[0] + "." + arr[2];
                            }

                            String invoice = objDefaultFormat.formatCellValue(row.getCell(2)).trim();

                            if (goodClass == null) {
                                goodClass = objDefaultFormat.formatCellValue(row.getCell(7)).trim();
                            }

                            String weight = objDefaultFormat.formatCellValue(row.getCell(4)).trim();
                            if (weight.equals("") | weight.equals("0") | weight.equals("0,00") | weight.equals("0.00")
                                    | weight.equals("0.0") | weight.equals("0,0") | weight.equals("0,000") | weight.equals("0.000")) {
                                weight = "0";
                            } else if (row.getCell(4).getCellType() == Cell.CELL_TYPE_FORMULA) {

                                double tempW = row.getCell(4).getNumericCellValue();
                                weight = Double.toString(tempW);
                                weight = weight.replace('.', ',');

                            } else {

                                weight = weight.replace(',', '.');
                                double a = Double.parseDouble(weight);

                                weight = Double.toString(a);
                                weight = weight.replace('.', ',');

                                if ((weight.charAt(weight.length() - 1)) == ',') {
                                    weight = weight.substring(0, weight.length() - 1);
                                }

                            }

                            String consignor = objDefaultFormat.formatCellValue(row.getCell(5)).trim();
                            String warehouse = objDefaultFormat.formatCellValue(row.getCell(8)).trim();
                            Truck truck = new Truck(position, date, invoice, number, weight, consignor, goodClass, warehouse);
                            list.add(truck);

                            System.out.println(position + " " + date + " " + invoice + " " + number + " " + goodClass + " " + weight + " " + consignor + " " + warehouse);
                        } }}





        }}System.out.println("----------");
        input_document.close();
        return  list;}

    public static void saveNewRegister2(String fileOut, List<Truck> result) throws FileNotFoundException, IOException {

        List<Truck> forSave = result;


        FileInputStream input_document = new FileInputStream(new File(fileOut));
        XSSFWorkbook workbook = new XSSFWorkbook(input_document);
        XSSFSheet sheet =workbook.getSheet("Горох");

        XSSFCellStyle cellStyle = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontName(HSSFFont.FONT_ARIAL);
        font.setFontHeightInPoints((short)8);
        cellStyle.setFont(font);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        int lastRowW=0;
        int lastRowP=0;
        int lastRowB=0;
        XSSFCellStyle cellStyleC = workbook.createCellStyle();
        cellStyleC.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyleC.setFont(font);
        cellStyleC.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyleC.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyleC.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyleC.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        XSSFCellStyle cellStyleR = workbook.createCellStyle();
        cellStyleR.setAlignment(CellStyle.ALIGN_RIGHT);
        cellStyleR.setFont(font);
        cellStyleR.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyleR.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyleR.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyleR.setBorderLeft(HSSFCellStyle.BORDER_THIN);

        for(int i=0; i<forSave.size(); i++){

            Truck truck=forSave.get(i);
            if(truck.getGoodClass().equals("ФУР")|truck.getGoodClass().equals("3КЛ")|truck.getGoodClass().equals("2КЛ")|
                    truck.getGoodClass().equals("фур")|truck.getGoodClass().equals("3кл")|truck.getGoodClass().equals("2кл")
                    |truck.getGoodClass().equals("1-й")|truck.getGoodClass().equals("2-й")|truck.getGoodClass().equals("3-й")
                    |truck.getGoodClass().equals("фураж") |truck.getGoodClass().equals("ФУРАЖ")){
                sheet = workbook.getSheet("Пшеница");
                lastRowW = sheet.getPhysicalNumberOfRows();
                System.out.println(lastRowW);


                Row row=sheet.createRow(lastRowW);

                Cell cellPos= row.createCell(0);
                Cell cellDate= row.createCell(1);
                Cell cellInvoice= row.createCell(2);
                Cell cellNumber= row.createCell(3);
                Cell cellGood= row.createCell(6);
                //Cell cellGoodPayFor = row.createCell(7);
                Cell cellWeight =row.createCell(4);
                Cell cellCons = row.createCell(5);
                Cell cellWarehouse = row.createCell(7);

                cellPos.setCellStyle(cellStyleR);
                cellDate.setCellStyle(cellStyle);
                cellInvoice.setCellStyle(cellStyle);
                cellNumber.setCellStyle(cellStyle);
                cellGood.setCellStyle(cellStyleC);
                cellWeight.setCellStyle(cellStyleR);
                cellCons.setCellStyle(cellStyle);
                cellWarehouse.setCellStyle(cellStyleC);

                cellPos.setCellValue(truck.getPosition());
                cellDate.setCellValue(truck.getDate());
                cellInvoice.setCellValue(truck.getInvoice());
                cellNumber.setCellValue(truck.getNumber());
                cellGood.setCellValue(truck.getGoodClass());
                cellWeight.setCellValue(Double.parseDouble(truck.getWeight().replace(",",".")));
                cellCons.setCellValue(truck.getConsignor());
                cellWarehouse.setCellValue(truck.getWarehouse());
                ++lastRowW;}}

        for(int i=0; i<forSave.size(); i++) {
            Truck truck = forSave.get(i);
            if (truck.getGoodClass().equals("горох")) {
                sheet = workbook.getSheet("Горох");
                lastRowP = sheet.getPhysicalNumberOfRows();
                System.out.println(lastRowP);
                Row row=sheet.createRow(lastRowP);

                Cell cellPos= row.createCell(0);
                Cell cellDate= row.createCell(1);
                Cell cellInvoice= row.createCell(2);
                Cell cellNumber= row.createCell(3);
                Cell cellGood= row.createCell(6);
                Cell cellWeight =row.createCell(4);
                Cell cellCons = row.createCell(5);
                Cell cellWarehouse = row.createCell(7);

                cellPos.setCellStyle(cellStyleR);
                cellDate.setCellStyle(cellStyle);
                cellInvoice.setCellStyle(cellStyle);
                cellNumber.setCellStyle(cellStyle);
                cellGood.setCellStyle(cellStyleC);
                cellWeight.setCellStyle(cellStyleR);
                cellCons.setCellStyle(cellStyle);
                cellWarehouse.setCellStyle(cellStyleC);

                cellPos.setCellValue(truck.getPosition());
                cellDate.setCellValue(truck.getDate());
                cellInvoice.setCellValue(truck.getInvoice());
                cellNumber.setCellValue(truck.getNumber());
                cellGood.setCellValue(truck.getGoodClass());
                cellWeight.setCellValue(Double.parseDouble(truck.getWeight().replace(",",".")));
                cellCons.setCellValue(truck.getConsignor());
                cellWarehouse.setCellValue(truck.getWarehouse());
                ++lastRowP;}

        }
        for(int i=0; i<forSave.size(); i++) {

            Truck truck = forSave.get(i);
            if(truck.getGoodClass().equals("ячмень")){
                sheet = workbook.getSheet("Ячмень");
                lastRowB = sheet.getPhysicalNumberOfRows();
                System.out.println(lastRowB);

                Row row=sheet.createRow(lastRowB);

                Cell cellPos= row.createCell(0);
                Cell cellDate= row.createCell(1);
                Cell cellInvoice= row.createCell(2);
                Cell cellNumber= row.createCell(3);
                Cell cellGood= row.createCell(6);

                Cell cellWeight =row.createCell(4);
                Cell cellCons = row.createCell(5);
                Cell cellWarehouse = row.createCell(7);

                cellPos.setCellStyle(cellStyleR);
                cellDate.setCellStyle(cellStyle);
                cellInvoice.setCellStyle(cellStyle);
                cellNumber.setCellStyle(cellStyle);
                cellGood.setCellStyle(cellStyleC);
                cellWeight.setCellStyle(cellStyleR);
                cellCons.setCellStyle(cellStyle);
                cellWarehouse.setCellStyle(cellStyleC);

                cellPos.setCellValue(truck.getPosition());
                cellDate.setCellValue(truck.getDate());
                cellInvoice.setCellValue(truck.getInvoice());
                cellNumber.setCellValue(truck.getNumber());
                cellGood.setCellValue(truck.getGoodClass());
                cellWeight.setCellValue(Double.parseDouble(truck.getWeight().replace(",",".")));
                cellCons.setCellValue(truck.getConsignor());
                cellWarehouse.setCellValue(truck.getWarehouse());
                ++lastRowB;}

        }
        for(int i=0; i<forSave.size(); i++) {

            Truck truck = forSave.get(i);
            if(truck.getGoodClass().equals("ШРОТ")|truck.getGoodClass().equals("шрот")){
                sheet = workbook.getSheet("Шрот");
                lastRowB = sheet.getPhysicalNumberOfRows();
                System.out.println(lastRowB);

                Row row=sheet.createRow(lastRowB);

                Cell cellPos= row.createCell(0);
                Cell cellDate= row.createCell(1);
                Cell cellInvoice= row.createCell(2);
                Cell cellNumber= row.createCell(3);
                Cell cellGood= row.createCell(6);

                Cell cellWeight =row.createCell(4);
                Cell cellCons = row.createCell(5);
                Cell cellWarehouse = row.createCell(7);


                cellPos.setCellStyle(cellStyleR);
                cellDate.setCellStyle(cellStyle);
                cellInvoice.setCellStyle(cellStyle);
                cellNumber.setCellStyle(cellStyle);
                cellGood.setCellStyle(cellStyleC);
                cellWeight.setCellStyle(cellStyleR);
                cellCons.setCellStyle(cellStyle);
                cellWarehouse.setCellStyle(cellStyleC);

                cellPos.setCellValue(truck.getPosition());
                cellDate.setCellValue(truck.getDate());
                cellInvoice.setCellValue(truck.getInvoice());
                cellNumber.setCellValue(truck.getNumber());
                cellGood.setCellValue(truck.getGoodClass());
                cellWeight.setCellValue(Double.parseDouble(truck.getWeight().replace(",",".")));
                cellCons.setCellValue(truck.getConsignor());
                cellWarehouse.setCellValue(truck.getWarehouse());
                ++lastRowB;}

        }
        for(int i=0; i<forSave.size(); i++) {

            Truck truck = forSave.get(i);
            if(truck.getGoodClass().equals("ШРОТ")|truck.getGoodClass().equals("шрот")){
                sheet = workbook.getSheet("Шрот");
                lastRowB = sheet.getPhysicalNumberOfRows();
                System.out.println(lastRowB);

                Row row=sheet.createRow(lastRowB);

                Cell cellPos= row.createCell(0);
                Cell cellDate= row.createCell(1);
                Cell cellInvoice= row.createCell(2);
                Cell cellNumber= row.createCell(3);
                Cell cellGood= row.createCell(6);
                Cell cellWeight =row.createCell(4);
                Cell cellCons = row.createCell(5);
                Cell cellWarehouse = row.createCell(7);

                cellPos.setCellStyle(cellStyleR);
                cellDate.setCellStyle(cellStyle);
                cellInvoice.setCellStyle(cellStyle);
                cellNumber.setCellStyle(cellStyle);
                cellGood.setCellStyle(cellStyleC);
                cellWeight.setCellStyle(cellStyleR);
                cellCons.setCellStyle(cellStyle);
                cellWarehouse.setCellStyle(cellStyleC);

                cellPos.setCellValue(truck.getPosition());
                cellDate.setCellValue(truck.getDate());
                cellInvoice.setCellValue(truck.getInvoice());
                cellNumber.setCellValue(truck.getNumber());
                cellGood.setCellValue(truck.getGoodClass());
                cellWeight.setCellValue(Double.parseDouble(truck.getWeight().replace(",",".")));
                cellCons.setCellValue(truck.getConsignor());
                cellWarehouse.setCellValue(truck.getWarehouse());
                ++lastRowB;}

        }
        for(int i=0; i<forSave.size(); i++) {

            Truck truck = forSave.get(i);
            if(truck.getGoodClass().equals("кукуруза")|truck.getGoodClass().equals("КУКУРУЗА")){
                sheet = workbook.getSheet("Кукуруза");
                lastRowB = sheet.getPhysicalNumberOfRows();
                System.out.println(lastRowB);

                Row row=sheet.createRow(lastRowB);

                Cell cellPos= row.createCell(0);
                Cell cellDate= row.createCell(1);
                Cell cellInvoice= row.createCell(2);
                Cell cellNumber= row.createCell(3);
                Cell cellGood= row.createCell(6);

                Cell cellWeight =row.createCell(4);
                Cell cellCons = row.createCell(5);
                Cell cellWarehouse = row.createCell(7);


                cellPos.setCellStyle(cellStyle);
                cellDate.setCellStyle(cellStyle);
                cellInvoice.setCellStyle(cellStyle);
                cellNumber.setCellStyle(cellStyle);
                cellGood.setCellStyle(cellStyleC);
                cellWeight.setCellStyle(cellStyle);
                cellCons.setCellStyle(cellStyle);
                cellWarehouse.setCellStyle(cellStyleC);

                cellPos.setCellValue(truck.getPosition());
                cellDate.setCellValue(truck.getDate());
                cellInvoice.setCellValue(truck.getInvoice());
                cellNumber.setCellValue(truck.getNumber());
                cellGood.setCellValue(truck.getGoodClass());
                cellWeight.setCellValue(Double.parseDouble(truck.getWeight().replace(",",".")));
                cellCons.setCellValue(truck.getConsignor());
                cellWarehouse.setCellValue(truck.getWarehouse());
                ++lastRowB;}

        }
        for(int i=0; i<forSave.size(); i++) {

            Truck truck = forSave.get(i);
            if(truck.getGoodClass().equals("СОЯ")|truck.getGoodClass().equals("соя")){
                sheet = workbook.getSheet("Соя");
                lastRowB = sheet.getPhysicalNumberOfRows();
                System.out.println(lastRowB);

                Row row=sheet.createRow(lastRowB);

                Cell cellPos= row.createCell(0);
                Cell cellDate= row.createCell(1);
                Cell cellInvoice= row.createCell(2);
                Cell cellNumber= row.createCell(3);
                Cell cellGood= row.createCell(6);
                Cell cellWeight =row.createCell(4);
                Cell cellCons = row.createCell(5);
                Cell cellWarehouse = row.createCell(7);

                cellPos.setCellStyle(cellStyleR);
                cellDate.setCellStyle(cellStyle);
                cellInvoice.setCellStyle(cellStyle);
                cellNumber.setCellStyle(cellStyle);
                cellGood.setCellStyle(cellStyleC);
                cellWeight.setCellStyle(cellStyleR);
                cellCons.setCellStyle(cellStyle);
                cellWarehouse.setCellStyle(cellStyleC);

                cellPos.setCellValue(truck.getPosition());
                cellDate.setCellValue(truck.getDate());
                cellInvoice.setCellValue(truck.getInvoice());
                cellNumber.setCellValue(truck.getNumber());
                cellGood.setCellValue(truck.getGoodClass());
                cellWeight.setCellValue(Double.parseDouble(truck.getWeight().replace(",",".")));
                cellCons.setCellValue(truck.getConsignor());
                cellWarehouse.setCellValue(truck.getWarehouse());
                ++lastRowB;}
        }

        FileOutputStream output_file =new FileOutputStream(new File(fileOut));


        workbook.write(output_file);


        input_document.close();
        output_file.close();

    }


    public static List<Truck> getGlobal(String fileIn) throws IOException {
        List<Truck> list = new ArrayList<Truck>();
        HSSFWorkbook excelInReport = new HSSFWorkbook(new FileInputStream(fileIn));
        DataFormatter objDefaultFormat = new DataFormatter();
        HSSFSheet sheet = excelInReport.getSheet("Вагоны, Машины");
        Iterator<Row> objIterator = sheet.rowIterator();
        while (objIterator.hasNext()) {

            Row row = objIterator.next();
            if (row.getRowNum()>=6) {

                String position = objDefaultFormat.formatCellValue(row.getCell(0));
                if(position.equals("Итого")){System.out.println("----------");break;}
                String date = objDefaultFormat.formatCellValue(row.getCell(2));
                if (date.contains("/")){
                    String dateTemp = date.replace("/","." );
                    StringTokenizer st = new StringTokenizer(dateTemp);
                    String arr[]=dateTemp.split("\\.");
                    if(arr[1].length()<2){arr[1]="0"+arr[1];}
                    if(arr[0].length()<2){arr[0]="0"+arr[0];}
                    if(arr[2].length()==2){arr[2]="20"+arr[2];}
                    date=arr[1]+"."+arr[0]+"."+arr[2];}
                String invoice = objDefaultFormat.formatCellValue(row.getCell(3)).trim();
                String number = objDefaultFormat.formatCellValue(row.getCell(4)).trim();
                String goodClass = objDefaultFormat.formatCellValue(row.getCell(5)).trim();
                if(goodClass.equals("Пшениця")){goodClass=objDefaultFormat.formatCellValue(row.getCell(6)).trim();}
                String weight = objDefaultFormat.formatCellValue(row.getCell(14)).trim();
                if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                if ((weight.charAt(weight.length()-1))==','){ weight=weight.substring(0,weight.length()-1);}
                String consignor = objDefaultFormat.formatCellValue(row.getCell(18)).trim();
                String warehouse= "Карготранс";
                Truck truck= new Truck(position,date,invoice,number,weight,consignor,goodClass,warehouse);
                list.add(truck);

                System.out.println(position + " " + date+ " " + invoice+ " " + number+" "+goodClass+" "+weight+" "+consignor+" "+warehouse);
            }

        }excelInReport.close();
        return  list;}

    public static List<Truck> getUmaAgro(String fileIn) throws IOException {
        List<Truck> list = new ArrayList<Truck>();
        DataFormatter objDefaultFormat = new DataFormatter();

        FileInputStream input_document = new FileInputStream(new File(fileIn));
        XSSFWorkbook workbook = new XSSFWorkbook(input_document);
        XSSFSheet sheetW = workbook.getSheet("пшеница");
        XSSFSheet sheetSH = workbook.getSheet("шрот");
        XSSFSheet sheetP = workbook.getSheet("ГОРОХ");
        XSSFSheet sheetCorn = workbook.getSheet("кукуруза");
        XSSFSheet sheetBar = workbook.getSheet("ячмень");
        XSSFSheet sheetSoy = workbook.getSheet("СОЯ");
        Iterator<Row> objIterator = sheetW.rowIterator();
        while (objIterator.hasNext()) {

            Row row = objIterator.next();
            if (row.getRowNum()>508) {

                String position = objDefaultFormat.formatCellValue(row.getCell(0));
                if(position.equals("")){System.out.println("----------");break;}
                String date = objDefaultFormat.formatCellValue(row.getCell(2));
                     if (date.contains("/")){
               String dateTemp = date.replace("/","." );
                    StringTokenizer st = new StringTokenizer(dateTemp);
                    String arr[]=dateTemp.split("\\.");

                    if(arr[1].length()<2){arr[1]="0"+arr[1];}
                    if(arr[0].length()<2){arr[0]="0"+arr[0];}
                    if(arr[2].length()==2){arr[2]="20"+arr[2];}
                    date=arr[1]+"."+arr[0]+"."+arr[2];}
                String invoice = objDefaultFormat.formatCellValue(row.getCell(3)).trim();
                String number = objDefaultFormat.formatCellValue(row.getCell(4)).trim();

                String consignor =  objDefaultFormat.formatCellValue(row.getCell(12)).trim();

                String goodClass = objDefaultFormat.formatCellValue(row.getCell(13)).trim().toUpperCase();
                if (goodClass.equals("3КЛ")|goodClass.equals("3 КЛ")){goodClass="3-й";}
                if (goodClass.equals("2КЛ")|goodClass.equals("2 КЛ")){goodClass="2-й";}
                if (goodClass.equals("1КЛ")|goodClass.equals("1 КЛ")){goodClass="1-й";}
                if (goodClass.equals("ФУРАЖ")|goodClass.equals("ФУР")){goodClass="фураж";}

                String weightUnForm = objDefaultFormat.formatCellValue(row.getCell(10)).trim();
                String weight=null;

                if(weightUnForm.length()>=3){
                    weight=new StringBuffer(weightUnForm).insert(weightUnForm.length()-3, ",").toString();

                     if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))==','){ weight=weight.substring(0,weight.length()-1);}

                    weight= weight.replace(',','.');
                    double a= Double.parseDouble(weight);
                    weight = Double.toString(a);
                    weight= weight.replace('.',',');


                }

                String warehouse= objDefaultFormat.formatCellValue(row.getCell(14)).trim();
                if(warehouse.equals("15")){warehouse="15скл";}
                Truck truck= new Truck(position,date,invoice,number,weight,consignor,goodClass,warehouse);
                list.add(truck);

                System.out.println(position + " " + date+ " " + invoice+ " " + number+" "+goodClass+" "+weight+" "+consignor+" "+warehouse);
            }

        }

        Iterator<Row> objIteratorSH = sheetSH.rowIterator();
        while (objIteratorSH.hasNext()) {

            Row row = objIteratorSH.next();
            if (row.getRowNum()>19) {

                String position = objDefaultFormat.formatCellValue(row.getCell(0));
                if(position.equals("")){System.out.println("----------");break;}
                String date = objDefaultFormat.formatCellValue(row.getCell(2));
                if (date.contains("/")){
                    String dateTemp = date.replace("/","." );
                    StringTokenizer st = new StringTokenizer(dateTemp);
                    String arr[]=dateTemp.split("\\.");

                    if(arr[1].length()<2){arr[1]="0"+arr[1];}
                    if(arr[0].length()<2){arr[0]="0"+arr[0];}
                    if(arr[2].length()==2){arr[2]="20"+arr[2];}
                    date=arr[1]+"."+arr[0]+"."+arr[2];}
                String invoice = objDefaultFormat.formatCellValue(row.getCell(3)).trim();
                String number = objDefaultFormat.formatCellValue(row.getCell(4)).trim();

                String consignor =  objDefaultFormat.formatCellValue(row.getCell(12)).trim();

                String goodClass = objDefaultFormat.formatCellValue(row.getCell(13)).trim();

                String weightUnForm = objDefaultFormat.formatCellValue(row.getCell(10)).trim();
                String weight=null;
                if(weightUnForm.equals("")|weightUnForm.equals("0")){ weight="0";}

                if(weightUnForm.length()>=3){
                    weight=new StringBuffer(weightUnForm).insert(weightUnForm.length()-3, ",").toString();

                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))==','){ weight=weight.substring(0,weight.length()-1);}

                    weight= weight.replace(',','.');
                    double a= Double.parseDouble(weight);
                    weight = Double.toString(a);
                    weight= weight.replace('.',',');

                }

                String warehouse= objDefaultFormat.formatCellValue(row.getCell(14)).trim();
                if(warehouse.equals("15")){warehouse="15скл";}
                Truck truck= new Truck(position,date,invoice,number,weight,consignor,goodClass,warehouse);
                list.add(truck);

                System.out.println(position + " " + date+ " " + invoice+ " " + number+" "+goodClass+" "+weight+" "+consignor+" "+warehouse);
            }

        }
        Iterator<Row> objIteratorP = sheetP.rowIterator();
        while (objIteratorP.hasNext()) {

            Row row = objIteratorP.next();
            if (row.getRowNum()>82) {

                String position = objDefaultFormat.formatCellValue(row.getCell(0));
                if(position.equals("")){System.out.println("----------");break;}
                String date = objDefaultFormat.formatCellValue(row.getCell(2));
                if (date.contains("/")){
                    String dateTemp = date.replace("/","." );
                    StringTokenizer st = new StringTokenizer(dateTemp);
                    String arr[]=dateTemp.split("\\.");

                    if(arr[1].length()<2){arr[1]="0"+arr[1];}
                    if(arr[0].length()<2){arr[0]="0"+arr[0];}
                    if(arr[2].length()==2){arr[2]="20"+arr[2];}
                    date=arr[1]+"."+arr[0]+"."+arr[2];}
                String invoice = objDefaultFormat.formatCellValue(row.getCell(3)).trim();
                String number = objDefaultFormat.formatCellValue(row.getCell(4)).trim();

                String consignor =  objDefaultFormat.formatCellValue(row.getCell(12)).trim();

                String goodClass = objDefaultFormat.formatCellValue(row.getCell(13)).trim();

                String weight=null;
                String weightUnForm = objDefaultFormat.formatCellValue(row.getCell(10)).trim();
                if(weightUnForm.equals("")|weightUnForm.equals("0")){ weight="0";}


                if(weightUnForm.length()>=3){
                    weight=new StringBuffer(weightUnForm).insert(weightUnForm.length()-3, ",").toString();

                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))==','){ weight=weight.substring(0,weight.length()-1);}

                    weight= weight.replace(',','.');
                    double a= Double.parseDouble(weight);
                    weight = Double.toString(a);
                    weight= weight.replace('.',',');

                }

                String warehouse= objDefaultFormat.formatCellValue(row.getCell(14)).trim();
                if(warehouse.equals("15")){warehouse="15скл";}
                Truck truck= new Truck(position,date,invoice,number,weight,consignor,goodClass,warehouse);
                list.add(truck);

                System.out.println(position + " " + date+ " " + invoice+ " " + number+" "+goodClass+" "+weight+" "+consignor+" "+warehouse);
            }

        }
        Iterator<Row> objIteratorCorn = sheetCorn.rowIterator();
        while (objIteratorCorn.hasNext()) {

            Row row = objIteratorCorn.next();
            if (row.getRowNum()>22) {

                String position = objDefaultFormat.formatCellValue(row.getCell(0));
                if(position.equals("")){System.out.println("----------");break;}
                String date = objDefaultFormat.formatCellValue(row.getCell(2));
                if (date.contains("/")){
                    String dateTemp = date.replace("/","." );
                    StringTokenizer st = new StringTokenizer(dateTemp);
                    String arr[]=dateTemp.split("\\.");

                    if(arr[1].length()<2){arr[1]="0"+arr[1];}
                    if(arr[0].length()<2){arr[0]="0"+arr[0];}
                    if(arr[2].length()==2){arr[2]="20"+arr[2];}
                    date=arr[1]+"."+arr[0]+"."+arr[2];}
                String invoice = objDefaultFormat.formatCellValue(row.getCell(3)).trim();
                String number = objDefaultFormat.formatCellValue(row.getCell(4)).trim();

                String consignor =  objDefaultFormat.formatCellValue(row.getCell(12)).trim();

                String goodClass = objDefaultFormat.formatCellValue(row.getCell(13)).trim();

                String weight=null;
                String weightUnForm = objDefaultFormat.formatCellValue(row.getCell(10)).trim();
                if(weightUnForm.equals("")|weightUnForm.equals("0")){ weight="0";}


                if(weightUnForm.length()>=3){
                    weight=new StringBuffer(weightUnForm).insert(weightUnForm.length()-3, ",").toString();

                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))==','){ weight=weight.substring(0,weight.length()-1);}

                    weight= weight.replace(',','.');
                    double a= Double.parseDouble(weight);
                    weight = Double.toString(a);
                    weight= weight.replace('.',',');

                }

                String warehouse= objDefaultFormat.formatCellValue(row.getCell(14)).trim();
                if(warehouse.equals("15")){warehouse="15скл";}
                Truck truck= new Truck(position,date,invoice,number,weight,consignor,goodClass,warehouse);
                list.add(truck);

                System.out.println(position + " " + date+ " " + invoice+ " " + number+" "+goodClass+" "+weight+" "+consignor+" "+warehouse);
            }

        }
        Iterator<Row> objIteratorBar = sheetBar.rowIterator();
        while (objIteratorBar.hasNext()) {

            Row row = objIteratorBar.next();
            if (row.getRowNum()>130) {

                String position = objDefaultFormat.formatCellValue(row.getCell(0));
                if(position.equals("")){System.out.println("----------");break;}
                String date = objDefaultFormat.formatCellValue(row.getCell(2));
                if (date.contains("/")){
                    String dateTemp = date.replace("/","." );
                    StringTokenizer st = new StringTokenizer(dateTemp);
                    String arr[]=dateTemp.split("\\.");

                    if(arr[1].length()<2){arr[1]="0"+arr[1];}
                    if(arr[0].length()<2){arr[0]="0"+arr[0];}
                    if(arr[2].length()==2){arr[2]="20"+arr[2];}
                    date=arr[1]+"."+arr[0]+"."+arr[2];}
                String invoice = objDefaultFormat.formatCellValue(row.getCell(3)).trim();
                String number = objDefaultFormat.formatCellValue(row.getCell(4)).trim();

                String consignor =  objDefaultFormat.formatCellValue(row.getCell(12)).trim();

                String goodClass = objDefaultFormat.formatCellValue(row.getCell(13)).trim();

                String weight=null;
                String weightUnForm = objDefaultFormat.formatCellValue(row.getCell(10)).trim();
                if(weightUnForm.equals("")|weightUnForm.equals("0")){ weight="0";}


                if(weightUnForm.length()>=3){
                    weight=new StringBuffer(weightUnForm).insert(weightUnForm.length()-3, ",").toString();

                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))==','){ weight=weight.substring(0,weight.length()-1);}

                    weight= weight.replace(',','.');
                    double a= Double.parseDouble(weight);
                    weight = Double.toString(a);
                    weight= weight.replace('.',',');

                }

                String warehouse= objDefaultFormat.formatCellValue(row.getCell(14)).trim();
                if(warehouse.equals("15")){warehouse="15скл";}
                Truck truck= new Truck(position,date,invoice,number,weight,consignor,goodClass,warehouse);
                list.add(truck);

                System.out.println(position + " " + date+ " " + invoice+ " " + number+" "+goodClass+" "+weight+" "+consignor+" "+warehouse);
            }

        }
        Iterator<Row> objIteratorSoy = sheetSoy.rowIterator();
        while (objIteratorSoy.hasNext()) {

            Row row = objIteratorSoy.next();
            if (row.getRowNum()>13) {

                String position = objDefaultFormat.formatCellValue(row.getCell(0));
                if(position.equals("")){System.out.println("----------");break;}
                String date = objDefaultFormat.formatCellValue(row.getCell(2));
                if (date.contains("/")){
                    String dateTemp = date.replace("/","." );
                    StringTokenizer st = new StringTokenizer(dateTemp);
                    String arr[]=dateTemp.split("\\.");

                    if(arr[1].length()<2){arr[1]="0"+arr[1];}
                    if(arr[0].length()<2){arr[0]="0"+arr[0];}
                    if(arr[2].length()==2){arr[2]="20"+arr[2];}
                    date=arr[1]+"."+arr[0]+"."+arr[2];}
                String invoice = objDefaultFormat.formatCellValue(row.getCell(3)).trim();
                String number = objDefaultFormat.formatCellValue(row.getCell(4)).trim();

                String consignor =  objDefaultFormat.formatCellValue(row.getCell(12)).trim();

                String goodClass = objDefaultFormat.formatCellValue(row.getCell(13)).trim();

                String weight=null;
                String weightUnForm = objDefaultFormat.formatCellValue(row.getCell(10)).trim();
                if(weightUnForm.equals("")|weightUnForm.equals("0")){ weight="0";}


                if(weightUnForm.length()>=3){
                    weight=new StringBuffer(weightUnForm).insert(weightUnForm.length()-3, ",").toString();

                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}
                    if ((weight.charAt(weight.length()-1))==','){ weight=weight.substring(0,weight.length()-1);}

                    weight= weight.replace(',','.');
                    double a= Double.parseDouble(weight);
                    weight = Double.toString(a);
                    weight= weight.replace('.',',');

                }

                String warehouse= objDefaultFormat.formatCellValue(row.getCell(14)).trim();
                if(warehouse.equals("15")){warehouse="15скл";}
                Truck truck= new Truck(position,date,invoice,number,weight,consignor,goodClass,warehouse);
                list.add(truck);

                System.out.println(position + " " + date+ " " + invoice+ " " + number+" "+goodClass+" "+weight+" "+consignor+" "+warehouse);
            }

        }
        input_document.close();
        return  list;}

    public static List<Truck> getMainSummarry(String fileOut) throws IOException,InvalidFormatException,OpenXML4JException {
        List<Truck> list = new ArrayList<Truck>();
        DataFormatter objDefaultFormat = new DataFormatter();
        FileInputStream input_document = new FileInputStream(new File(fileOut));
        XSSFWorkbook my_xlsx_workbook = new XSSFWorkbook(input_document);
        XSSFSheet sheet = my_xlsx_workbook.getSheet("Сводка");
        Iterator<Row> objIterator = sheet.rowIterator();
        while (objIterator.hasNext()) {

            Row row = objIterator.next();
            if (row.getRowNum()>=2) {

                String position = objDefaultFormat.formatCellValue(row.getCell(0)).trim();
                String goodClass = objDefaultFormat.formatCellValue(row.getCell(6)).trim();
                String number = objDefaultFormat.formatCellValue(row.getCell(3)).trim();
                if(position.equals("")&goodClass.equals("")){System.out.println("----no truck position-");break;}
                String date = objDefaultFormat.formatCellValue(row.getCell(1)).trim();
                if (date.contains("/")){
                    String dateTemp = date.replace("/","." ).trim();
                    StringTokenizer st = new StringTokenizer(dateTemp);
                    String arr[]=dateTemp.split("\\.");
                //    System.out.println(arr[1]+"."+arr[0]+"."+arr[2]);
                    if(arr[1].length()<2){arr[1]="0"+arr[1];}
                    if(arr[0].length()<2){arr[0]="0"+arr[0];}
                    if(arr[2].length()==2){arr[2]="20"+arr[2];}
                    date=arr[1]+"."+arr[0]+"."+arr[2];
                }

                String invoice = objDefaultFormat.formatCellValue(row.getCell(2)).trim();


                if(goodClass==null){
                    goodClass=objDefaultFormat.formatCellValue(row.getCell(7)).trim();
                }

                String weight = objDefaultFormat.formatCellValue(row.getCell(4)).trim();
                if(weight.equals("")|weight.equals("0")|weight.equals("0,00")|weight.equals("0.00")
                        |weight.equals("0.0")|weight.equals("0,0")|weight.equals("0,000")|weight.equals("0.000")){ weight="0";}

                else if(row.getCell(4).getCellType() == Cell.CELL_TYPE_FORMULA) {

                    double tempW = row.getCell(4).getNumericCellValue();
                    weight=Double.toString(tempW);
                   weight= weight.replace('.',',');

                }
                else{

                    weight= weight.replace(',','.');
                   double a= Double.parseDouble(weight);

                    weight = Double.toString(a);
                    weight= weight.replace('.',',');
                    if ((weight.charAt(weight.length()-1))=='0'){ weight=weight.substring(0,weight.length()-1);}

                }

                String consignor = objDefaultFormat.formatCellValue(row.getCell(5)).trim();
                String warehouse= objDefaultFormat.formatCellValue(row.getCell(8)).trim();
                Truck truck= new Truck(position,date,invoice,number,weight,consignor,goodClass,warehouse);
                list.add(truck);

                System.out.println(position + " " + date+ " " + invoice+ " " + number+" "+goodClass+" "+weight+" "+consignor+" "+warehouse);
            }

        }System.out.println("----------");
        input_document.close();
        return  list;}

    public static List<String> getMainSummarryCons(List<Truck> listIn) throws IOException,InvalidFormatException,OpenXML4JException {

           List<String> consNames = new ArrayList<String>();
            List<String> consNamesForComb = new ArrayList<String>();


                for (Truck a:listIn){
                    consNames.add(a.getConsignor());                }

                for (String a:consNames ) {
                    if(!consNamesForComb.contains(a))
                    {
                        consNamesForComb.add(a);
                        System.out.println(a);}
}
        return consNamesForComb; }


    public static List<Truck> getRegister(String fileOut) throws IOException,InvalidFormatException,OpenXML4JException {
        List<Truck> list = new ArrayList<Truck>();
        DataFormatter objDefaultFormat = new DataFormatter();
        FileInputStream input_document = new FileInputStream(new File(fileOut));
        XSSFWorkbook my_xlsx_workbook = new XSSFWorkbook(input_document);
        XSSFSheet sheet = my_xlsx_workbook.getSheet("Горох");
        XSSFSheet sheet2 = my_xlsx_workbook.getSheet("Пшеница");
        XSSFSheet sheet3 = my_xlsx_workbook.getSheet("Ячмень");
        XSSFSheet sheet4 = my_xlsx_workbook.getSheet("Шрот");

        Iterator<Row> objIterator = sheet.rowIterator();
        Iterator<Row> objIterator2 = sheet2.rowIterator();
        Iterator<Row> objIterator3 = sheet3.rowIterator();
        Iterator<Row> objIterator4 = sheet4.rowIterator();

        while (objIterator.hasNext()) {
            Row row = objIterator.next();
            if (row.getRowNum()>=2) {
                if(objDefaultFormat.formatCellValue(row.getCell(4)).indexOf("SUM") != -1){

                if(objIterator.hasNext()){
                        row=objIterator.next();
                    }
                    if(!objIterator.hasNext()){
                        break;
                    }
                }
                if(objDefaultFormat.formatCellValue(row.getCell(0)).trim().equals("")&
                        objDefaultFormat.formatCellValue(row.getCell(4)).trim().equals("")&
                        objDefaultFormat.formatCellValue(row.getCell(1)).trim().equals("")
                        ){
                   System.out.println("----no ---");
                   // row=objIterator.next();
                    break;
                   }

                String position = objDefaultFormat.formatCellValue(row.getCell(0)).trim();
                String weight = objDefaultFormat.formatCellValue(row.getCell(4)).trim();


                String date = objDefaultFormat.formatCellValue(row.getCell(1)).trim();
                if (date.contains("/")){
                    String dateTemp = date.replace("/","." ).trim();
                    StringTokenizer st = new StringTokenizer(dateTemp);
                    String arr[]=dateTemp.split("\\.");
                    //    System.out.println(arr[1]+"."+arr[0]+"."+arr[2]);
                    if(arr[1].length()<2){arr[1]="0"+arr[1];}
                    if(arr[0].length()<2){arr[0]="0"+arr[0];}
                    if(arr[2].length()==2){arr[2]="20"+arr[2];}
                    date=arr[1]+"."+arr[0]+"."+arr[2];
                }

                String invoice = objDefaultFormat.formatCellValue(row.getCell(2)).trim();
                String number = objDefaultFormat.formatCellValue(row.getCell(3)).trim();
                String goodClass = objDefaultFormat.formatCellValue(row.getCell(6)).trim();
                if(goodClass==null){
                    goodClass=objDefaultFormat.formatCellValue(row.getCell(7));
                }

                String consignor = objDefaultFormat.formatCellValue(row.getCell(5));
                String warehouse= objDefaultFormat.formatCellValue(row.getCell(7));
                Truck truck= new Truck(position,date,invoice,number,weight,consignor,goodClass,warehouse);
                list.add(truck);

                System.out.println(position + " " + date+ " " + invoice+ " " + number+" "+goodClass+" "+weight+" "+consignor+" "+warehouse);
            }

        }System.out.println("----------");

        while (objIterator2.hasNext()) {
            Row row = objIterator2.next();
            if (row.getRowNum()>=2) {
                if(objDefaultFormat.formatCellValue(row.getCell(4)).indexOf("SUM") != -1){

                    if(objIterator2.hasNext()){
                        row=objIterator2.next();
                    }
                    if(!objIterator2.hasNext()){
                        break;
                    }
                }
                if(objDefaultFormat.formatCellValue(row.getCell(0)).trim().equals("")&
                        objDefaultFormat.formatCellValue(row.getCell(4)).trim().equals("")){
                    System.out.println("----no ---");
                    row=objIterator2.next();
                }

                String position = objDefaultFormat.formatCellValue(row.getCell(0)).trim();
                String weight = objDefaultFormat.formatCellValue(row.getCell(4)).trim();


                String date = objDefaultFormat.formatCellValue(row.getCell(1)).trim();
                if (date.contains("/")){
                    String dateTemp = date.replace("/","." ).trim();
                    StringTokenizer st = new StringTokenizer(dateTemp);
                    String arr[]=dateTemp.split("\\.");
                    //    System.out.println(arr[1]+"."+arr[0]+"."+arr[2]);
                    if(arr[1].length()<2){arr[1]="0"+arr[1];}
                    if(arr[0].length()<2){arr[0]="0"+arr[0];}
                    if(arr[2].length()==2){arr[2]="20"+arr[2];}
                    date=arr[1]+"."+arr[0]+"."+arr[2];
                }

                String invoice = objDefaultFormat.formatCellValue(row.getCell(2)).trim();
                String number = objDefaultFormat.formatCellValue(row.getCell(3)).trim();
                String goodClass = objDefaultFormat.formatCellValue(row.getCell(6)).trim().toUpperCase();
                if(goodClass==null){
                    goodClass=objDefaultFormat.formatCellValue(row.getCell(7)).toUpperCase();
                }

                String consignor = objDefaultFormat.formatCellValue(row.getCell(5));
                String warehouse= objDefaultFormat.formatCellValue(row.getCell(7));
                Truck truck= new Truck(position,date,invoice,number,weight,consignor,goodClass,warehouse);
                list.add(truck);

                System.out.println(position + " " + date+ " " + invoice+ " " + number+" "+goodClass+" "+weight+" "+consignor+" "+warehouse);
            }

        }System.out.println("----------");
        while (objIterator3.hasNext()) {
            Row row = objIterator3.next();
            if (row.getRowNum()>=2) {
                if(objDefaultFormat.formatCellValue(row.getCell(4)).indexOf("SUM") != -1){

                    if(objIterator3.hasNext()){
                        row=objIterator3.next();
                    }
                    if(!objIterator3.hasNext()){
                        break;
                    }
                }
                if(objDefaultFormat.formatCellValue(row.getCell(0)).trim().equals("")&
                        objDefaultFormat.formatCellValue(row.getCell(4)).trim().equals("")){
                    System.out.println("----no ---");
                    row=objIterator3.next();
                }

                String position = objDefaultFormat.formatCellValue(row.getCell(0)).trim();
                String weight = objDefaultFormat.formatCellValue(row.getCell(4)).trim();
                String date = objDefaultFormat.formatCellValue(row.getCell(1)).trim();
                if (date.contains("/")){
                    String dateTemp = date.replace("/","." ).trim();
                    StringTokenizer st = new StringTokenizer(dateTemp);
                    String arr[]=dateTemp.split("\\.");
                    //    System.out.println(arr[1]+"."+arr[0]+"."+arr[2]);
                    if(arr[1].length()<2){arr[1]="0"+arr[1];}
                    if(arr[0].length()<2){arr[0]="0"+arr[0];}
                    if(arr[2].length()==2){arr[2]="20"+arr[2];}
                    date=arr[1]+"."+arr[0]+"."+arr[2];
                }

                String invoice = objDefaultFormat.formatCellValue(row.getCell(2)).trim();
                String number = objDefaultFormat.formatCellValue(row.getCell(3)).trim();
                String goodClass = objDefaultFormat.formatCellValue(row.getCell(6)).trim();
                if(goodClass==null){
                    goodClass=objDefaultFormat.formatCellValue(row.getCell(7));
                }

                String consignor = objDefaultFormat.formatCellValue(row.getCell(5));
                String warehouse= objDefaultFormat.formatCellValue(row.getCell(7));
                Truck truck= new Truck(position,date,invoice,number,weight,consignor,goodClass,warehouse);
                list.add(truck);

                System.out.println(position + " " + date+ " " + invoice+ " " + number+" "+goodClass+" "+weight+" "+consignor+" "+warehouse);
            }

        }System.out.println("----------");

        while (objIterator4.hasNext()) {
            Row row = objIterator4.next();
            if (row.getRowNum()>=2) {
                if(objDefaultFormat.formatCellValue(row.getCell(4)).indexOf("SUM") != -1){

                    if(objIterator4.hasNext()){
                        row=objIterator4.next();
                    }
                    if(!objIterator4.hasNext()){
                        break;
                    }
                }
                if(objDefaultFormat.formatCellValue(row.getCell(0)).trim().equals("")&
                        objDefaultFormat.formatCellValue(row.getCell(4)).trim().equals("")){
                    System.out.println("----no ---");
                    row=objIterator3.next();
                }

                String position = objDefaultFormat.formatCellValue(row.getCell(0)).trim();
                String weight = objDefaultFormat.formatCellValue(row.getCell(4)).trim();
                String date = objDefaultFormat.formatCellValue(row.getCell(1)).trim();
                if (date.contains("/")){
                    String dateTemp = date.replace("/","." ).trim();
                    StringTokenizer st = new StringTokenizer(dateTemp);
                    String arr[]=dateTemp.split("\\.");
                    //    System.out.println(arr[1]+"."+arr[0]+"."+arr[2]);
                    if(arr[1].length()<2){arr[1]="0"+arr[1];}
                    if(arr[0].length()<2){arr[0]="0"+arr[0];}
                    if(arr[2].length()==2){arr[2]="20"+arr[2];}
                    date=arr[1]+"."+arr[0]+"."+arr[2];
                }

                String invoice = objDefaultFormat.formatCellValue(row.getCell(2)).trim();
                String number = objDefaultFormat.formatCellValue(row.getCell(3)).trim();
                String goodClass = objDefaultFormat.formatCellValue(row.getCell(6)).trim();
                if(goodClass==null){
                    goodClass=objDefaultFormat.formatCellValue(row.getCell(7));
                }

                String consignor = objDefaultFormat.formatCellValue(row.getCell(5));
                String warehouse= objDefaultFormat.formatCellValue(row.getCell(7));
                Truck truck= new Truck(position,date,invoice,number,weight,consignor,goodClass,warehouse);
                list.add(truck);

                System.out.println(position + " " + date+ " " + invoice+ " " + number+" "+goodClass+" "+weight+" "+consignor+" "+warehouse);
            }

        }System.out.println("----------");
        input_document.close();
        return  list;}

    public static List<Truck> compareRandS(List<Truck> reportIn,List<Truck> mainSum){
       List<Truck> result = new ArrayList<Truck>();
        for (Truck truck:reportIn ) {

     //   }
      //  for(int i=0; i<report.size(); i++) {
            boolean contains=false;
               for(int j=0; j<mainSum.size(); j++) {

                //   if (truck.getPosition().equals(mainSum.get(j).getPosition())) {
                       if (truck.getDate().equals(mainSum.get(j).getDate())) {

                           Double reportT=Double.parseDouble(truck.getWeight().replace(',','.'));
                           Double sumT=Double.parseDouble(mainSum.get(j).getWeight().replace(',','.'));

                           if (reportT.equals(sumT)) {
                               if (truck.getInvoice().equals(mainSum.get(j).getInvoice())) {
                                 //  if (truck.getConsignor().equals(mainSum.get(j).getConsignor())) {
                                       if (truck.getGoodClass().toUpperCase().equals(mainSum.get(j).getGoodClass().toUpperCase())) {
                                           if (truck.getWarehouse().equals(mainSum.get(j).getWarehouse())) {
                                               contains = true;

                                           }
                                     }
                               //   }
                              }
                          }
                       }
                  //}


               }
            if (!contains) {
                result.add(truck);
                System.out.println(truck.getPosition()+" "
                        +truck.getDate()+" "
                        +truck.getInvoice()+" "
                        +truck.getNumber()+" "
                        +truck.getGoodClass()+" "
                        +truck.getWeight()+" "
                        +truck.getConsignor() +" "

                         +truck.getWarehouse());
            }
    } return result;
    }

    public static void saveNewSummary(String fileOut, List<Truck> result) throws FileNotFoundException, IOException {
        FileInputStream input_document = new FileInputStream(new File(fileOut));
        XSSFWorkbook workbook = new XSSFWorkbook(input_document);
        XSSFSheet sheet = workbook.getSheet("Сводка");

         int lastRow =sheet.getPhysicalNumberOfRows();

        XSSFCellStyle cellStyle = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontName(HSSFFont.FONT_ARIAL);
        font.setFontHeightInPoints((short)8);
        cellStyle.setFont(font);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);

        XSSFCellStyle cellStyleR = workbook.createCellStyle();
        cellStyleR.setAlignment(CellStyle.ALIGN_RIGHT);
        cellStyleR.setFont(font);
        cellStyleR.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyleR.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyleR.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyleR.setBorderLeft(HSSFCellStyle.BORDER_THIN);

        XSSFCellStyle cellStyleM = workbook.createCellStyle();
        cellStyleM.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyleM.setFont(font);
        cellStyleM.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyleM.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyleM.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyleM.setBorderLeft(HSSFCellStyle.BORDER_THIN);

    for(int i=0; i<result.size(); i++){
        Truck truck=result.get(i);
        Row row=sheet.createRow(lastRow);

        Cell cellPos= row.createCell(0);

        Cell cellDate= row.createCell(1);
        Cell cellInvoice= row.createCell(2);
        Cell cellNumber= row.createCell(3);
        Cell cellGood= row.createCell(6);
        Cell cellGoodPayFor = row.createCell(7);
        Cell cellWeight =row.createCell(4);
        Cell cellCons = row.createCell(5);
        Cell cellWarehouse = row.createCell(8);

        cellPos.setCellStyle(cellStyleR);
        cellDate.setCellStyle(cellStyle);
        cellInvoice.setCellStyle(cellStyle);
        cellNumber.setCellStyle(cellStyle);
        cellGood.setCellStyle(cellStyleM);
        cellGoodPayFor.setCellStyle(cellStyleM);
        cellWeight.setCellStyle(cellStyleR);
        cellWeight.setCellType(Cell.CELL_TYPE_NUMERIC);
        cellCons.setCellStyle(cellStyle);
        cellWarehouse.setCellStyle(cellStyleM);

        cellPos.setCellValue(truck.getPosition());
        cellDate.setCellValue(truck.getDate());
        cellInvoice.setCellValue(truck.getInvoice());
        cellNumber.setCellValue(truck.getNumber());
        cellGood.setCellValue(truck.getGoodClass());
        cellGoodPayFor.setCellValue(truck.getGoodClass());
        cellWeight.setCellValue(Double.parseDouble(truck.getWeight().replace(",",".")));
        cellCons.setCellValue(truck.getConsignor());
        cellWarehouse.setCellValue(truck.getWarehouse());
        ++lastRow;}
        System.out.println(lastRow);
        FileOutputStream output_file =new FileOutputStream(new File(fileOut));
        workbook.write(output_file);
        input_document.close();

}

    public static void saveNewRegister(String fileOut, List<Truck> result,String cons) throws FileNotFoundException, IOException {
        List<Truck> forSave = new ArrayList<>();
        if(cons.equals("ООО\"АСТОРИЯ-А\"") |cons.equals("Астория - А ООО")){
            for (Truck truck:result){
                if(truck.getConsignor().equals("ООО\"АСТОРИЯ-А\"")|truck.getConsignor().equals("Астория - А ООО")){
                    forSave.add(truck);

                }
            }
        }
        if(cons.equals("\"КАСАДА УКРАИНА\"") |cons.equals("Касада Украина ООО")){
            for (Truck truck:result){
                if(truck.getConsignor().equals("\"КАСАДА УКРАИНА\"")|truck.getConsignor().equals("Касада Украина ООО")){
                    forSave.add(truck);

                }
            }
        }
      else{
        for (Truck truck:result) {

            if (truck.getConsignor().equals(cons)) {

                forSave.add(truck);
            }
        }
        }
        {}


        FileInputStream input_document = new FileInputStream(new File(fileOut));
        XSSFWorkbook workbook = new XSSFWorkbook(input_document);
        XSSFSheet sheet =workbook.getSheet("Горох");

        XSSFCellStyle cellStyle = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontName(HSSFFont.FONT_ARIAL);
        font.setFontHeightInPoints((short)8);
        cellStyle.setFont(font);
        cellStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        int lastRowW=0;
        int lastRowP=0;
        int lastRowB=0;
        XSSFCellStyle cellStyleC = workbook.createCellStyle();
        cellStyleC.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyleC.setFont(font);
        cellStyleC.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyleC.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyleC.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyleC.setBorderLeft(HSSFCellStyle.BORDER_THIN);

        XSSFCellStyle cellStyleR = workbook.createCellStyle();
        cellStyleR.setAlignment(CellStyle.ALIGN_RIGHT);
        cellStyleR.setFont(font);
        cellStyleR.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        cellStyleR.setBorderTop(HSSFCellStyle.BORDER_THIN);
        cellStyleR.setBorderRight(HSSFCellStyle.BORDER_THIN);
        cellStyleR.setBorderLeft(HSSFCellStyle.BORDER_THIN);



        for(int i=0; i<forSave.size(); i++){

            Truck truck=forSave.get(i);
            if(truck.getGoodClass().equals("ФУР")|truck.getGoodClass().equals("3КЛ")|truck.getGoodClass().equals("2КЛ")|
                    truck.getGoodClass().equals("фур")|truck.getGoodClass().equals("3кл")|truck.getGoodClass().equals("2кл")
                    |truck.getGoodClass().equals("1-й")|truck.getGoodClass().equals("2-й")|truck.getGoodClass().equals("3-й")
                    |truck.getGoodClass().equals("фураж") |truck.getGoodClass().equals("ФУРАЖ")){
                sheet = workbook.getSheet("Пшеница");
                lastRowW = sheet.getPhysicalNumberOfRows();
                System.out.println(lastRowW);


            Row row=sheet.createRow(lastRowW);

            Cell cellPos= row.createCell(0);
            Cell cellDate= row.createCell(1);
            Cell cellInvoice= row.createCell(2);
            Cell cellNumber= row.createCell(3);
            Cell cellGood= row.createCell(6);
            //Cell cellGoodPayFor = row.createCell(7);
            Cell cellWeight =row.createCell(4);
            Cell cellCons = row.createCell(5);
            Cell cellWarehouse = row.createCell(7);


            cellPos.setCellStyle(cellStyleR);
            cellDate.setCellStyle(cellStyle);
            cellInvoice.setCellStyle(cellStyle);
            cellNumber.setCellStyle(cellStyle);
            cellGood.setCellStyle(cellStyleC);
            cellWeight.setCellStyle(cellStyleR);
            cellCons.setCellStyle(cellStyle);
            cellWarehouse.setCellStyle(cellStyleC);

            cellPos.setCellValue(truck.getPosition());
            cellDate.setCellValue(truck.getDate());
            cellInvoice.setCellValue(truck.getInvoice());
            cellNumber.setCellValue(truck.getNumber());
            cellGood.setCellValue(truck.getGoodClass());
            cellWeight.setCellValue(Double.parseDouble(truck.getWeight().replace(",",".")));
            cellCons.setCellValue(truck.getConsignor());
            cellWarehouse.setCellValue(truck.getWarehouse());
            ++lastRowW;}}

        for(int i=0; i<forSave.size(); i++) {
            Truck truck = forSave.get(i);
            if (truck.getGoodClass().equals("горох")) {
                sheet = workbook.getSheet("Горох");
                lastRowP = sheet.getPhysicalNumberOfRows();
                System.out.println(lastRowP);
                Row row=sheet.createRow(lastRowP);

                Cell cellPos= row.createCell(0);
                Cell cellDate= row.createCell(1);
                Cell cellInvoice= row.createCell(2);
                Cell cellNumber= row.createCell(3);
                Cell cellGood= row.createCell(6);
                //Cell cellGoodPayFor = row.createCell(7);
                Cell cellWeight =row.createCell(4);
                Cell cellCons = row.createCell(5);
                Cell cellWarehouse = row.createCell(7);

                cellPos.setCellStyle(cellStyleR);
                cellDate.setCellStyle(cellStyle);
                cellInvoice.setCellStyle(cellStyle);
                cellNumber.setCellStyle(cellStyle);
                cellGood.setCellStyle(cellStyleC);
                cellWeight.setCellStyle(cellStyleR);
                cellCons.setCellStyle(cellStyle);
                cellWarehouse.setCellStyle(cellStyleC);

                cellPos.setCellValue(truck.getPosition());
                cellDate.setCellValue(truck.getDate());
                cellInvoice.setCellValue(truck.getInvoice());
                cellNumber.setCellValue(truck.getNumber());
                cellGood.setCellValue(truck.getGoodClass());
                cellWeight.setCellValue(Double.parseDouble(truck.getWeight().replace(",",".")));
                cellCons.setCellValue(truck.getConsignor());
                cellWarehouse.setCellValue(truck.getWarehouse());
                ++lastRowP;}

            }
        for(int i=0; i<forSave.size(); i++) {

    Truck truck = forSave.get(i);
        if(truck.getGoodClass().equals("ячмень")){
            sheet = workbook.getSheet("Ячмень");
            lastRowB = sheet.getPhysicalNumberOfRows();
            System.out.println(lastRowB);

        Row row=sheet.createRow(lastRowB);

        Cell cellPos= row.createCell(0);
        Cell cellDate= row.createCell(1);
        Cell cellInvoice= row.createCell(2);
        Cell cellNumber= row.createCell(3);
        Cell cellGood= row.createCell(6);

        Cell cellWeight =row.createCell(4);
        Cell cellCons = row.createCell(5);
        Cell cellWarehouse = row.createCell(7);


            cellPos.setCellStyle(cellStyleR);
            cellDate.setCellStyle(cellStyle);
            cellInvoice.setCellStyle(cellStyle);
            cellNumber.setCellStyle(cellStyle);
            cellGood.setCellStyle(cellStyleC);
            cellWeight.setCellStyle(cellStyleR);
            cellCons.setCellStyle(cellStyle);
            cellWarehouse.setCellStyle(cellStyleC);


        cellPos.setCellValue(truck.getPosition());
        cellDate.setCellValue(truck.getDate());
        cellInvoice.setCellValue(truck.getInvoice());
        cellNumber.setCellValue(truck.getNumber());
        cellGood.setCellValue(truck.getGoodClass());
        cellWeight.setCellValue(Double.parseDouble(truck.getWeight().replace(",",".")));
        cellCons.setCellValue(truck.getConsignor());
        cellWarehouse.setCellValue(truck.getWarehouse());
        ++lastRowB;}

        }


        for(int i=0; i<forSave.size(); i++) {

            Truck truck = forSave.get(i);
            if(truck.getGoodClass().equals("ШРОТ")|truck.getGoodClass().equals("шрот")){
                sheet = workbook.getSheet("Шрот");
                lastRowB = sheet.getPhysicalNumberOfRows();
                System.out.println(lastRowB);

                Row row=sheet.createRow(lastRowB);

                Cell cellPos= row.createCell(0);
                Cell cellDate= row.createCell(1);
                Cell cellInvoice= row.createCell(2);
                Cell cellNumber= row.createCell(3);
                Cell cellGood= row.createCell(6);
                Cell cellWeight =row.createCell(4);
                Cell cellCons = row.createCell(5);
                Cell cellWarehouse = row.createCell(7);

                cellPos.setCellStyle(cellStyleR);
                cellDate.setCellStyle(cellStyle);
                cellInvoice.setCellStyle(cellStyle);
                cellNumber.setCellStyle(cellStyle);
                cellGood.setCellStyle(cellStyleC);
                cellWeight.setCellStyle(cellStyleR);
                cellCons.setCellStyle(cellStyle);
                cellWarehouse.setCellStyle(cellStyleC);

                cellPos.setCellValue(truck.getPosition());
                cellDate.setCellValue(truck.getDate());
                cellInvoice.setCellValue(truck.getInvoice());
                cellNumber.setCellValue(truck.getNumber());
                cellGood.setCellValue(truck.getGoodClass());
                cellWeight.setCellValue(Double.parseDouble(truck.getWeight().replace(",",".")));
                cellCons.setCellValue(truck.getConsignor());
                cellWarehouse.setCellValue(truck.getWarehouse());
                ++lastRowB;}

        }
        for(int i=0; i<forSave.size(); i++) {

            Truck truck = forSave.get(i);
            if(truck.getGoodClass().equals("кукуруза")|truck.getGoodClass().equals("КУКУРУЗА")){
                sheet = workbook.getSheet("Кукуруза");
                lastRowB = sheet.getPhysicalNumberOfRows();
                System.out.println(lastRowB);

                Row row=sheet.createRow(lastRowB);

                Cell cellPos= row.createCell(0);
                Cell cellDate= row.createCell(1);
                Cell cellInvoice= row.createCell(2);
                Cell cellNumber= row.createCell(3);
                Cell cellGood= row.createCell(6);

                Cell cellWeight =row.createCell(4);
                Cell cellCons = row.createCell(5);
                Cell cellWarehouse = row.createCell(7);


                cellPos.setCellStyle(cellStyleR);
                cellDate.setCellStyle(cellStyle);
                cellInvoice.setCellStyle(cellStyle);
                cellNumber.setCellStyle(cellStyle);
                cellGood.setCellStyle(cellStyleC);
                cellWeight.setCellStyle(cellStyleR);
                cellCons.setCellStyle(cellStyle);
                cellWarehouse.setCellStyle(cellStyleC);

                cellPos.setCellValue(truck.getPosition());
                cellDate.setCellValue(truck.getDate());
                cellInvoice.setCellValue(truck.getInvoice());
                cellNumber.setCellValue(truck.getNumber());
                cellGood.setCellValue(truck.getGoodClass());
                cellWeight.setCellValue(Double.parseDouble(truck.getWeight().replace(",",".")));
                cellCons.setCellValue(truck.getConsignor());
                cellWarehouse.setCellValue(truck.getWarehouse());
                ++lastRowB;}

        }
        for(int i=0; i<forSave.size(); i++) {

            Truck truck = forSave.get(i);
            if(truck.getGoodClass().equals("СОЯ")|truck.getGoodClass().equals("соя")){
                sheet = workbook.getSheet("Соя");
                lastRowB = sheet.getPhysicalNumberOfRows();
                System.out.println(lastRowB);

                Row row=sheet.createRow(lastRowB);

                Cell cellPos= row.createCell(0);
                Cell cellDate= row.createCell(1);
                Cell cellInvoice= row.createCell(2);
                Cell cellNumber= row.createCell(3);
                Cell cellGood= row.createCell(6);
                Cell cellWeight =row.createCell(4);
                Cell cellCons = row.createCell(5);
                Cell cellWarehouse = row.createCell(7);

                cellPos.setCellStyle(cellStyleR);
                cellDate.setCellStyle(cellStyle);
                cellInvoice.setCellStyle(cellStyle);
                cellNumber.setCellStyle(cellStyle);
                cellGood.setCellStyle(cellStyleC);
                cellWeight.setCellStyle(cellStyleR);
                cellCons.setCellStyle(cellStyle);
                cellWarehouse.setCellStyle(cellStyleC);

                cellPos.setCellValue(truck.getPosition());
                cellDate.setCellValue(truck.getDate());
                cellInvoice.setCellValue(truck.getInvoice());
                cellNumber.setCellValue(truck.getNumber());
                cellGood.setCellValue(truck.getGoodClass());
                cellWeight.setCellValue(Double.parseDouble(truck.getWeight().replace(",",".")));
                cellCons.setCellValue(truck.getConsignor());
                cellWarehouse.setCellValue(truck.getWarehouse());
                ++lastRowB;}
        }

        FileOutputStream output_file =new FileOutputStream(new File(fileOut));


        workbook.write(output_file);


        input_document.close();
        output_file.close();

    }

}