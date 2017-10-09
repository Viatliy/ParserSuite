
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class Main {


    static final String[] nameCollum = {"Название", "Сайт", "Почта", "Телефон", "Факс", "Адрес", "Представитель"};
    static final String pathCreateFile = "C:\\Users\\d\\Desktop\\test.xls";
    static final String alphavit[] = {" ", "1", "3", "5", "@", "A", "B", "C", "D", "E", "F",
            "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q",
            "R", "S", "T", "U", "V", "W", "Y", "Z"};
    static final String baseURL = "http://www.ic.gc.ca";
    static final String beforeLetterURL = "/app/ccc/sld/cmpny.do?letter=";
    static final String afterLetterURL = "&lang=fre&profileId=1921&naics=621#";


    static void createExel(Sheet sheet) {
        for (int i = 0; i < 20; i++) {
            if (i < 5) {
                sheet.setColumnWidth(i, 6000);
            } else if (i == 5) {
                sheet.setColumnWidth(i, 10000);
            } else if (i > 5) {
                sheet.setColumnWidth(i, 15000);
            }
        }
        Row row = sheet.createRow(0);
        int cnt = 1;
        for (int i = 0; i < nameCollum.length + 9; i++) {
            Cell cell1 = row.createCell(i);
            if (i < nameCollum.length - 1) {
                cell1.setCellValue(nameCollum[i]);
            } else {
                cell1.setCellValue(nameCollum[nameCollum.length - 1] + " " + cnt);
                cnt++;
            }
        }
    }

    static String getInformation(Document doc, String key, String value) {
        Elements tempInfo = doc.getElementsByAttributeValue(key, value);
        String tempInf = "";
        if (value.equals("mrgn-bttm-0")) {
            Element tempAdr = tempInfo.get(0);
            tempInf = tempAdr.html().replaceAll("<br>", "\n");
        } else {
            for (int j = 0; j < tempInfo.size(); j++) {
                Element temp = tempInfo.get(j);
                tempInf += temp.text() + "\n";
            }
        }
        return tempInf;

    }

    static void writeInfo(Row row1, String name, String webAdr, String eM, String tel, String telT, String fax, String adr) {
        Cell cell11 = row1.createCell(0);
        cell11.setCellValue(name);
        Cell cell21 = row1.createCell(1);
        cell21.setCellValue(webAdr);
        Cell cell31 = row1.createCell(2);
        cell31.setCellValue(eM);
        Cell cell41 = row1.createCell(3);
        cell41.setCellValue(tel + "\n" + telT);
        Cell cell51 = row1.createCell(4);
        cell51.setCellValue(fax);
        Cell cell61 = row1.createCell(5);
        cell61.setCellValue(adr);
    }

    static void getAndWriteRepresentatives(Document doc1, Row row1, String key, String value) {
        ArrayList<String> allInfoPerson = new ArrayList<>();
        Elements informationPerson = doc1.getElementsByAttributeValue(key, value);
        Elements tempInfoPerson = informationPerson.get(0).children();
        int k = 0, с = 0;
        for (int j = 0; j < tempInfoPerson.size(); j = k) {
            Element tempEl = tempInfoPerson.get(j);
            String className = tempEl.child(0).className();
            String namePerson = tempEl.child(0).text();
            k++;
            String tempInfo = "";
            tempInfo += namePerson + "\n";
            while ("col-md-3".equals(className)) {
                if (k == tempInfoPerson.size()) {
                    break;
                } else {
                    Element tempEl1 = tempInfoPerson.get(k);
                    if ("col-md-5".equals(tempEl1.child(0).className())) {
                        for (int l = 0; l < tempInfoPerson.get(k).children().size(); l++) {
                            if (l % 2 == 0) {
                                tempInfo += tempEl1.child(l).text() + " ";
                            } else {
                                tempInfo += tempEl1.child(l).text() + "\n";
                            }
                        }
                    } else if ("col-md-3".equals(tempEl1.child(0).className())) {
                        break;
                    }
                    k++;
                }
            }
            allInfoPerson.add(tempInfo);
            Cell cell71 = row1.createCell(6 + с);
            cell71.setCellValue(tempInfo);
            с++;
        }
    }


    public static void main(String[] args) throws IOException {
        FileOutputStream file = new FileOutputStream(pathCreateFile);
        String str_url = "";
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("Company");
        createExel(sheet);

        int countCompany = 1;
        for (int y = 0; y < alphavit.length; y++) {
            System.out.println(y);
            if (alphavit[y].equals(" ")) {
                str_url = baseURL + beforeLetterURL + "+" + afterLetterURL;
            } else {
                str_url = baseURL + beforeLetterURL + alphavit[y] + afterLetterURL + alphavit[y];
            }
            Document doc = Jsoup.connect(str_url).get();
            Elements listCompany = doc.getElementsByAttributeValue("class", "mrgn-bttm-sm");
            for (int i = 0; i < listCompany.size(); i++) {
                Element temp = listCompany.get(i);
                String href = temp.child(0).attr("href");
                Document doc1 = Jsoup.connect(baseURL + href).get();
                Row row1 = sheet.createRow(countCompany);
                writeInfo(row1, getInformation(doc1, "class", "col-sm-9"),
                        getInformation(doc1, "title", "Adresse du site web"),
                        getInformation(doc1, "title", "Adresse courriel"),
                        getInformation(doc1, "title", "Téléphone"),
                        getInformation(doc1, "title", "Téléphone sans frais"),
                        getInformation(doc1, "title", "Téléphone sans frais"),
                        getInformation(doc1, "class", "mrgn-bttm-0"));
                getAndWriteRepresentatives(doc1, row1, "class", "container-fluid");
                countCompany++;
            }
        }
        wb.write(file);
    }
}