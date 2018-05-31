import com.sun.corba.se.spi.orbutil.threadpool.Work;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.formula.ConditionalFormattingEvaluator;
import org.apache.poi.ss.formula.EvaluationConditionalFormatRule;
import org.apache.poi.ss.formula.WorkbookEvaluatorProvider;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold.RangeType;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting.IconSet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.omg.PortableInterceptor.INACTIVE;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class Main {
    private static Workbook workbook = new HSSFWorkbook();
    private static CellStyle exam1Style = workbook.createCellStyle();
    private static CellStyle exam2Style = workbook.createCellStyle();
    private static CellStyle distrStyle = workbook.createCellStyle();
    private static CellStyle removedStyle = workbook.createCellStyle();
    private static CellStyle blackStyle = workbook.createCellStyle();
    private static HashMap<String,Integer> examene1 = new HashMap<>();
    private static HashMap<String,Integer> examene2 = new HashMap<>();
    private static HashMap<String,Integer> distribuite = new HashMap<>();
    private static HashMap<String, Integer> limite = new HashMap<>();
    private static Sheet mySheet;
    private static int i=0;
    private static int j=0;
    public static void initWorkbook(Workbook workbook){
        exam1Style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index);
        exam1Style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        exam2Style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.index);
        exam2Style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        distrStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.index);
        distrStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        removedStyle.setFillForegroundColor(IndexedColors.RED.index);
        removedStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        blackStyle.setFillForegroundColor(IndexedColors.BLACK.index);
        blackStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        i=0;
        j=0;
        Row myRow = mySheet.createRow(i);
        myRow.createCell(j++).setCellValue("Nr.Crt.");
        myRow.createCell(j++).setCellValue("Nume");
        myRow.createCell(j++).setCellValue("Grupa");
        myRow.createCell(j++).setCellValue("Media");
        for(int k=1;k<=9;k++){
            myRow.createCell(j++).setCellValue(""+k);
        }
        myRow.createCell(j++).setCellStyle(blackStyle);
        myRow.createCell(j++).setCellValue("IEP");
        myRow.createCell(j++).setCellValue("TD");
        myRow.createCell(j++).setCellValue("MS");
        myRow.createCell(j++).setCellValue("VVS");
        myRow.createCell(j++).setCellValue("TPAC");
        myRow.createCell(j++).setCellValue("SMA");
        myRow.createCell(j++).setCellValue("APND");
        myRow.createCell(j++).setCellValue("AIAW");
        myRow.createCell(j++).setCellValue("PBD");
        myRow.createCell(j++).setCellStyle(blackStyle);
        for(int k=1;k<=5;k++){
            myRow.createCell(j++).setCellValue(""+k);
        }
        myRow.createCell(j++).setCellStyle(blackStyle);
        myRow.createCell(j++).setCellValue("SSC");
        myRow.createCell(j++).setCellValue("CRF");
        myRow.createCell(j++).setCellValue("CHS");
        myRow.createCell(j++).setCellValue("SIPAC");
        myRow.createCell(j++).setCellValue("PCBE");
        myRow.createCell(j++).setCellStyle(blackStyle);
        for(int k=1;k<=10;k++){
            myRow.createCell(j++).setCellValue(""+k);
        }
        myRow.createCell(j++).setCellStyle(blackStyle);
        myRow.createCell(j++).setCellValue("CR");
        myRow.createCell(j++).setCellValue("FSC");
        myRow.createCell(j++).setCellValue("SE");
        myRow.createCell(j++).setCellValue("MPS");
        myRow.createCell(j++).setCellValue("EPSC");
        myRow.createCell(j++).setCellValue("PD");
        myRow.createCell(j++).setCellValue("CES");
        myRow.createCell(j++).setCellValue("SM");
        myRow.createCell(j++).setCellValue("LFA");
        myRow.createCell(j++).setCellValue("STD");
        myRow.createCell(j++).setCellStyle(blackStyle);
        i++;
        /*limite.put("IEP",93);
        limite.put("TD",78);
        limite.put("MS",77);
        limite.put("VVS",62);
        limite.put("TPAC",93);
        limite.put("SMA",93);
        limite.put("APND",78);
        limite.put("AIAW",77);
        limite.put("PBD",93);
        limite.put("SSC",46);
        limite.put("CRF",93);
        limite.put("CHS",93);
        limite.put("SIPAC",62);
        limite.put("PCBE",78);
        limite.put("CR",62);
        limite.put("FSC",78);
        limite.put("SE",62);
        limite.put("MPS",77);
        limite.put("EPSC",78);
        limite.put("PD",77);
        limite.put("CES",78);
        limite.put("SM",77);
        limite.put("LFA",78);
        limite.put("STD",77);
        limite.put("IEP"  ,2);
        limite.put("TD"   ,2);
        limite.put("MS"   ,2);
        limite.put("VVS"  ,2);
        limite.put("TPAC" ,2);
        limite.put("SMA"  ,2);
        limite.put("APND" ,2);
        limite.put("AIAW" ,2);
        limite.put("PBD"  ,2);
        limite.put("SSC"  ,3);
        limite.put("CRF"  ,3);
        limite.put("CHS"  ,3);
        limite.put("SIPAC",3);
        limite.put("PCBE" ,3);
        limite.put("CR"   ,2);
        limite.put("FSC"  ,2);
        limite.put("SE"   ,2);
        limite.put("MPS"  ,2);
        limite.put("EPSC" ,2);
        limite.put("PD"   ,2);
        limite.put("CES"  ,2);
        limite.put("SM"   ,2);
        limite.put("LFA"  ,2);
        limite.put("STD"  ,2);*/
    }
    private static ArrayList<StudOptionale> solStud = new ArrayList<>();
    public static void scrieExcel(ArrayList<StudOptionale> students){
        for(StudOptionale student: students) {
            j=0;
            Row myRow = mySheet.createRow(i);
            myRow.createCell(j++).setCellValue(i++); //Nr.Crt.
            myRow.createCell(j++).setCellValue(student.getName()); //Nume
            myRow.createCell(j++).setCellValue(student.getGroup()); //Grupa
            myRow.createCell(j++).setCellValue(student.getMedia()); //Media
            for (String optionala : student.getEx1Order()) {
                if (student.isAles(optionala)) {
                    Cell myCell = myRow.createCell(j++);
                    myCell.setCellValue(optionala);
                    myCell.setCellStyle(exam1Style);
                } else if (examene1.get(optionala) == limite.get(optionala)) {
                    Cell myCell = myRow.createCell(j++);
                    myCell.setCellValue(optionala);
                    myCell.setCellStyle(removedStyle);
                } else {
                    myRow.createCell(j++).setCellValue(optionala);
                }
            }
            myRow.createCell(j++).setCellStyle(blackStyle);
            String[] auxArray = {"IEP", "TD", "MS", "VVS", "TPAC", "SMA", "APND", "AIAW", "PBD"};
            for (String aux : auxArray) {
                if (examene1.get(aux) != limite.get(aux)) {
                    myRow.createCell(j++).setCellValue(aux + " " + (limite.get(aux) - examene1.get(aux)));
                } else {
                    Cell myCell = myRow.createCell(j++);
                    myCell.setCellValue("[MAX]");
                    myCell.setCellStyle(removedStyle);
                }
            }
            myRow.createCell(j++).setCellStyle(blackStyle);
            for (String optionala : student.getDistrOrder()) {
                if (student.isAles(optionala)) {
                    Cell myCell = myRow.createCell(j++);
                    myCell.setCellValue(optionala);
                    myCell.setCellStyle(distrStyle);
                } else if (examene1.get(optionala) == limite.get(optionala)) {
                    Cell myCell = myRow.createCell(j++);
                    myCell.setCellValue(optionala);
                    myCell.setCellStyle(removedStyle);
                } else {
                    myRow.createCell(j++).setCellValue(optionala);
                }
            }
            myRow.createCell(j++).setCellStyle(blackStyle);
            String[] auxArray1 = {"SSC", "CRF", "CHS", "SIPAC", "PCBE"};
            for (String aux : auxArray1) {
                if (distribuite.get(aux) != limite.get(aux)) {
                    myRow.createCell(j++).setCellValue(aux + " " + (limite.get(aux) - distribuite.get(aux)));
                } else {
                    Cell myCell = myRow.createCell(j++);
                    myCell.setCellValue("[MAX]");
                    myCell.setCellStyle(removedStyle);
                }
            }
            myRow.createCell(j++).setCellStyle(blackStyle);
            for (String optionala : student.getEx2Order()) {
                if (student.isAles(optionala)) {
                    Cell myCell = myRow.createCell(j++);
                    myCell.setCellValue(optionala);
                    myCell.setCellStyle(exam2Style);
                } else if (examene1.get(optionala) == limite.get(optionala)) {
                    Cell myCell = myRow.createCell(j++);
                    myCell.setCellValue(optionala);
                    myCell.setCellStyle(removedStyle);
                } else {
                    myRow.createCell(j++).setCellValue(optionala);
                }
            }
            myRow.createCell(j++).setCellStyle(blackStyle);
            String[] auxArray2 = {"CR", "FSC", "SE", "MPS", "EPSC", "PD", "CES", "SM", "LFA", "STD"};
            for (String aux : auxArray2) {
                if (examene2.get(aux) != limite.get(aux)) {
                    myRow.createCell(j++).setCellValue(aux + " " + (limite.get(aux) - examene2.get(aux)));
                } else {
                    Cell myCell = myRow.createCell(j++);
                    myCell.setCellValue("[MAX]");
                    myCell.setCellStyle(removedStyle);
                }
            }
            myRow.createCell(j++).setCellStyle(blackStyle);
        }
        System.out.println("\n\n");
        String fileWrite = "test_elective_results.xls";
        try {
            FileOutputStream out = new FileOutputStream(fileWrite);
            workbook.write(out);
            out.close();
            workbook.close();
        } catch (Exception e){
            System.out.println(e.getMessage());
            e.printStackTrace();
        }

    }
    public static boolean alegeri(StudOptionale student) throws Exception{
        int lim;
        if(student.getEx1Sterse() || student.getEx2Sterse() || student.getDistrSterse()){
            System.out.println("Stud: "+student.getName()+" no more to generate");
            return false;
        }
        System.out.println(student.getName());
        System.out.println("Ex1Alese: "+student.getEx1Alese()+" Ex2Alese: "+student.getEx2Alese()+ " DistrAlese: "+student.getDistrAlese());
        System.out.print(student.getName()+" Ex1: ");
        for(String optionala: student.getEx1Order()) {
            System.out.print(" Optionala: "+optionala+" aleasa: "+student.isAles(optionala)+
                    " stearsa: "+student.isSters(optionala)+" 4alese: "+student.getEx1Alese());
            if (!student.isAles(optionala) && !student.isSters(optionala) && student.getEx1Alese() != 4) {
                lim = examene1.get(optionala);
                if (lim < limite.get(optionala)) {
                    student.alegeOpt(optionala);
                    examene1.replace(optionala, lim + 1);
                    System.out.print(" "+(limite.get(optionala)-examene1.get(optionala))+"\n");
                } else {
                    System.out.print(" MAX \n");
                }
            } else {
                System.out.println();
            }
        }
        System.out.print(" Distr: ");
        for (String optionala : student.getDistrOrder()) {
            System.out.print(" Optionala: "+optionala+" aleasa: "+student.isAles(optionala)+
                    " stearsa: "+student.isSters(optionala)+" 2alese: "+student.getDistrAlese());
            if (!student.isAles(optionala) && !student.isSters(optionala) && student.getDistrAlese() != 2) {
                lim = distribuite.get(optionala);
                if (lim < limite.get(optionala)) {
                    student.alegeOpt(optionala);
                    distribuite.replace(optionala, lim + 1);
                    System.out.print(" "+(limite.get(optionala)-distribuite.get(optionala))+"\n");
                } else {
                    System.out.print(" MAX \n");
                }
            } else {
                System.out.println();
            }
        }
        System.out.print(" Ex2: ");
        for (String optionala : student.getEx2Order()) {
            System.out.print(" Optionala: "+optionala+" aleasa: "+student.isAles(optionala)+
                    " stearsa: "+student.isSters(optionala)+" 4alese: "+student.getEx2Alese());
            if (!student.isAles(optionala) && !student.isSters(optionala) && student.getEx2Alese() != 4) {
                lim = examene2.get(optionala);
                if (lim < limite.get(optionala)) {
                    // System.out.println("Stud: "+student.getName()+" S-a ales "+optionala+"isAles: "+
                    //        student.isAles(optionala)+" !isSters: "+!student.isSters(optionala));
                    student.alegeOpt(optionala);
                    examene2.replace(optionala, lim + 1);
                    System.out.print(" "+(limite.get(optionala)-examene2.get(optionala))+"\n");
                } else {
                    System.out.print(" MAX \n");
                }
            } else {
                System.out.println();
            }
        }
        System.out.println(" \n");
        return true;
    }
    public static void generateElectives(String file) {
        try {
            POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file));
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            HSSFSheet sheet = wb.getSheetAt(0);
            HSSFRow row;
            HSSFCell cell;

            int rows; // No of rows
            rows = sheet.getPhysicalNumberOfRows();

            int cols = 0; // No of columns
            int tmp = 0;

            // This trick ensures that we get the data properly even if it doesn't start from first few rows
            for (int i = 0; i < 10 || i < rows; i++) {
                row = sheet.getRow(i);
                if (row != null) {
                    tmp = sheet.getRow(i).getPhysicalNumberOfCells();
                    if (tmp > cols) cols = tmp;
                }
            }
            Students students = new Students();
            ArrayList<String> optionale = new ArrayList<String>();
            for (int r = 0; r < rows; r++) {
                row = sheet.getRow(r);
                if (row != null) {
                    /**CREATING STUDENT */
                    Student student = null;
                    if (r != 0) {
                        student = new Student();
                    }
                    for (int c = 0; c < cols; c++) {
                        cell = row.getCell((short) c);
                        if (cell != null) {
                            // Your code here
                            if (r != 0) {
                                if (c == 0) {
                                    /**STUDENT NAME */
                                    student.setName(cell.getStringCellValue());
                                } else if (c == 1) {
                                    /**STUDENT GROUP */
                                    student.setGroup(cell.getStringCellValue());
                                } else if (c == 2) {
                                    /**STUDENT MEDIA */
                                    student.setMedia(cell.getNumericCellValue());
                                } else if (cell.getStringCellValue() != "") {
                                    student.addOption(cell.getStringCellValue());
                                }
                            } else if (c > 2) {
                                optionale.add(cell.getStringCellValue());
                            }
                        }
                    }
                    //System.out.println(student.toString());
                    if (student != null) {
                        students.addStudent(student);
                    }
                }
            }
            ArrayList<Elective> electives = new ArrayList<Elective>();
            Collections.sort(optionale, optionaleAscending);
            for (String opt : optionale) {
                Elective e = new Elective(opt);
                e.setMaxStudents(200);
                for (Student s : students.getStudentList()) {
                    for (String option : s.getOptions()) {
                        if (option.equals(opt)) {
                            //System.out.println("Student " + s.getName() + " was added at elective " + e.getElectiveName());
                            e.addStudent(s);
                        }
                    }
                }
                electives.add(e);
            }
            //System.out.println(electives);
            Workbook wbWrite;
            wbWrite = new HSSFWorkbook();
            String fileWrite = "electives.xls";
            for (Elective e : electives) {
                //ArrayList<Student> std = e.getStudents();
                //System.out.println(std);
                //for (Student s : e.getStudents()) {
                //    System.out.println(s.toString());
                //}
                Sheet sheetWrite = wbWrite.createSheet(e.getElectiveName());
                int i = 1;
                Row r = sheetWrite.createRow(0);
                r.createCell(0).setCellValue("Nr.Crt.");
                r.createCell(1).setCellValue("Nume");
                r.createCell(2).setCellValue("Grupa");
                r.createCell(3).setCellValue("Media");
                for (Student s : e.getStudents()) {
                    int j = 0;
                    //System.out.println(s.getName()+" "+s.getGroup()+ " "+s.getMedia());
                    Row rnew = sheetWrite.createRow(i);
                    rnew.createCell(j++).setCellValue(i);
                    rnew.createCell(j++).setCellValue(s.getName());
                    rnew.createCell(j++).setCellValue(s.getGroup());
                    rnew.createCell(j++).setCellValue(s.getMedia());
                    i++;
                }
            }
            FileOutputStream out = new FileOutputStream(fileWrite);
            wbWrite.write(out);
            out.close();
            wbWrite.close();


            //System.out.println(optionale.toString());
            //System.out.println("\n\nSTUDENTS: "+students.toString());
            //students.sortNameAscending();
            //students.sortMediaDescending();
            //System.out.println(students.toString());
            //students.sortNameDescending();
            //students.sortMediaAscending();
            //System.out.println(students.toString());
        } catch (Exception ioe) {
            System.out.println(ioe.getMessage());
            ioe.printStackTrace();
        }
    }
    public static void repartizeaza(String file, int []opt, int optMaxNumber){
        try {
            POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file));
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            HSSFSheet sheet = wb.getSheetAt(0);
            HSSFRow row;
            HSSFCell cell;

            int rows; // No of rows
            rows = sheet.getPhysicalNumberOfRows();

            int cols = 0; // No of columns
            int tmp = 0;

            // This trick ensures that we get the data properly even if it doesn't start from first few rows
            for (int i = 0; i < 10 || i < rows; i++) {
                row = sheet.getRow(i);
                if (row != null) {
                    tmp = sheet.getRow(i).getPhysicalNumberOfCells();
                    if (tmp > cols) cols = tmp;
                }
            }
            /**FORMAT: Nr.Crt. | NUME | GRUPA | MEDIA | OPT1 | OPT2 | OPT3 | OPT4*/
            Students students = new Students();
            int []opCount = new int[opt.length];
            for(int opc:opCount){
                opc = 0;
            }
            //for(int opc:opCount){
            //    System.out.println(opc);
            //}

            int aux = cols+1;
            for (int r = 0; r < rows; r++) {
                row = sheet.getRow(r);
                if (row != null) {
                    /**CREATING STUDENT */
                    Student student = null;
                    if (r != 0) {
                        student = new Student();
                    } else {
                        row.createCell(cols).setCellValue("Repartizare");
                        //System.out.println("columns: "+cols);
                    }
                    for (int c = 0; c < cols; c++) {
                        //System.out.println("Columns: "+cols);
                        cell = row.getCell((short) c);
                        if (cell != null) {
                            // Your code here
                            if (r != 0) {
                                if (c == 1) {
                                    /**STUDENT NAME */
                                    //System.out.println("aici");
                                    //System.out.println(cell.getStringCellValue());
                                    student.setName(cell.getStringCellValue());
                                } else if (c == 2) {
                                    /**STUDENT GROUP */
                                    //System.out.println("acolo");
                                    //System.out.println(cell.getStringCellValue());
                                    student.setGroup(cell.getStringCellValue());
                                } else if (c == 3) {
                                    /**STUDENT MEDIA */
                                    //System.out.println(cell.getNumericCellValue());
                                    student.setMedia(cell.getNumericCellValue());
                                } //else if (cell.getStringCellValue() != "") {
                                //student.addOption(cell.getStringCellValue());
                                //}
                                else if (c>3){
                                    if(!cell.getStringCellValue().equals("")) {
                                        //System.out.println("Adaug optiune "+cell.getStringCellValue());
                                        //System.out.println(cell.getStringCellValue());
                                        student.addOption(cell.getStringCellValue());
                                    }
                                }
                            }
                        }
                    }
                    //System.out.println(student.toString());
                    if (student != null) {
                        students.addStudent(student);
                        //TODO add here code
                        //System.out.println("Student name: "+ student.getName()+" Options: ");
                        for (String option : student.getOptions()) {
                            //System.out.print(option+" ");
                            boolean toBreak = false;
                            for(int optiune:opt) {
                                CharSequence grupa = "Grupa "+optiune;
                                //System.out.println(grupa);
                                if (option.contains(grupa)) {
                                    //System.out.println(option + "GRUPA 1");
                                    if (opCount[optiune-1] < optMaxNumber) {
                                        row.createCell(cols).setCellValue(option);
                                        opCount[optiune-1]++;
                                        row.createCell(aux).setCellValue("Gr"+optiune+":" + opCount[optiune-1]);
                                        toBreak=true;
                                        break;
                                    }
                                }
                            }
                            if(toBreak){
                                break;
                            }
                        }
                        //System.out.println();
                    }
                }
            }
            for(int optiune: opt){
                String grupa = "Gr"+optiune+": "+opCount[optiune-1];
                System.out.println(grupa);
            }
            String fileWrite = file;
            FileOutputStream out = new FileOutputStream(fileWrite);
            wb.write(out);
            out.close();
            wb.close();
            //System.out.println(students.toString());
        } catch (Exception ioe) {
            System.out.println(ioe.getMessage());
            ioe.printStackTrace();
        }
    }
    public static boolean checkOpt(String[] exOrd,HashMap<String,Integer> limite,HashMap<String,Integer> examene,int exMin,
                                   double nrStud){
        int nrExamene=examene.size();
        double maxStud = (185-nrStud)*exMin;
        for(String o:exOrd){
            if(examene.get(o)==limite.get(o)) {
                //System.out.print(" [" + o + "=MAX!]");
                nrExamene--;
            } else {
                maxStud-=(limite.get(o)-examene.get(o));
                //System.out.print(" [" + o + "=" + (limite.get(o) - examene.get(o)) + "]");
            }
        }
        //if(maxStud>0) {
        //    System.out.println("Not good maxStud");
        //    return false;
        //}
        if(nrExamene<exMin){
            //System.out.println("FALSE nrExamene: "+nrExamene+" vs exMin: "+exMin);
            return false;
        } else {
            //System.out.println("TRUE nrExamene: "+nrExamene+" vs exMin: "+exMin);
            return true;
        }
        //System.out.println();
    }
    public static void afiseazaOptionaleMax(String[] exOrd,HashMap<String,Integer> limite,HashMap<String,Integer> examene){
        for(String o:exOrd){
            if(examene.get(o)==limite.get(o)) {
                System.out.print(" [" + o + "=MAX!]");
            } else {
                System.out.print(" [" + o + "=" + (limite.get(o)-examene.get(o))+"]");
            }
        }
        System.out.println();
    }
    static ArrayList<StudOptionale> students = new ArrayList<>();
    static ArrayList<String> e1 = new ArrayList<>();
    static ArrayList<String> e2 = new ArrayList<>();
    static ArrayList<String> d = new ArrayList<>();
    static final double stuPeGrupa = 15.5;
    static final int maxGrupError = 2;
    public static boolean genereazaDistribuite(){
        j=0;
        i=0;
        Row myRow = mySheet.createRow(i);
        myRow.createCell(j++).setCellValue("Nr.Crt.");
        myRow.createCell(j++).setCellValue("Nume");
        myRow.createCell(j++).setCellValue("Grupa");
        myRow.createCell(j++).setCellValue("Media");
        for(int k=1;k<=5;k++){
            myRow.createCell(j++).setCellValue(""+k);
        }
        myRow.createCell(j++).setCellStyle(blackStyle);
        myRow.createCell(j++).setCellValue("SSC "+limite.get("SSC")+"-Gr. "+(int)(limite.get("SSC")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("CRF "+limite.get("CRF")+"-Gr. "+(int)(limite.get("CRF")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("CHS "+limite.get("CHS")+"-Gr. "+(int)(limite.get("CHS")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("SIPAC "+limite.get("SIPAC")+"-Gr. "+(int)(limite.get("SIPAC")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("PCBE "+limite.get("PCBE")+"-Gr. "+(int)(limite.get("PCBE")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellStyle(blackStyle);
        i++;
        for(String materie:d){
            distribuite.replace(materie,0);
        }
        try {
            for (StudOptionale student : students) {
                j = 0;
                myRow = mySheet.createRow(i);
                myRow.createCell(j++).setCellValue(i++); //Nr.Crt.
                myRow.createCell(j++).setCellValue(student.getName()); //Nume
                myRow.createCell(j++).setCellValue(student.getGroup()); //Grupa
                myRow.createCell(j++).setCellValue(student.getMedia()); //Media
                String[] distrOrdine = student.getOrderedElectives(distribuite, d);
                int distrAlese = 0;
                for (String materie : distrOrdine) {
                    int lim = distribuite.get(materie);
                    if (distrAlese < 2) {
                        if (lim < limite.get(materie)) {
                            distribuite.replace(materie, lim + 1);
                            Cell myCell = myRow.createCell(j++);
                            myCell.setCellValue(materie);
                            myCell.setCellStyle(distrStyle);
                            distrAlese++;
                        } else {
                            Cell myCell = myRow.createCell(j++);
                            myCell.setCellValue(materie);
                            myCell.setCellStyle(removedStyle);
                        }
                    } else {
                        myRow.createCell(j++).setCellValue(materie);
                    }
                }
                if(distrAlese!=2){
                    return false;
                }
                myRow.createCell(j++).setCellStyle(blackStyle);
                String[] auxArray = {"SSC", "CRF", "CHS", "SIPAC", "PCBE"};
                for (String aux : auxArray) {
                    if (!distribuite.get(aux).equals(limite.get(aux))) {
                        myRow.createCell(j++).setCellValue(aux + " " + (limite.get(aux) - distribuite.get(aux)));
                    } else {
                        Cell myCell = myRow.createCell(j++);
                        myCell.setCellValue("[MAX]");
                        myCell.setCellStyle(removedStyle);
                    }
                }
                myRow.createCell(j++).setCellStyle(blackStyle);
            }
            String fileWrite = "test_elective_results.xls";
            FileOutputStream out = new FileOutputStream(fileWrite);
            workbook.write(out);
            out.close();
        } catch (Exception e) {
            System.out.println(e.getMessage());
            e.printStackTrace();
        }
        return true;
    }
    public static boolean genereazaExamene1(){
        i=0;
        j=0;
        Row myRow = mySheet.createRow(i);
        myRow.createCell(j++).setCellValue("Nr.Crt.");
        myRow.createCell(j++).setCellValue("Nume");
        myRow.createCell(j++).setCellValue("Grupa");
        myRow.createCell(j++).setCellValue("Media");
        for(int k=1;k<=9;k++){
            myRow.createCell(j++).setCellValue(""+k);
        }
        myRow.createCell(j++).setCellStyle(blackStyle);
        myRow.createCell(j++).setCellValue("IEP "+limite.get("IEP")+"-Gr. "+(int)(limite.get("IEP")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("TD "+limite.get("TD")+"-Gr. "+(int)(limite.get("TD")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("MS "+limite.get("MS")+"-Gr. "+(int)(limite.get("MS")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("VVS "+limite.get("VVS")+"-Gr. "+(int)(limite.get("VVS")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("TPAC "+limite.get("TPAC")+"-Gr. "+(int)(limite.get("TPAC")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("SMA "+limite.get("SMA")+"-Gr. "+(int)(limite.get("SMA")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("APND "+limite.get("APND")+"-Gr. "+(int)(limite.get("APND")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("AIAW "+limite.get("AIAW")+"-Gr. "+(int)(limite.get("AIAW")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("PBD "+limite.get("PBD")+"-Gr. "+(int)(limite.get("PBD")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellStyle(blackStyle);
        i++;
        for(String materie:e1){
            examene1.replace(materie,0);
        }
        //System.out.println("\n\n");
        //for(String materie:e1){
        //    System.out.print(examene1.get(materie)+" ");
        //}
        //System.out.println();
        try {
            for (StudOptionale student : students) {
                j = 0;
                myRow = mySheet.createRow(i);
                myRow.createCell(j++).setCellValue(i++); //Nr.Crt.
                myRow.createCell(j++).setCellValue(student.getName()); //Nume
                myRow.createCell(j++).setCellValue(student.getGroup()); //Grupa
                myRow.createCell(j++).setCellValue(student.getMedia()); //Media
                String[] ex1Ordine = student.getOrderedElectives(examene1, e1);
                int ex1Alese = 0;
                for (String materie : ex1Ordine) {
                    int lim = examene1.get(materie);
                    if (ex1Alese < 4) {
                        if (lim < limite.get(materie)) {
                            examene1.replace(materie, lim + 1);
                            Cell myCell = myRow.createCell(j++);
                            myCell.setCellValue(materie);
                            myCell.setCellStyle(exam1Style);
                            ex1Alese++;
                        } else {
                            Cell myCell = myRow.createCell(j++);
                            myCell.setCellValue(materie);
                            myCell.setCellStyle(removedStyle);
                        }
                    } else {
                        myRow.createCell(j++).setCellValue(materie);
                    }
                }
                //for(String materie:e1){
                //    System.out.print((limite.get(materie)-examene1.get(materie))+" ");
                //}
                //System.out.println();
                if(ex1Alese!=4){
                    //System.out.println(student.getName()+" "+ex1Alese);
                    return false;
                }
                myRow.createCell(j++).setCellStyle(blackStyle);
                String[] auxArray = {"IEP", "TD", "MS", "VVS", "TPAC", "SMA", "APND", "AIAW", "PBD"};
                for (String aux : auxArray) {
                    if (!examene1.get(aux).equals(limite.get(aux))) {
                        myRow.createCell(j++).setCellValue(aux + " " + (limite.get(aux) - examene1.get(aux)));
                    } else {
                        Cell myCell = myRow.createCell(j++);
                        myCell.setCellValue("[MAX]");
                        myCell.setCellStyle(removedStyle);
                    }
                }
                myRow.createCell(j++).setCellStyle(blackStyle);
            }
            String fileWrite = "test_elective_results.xls";
            FileOutputStream out = new FileOutputStream(fileWrite);
            workbook.write(out);
            out.close();
        } catch (Exception e) {
            System.out.println(e.getMessage());
            e.printStackTrace();
        }
        return true;
    }
    public static boolean genereazaExamene2(){
        i=0;
        j=0;
        Row myRow = mySheet.createRow(i);
        myRow.createCell(j++).setCellValue("Nr.Crt.");
        myRow.createCell(j++).setCellValue("Nume");
        myRow.createCell(j++).setCellValue("Grupa");
        myRow.createCell(j++).setCellValue("Media");
        for(int k=1;k<=10;k++){
            myRow.createCell(j++).setCellValue(""+k);
        }
        myRow.createCell(j++).setCellStyle(blackStyle);
        myRow.createCell(j++).setCellValue("CR "+limite.get("CR")+"-Gr. "+(int)(limite.get("CR")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("FSC "+limite.get("FSC")+"-Gr. "+(int)(limite.get("FSC")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("SE "+limite.get("SE")+"-Gr. "+(int)(limite.get("SE")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("MPS "+limite.get("MPS")+"-Gr. "+(int)(limite.get("MPS")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("EPSC "+limite.get("EPSC")+"-Gr. "+(int)(limite.get("EPSC")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("PD "+limite.get("PD")+"-Gr. "+(int)(limite.get("PD")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("CES "+limite.get("CES")+"-Gr. "+(int)(limite.get("CES")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("SM "+limite.get("SM")+"-Gr. "+(int)(limite.get("SM")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("LFA "+limite.get("LFA")+"-Gr. "+(int)(limite.get("LFA")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellValue("STD "+limite.get("STD")+"-Gr. "+(int)(limite.get("STD")/stuPeGrupa+0.25));
        myRow.createCell(j++).setCellStyle(blackStyle);
        i++;
        for(String materie:e2){
            examene2.replace(materie,0);
        }
        try {
            for (StudOptionale student : students) {
                j = 0;
                myRow = mySheet.createRow(i);
                myRow.createCell(j++).setCellValue(i++); //Nr.Crt.
                myRow.createCell(j++).setCellValue(student.getName()); //Nume
                myRow.createCell(j++).setCellValue(student.getGroup()); //Grupa
                myRow.createCell(j++).setCellValue(student.getMedia()); //Media
                String[] ex2Ordine = student.getOrderedElectives(examene2, e2);
                int ex2Alese = 0;
                for (String materie : ex2Ordine) {
                    int lim = examene2.get(materie);
                    if (ex2Alese < 4) {
                        if (lim < limite.get(materie)) {
                            examene2.replace(materie, lim + 1);
                            Cell myCell = myRow.createCell(j++);
                            myCell.setCellValue(materie);
                            myCell.setCellStyle(exam2Style);
                            ex2Alese++;
                        } else {
                            Cell myCell = myRow.createCell(j++);
                            myCell.setCellValue(materie);
                            myCell.setCellStyle(removedStyle);
                        }
                    } else {
                        myRow.createCell(j++).setCellValue(materie);
                    }
                }
                if(ex2Alese!=4){
                    return false;
                }
                myRow.createCell(j++).setCellStyle(blackStyle);
                String[] auxArray = {"CR", "FSC", "SE", "MPS", "EPSC", "PD", "CES", "SM", "LFA", "STD"};
                for (String aux : auxArray) {
                    if (!examene2.get(aux).equals(limite.get(aux))) {
                        myRow.createCell(j++).setCellValue(aux + " " + (limite.get(aux) - examene2.get(aux)));
                    } else {
                        Cell myCell = myRow.createCell(j++);
                        myCell.setCellValue("[MAX]");
                        myCell.setCellStyle(removedStyle);
                    }
                }
                myRow.createCell(j++).setCellStyle(blackStyle);
            }
            String fileWrite = "test_elective_results.xls";
            FileOutputStream out = new FileOutputStream(fileWrite);
            workbook.write(out);
            out.close();
            workbook.close();
        } catch (Exception e) {
            System.out.println(e.getMessage());
            e.printStackTrace();
        }
        return true;
    }
    public static void optionaleRepartizare(String file){
        try {
            POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(file));
            HSSFWorkbook wb = new HSSFWorkbook(fs);
            HSSFSheet sheet = wb.getSheetAt(0);
            HSSFRow row;
            HSSFCell cell;

            int rows; // No of rows
            rows = sheet.getPhysicalNumberOfRows();

            int cols = 0; // No of columns
            int tmp = 0;

            // This trick ensures that we get the data properly even if it doesn't start from first few rows
            for (int i = 0; i < 10 || i < rows; i++) {
                row = sheet.getRow(i);
                if (row != null) {
                    tmp = sheet.getRow(i).getPhysicalNumberOfCells();
                    if (tmp > cols) cols = tmp;
                }
            }
            /**FORMAT: Nr.Crt. | NUME | GRUPA | MEDIA | OPT1 | OPT2 | OPT3 | OPT4*/


            //i=0;
            //j=0;
            //Row myRow = mySheet.createRow(i);
            //myRow.createCell(j++).setCellValue("Nr.Crt.");
            //myRow.createCell(j++).setCellValue("Nume");
            //myRow.createCell(j++).setCellValue("Grupa");
            //myRow.createCell(j++).setCellValue("Media");
            //for(int k=1;k<=9;k++){
            //    myRow.createCell(j++).setCellValue(""+k);
            //}
            //myRow.createCell(j++).setCellStyle(blackStyle);
            //myRow.createCell(j++).setCellValue("IEP");
            //myRow.createCell(j++).setCellValue("TD");
            //myRow.createCell(j++).setCellValue("MS");
            //myRow.createCell(j++).setCellValue("VVS");
            //myRow.createCell(j++).setCellValue("TPAC");
            //myRow.createCell(j++).setCellValue("SMA");
            //myRow.createCell(j++).setCellValue("APND");
            //myRow.createCell(j++).setCellValue("AIAW");
            //myRow.createCell(j++).setCellValue("PBD");
            //myRow.createCell(j++).setCellStyle(blackStyle);
            //for(int k=1;k<=5;k++){
            //    myRow.createCell(j++).setCellValue(""+k);
            //}
            //myRow.createCell(j++).setCellStyle(blackStyle);
            //myRow.createCell(j++).setCellValue("SSC");
            //myRow.createCell(j++).setCellValue("CRF");
            //myRow.createCell(j++).setCellValue("CHS");
            //myRow.createCell(j++).setCellValue("SIPAC");
            //myRow.createCell(j++).setCellValue("PCBE");
            //myRow.createCell(j++).setCellStyle(blackStyle);
            //for(int k=1;k<=10;k++){
            //    myRow.createCell(j++).setCellValue(""+k);
            //}
            //myRow.createCell(j++).setCellStyle(blackStyle);
            //myRow.createCell(j++).setCellValue("CR");
            //myRow.createCell(j++).setCellValue("FSC");
            //myRow.createCell(j++).setCellValue("SE");
            //myRow.createCell(j++).setCellValue("MPS");
            //myRow.createCell(j++).setCellValue("EPSC");
            //myRow.createCell(j++).setCellValue("PD");
            //myRow.createCell(j++).setCellValue("CES");
            //myRow.createCell(j++).setCellValue("SM");
            //myRow.createCell(j++).setCellValue("LFA");
            //myRow.createCell(j++).setCellValue("STD");
            //myRow.createCell(j++).setCellStyle(blackStyle);
            //i++;
            for ( int r=0;r < rows; r++) {
                row = sheet.getRow(r);
                if (row != null) {
                    /**CREATING STUDENT */
                    StudOptionale student = null;
                    if (r != 0) {
                        student = new StudOptionale();
                    }
                    for (int c = 0; c < cols; c++) {
                        //System.out.println("Columns: "+cols);
                        cell = row.getCell((short) c);
                        if (cell != null) {
                            // Your code here
                            if (r != 0) {
                                if(c==0) {
                                    /**NR.CRT.*/
                                    student.setNrCrt(cell.getNumericCellValue());
                                }else if (c == 1) {
                                    /**STUDENT NAME */
                                    //System.out.println("aici");
                                    //System.out.println(cell.getStringCellValue());
                                    student.setName(cell.getStringCellValue());
                                } else if (c == 2) {
                                    /**STUDENT GROUP */
                                    //System.out.println("acolo");
                                    //System.out.println(cell.getStringCellValue());
                                    student.setGroup(cell.getStringCellValue());
                                } else if (c == 3) {
                                    /**STUDENT MEDIA */
                                    //System.out.println(cell.getNumericCellValue());
                                    //System.out.println(cell.getStringCellValue());
                                    student.setMedia(cell.getNumericCellValue());
                                } //else if (cell.getStringCellValue() != "") {
                                //student.addOption(cell.getStringCellValue());
                                //}
                                else if (c>3){
                                    if(!cell.getStringCellValue().equals("")) {
                                        //System.out.println("Adaug optiune "+cell.getStringCellValue());
                                        int value=c;
                                        if(c<13) {
                                            value-=3;
                                            if(!examene1.containsKey(cell.getStringCellValue())) {
                                                examene1.put(cell.getStringCellValue(),0);
                                                e1.add(cell.getStringCellValue());
                                            }
                                            //System.out.println(cell.getStringCellValue() + " " + value);
                                        } else if(c<18){
                                            value-=12;
                                            if(!distribuite.containsKey(cell.getStringCellValue())) {
                                                distribuite.put(cell.getStringCellValue(),0);
                                                d.add(cell.getStringCellValue());
                                            }
                                            //System.out.println(cell.getStringCellValue() + " " + value);
                                        } else {
                                            value-=17;
                                            if(!examene2.containsKey(cell.getStringCellValue())) {
                                                examene2.put(cell.getStringCellValue(),0);
                                                e2.add(cell.getStringCellValue());
                                            }
                                            //System.out.println(cell.getStringCellValue() + " " + value);
                                        }
                                        student.addOption(cell.getStringCellValue(),value);
                                    }
                                }
                            }
                        }
                    }
                    if(student!=null && !student.getName().equals("")) {
                        students.add(student);
                        student.initAleseSterse();
                    }
                        /*//TODO add code here
                        j=0;
                        myRow = mySheet.createRow(i);
                        myRow.createCell(j++).setCellValue(i++); //Nr.Crt.
                        myRow.createCell(j++).setCellValue(student.getName()); //Nume
                        myRow.createCell(j++).setCellValue(student.getGroup()); //Grupa
                        myRow.createCell(j++).setCellValue(student.getMedia()); //Media
                        String[] ex1Ordine = student.getOrderedElectives(examene1, e1);
                        String[] distrOrdine = student.getOrderedElectives(distribuite,d);
                        String[] ex2Ordine = student.getOrderedElectives(examene2,e2);
                        int ex1Alese=0;
                        for(String materie:ex1Ordine){
                            int lim = examene1.get(materie);
                                if (ex1Alese < 4) {
                                    //System.out.println("FirstCheck"+student.getName());
                                    if (lim < limite.get(materie)) {
                                        examene1.replace(materie, lim + 1);
                                        student.alegeOpt(materie);
                                        Cell myCell = myRow.createCell(j++);
                                        myCell.setCellValue(materie);
                                        myCell.setCellStyle(exam1Style);
                                        ex1Alese++;
                                        //System.out.println("SecondCheck"+student.getName());
                                    } else {
                                        Cell myCell = myRow.createCell(j++);
                                        myCell.setCellValue(materie);
                                        myCell.setCellStyle(removedStyle);
                                    }
                                } else {
                                    myRow.createCell(j++).setCellValue(materie);
                                }
                        }
                        myRow.createCell(j++).setCellStyle(blackStyle);
                        String[] auxArray = {"IEP", "TD", "MS", "VVS", "TPAC", "SMA", "APND", "AIAW", "PBD"};
                        for (String aux : auxArray) {
                            if (examene1.get(aux) != limite.get(aux)) {
                                myRow.createCell(j++).setCellValue(aux + " " + (limite.get(aux) - examene1.get(aux)));
                            } else {
                                Cell myCell = myRow.createCell(j++);
                                myCell.setCellValue("[MAX]");
                                myCell.setCellStyle(removedStyle);
                            }
                        }
                        myRow.createCell(j++).setCellStyle(blackStyle);
                        int distrAlese=0;
                        for (String materie : distrOrdine) {
                            int lim = distribuite.get(materie);
                            if (distrAlese < 2) {
                                //System.out.println("FirstCheck" + student.getName());
                                if (lim < limite.get(materie)) {
                                    distribuite.replace(materie, lim + 1);
                                    student.alegeOpt(materie);
                                    //System.out.println("SecondCheck"+ student.getName());
                                    Cell myCell = myRow.createCell(j++);
                                    myCell.setCellValue(materie);
                                    myCell.setCellStyle(distrStyle);
                                    distrAlese++;
                                } else {
                                    Cell myCell = myRow.createCell(j++);
                                    myCell.setCellValue(materie);
                                    myCell.setCellStyle(removedStyle);
                                }
                            } else {
                                myRow.createCell(j++).setCellValue(materie);
                            }
                        }
                        myRow.createCell(j++).setCellStyle(blackStyle);
                        String[] auxArray1 = {"SSC", "CRF", "CHS", "SIPAC", "PCBE"};
                        for (String aux : auxArray1) {
                            if (distribuite.get(aux) != limite.get(aux)) {
                                myRow.createCell(j++).setCellValue(aux + " " + (limite.get(aux) - distribuite.get(aux)));
                            } else {
                                Cell myCell = myRow.createCell(j++);
                                myCell.setCellValue("[MAX]");
                                myCell.setCellStyle(removedStyle);
                            }
                        }
                        myRow.createCell(j++).setCellStyle(blackStyle);
                        int ex2Alese=0;
                        for(String materie:ex2Ordine){
                            int lim = examene2.get(materie);
                            if(ex2Alese<4) {
                                //System.out.println("FirstCheck"+student.getName());
                                if (lim < limite.get(materie)) {
                                    examene2.replace(materie, lim + 1);
                                    student.alegeOpt(materie);
                                    //System.out.println("SecondCheck"+student.getName());
                                    Cell myCell = myRow.createCell(j++);
                                    myCell.setCellValue(materie);
                                    myCell.setCellStyle(exam2Style);
                                    ex2Alese++;
                                } else {
                                    Cell myCell = myRow.createCell(j++);
                                    myCell.setCellValue(materie);
                                    myCell.setCellStyle(removedStyle);
                                }

                            } else {
                                myRow.createCell(j++).setCellValue(materie);
                            }
                        }
                        myRow.createCell(j++).setCellStyle(blackStyle);
                        String[] auxArray2 = {"CR", "FSC", "SE", "MPS", "EPSC", "PD", "CES", "SM", "LFA", "STD"};
                        for (String aux : auxArray2) {
                            if (examene2.get(aux) != limite.get(aux)) {
                                myRow.createCell(j++).setCellValue(aux + " " + (limite.get(aux) - examene2.get(aux)));
                            } else {
                                Cell myCell = myRow.createCell(j++);
                                myCell.setCellValue("[MAX]");
                                myCell.setCellStyle(removedStyle);
                            }
                        }
                        myRow.createCell(j++).setCellStyle(blackStyle);
                        student.setEx1Order(ex1Ordine);
                        student.setEx2Order(ex2Ordine);
                        student.setDistrOrder(distrOrdine);
                        System.out.println(student.getName()+" Distr: "+student.getDistrAlese());*/
                }
            }
            wb.close();
        } catch (Exception ioe) {
            System.out.println(ioe.getMessage());
            ioe.printStackTrace();
        }
    }
    public static Comparator<String> optionaleAscending = new Comparator<String>() {

        public int compare(String s1, String s2) {
            String a = s1.toUpperCase();
            String b = s2.toUpperCase();
            //System.out.print("Student1: "+studentName1);
            //System.out.print(" Student2: "+studentName2);
            //System.out.println(" CompareToResult: "+studentName1.compareTo(studentName2));
            //ascending order
            //studentName1>studentName2 => >0;
            //studentName1<studentName2 => <0;
            return a.compareTo(b);

            //descending order
            //return StudentName2.compareTo(StudentName1);
        }
    };
    public static boolean grupeDistribuite(){
        mySheet = workbook.createSheet("Distribuite");
        //initWorkbook(workbook);
        int initialSSC=3;
        int initialCRF=6;
        int initialCHS=6;
        int initialSIPAC=4;
        int initialPCBE=5;
        int grupeMax = 7;
        int grupeMin = 2;
        int min = Integer.MAX_VALUE;
        int max = Integer.MIN_VALUE;
        int noZero=0;
        int countChanges=0;
        int maxChange=Integer.MIN_VALUE;
        int mina=-1;
        int minb=-1;
        int minc=-1;
        int mind=-1;
        int mine=-1;
        int suma=0;
        String[] auxArray = {"SSC", "CRF", "CHS", "SIPAC", "PCBE"};
        for(int ssc=grupeMin;ssc<=grupeMax;ssc++){
            for(int crf=grupeMin;crf<=grupeMax;crf++){
                for(int chs=grupeMin;chs<=grupeMax;chs++){
                    for(int sipac=grupeMin;sipac<=grupeMax;sipac++){
                        for(int pcbe=grupeMin;pcbe<=grupeMax;pcbe++){
                            limite.replace("SSC",   (int) (ssc    *stuPeGrupa));
                            limite.replace("CRF",   (int) (crf    *stuPeGrupa));
                            limite.replace("CHS",   (int) (chs    *stuPeGrupa));
                            limite.replace("SIPAC", (int) (sipac  *stuPeGrupa));
                            limite.replace("PCBE",  (int) (pcbe   *stuPeGrupa));
                            //System.out.println(ssc+" "+crf+" "+chs+" "+sipac+" "+pcbe+" end: "+end);
                            int sum =ssc+crf+chs+sipac+pcbe;
                            if(sum==24) {
                                if (genereazaDistribuite()) {
                                    int a = ssc - initialSSC;
                                    int b = crf - initialCRF;
                                    int c = chs - initialCHS;
                                    int d = sipac - initialSIPAC;
                                    int e = pcbe - initialPCBE;
                                    if(a==0){
                                        countChanges++;
                                    }
                                    if(b==0){
                                        countChanges++;
                                    }
                                    if(c==0){
                                        countChanges++;
                                    }
                                    if(d==0){
                                        countChanges++;
                                    }
                                    if(e==0){
                                        countChanges++;
                                    }
                                    if(Math.abs(a)<=maxGrupError &&Math.abs(b)<=maxGrupError &&
                                            Math.abs(c)<=maxGrupError &&Math.abs(d)<=maxGrupError &&
                                            Math.abs(e)<=maxGrupError) {
                                        for (String aux : auxArray) {
                                            suma = suma + limite.get(aux) - distribuite.get(aux);
                                            if(limite.get(aux) - distribuite.get(aux)==0){
                                                //noZero++;
                                            }
                                        }
                                        //System.out.println(suma);
                                        //noZero>=max &&
                                        if ( suma<=min && countChanges>=maxChange) {
                                            min = suma;
                                            //max = noZero;
                                            maxChange=countChanges;
                                            mina = a;
                                            minb = b;
                                            minc = c;
                                            mind = d;
                                            mine = e;
                                        }
                                        suma = 0;
                                        noZero=0;
                                    }
                                    countChanges=0;
                                }
                            }
                        }
                    }
                }
            }
        }
        System.out.println("BEST "+min+" noZero:"+maxChange+": SSC: " + mina +"+"+initialSSC+"="+(mina+initialSSC) + " CRF: " + minb +"+"+initialCRF+
                "="+(minb+initialCRF)+ " CHS: " + minc +"+"+initialCHS+"="+(minc+initialCHS)+
                " SIPAC: " + mind +"+"+initialSIPAC+"="+(mind+initialSIPAC)+
                " PCBE: " + mine+"+"+initialPCBE+"="+(mine+initialPCBE));
        limite.replace("SSC",   (int) ((mina+initialSSC)  *stuPeGrupa));
        limite.replace("CRF",   (int) ((minb+initialCRF)  *stuPeGrupa));
        limite.replace("CHS",   (int) ((minc+initialCHS)  *stuPeGrupa));
        limite.replace("SIPAC", (int) ((mind+initialSIPAC)  *stuPeGrupa));
        limite.replace("PCBE",  (int) ((mine+initialPCBE)  *stuPeGrupa));
        genereazaDistribuite();
        return true;
    }
    public static boolean grupeExamene1(){
        mySheet = workbook.createSheet("Examene 1");
        //initWorkbook(workbook);
        int initialIEP=6;
        int initialTD=5;
        int initialMS=5;
        int initialVVS=4;
        int initialTPAC=6;
        int initialSMA=6;
        int initialAPND=5;
        int initialAIAW=5;
        int initialPBD=4;
        int grupeMax = 7;
        int grupeMin = 3;
        int min = Integer.MAX_VALUE;
        int max = Integer.MIN_VALUE;
        int noZero=0;
        int countChanges=0;
        int maxChanges=Integer.MIN_VALUE;
        int mina=-1;
        int minb=-1;
        int minc=-1;
        int mind=-1;
        int mine=-1;
        int minf=-1;
        int ming=-1;
        int minh=-1;
        int mink=-1;
        int suma=0;
        String[] auxArray = {"IEP", "TD", "MS", "VVS", "TPAC", "SMA", "APND", "AIAW", "PBD"};
        for(int iep=grupeMin;iep<=grupeMax;iep++){
            for(int td=grupeMin;td<=grupeMax;td++){
                for(int ms=grupeMin;ms<=grupeMax;ms++){
                    for(int vvs=grupeMin;vvs<=grupeMax;vvs++){
                        for(int tpac=grupeMin;tpac<=grupeMax;tpac++){
                            for(int sma=grupeMin;sma<=grupeMax;sma++) {
                                for (int apnd = grupeMin; apnd<=grupeMax; apnd++) {
                                    for (int aiaw = grupeMin; aiaw<=grupeMax; aiaw++) {
                                        for (int pbd = grupeMin; pbd<=grupeMax; pbd++) {
                                            limite.replace("IEP", (int)    (iep  * stuPeGrupa));
                                            limite.replace("TD", (int)     (td   * stuPeGrupa));
                                            limite.replace("MS", (int)     (ms   * stuPeGrupa));
                                            limite.replace("VVS", (int)    (vvs  * stuPeGrupa));
                                            limite.replace("TPAC", (int)   (tpac * stuPeGrupa));
                                            limite.replace("SMA",   (int)  (sma  * stuPeGrupa));
                                            limite.replace("APND",   (int) (apnd * stuPeGrupa));
                                            limite.replace("AIAW", (int)   (aiaw * stuPeGrupa));
                                            limite.replace("PBD",  (int)   (pbd  * stuPeGrupa));
                                            //System.out.println(ssc+" "+crf+" "+chs+" "+sipac+" "+pcbe+" end: "+end);
                                            int sum = iep + td + ms + vvs + tpac + sma + apnd + aiaw + pbd;
                                            if (sum == 48) {
                                                if (genereazaExamene1()) {
                                                    int a = iep - initialIEP;
                                                    int b = td - initialTD;
                                                    int c = ms - initialMS;
                                                    int d = vvs - initialVVS;
                                                    int e = tpac - initialTPAC;
                                                    int f = sma - initialSMA;
                                                    int g = apnd - initialAPND;
                                                    int h = aiaw - initialAIAW;
                                                    int k = pbd - initialPBD;
                                                    if(a==0){
                                                        countChanges++;
                                                    }
                                                    if(b==0){
                                                        countChanges++;
                                                    }
                                                    if(c==0){
                                                        countChanges++;
                                                    }
                                                    if(d==0){
                                                        countChanges++;
                                                    }
                                                    if(e==0){
                                                        countChanges++;
                                                    }
                                                    if(f==0){
                                                        countChanges++;
                                                    }
                                                    if(g==0){
                                                        countChanges++;
                                                    }
                                                    if(h==0){
                                                        countChanges++;
                                                    }
                                                    if(k==0){
                                                        countChanges++;
                                                    }
                                                    if(Math.abs(a)<=maxGrupError && Math.abs(b)<=maxGrupError &&
                                                            Math.abs(c)<=maxGrupError && Math.abs(d)<=maxGrupError &&
                                                            Math.abs(e)<=maxGrupError && Math.abs(f)<=maxGrupError &&
                                                            Math.abs(g)<=maxGrupError && Math.abs(h)<=maxGrupError &&
                                                            Math.abs(k)<=maxGrupError) {
                                                        for (String aux : auxArray) {
                                                            suma = suma + limite.get(aux) - examene1.get(aux);
                                                            if(limite.get(aux) - examene1.get(aux)==0){
                                                                //noZero++;
                                                            }
                                                        }
                                                        //System.out.println(suma);
                                                        //noZero>=max &&
                                                        if (suma <= min && countChanges>=maxChanges) {
                                                            min = suma;
                                                            //max = noZero;
                                                            maxChanges=countChanges;
                                                            mina = a;
                                                            minb = b;
                                                            minc = c;
                                                            mind = d;
                                                            mine = e;
                                                            minf = f;
                                                            ming = g;
                                                            minh = h;
                                                            mink = k;
                                                        }
                                                        suma = 0;
                                                        //noZero = 0;
                                                    }
                                                    countChanges=0;
                                                }
                                                //System.out.println(genereazaExamene1());
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        System.out.println("BEST "+min+" noZero:"+maxChanges+": IEP: " + mina +"+"+initialIEP+"="+(mina+initialIEP)+
                " TD: " + minb +"+"+initialTD+"="+(minb+initialTD)+ " MS: " + minc +"+"+initialMS+"="+(minc+initialMS)+
                " VVS: " + mind +"+"+initialVVS+"="+(mind+initialVVS)+ " TPAC: " + mine +"+"+initialTPAC+"="+(mine+initialTPAC)+
                " SMA: " + minf +"+"+initialSMA+"="+(minf+initialSMA)+ " APND: " + ming +"+"+initialAPND+"="+(ming+initialAPND)+
                " AIAW: " + minh +"+"+initialAIAW+"="+(minh+initialAIAW)+ " PBD: " + mink+"+"+initialPBD+"="+(mink+initialPBD));
        limite.replace("IEP", (int)    ((mina+initialIEP)   * stuPeGrupa));
        limite.replace("TD", (int)     ((minb+initialTD)    * stuPeGrupa));
        limite.replace("MS", (int)     ((minc+initialMS)    * stuPeGrupa));
        limite.replace("VVS", (int)    ((mind+initialVVS)   * stuPeGrupa));
        limite.replace("TPAC", (int)   ((mine+initialTPAC)  * stuPeGrupa));
        limite.replace("SMA",   (int)  ((minf+initialSMA)   * stuPeGrupa));
        limite.replace("APND",   (int) ((ming+initialAPND)  * stuPeGrupa));
        limite.replace("AIAW", (int)   ((minh+initialAIAW)  * stuPeGrupa));
        limite.replace("PBD",  (int)   ((mink+initialPBD)   * stuPeGrupa));
        genereazaExamene1();
        return true;
    }
    public static boolean grupeExamene2(){
        mySheet = workbook.createSheet("Examene 2");
        //initWorkbook(workbook);
        int initialCR=4;
        int initialFSC=5;
        int initialSE=4;
        int initialMPS=5;
        int initialEPSC=5;
        int initialPD=5;
        int initialCES=5;
        int initialSM=5;
        int initialLFA=5;
        int initialSTD=5;
        int grupeMax = 6;
        int grupeMin = 3;
        int min = Integer.MAX_VALUE;
        int max = Integer.MIN_VALUE;
        int noZero=0;
        int countChanges=0;
        int maxChanges=Integer.MIN_VALUE;
        int mina=-1;
        int minb=-1;
        int minc=-1;
        int mind=-1;
        int mine=-1;
        int minf=-1;
        int ming=-1;
        int minh=-1;
        int mink=-1;
        int minl=-1;
        int suma=0;
        String[] auxArray = {"CR", "FSC", "SE", "MPS", "EPSC", "PD", "CES", "SM", "LFA", "STD"};
        for(int cr=grupeMin;cr<=grupeMax;cr++){
            for(int fsc=grupeMin;fsc<=grupeMax;fsc++){
                for(int se=grupeMin;se<=grupeMax;se++){
                    for(int mps=grupeMin;mps<=grupeMax;mps++){
                        for(int epsc=grupeMin;epsc<=grupeMax;epsc++){
                            for(int pd=grupeMin;pd<=grupeMax;pd++) {
                                for (int ces = grupeMin; ces<=grupeMax; ces++) {
                                    for (int sm = grupeMin; sm<=grupeMax; sm++) {
                                        for (int lfa = grupeMin; lfa<=grupeMax; lfa++) {
                                            for (int std = grupeMin; std<=grupeMax; std++) {
                                                limite.replace("CR", (int) (cr * stuPeGrupa));
                                                limite.replace("FSC", (int) (fsc * stuPeGrupa));
                                                limite.replace("SE", (int) (se * stuPeGrupa));
                                                limite.replace("MPS", (int) (mps * stuPeGrupa));
                                                limite.replace("EPSC", (int) (epsc * stuPeGrupa));
                                                limite.replace("PD", (int) (pd * stuPeGrupa));
                                                limite.replace("CES", (int) (ces * stuPeGrupa));
                                                limite.replace("SM", (int) (sm * stuPeGrupa));
                                                limite.replace("LFA", (int) (lfa * stuPeGrupa));
                                                limite.replace("STD", (int) (std * stuPeGrupa));
                                                //System.out.println(ssc+" "+crf+" "+chs+" "+sipac+" "+pcbe+" end: "+end);
                                                int sum = cr + fsc + se + mps + epsc + pd + ces + sm + lfa + std;
                                                if (sum == 48) {
                                                    if (genereazaExamene2()) {
                                                        int a = cr - initialCR;
                                                        int b = fsc - initialFSC;
                                                        int c = se - initialSE;
                                                        int d = mps - initialMPS;
                                                        int e = epsc - initialEPSC;
                                                        int f = pd - initialPD;
                                                        int g = ces - initialCES;
                                                        int h = sm - initialSM;
                                                        int k = lfa - initialLFA;
                                                        int l = std - initialSTD;
                                                        if(a==0){
                                                            countChanges++;
                                                        }
                                                        if(b==0){
                                                            countChanges++;
                                                        }
                                                        if(c==0){
                                                            countChanges++;
                                                        }
                                                        if(d==0){
                                                            countChanges++;
                                                        }
                                                        if(e==0){
                                                            countChanges++;
                                                        }
                                                        if(f==0){
                                                            countChanges++;
                                                        }
                                                        if(g==0){
                                                            countChanges++;
                                                        }
                                                        if(h==0){
                                                            countChanges++;
                                                        }
                                                        if(k==0){
                                                            countChanges++;
                                                        }
                                                        if(l==0){
                                                            countChanges++;
                                                        }
                                                        if(Math.abs(a)<=maxGrupError && Math.abs(b)<=maxGrupError &&
                                                                Math.abs(c)<=maxGrupError && Math.abs(d)<=maxGrupError &&
                                                                Math.abs(e)<=maxGrupError && Math.abs(f)<=maxGrupError &&
                                                                Math.abs(g)<=maxGrupError && Math.abs(h)<=maxGrupError &&
                                                                Math.abs(k)<=maxGrupError && Math.abs(l)<=maxGrupError) {
                                                            for (String aux : auxArray) {
                                                                suma = suma + limite.get(aux) - examene2.get(aux);
                                                                if(limite.get(aux) - examene2.get(aux)==0){
                                                                    //noZero++;
                                                                }
                                                            }
                                                            //System.out.println(suma);
                                                            //noZero>=max &&
                                                            if ( suma <= min && countChanges>=maxChanges) {
                                                                min = suma;
                                                                //max = noZero;
                                                                maxChanges=countChanges;
                                                                mina = a;
                                                                minb = b;
                                                                minc = c;
                                                                mind = d;
                                                                mine = e;
                                                                minf = f;
                                                                ming = g;
                                                                minh = h;
                                                                mink = k;
                                                                minl = l;
                                                            }
                                                            suma = 0;
                                                            //noZero = 0;
                                                        }
                                                        countChanges=0;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        System.out.println("BEST "+min+" noZero:"+maxChanges+": CR: " + mina+"+"+initialCR+"="+(mina+initialCR) +
                " FSC: " + minb +"+"+initialFSC+"="+(minb+initialFSC)+ " SE: " + minc +"+"+initialSE+"="+(minc+initialSE)+
                " MPS: " + mind +"+"+initialMPS+"="+(mind+initialMPS)+ " EPSC: " + mine +"+"+initialEPSC+"="+(mine+initialEPSC)+
                " PD: " + minf +"+"+initialPD+"="+(minf+initialPD)+ " CES: " + ming +"+"+initialCES+"="+(ming+initialCES)+
                " SM: " + minh +"+"+initialSM+"="+(minh+initialSM)+ " LFA: " + mink +"+"+initialLFA+"="+(mink+initialLFA)+
                " STD: " + minl+"+"+initialSTD+"="+(minl+initialSTD));
        limite.replace("CR", (int)  ((mina+initialCR)   * stuPeGrupa));
        limite.replace("FSC", (int) ((minb+initialFSC)  * stuPeGrupa));
        limite.replace("SE", (int)  ((minc+initialSE)   * stuPeGrupa));
        limite.replace("MPS", (int) ((mind+initialMPS)  * stuPeGrupa));
        limite.replace("EPSC", (int)((mine+initialEPSC) * stuPeGrupa));
        limite.replace("PD", (int)  ((minf+initialPD)   * stuPeGrupa));
        limite.replace("CES", (int) ((ming+initialCES)  * stuPeGrupa));
        limite.replace("SM", (int)  ((minh+initialSM)   * stuPeGrupa));
        limite.replace("LFA", (int) ((mink+initialLFA)  * stuPeGrupa));
        limite.replace("STD", (int) ((minl+initialSTD)  * stuPeGrupa));
        genereazaExamene2();
        return true;
    }
    public static void main(String[] args) throws IOException {
        //Workbook wb;
        //generateElectives("optionaleAn4.xls");
        //int []opt = {1,2,3,4,5,6};
        //int optMaxNumber = 18;
        //repartizeaza("test_elective.xls",opt, optMaxNumber);
        exam1Style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.index);
        exam1Style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        exam2Style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.index);
        exam2Style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        distrStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.index);
        distrStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        removedStyle.setFillForegroundColor(IndexedColors.RED.index);
        removedStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        blackStyle.setFillForegroundColor(IndexedColors.BLACK.index);
        blackStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        limite.put("IEP"  ,93 );
        limite.put("TD"   ,93 );
        limite.put("MS"   ,93 );
        limite.put("VVS"  ,93 );
        limite.put("TPAC" ,93 );
        limite.put("SMA"  ,93 );
        limite.put("APND" ,93 );
        limite.put("AIAW" ,93 );
        limite.put("PBD"  ,93 );//Distr

        limite.put("SSC"  ,47); //46
        limite.put("CRF"  ,62); //93
        limite.put("CHS"  ,93); //93
        limite.put("SIPAC",90); //62
        limite.put("PCBE" ,78); //78 //Examene2

        limite.put("CR"   ,62);
        limite.put("FSC"  ,78);
        limite.put("SE"   ,62);
        limite.put("MPS"  ,77);
        limite.put("EPSC" ,78);
        limite.put("PD"   ,77);
        limite.put("CES"  ,78);
        limite.put("SM"   ,77);
        limite.put("LFA"  ,78);
        limite.put("STD"  ,77);

        optionaleRepartizare("test_elective.xls");
        grupeDistribuite();
        grupeExamene1();
        //System.out.println(genereazaExamene1());
        grupeExamene2();
        /**
         * int initialSSC=  (int) (3*15.5);
         *         int initialCRF=  (int) (6*15.5);
         *         int initialCHS=  (int) (6*15.5);
         *         int initialSIPAC=(int) (4*15.5);
         *         int initialPCBE= (int) (5*15.5);
         *         int grupeMax = (int) (7*15.5);
         *         int grupeMin = (int) (2*15.5);
         *         int eroare=10;
         *         for(int ssc=initialSSC-eroare;ssc<initialSSC+eroare;ssc++){
         *             for(int crf=initialCRF-eroare;crf<initialCRF+eroare;crf++){
         *                 for(int chs=initialCHS-eroare;chs<initialCHS+eroare;chs++){
         *                     for(int sipac=initialSIPAC-eroare;sipac<initialSIPAC+eroare;sipac++){
         *                         for(int pcbe=initialSIPAC-eroare;pcbe<initialSIPAC+eroare;pcbe++){
         *                             limite.replace("SSC",   ssc);//(int) (ssc    *15.5));
         *                             limite.replace("CRF",   crf);//(int) (crf    *15.5));
         *                             limite.replace("CHS",   chs);//(int) (chs    *15.5));
         *                             limite.replace("SIPAC", sipac);//(int) (sipac  *15.5));
         *                             limite.replace("PCBE",  pcbe);//(int) (pcbe   *15.5));
         *                             //System.out.println(ssc+" "+crf+" "+chs+" "+sipac+" "+pcbe+" end: "+end);
         *                             if(genereazaDistribuite()){
         *                                 int sum =ssc+crf+chs+sipac+pcbe;
         *                                 if(sum==(int)(24*15.5)) {
         *                                     int a=ssc-initialSSC;
         *                                     int b=crf-initialCRF;
         *                                     int c=chs-initialCHS;
         *                                     int d=sipac-initialSIPAC;
         *                                     int e=pcbe-initialPCBE;
         *                                     System.out.println("GOOD "+sum+": SSC: " + a + " CRF: " + b +
         *                                             " CHS: " + c + " SIPAC: " + d +
         *                                             " PCBE: " + e);
         *                                 }
         *                                 //else {
         *                                 //    System.out.println("BAD "+sum+": SSC: " + ssc + " CRF: " + crf +
         *                                 //            " CHS: " + chs + " SIPAC: " + sipac +
         *                                 //            " PCBE: " + pcbe);
         *                                 //}
         *                             }
         *                             //System.out.println(ssc+""+crf+chs+sipac+pcbe);
         *                         }
         *                     }
         *                 }
         *             }
         *         }*/
        //if(args.length > 0 && args[0].equals("-xls")) {
        //wb = new HSSFWorkbook();
        //} else {
        //    wb = new XSSFWorkbook();
        //}
        /*sameCell(wb.createSheet("Same Cell"));
        multiCell(wb.createSheet("MultiCell"));
        overlapping(wb.createSheet("Overlapping"));
        errors(wb.createSheet("Errors"));
        hideDupplicates(wb.createSheet("Hide Dups"));
        formatDuplicates(wb.createSheet("Duplicates"));
        inList(wb.createSheet("In List"));
        expiry(wb.createSheet("Expiry"));
        shadeAlt(wb.createSheet("Shade Alt"));
        shadeBands(wb.createSheet("Shade Bands"));
        iconSets(wb.createSheet("Icon Sets"));
        colourScales(wb.createSheet("Colour Scales"));
        dataBars(wb.createSheet("Data Bars"));

        // print overlapping rule results
        evaluateRules(wb, "Overlapping");
        */
        // Write the output to a file
        //String file = "electives.xls";
        //if(wb instanceof XSSFWorkbook) {
        //    file += "x";
        //}
        //FileOutputStream out = new FileOutputStream(file);
        //wb.write(out);
        //out.close();
        //System.out.println("Generated: " + file);
        //wb.close();
    }

    /**
     * Highlight cells based on their values
     */
    static void sameCell(Sheet sheet) {
        sheet.createRow(0).createCell(0).setCellValue(84);
        sheet.createRow(1).createCell(0).setCellValue(74);
        sheet.createRow(2).createCell(0).setCellValue(50);
        sheet.createRow(3).createCell(0).setCellValue(51);
        sheet.createRow(4).createCell(0).setCellValue(49);
        sheet.createRow(5).createCell(0).setCellValue(41);

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Cell Value Is   greater than  70   (Blue Fill)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule(ComparisonOperator.GT, "70");
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(IndexedColors.BLUE.index);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        // Condition 2: Cell Value Is  less than      50   (Green Fill)
        ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.LT, "50");
        PatternFormatting fill2 = rule2.createPatternFormatting();
        fill2.setFillBackgroundColor(IndexedColors.GREEN.index);
        fill2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A1:A6")
        };

        sheetCF.addConditionalFormatting(regions, rule1, rule2);

        sheet.getRow(0).createCell(2).setCellValue("<== Condition 1: Cell Value Is greater than 70 (Blue Fill)");
        sheet.getRow(4).createCell(2).setCellValue("<== Condition 2: Cell Value Is less than 50 (Green Fill)");
    }

    /**
     * Highlight multiple cells based on a formula
     */
    static void multiCell(Sheet sheet) {
        // header row
        Row row0 = sheet.createRow(0);
        row0.createCell(0).setCellValue("Units");
        row0.createCell(1).setCellValue("Cost");
        row0.createCell(2).setCellValue("Total");

        Row row1 = sheet.createRow(1);
        row1.createCell(0).setCellValue(71);
        row1.createCell(1).setCellValue(29);
        row1.createCell(2).setCellValue(2059);

        Row row2 = sheet.createRow(2);
        row2.createCell(0).setCellValue(85);
        row2.createCell(1).setCellValue(29);
        row2.createCell(2).setCellValue(2059);

        Row row3 = sheet.createRow(3);
        row3.createCell(0).setCellValue(71);
        row3.createCell(1).setCellValue(29);
        row3.createCell(2).setCellValue(2059);

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Formula Is   =$B2>75   (Blue Fill)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("$A2>75");
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(IndexedColors.BLUE.index);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A2:C4")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.getRow(2).createCell(4).setCellValue("<== Condition 1: Formula Is =$B2>75   (Blue Fill)");
    }

    /**
     * Multiple conditional formatting rules can apply to
     *  one cell, some combining, some beating others.
     * Done in order of the rules added to the
     *  SheetConditionalFormatting object
     */
    static void overlapping(Sheet sheet) {
        for (int i=0; i<40; i++) {
            int rn = i+1;
            Row r = sheet.createRow(i);
            r.createCell(0).setCellValue("This is row " + rn + " (" + i + ")");
            String str = "";
            if (rn%2 == 0) {
                str = str + "even ";
            }
            if (rn%3 == 0) {
                str = str + "x3 ";
            }
            if (rn%5 == 0) {
                str = str + "x5 ";
            }
            if (rn%10 == 0) {
                str = str + "x10 ";
            }
            if (str.length() == 0) {
                str = "nothing special...";
            }
            r.createCell(1).setCellValue("It is " + str);
        }
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);

        sheet.getRow(1).createCell(3).setCellValue("Even rows are blue");
        sheet.getRow(2).createCell(3).setCellValue("Multiples of 3 have a grey background");
        sheet.getRow(4).createCell(3).setCellValue("Multiples of 5 are bold");
        sheet.getRow(9).createCell(3).setCellValue("Multiples of 10 are red (beats even)");

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Row divides by 10, red (will beat #1)
        ConditionalFormattingRule rule1 =
                sheetCF.createConditionalFormattingRule("MOD(ROW(),10)=0");
        FontFormatting font1 = rule1.createFontFormatting();
        font1.setFontColorIndex(IndexedColors.RED.index);

        // Condition 2: Row is even, blue
        ConditionalFormattingRule rule2 =
                sheetCF.createConditionalFormattingRule("MOD(ROW(),2)=0");
        FontFormatting font2 = rule2.createFontFormatting();
        font2.setFontColorIndex(IndexedColors.BLUE.index);

        // Condition 3: Row divides by 5, bold
        ConditionalFormattingRule rule3 =
                sheetCF.createConditionalFormattingRule("MOD(ROW(),5)=0");
        FontFormatting font3 = rule3.createFontFormatting();
        font3.setFontStyle(false, true);

        // Condition 4: Row divides by 3, grey background
        ConditionalFormattingRule rule4 =
                sheetCF.createConditionalFormattingRule("MOD(ROW(),3)=0");
        PatternFormatting fill4 = rule4.createPatternFormatting();
        fill4.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.index);
        fill4.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        // Apply
        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A1:F41")
        };

        sheetCF.addConditionalFormatting(regions, rule1);
        sheetCF.addConditionalFormatting(regions, rule2);
        sheetCF.addConditionalFormatting(regions, rule3);
        sheetCF.addConditionalFormatting(regions, rule4);
    }

    /**
     *  Use Excel conditional formatting to check for errors,
     *  and change the font colour to match the cell colour.
     *  In this example, if formula result is  #DIV/0! then it will have white font colour.
     */
    static void errors(Sheet sheet) {
        sheet.createRow(0).createCell(0).setCellValue(84);
        sheet.createRow(1).createCell(0).setCellValue(0);
        sheet.createRow(2).createCell(0).setCellFormula("ROUND(A1/A2,0)");
        sheet.createRow(3).createCell(0).setCellValue(0);
        sheet.createRow(4).createCell(0).setCellFormula("ROUND(A6/A4,0)");
        sheet.createRow(5).createCell(0).setCellValue(41);

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Formula Is   =ISERROR(C2)   (White Font)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("ISERROR(A1)");
        FontFormatting font = rule1.createFontFormatting();
        font.setFontColorIndex(IndexedColors.WHITE.index);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A1:A6")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.getRow(2).createCell(1).setCellValue("<== The error in this cell is hidden. Condition: Formula Is   =ISERROR(C2)   (White Font)");
        sheet.getRow(4).createCell(1).setCellValue("<== The error in this cell is hidden. Condition: Formula Is   =ISERROR(C2)   (White Font)");
    }

    /**
     * Use Excel conditional formatting to hide the duplicate values,
     * and make the list easier to read. In this example, when the table is sorted by Region,
     * the second (and subsequent) occurences of each region name will have white font colour.
     */
    static void hideDupplicates(Sheet sheet) {
        sheet.createRow(0).createCell(0).setCellValue("City");
        sheet.createRow(1).createCell(0).setCellValue("Boston");
        sheet.createRow(2).createCell(0).setCellValue("Boston");
        sheet.createRow(3).createCell(0).setCellValue("Chicago");
        sheet.createRow(4).createCell(0).setCellValue("Chicago");
        sheet.createRow(5).createCell(0).setCellValue("New York");

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Formula Is   =A2=A1   (White Font)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("A2=A1");
        FontFormatting font = rule1.createFontFormatting();
        font.setFontColorIndex(IndexedColors.WHITE.index);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A2:A6")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.getRow(1).createCell(1).setCellValue("<== the second (and subsequent) " +
                "occurences of each region name will have white font colour.  " +
                "Condition: Formula Is   =A2=A1   (White Font)");
    }

    /**
     * Use Excel conditional formatting to highlight duplicate entries in a column.
     */
    static void formatDuplicates(Sheet sheet) {
        sheet.createRow(0).createCell(0).setCellValue("Code");
        sheet.createRow(1).createCell(0).setCellValue(4);
        sheet.createRow(2).createCell(0).setCellValue(3);
        sheet.createRow(3).createCell(0).setCellValue(6);
        sheet.createRow(4).createCell(0).setCellValue(3);
        sheet.createRow(5).createCell(0).setCellValue(5);
        sheet.createRow(6).createCell(0).setCellValue(8);
        sheet.createRow(7).createCell(0).setCellValue(0);
        sheet.createRow(8).createCell(0).setCellValue(2);
        sheet.createRow(9).createCell(0).setCellValue(8);
        sheet.createRow(10).createCell(0).setCellValue(6);

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Formula Is   =A2=A1   (White Font)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("COUNTIF($A$2:$A$11,A2)>1");
        FontFormatting font = rule1.createFontFormatting();
        font.setFontStyle(false, true);
        font.setFontColorIndex(IndexedColors.BLUE.index);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A2:A11")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.getRow(2).createCell(1).setCellValue("<== Duplicates numbers in the column are highlighted.  " +
                "Condition: Formula Is =COUNTIF($A$2:$A$11,A2)>1   (Blue Font)");
    }

    /**
     * Use Excel conditional formatting to highlight items that are in a list on the worksheet.
     */
    static void inList(Sheet sheet) {
        sheet.createRow(0).createCell(0).setCellValue("Codes");
        sheet.createRow(1).createCell(0).setCellValue("AA");
        sheet.createRow(2).createCell(0).setCellValue("BB");
        sheet.createRow(3).createCell(0).setCellValue("GG");
        sheet.createRow(4).createCell(0).setCellValue("AA");
        sheet.createRow(5).createCell(0).setCellValue("FF");
        sheet.createRow(6).createCell(0).setCellValue("XX");
        sheet.createRow(7).createCell(0).setCellValue("CC");

        sheet.getRow(0).createCell(2).setCellValue("Valid");
        sheet.getRow(1).createCell(2).setCellValue("AA");
        sheet.getRow(2).createCell(2).setCellValue("BB");
        sheet.getRow(3).createCell(2).setCellValue("CC");

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Formula Is   =A2=A1   (White Font)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("COUNTIF($C$2:$C$4,A2)");
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(IndexedColors.LIGHT_BLUE.index);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A2:A8")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.getRow(2).createCell(3).setCellValue("<== Use Excel conditional formatting to highlight items that are in a list on the worksheet");
    }

    /**
     *  Use Excel conditional formatting to highlight payments that are due in the next thirty days.
     *  In this example, Due dates are entered in cells A2:A4.
     */
    static void expiry(Sheet sheet) {
        CellStyle style = sheet.getWorkbook().createCellStyle();
        style.setDataFormat((short)BuiltinFormats.getBuiltinFormat("d-mmm"));

        sheet.createRow(0).createCell(0).setCellValue("Date");
        sheet.createRow(1).createCell(0).setCellFormula("TODAY()+29");
        sheet.createRow(2).createCell(0).setCellFormula("A2+1");
        sheet.createRow(3).createCell(0).setCellFormula("A3+1");

        for(int rownum = 1; rownum <= 3; rownum++) {
            sheet.getRow(rownum).getCell(0).setCellStyle(style);
        }

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Formula Is   =A2=A1   (White Font)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("AND(A2-TODAY()>=0,A2-TODAY()<=30)");
        FontFormatting font = rule1.createFontFormatting();
        font.setFontStyle(false, true);
        font.setFontColorIndex(IndexedColors.BLUE.index);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A2:A4")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.getRow(0).createCell(1).setCellValue("Dates within the next 30 days are highlighted");
    }

    /**
     * Use Excel conditional formatting to shade alternating rows on the worksheet
     */
    static void shadeAlt(Sheet sheet) {
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        // Condition 1: Formula Is   =A2=A1   (White Font)
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("MOD(ROW(),2)");
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.index);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A1:Z100")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.createRow(0).createCell(1).setCellValue("Shade Alternating Rows");
        sheet.createRow(1).createCell(1).setCellValue("Condition: Formula Is  =MOD(ROW(),2)   (Light Green Fill)");
    }

    /**
     * You can use Excel conditional formatting to shade bands of rows on the worksheet.
     * In this example, 3 rows are shaded light grey, and 3 are left with no shading.
     * In the MOD function, the total number of rows in the set of banded rows (6) is entered.
     */
    static void shadeBands(Sheet sheet) {
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule("MOD(ROW(),6)<3");
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.index);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A1:Z100")
        };

        sheetCF.addConditionalFormatting(regions, rule1);

        sheet.createRow(0).createCell(1).setCellValue("Shade Bands of Rows");
        sheet.createRow(1).createCell(1).setCellValue("Condition: Formula Is  =MOD(ROW(),6)<2   (Light Grey Fill)");
    }

    /**
     * Icon Sets / Multi-States allow you to have icons shown which vary
     *  based on the values, eg Red traffic light / Yellow traffic light /
     *  Green traffic light
     */
    static void iconSets(Sheet sheet) {
        sheet.createRow(0).createCell(0).setCellValue("Icon Sets");
        Row r = sheet.createRow(1);
        r.createCell(0).setCellValue("Reds");
        r.createCell(1).setCellValue(0);
        r.createCell(2).setCellValue(0);
        r.createCell(3).setCellValue(0);
        r = sheet.createRow(2);
        r.createCell(0).setCellValue("Yellows");
        r.createCell(1).setCellValue(5);
        r.createCell(2).setCellValue(5);
        r.createCell(3).setCellValue(5);
        r = sheet.createRow(3);
        r.createCell(0).setCellValue("Greens");
        r.createCell(1).setCellValue(10);
        r.createCell(2).setCellValue(10);
        r.createCell(3).setCellValue(10);

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        CellRangeAddress[] regions = { CellRangeAddress.valueOf("B1:B4") };
        ConditionalFormattingRule rule1 =
                sheetCF.createConditionalFormattingRule(IconSet.GYR_3_TRAFFIC_LIGHTS);
        IconMultiStateFormatting im1 = rule1.getMultiStateFormatting();
        im1.getThresholds()[0].setRangeType(RangeType.MIN);
        im1.getThresholds()[1].setRangeType(RangeType.PERCENT);
        im1.getThresholds()[1].setValue(33d);
        im1.getThresholds()[2].setRangeType(RangeType.MAX);
        sheetCF.addConditionalFormatting(regions, rule1);

        regions = new CellRangeAddress[] { CellRangeAddress.valueOf("C1:C4") };
        ConditionalFormattingRule rule2 =
                sheetCF.createConditionalFormattingRule(IconSet.GYR_3_FLAGS);
        IconMultiStateFormatting im2 = rule1.getMultiStateFormatting();
        im2.getThresholds()[0].setRangeType(RangeType.PERCENT);
        im2.getThresholds()[0].setValue(0d);
        im2.getThresholds()[1].setRangeType(RangeType.PERCENT);
        im2.getThresholds()[1].setValue(33d);
        im2.getThresholds()[2].setRangeType(RangeType.PERCENT);
        im2.getThresholds()[2].setValue(67d);
        sheetCF.addConditionalFormatting(regions, rule2);

        regions = new CellRangeAddress[] { CellRangeAddress.valueOf("D1:D4") };
        ConditionalFormattingRule rule3 =
                sheetCF.createConditionalFormattingRule(IconSet.GYR_3_SYMBOLS_CIRCLE);
        IconMultiStateFormatting im3 = rule1.getMultiStateFormatting();
        im3.setIconOnly(true);
        im3.getThresholds()[0].setRangeType(RangeType.MIN);
        im3.getThresholds()[1].setRangeType(RangeType.NUMBER);
        im3.getThresholds()[1].setValue(3d);
        im3.getThresholds()[2].setRangeType(RangeType.NUMBER);
        im3.getThresholds()[2].setValue(7d);
        sheetCF.addConditionalFormatting(regions, rule3);
    }

    /**
     * Color Scales / Colour Scales / Colour Gradients allow you shade the
     *  background colour of the cell based on the values, eg from Red to
     *  Yellow to Green.
     */
    static void colourScales(Sheet sheet) {
        sheet.createRow(0).createCell(0).setCellValue("Colour Scales");
        Row r = sheet.createRow(1);
        r.createCell(0).setCellValue("Red-Yellow-Green");
        for (int i=1; i<=7; i++) {
            r.createCell(i).setCellValue((i-1)*5);
        }
        r = sheet.createRow(2);
        r.createCell(0).setCellValue("Red-White-Blue");
        for (int i=1; i<=9; i++) {
            r.createCell(i).setCellValue((i-1)*5);
        }
        r = sheet.createRow(3);
        r.createCell(0).setCellValue("Blue-Green");
        for (int i=1; i<=16; i++) {
            r.createCell(i).setCellValue((i-1));
        }
        sheet.setColumnWidth(0, 5000);

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        CellRangeAddress[] regions = { CellRangeAddress.valueOf("B2:H2") };
        ConditionalFormattingRule rule1 =
                sheetCF.createConditionalFormattingColorScaleRule();
        ColorScaleFormatting cs1 = rule1.getColorScaleFormatting();
        cs1.getThresholds()[0].setRangeType(RangeType.MIN);
        cs1.getThresholds()[1].setRangeType(RangeType.PERCENTILE);
        cs1.getThresholds()[1].setValue(50d);
        cs1.getThresholds()[2].setRangeType(RangeType.MAX);
        ((ExtendedColor)cs1.getColors()[0]).setARGBHex("FFF8696B");
        ((ExtendedColor)cs1.getColors()[1]).setARGBHex("FFFFEB84");
        ((ExtendedColor)cs1.getColors()[2]).setARGBHex("FF63BE7B");
        sheetCF.addConditionalFormatting(regions, rule1);

        regions = new CellRangeAddress[] { CellRangeAddress.valueOf("B3:J3") };
        ConditionalFormattingRule rule2 =
                sheetCF.createConditionalFormattingColorScaleRule();
        ColorScaleFormatting cs2 = rule2.getColorScaleFormatting();
        cs2.getThresholds()[0].setRangeType(RangeType.MIN);
        cs2.getThresholds()[1].setRangeType(RangeType.PERCENTILE);
        cs2.getThresholds()[1].setValue(50d);
        cs2.getThresholds()[2].setRangeType(RangeType.MAX);
        ((ExtendedColor)cs2.getColors()[0]).setARGBHex("FFF8696B");
        ((ExtendedColor)cs2.getColors()[1]).setARGBHex("FFFCFCFF");
        ((ExtendedColor)cs2.getColors()[2]).setARGBHex("FF5A8AC6");
        sheetCF.addConditionalFormatting(regions, rule2);

        regions = new CellRangeAddress[] { CellRangeAddress.valueOf("B4:Q4") };
        ConditionalFormattingRule rule3=
                sheetCF.createConditionalFormattingColorScaleRule();
        ColorScaleFormatting cs3 = rule3.getColorScaleFormatting();
        cs3.setNumControlPoints(2);
        cs3.getThresholds()[0].setRangeType(RangeType.MIN);
        cs3.getThresholds()[1].setRangeType(RangeType.MAX);
        ((ExtendedColor)cs3.getColors()[0]).setARGBHex("FF5A8AC6");
        ((ExtendedColor)cs3.getColors()[1]).setARGBHex("FF63BE7B");
        sheetCF.addConditionalFormatting(regions, rule3);
    }

    /**
     * DataBars / Data-Bars allow you to have bars shown vary
     *  based on the values, from full to empty
     */
    static void dataBars(Sheet sheet) {
        sheet.createRow(0).createCell(0).setCellValue("Data Bars");
        Row r = sheet.createRow(1);
        r.createCell(1).setCellValue("Green Positive");
        r.createCell(2).setCellValue("Blue Mix");
        r.createCell(3).setCellValue("Red Negative");
        r = sheet.createRow(2);
        r.createCell(1).setCellValue(0);
        r.createCell(2).setCellValue(0);
        r.createCell(3).setCellValue(0);
        r = sheet.createRow(3);
        r.createCell(1).setCellValue(5);
        r.createCell(2).setCellValue(-5);
        r.createCell(3).setCellValue(-5);
        r = sheet.createRow(4);
        r.createCell(1).setCellValue(10);
        r.createCell(2).setCellValue(10);
        r.createCell(3).setCellValue(-10);
        r = sheet.createRow(5);
        r.createCell(1).setCellValue(5);
        r.createCell(2).setCellValue(5);
        r.createCell(3).setCellValue(-5);
        r = sheet.createRow(6);
        r.createCell(1).setCellValue(20);
        r.createCell(2).setCellValue(-10);
        r.createCell(3).setCellValue(-20);
        sheet.setColumnWidth(0, 3000);
        sheet.setColumnWidth(1, 5000);
        sheet.setColumnWidth(2, 5000);
        sheet.setColumnWidth(3, 5000);

        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

        ExtendedColor color = sheet.getWorkbook().getCreationHelper().createExtendedColor();
        color.setARGBHex("FF63BE7B");
        CellRangeAddress[] regions = { CellRangeAddress.valueOf("B2:B7") };
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule(color);
        DataBarFormatting db1 = rule1.getDataBarFormatting();
        db1.getMinThreshold().setRangeType(RangeType.MIN);
        db1.getMaxThreshold().setRangeType(RangeType.MAX);
        sheetCF.addConditionalFormatting(regions, rule1);

        color = sheet.getWorkbook().getCreationHelper().createExtendedColor();
        color.setARGBHex("FF5A8AC6");
        regions = new CellRangeAddress[] { CellRangeAddress.valueOf("C2:C7") };
        ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(color);
        DataBarFormatting db2 = rule2.getDataBarFormatting();
        db2.getMinThreshold().setRangeType(RangeType.MIN);
        db2.getMaxThreshold().setRangeType(RangeType.MAX);
        sheetCF.addConditionalFormatting(regions, rule2);

        color = sheet.getWorkbook().getCreationHelper().createExtendedColor();
        color.setARGBHex("FFF8696B");
        regions = new CellRangeAddress[] { CellRangeAddress.valueOf("D2:D7") };
        ConditionalFormattingRule rule3 = sheetCF.createConditionalFormattingRule(color);
        DataBarFormatting db3 = rule3.getDataBarFormatting();
        db3.getMinThreshold().setRangeType(RangeType.MIN);
        db3.getMaxThreshold().setRangeType(RangeType.MAX);
        sheetCF.addConditionalFormatting(regions, rule3);
    }

    /**
     * Print out a summary of the conditional formatting rules applied to cells on the given sheet.
     * Only cells with a matching rule are printed, and for those, all matching rules are sumarized.
     */
    static void evaluateRules(Workbook wb, String sheetName) {
        final WorkbookEvaluatorProvider wbEvalProv = (WorkbookEvaluatorProvider) wb.getCreationHelper().createFormulaEvaluator();
        final ConditionalFormattingEvaluator cfEval = new ConditionalFormattingEvaluator(wb, wbEvalProv);
        // if cell values have changed, clear cached format results
        cfEval.clearAllCachedValues();

        final Sheet sheet = wb.getSheet(sheetName);
        for (Row r : sheet) {
            for (Cell c : r) {
                final List<EvaluationConditionalFormatRule> rules = cfEval.getConditionalFormattingForCell(c);
                // check rules list for null, although current implementation will return an empty list, not null, then do what you want with results
                if (rules == null || rules.isEmpty()) continue;
                final CellReference ref = ConditionalFormattingEvaluator.getRef(c);
                if (rules.isEmpty()) continue;

                System.out.println("\n"
                        + ref.formatAsString()
                        + " has conditional formatting.");

                for (EvaluationConditionalFormatRule rule : rules) {
                    ConditionalFormattingRule cf = rule.getRule();

                    StringBuilder b = new StringBuilder();
                    b.append("\tRule ")
                            .append(rule.getFormattingIndex())
                            .append(": ");

                    // check for color scale
                    if (cf.getColorScaleFormatting() != null) {
                        b.append("\n\t\tcolor scale (caller must calculate bucket)");
                    }
                    // check for data bar
                    if (cf.getDataBarFormatting() != null) {
                        b.append("\n\t\tdata bar (caller must calculate bucket)");
                    }
                    // check for icon set
                    if (cf.getMultiStateFormatting() != null) {
                        b.append("\n\t\ticon set (caller must calculate icon bucket)");
                    }
                    // check for fill
                    if (cf.getPatternFormatting() != null) {
                        final PatternFormatting fill = cf.getPatternFormatting();
                        b.append("\n\t\tfill pattern ")
                                .append(fill.getFillPattern())
                                .append(" color index ")
                                .append(fill.getFillBackgroundColor());
                    }
                    // font stuff
                    if (cf.getFontFormatting() != null) {
                        final FontFormatting ff = cf.getFontFormatting();
                        b.append("\n\t\tfont format ")
                                .append("color index ")
                                .append(ff.getFontColorIndex());
                        if (ff.isBold()) b.append(" bold");
                        if (ff.isItalic()) b.append(" italic");
                        if (ff.isStruckout()) b.append(" strikeout");
                        b.append(" underline index ")
                                .append(ff.getUnderlineType());
                    }

                    System.out.println(b);
                }
            }
        }
    }
}
